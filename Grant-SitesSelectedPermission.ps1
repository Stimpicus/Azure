#Requires -Version 5.1

<#
.SYNOPSIS
    Grants an Entra ID app registration Sites.Selected access (Read or Read/Write) to one SharePoint site.

.DESCRIPTION
    Workflow:
      1. Temporarily makes the signed-in administrator a Site Collection Administrator on the target site.
      2. Grants the Microsoft Graph application permission Sites.Selected to the app registration
         (adds it to the app's configured permissions and grants admin consent).
      3. Grants the app 'read' or 'write' on the target site through the Graph /sites/{id}/permissions API.
      4. Removes the temporary Site Collection Administrator right. This runs in a finally block, so it is
         attempted even if steps 2-3 fail or the script is interrupted with Ctrl+C.

    Safety behaviour:
      - All inputs are validated (GUID format, UPN format, permission value) before anything connects.
      - The Graph session must be signed in as the same account as the supplied credentials, with every
        required scope, or the script stops.
      - The site is resolved by GUID and the match is checked, and locked sites (ReadOnly/NoAccess) are refused.
      - A full plan, including the current state, is shown and must be confirmed before any change.
        Use -WhatIf for a dry run (reads only, no changes) or -Confirm:$false to skip the prompt.
      - Steps are idempotent: anything already in place is left alone.
      - Step 4 only removes the admin right if this script added it. If you were already a Site Collection
        Administrator before the run, that pre-existing access is left untouched.
      - Every run is logged to a transcript in -LogDirectory.

.PARAMETER cred
    Your administrator credentials. Prompted for if omitted. The UPN identifies the account to elevate and
    is checked against the Graph sign-in. See NOTES on MFA.

.PARAMETER appGuid
    The Application (client) ID of the app registration. If you pass the app's Object ID instead, the
    script detects it and shows the matching Application ID in the plan.

.PARAMETER siteGuid
    The SharePoint site collection ID (the GUID in Graph's site id 'host,{siteGuid},{webGuid}', also shown
    by Get-SPOSite ... | Select Id, or /_api/site/id).

.PARAMETER perm
    'Read' or 'ReadWrite' ('Read/Write', 'RW' and 'Write' are also accepted). Maps to the Graph site
    permission roles 'read' and 'write'.

.PARAMETER AdminUrl
    Optional SharePoint admin center URL. Defaults to https://<tenant>-admin.sharepoint.com, derived from the
    tenant's root site. Set it for vanity domains or multi-geo tenants.

.PARAMETER LogDirectory
    Folder for the transcript log. Defaults to .\logs next to the current directory.

.EXAMPLE
    .\Grant-SitesSelectedPermission.ps1
    Prompts for everything, shows the plan, asks for confirmation.

.EXAMPLE
    .\Grant-SitesSelectedPermission.ps1 -appGuid 11111111-2222-3333-4444-555555555555 `
        -siteGuid 66666666-7777-8888-9999-000000000000 -perm Read -WhatIf
    Dry run: connects, resolves and reports current state, makes no changes.

.NOTES
    Requires:
      - Windows PowerShell 5.1 (recommended; the SharePoint Online module is most reliable there).
        PowerShell 7 works if the SharePoint module loads, possibly through Windows PowerShell compatibility.
      - Modules: Microsoft.Graph.Authentication (2.x) and Microsoft.Online.SharePoint.PowerShell
          Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
          Install-Module Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser
      - Role: Global Administrator (consenting to Microsoft Graph application permissions needs it, or
        Privileged Role Administrator plus SharePoint Administrator).
      - Delegated Graph scopes, consented at sign-in: Application.ReadWrite.All,
        AppRoleAssignment.ReadWrite.All, Sites.FullControl.All.

    About credentials and MFA:
      Microsoft Graph PowerShell does not accept a username and password for a user sign-in, so Graph always
      opens an interactive (modern auth) sign-in. $cred is used to identify the account, to verify the Graph
      sign-in matches it, and for Connect-SPOService. Connect-SPOService -Credential fails for accounts that
      require MFA (which should include every Global Administrator). When it fails, the script falls back to
      an interactive SharePoint sign-in, and you sign in again as the same account.
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [System.Management.Automation.PSCredential]$cred,
    [string]$appGuid,
    [string]$siteGuid,
    [string]$perm,
    [string]$AdminUrl,
    [string]$LogDirectory = (Join-Path -Path (Get-Location).Path -ChildPath 'logs')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Well-known Application ID of Microsoft Graph, the same in every tenant.
$GraphAppId  = '00000003-0000-0000-c000-000000000000'
$GraphScopes = @('Application.ReadWrite.All', 'AppRoleAssignment.ReadWrite.All', 'Sites.FullControl.All')

#region Helper functions

function Write-Step {
    param([string]$Message)
    Write-Host ''
    Write-Host "==> $Message" -ForegroundColor Cyan
}

function Read-GuidValue {
    # Returns a normalised GUID string. Prompts (repeatedly) until a valid, non-empty GUID is given.
    param([string]$Prompt, [string]$Value)
    while ($true) {
        if ([string]::IsNullOrWhiteSpace($Value)) { $Value = Read-Host -Prompt $Prompt }
        $parsed = [guid]::Empty
        if ([guid]::TryParse($Value.Trim(), [ref]$parsed) -and $parsed -ne [guid]::Empty) {
            return $parsed.ToString()
        }
        Write-Warning "'$Value' is not a valid GUID. Please try again."
        $Value = $null
    }
}

function ConvertTo-GraphSiteRole {
    # Maps the user's Read / Read-Write choice to the Graph site permission role. Returns $null if unrecognised.
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $null }
    switch -Regex ($Value.Trim()) {
        '^(r|read)$'                                     { return 'read' }
        '^(rw|read\s*/\s*write|read\s*-?\s*write|write)$' { return 'write' }
        default                                          { return $null }
    }
}

function Invoke-GraphGetAll {
    # GETs a Graph collection and follows @odata.nextLink so no results are silently dropped.
    param([string]$Uri)
    $items = New-Object System.Collections.Generic.List[object]
    while ($Uri) {
        $page = Invoke-MgGraphRequest -Method GET -Uri $Uri
        if ($page['value']) { foreach ($item in $page['value']) { $items.Add($item) } }
        $Uri = $page['@odata.nextLink']
    }
    # No leading comma: callers wrap the result in @(), and a comma would nest the array one level deeper.
    return $items.ToArray()
}

function Get-GraphFilterUri {
    # Builds a URI with a properly escaped $filter expression.
    param([string]$Path, [string]$Filter)
    return '{0}?$filter={1}' -f $Path, [uri]::EscapeDataString($Filter)
}

function Get-SiteCollectionIdFromGraphId {
    # Graph site ids look like 'contoso.sharepoint.com,{siteCollectionId},{webId}'.
    param([string]$GraphSiteId)
    $parts = $GraphSiteId -split ','
    if ($parts.Count -ge 2) { return $parts[1] }
    return $null
}

function Resolve-GraphSite {
    # Finds the site collection whose ID is $SiteCollectionId and returns its root web.
    param([string]$SiteCollectionId, [string]$Hostname)

    $select = '?$select=id,displayName,webUrl'

    # Fast paths: address the site directly by id.
    foreach ($path in @("v1.0/sites/$Hostname,$SiteCollectionId", "v1.0/sites/$SiteCollectionId")) {
        try {
            $site = Invoke-MgGraphRequest -Method GET -Uri ($path + $select)
            if ((Get-SiteCollectionIdFromGraphId $site['id']) -eq $SiteCollectionId) { return $site }
        }
        catch {
            Write-Verbose "Direct lookup '$path' failed: $($_.Exception.Message)"
        }
    }

    # Slow path: enumerate sites visible to the admin and match on the site collection id.
    Write-Verbose 'Direct lookup failed; searching all sites (this can take a while in large tenants).'
    $found = @(Invoke-GraphGetAll -Uri ('v1.0/sites?search=*&' + $select.TrimStart('?')) |
        Where-Object { (Get-SiteCollectionIdFromGraphId $_['id']) -eq $SiteCollectionId })
    if ($found.Count -eq 0) { return $null }

    # Subsites share the site collection id; the root web has the shortest URL.
    return $found | Sort-Object { $_['webUrl'].Length } | Select-Object -First 1
}

function Get-SiteAdminState {
    # Returns $true if $LoginName is a Site Collection Administrator on $SiteUrl, $false otherwise.
    # Fail-closed: any error (throttling, lost session) propagates to the caller instead of reading as "not admin".
    # Lists the site's users rather than looking one up, because a single lookup throws for a non-member and
    # that error can't be reliably told apart from a real failure.
    param([string]$SiteUrl, [string]$LoginName)
    $admins = @(Get-SPOUser -Site $SiteUrl -Limit All | Where-Object { $_.IsSiteAdmin })
    # LoginName may be the bare UPN or a claims value ending in |upn.
    $match = @($admins | Where-Object { $_.LoginName -eq $LoginName -or $_.LoginName -like "*|$LoginName" })
    return ($match.Count -gt 0)
}

function Invoke-WithRetry {
    # Retries a read that can lag behind a write while Entra ID / SharePoint replicate.
    param([scriptblock]$ScriptBlock, [int]$Attempts = 6, [int]$DelaySeconds = 5)
    for ($i = 1; $i -le $Attempts; $i++) {
        try {
            $result = & $ScriptBlock
            if ($result) { return $result }
        }
        catch {
            if ($i -eq $Attempts) { throw }
        }
        if ($i -lt $Attempts) { Start-Sleep -Seconds $DelaySeconds }
    }
    return $null
}

#endregion

# State tracked so the finally block can clean up only what this run did.
$transcriptStarted  = $false
$graphConnected     = $false
$spoConnected       = $false
$adminAddedByScript = $false
$siteUrl            = $null
$userUpn            = $null
$exitCode           = 0

try {
    #region Logging
    # -WhatIf/-Confirm:$false so a dry run still gets a log and isn't prompted for the log itself.
    New-Item -ItemType Directory -Path $LogDirectory -Force -WhatIf:$false -Confirm:$false | Out-Null
    $logFile = Join-Path $LogDirectory ('Grant-SitesSelected_{0:yyyyMMdd_HHmmss}.log' -f (Get-Date))
    Start-Transcript -Path $logFile -WhatIf:$false -Confirm:$false | Out-Null
    $transcriptStarted = $true
    #endregion

    #region Step 0: Gather and validate input (nothing connects until all input is valid)
    Write-Step 'Collecting input'

    if (-not $cred) {
        $cred = Get-Credential -Message 'Sign-in for the Global Administrator who will perform this change (UPN and password)'
    }
    if (-not $cred -or [string]::IsNullOrWhiteSpace($cred.UserName)) { throw 'No credentials were supplied.' }
    $userUpn = $cred.UserName.Trim()
    if ($userUpn -notmatch '^[^@\s\\]+@[^@\s]+\.[^@\s]+$') {
        throw "'$userUpn' is not a UPN (user@domain). Use your Entra ID sign-in name, not DOMAIN\user."
    }

    $appGuid  = Read-GuidValue -Prompt 'App registration Application (client) ID' -Value $appGuid
    $siteGuid = Read-GuidValue -Prompt 'SharePoint site collection ID'            -Value $siteGuid

    $graphRole = ConvertTo-GraphSiteRole $perm
    while (-not $graphRole) {
        if ($perm) { Write-Warning "'$perm' is not valid. Enter Read or ReadWrite." }
        $perm = Read-Host -Prompt 'Permission to grant (Read / ReadWrite)'
        $graphRole = ConvertTo-GraphSiteRole $perm
    }
    #endregion

    #region Modules
    Write-Step 'Loading modules'
    foreach ($module in 'Microsoft.Graph.Authentication', 'Microsoft.Online.SharePoint.PowerShell') {
        if (-not (Get-Module -ListAvailable -Name $module)) {
            throw "Required module '$module' is not installed. Run: Install-Module $module -Scope CurrentUser"
        }
    }
    Import-Module Microsoft.Graph.Authentication

    try {
        Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking -WarningAction SilentlyContinue
    }
    catch {
        if ($PSVersionTable.PSEdition -ne 'Core') { throw }
        Write-Warning 'SharePoint module did not load natively in PowerShell 7; retrying with Windows PowerShell compatibility.'
        Import-Module Microsoft.Online.SharePoint.PowerShell -UseWindowsPowerShell -DisableNameChecking -WarningAction SilentlyContinue
    }
    #endregion

    #region Connect to Microsoft Graph and verify the session
    Write-Step "Connecting to Microsoft Graph (sign in as $userUpn)"
    # Process scope keeps this session separate from any cached Graph session for another account.
    Connect-MgGraph -Scopes $GraphScopes -ContextScope Process -NoWelcome
    $graphConnected = $true

    $context = Get-MgContext
    if (-not $context -or -not $context.Account) { throw 'Graph connection did not return a signed-in account.' }
    if ($context.Account -ne $userUpn) {
        throw "Graph is signed in as '$($context.Account)', but the credentials are for '$userUpn'. Sign in as the same account."
    }
    $missingScopes = @($GraphScopes | Where-Object { $context.Scopes -notcontains $_ })
    if ($missingScopes.Count -gt 0) {
        throw "The Graph session is missing required scopes: $($missingScopes -join ', '). Consent was not granted."
    }
    Write-Host "Signed in to tenant $($context.TenantId) as $($context.Account)."
    #endregion

    #region Resolve the SharePoint site
    Write-Step "Resolving SharePoint site $siteGuid"
    $rootSite = Invoke-MgGraphRequest -Method GET -Uri 'v1.0/sites/root?$select=id,webUrl,siteCollection'
    $hostname = $rootSite['siteCollection']['hostname']
    if (-not $hostname) { throw 'Could not determine the SharePoint hostname from the tenant root site.' }

    $site = Resolve-GraphSite -SiteCollectionId $siteGuid -Hostname $hostname
    if (-not $site) { throw "No SharePoint site collection with ID $siteGuid was found (or it is not visible to $userUpn)." }
    $graphSiteId = $site['id']
    $siteUrl     = $site['webUrl']
    $siteName    = if ($site['displayName']) { $site['displayName'] } else { $siteUrl }
    Write-Host "Found: $siteName ($siteUrl)"

    if (-not $AdminUrl) {
        # contoso.sharepoint.com -> https://contoso-admin.sharepoint.com
        $AdminUrl = 'https://' + ($hostname -replace '^([^.]+)\.', '$1-admin.')
    }
    #endregion

    #region Resolve the app registration and its service principal
    Write-Step "Resolving app registration $appGuid"
    $application = @(Invoke-GraphGetAll -Uri (Get-GraphFilterUri 'v1.0/applications' "appId eq '$appGuid'")) | Select-Object -First 1

    if (-not $application) {
        # A common mistake is pasting the Object ID. Accept it, but switch to the real Application ID.
        try {
            $byObjectId = Invoke-MgGraphRequest -Method GET -Uri "v1.0/applications/$appGuid"
            Write-Warning "$appGuid is the app's Object ID. Using its Application (client) ID $($byObjectId['appId'])."
            $application = $byObjectId
            $appGuid     = $byObjectId['appId']
        }
        catch {
            Write-Verbose "No application object with Object ID $appGuid."
        }
    }

    $appSp = @(Invoke-GraphGetAll -Uri (Get-GraphFilterUri 'v1.0/servicePrincipals' "appId eq '$appGuid'")) | Select-Object -First 1

    if (-not $application -and -not $appSp) {
        throw "No app registration or enterprise application with ID $appGuid exists in this tenant."
    }
    if (-not $application) {
        # Multi-tenant app registered in another tenant: consent still works through its service principal here.
        Write-Warning 'The app registration is not in this tenant (it is a multi-tenant or third-party app). Its configured-permissions list will not be updated; consent is granted to its service principal here.'
    }
    $appName = if ($application) { $application['displayName'] } else { $appSp['displayName'] }

    $graphSp = @(Invoke-GraphGetAll -Uri (Get-GraphFilterUri 'v1.0/servicePrincipals' "appId eq '$GraphAppId'")) | Select-Object -First 1
    if (-not $graphSp) { throw 'Could not find the Microsoft Graph service principal in this tenant.' }

    $sitesSelectedRole = @($graphSp['appRoles'] | Where-Object {
            $_['value'] -eq 'Sites.Selected' -and @($_['allowedMemberTypes']) -contains 'Application'
        }) | Select-Object -First 1
    if (-not $sitesSelectedRole) { throw 'The Sites.Selected application role was not found on Microsoft Graph.' }
    #endregion

    #region Read current state (so the plan is accurate and every step is idempotent)
    Write-Step 'Reading current state'

    # Is Sites.Selected already in the app's configured permissions?
    $requiredAccess = @()
    if ($application -and $application['requiredResourceAccess']) { $requiredAccess = @($application['requiredResourceAccess']) }
    $graphAccessEntry = $requiredAccess | Where-Object { $_['resourceAppId'] -eq $GraphAppId } | Select-Object -First 1
    $alreadyConfigured = [bool]($graphAccessEntry -and
        (@($graphAccessEntry['resourceAccess']) | Where-Object { $_['id'] -eq $sitesSelectedRole['id'] -and $_['type'] -eq 'Role' }))

    # Has admin consent (app role assignment) already been granted?
    $alreadyConsented = $false
    if ($appSp) {
        $assignments = Invoke-GraphGetAll -Uri "v1.0/servicePrincipals/$($appSp['id'])/appRoleAssignments"
        $alreadyConsented = [bool](@($assignments) | Where-Object {
                $_['resourceId'] -eq $graphSp['id'] -and $_['appRoleId'] -eq $sitesSelectedRole['id']
            })
    }

    # Does the app already have a permission on this site?
    $existingSitePerm = $null
    foreach ($p in (Invoke-GraphGetAll -Uri "v1.0/sites/$graphSiteId/permissions")) {
        $identities = @($p['grantedToIdentitiesV2']) + @($p['grantedToIdentities']) | Where-Object { $_ }
        if (@($identities | Where-Object { $_['application'] -and $_['application']['id'] -eq $appGuid }).Count -gt 0) {
            # The list call may omit roles, so read the full permission.
            $existingSitePerm = Invoke-MgGraphRequest -Method GET -Uri "v1.0/sites/$graphSiteId/permissions/$($p['id'])"
            break
        }
    }
    # @() outside the if: assigning an if-statement's output unrolls a one-element array to a scalar.
    $existingRoles = @(if ($existingSitePerm) { $existingSitePerm['roles'] })

    # Connect to SharePoint Online admin to check and manage Site Collection Administrators.
    Write-Step "Connecting to SharePoint Online admin ($AdminUrl)"
    try {
        Connect-SPOService -Url $AdminUrl -Credential $cred
    }
    catch {
        Write-Warning "Credential sign-in to SharePoint failed ($($_.Exception.Message.Trim())). This is expected for MFA-enabled accounts. Opening interactive sign-in; sign in as $userUpn."
        Connect-SPOService -Url $AdminUrl
    }
    $spoConnected = $true

    $spoSite = Get-SPOSite -Identity $siteUrl
    if ($spoSite.LockState -and $spoSite.LockState -ne 'Unlock') {
        throw "Site $siteUrl is locked ($($spoSite.LockState)). Refusing to change it."
    }
    $wasAlreadyAdmin = Get-SiteAdminState -SiteUrl $siteUrl -LoginName $userUpn
    #endregion

    #region Show the plan and confirm
    $roleLabel = if ($graphRole -eq 'read') { 'Read' } else { 'Read/Write' }
    Write-Step 'Plan'
    Write-Host "  Admin account : $userUpn"
    Write-Host "  Site          : $siteName"
    Write-Host "                  $siteUrl"
    Write-Host "  App           : $appName (Application ID $appGuid)"
    Write-Host "  Site access   : $roleLabel (Graph role '$graphRole')"
    Write-Host ''
    Write-Host ('  1. Site Collection Admin  : ' + $(if ($wasAlreadyAdmin) { 'already admin - no change (and will NOT be removed in step 4)' } else { 'ADD temporarily' }))
    Write-Host ('  2a. Configured permission : ' + $(if (-not $application) { 'n/a (app registered in another tenant)' } elseif ($alreadyConfigured) { 'Sites.Selected already listed - no change' } else { 'ADD Microsoft Graph Sites.Selected (Application)' }))
    Write-Host ('  2b. Admin consent         : ' + $(if ($alreadyConsented) { 'already granted - no change' } elseif (-not $appSp) { 'CREATE service principal, then GRANT' } else { 'GRANT' }))
    Write-Host ('  3. Site permission        : ' + $(
            if (-not $existingSitePerm) { "GRANT '$graphRole'" }
            elseif ($existingRoles.Count -eq 1 -and $existingRoles[0] -eq $graphRole) { "already '$graphRole' - no change" }
            else { "CHANGE '$($existingRoles -join ',')' -> '$graphRole'" }))
    Write-Host ('  4. Site Collection Admin  : ' + $(if ($wasAlreadyAdmin) { 'leave in place (pre-existing)' } else { 'REMOVE' }))

    if ($existingRoles | Where-Object { $_ -in @('fullcontrol', 'owner') }) {
        Write-Warning "The app currently has '$($existingRoles -join ',')' on this site. Continuing will REDUCE it to '$graphRole'."
    }

    if (-not $PSCmdlet.ShouldProcess("$siteUrl / app $appGuid", "Grant Sites.Selected '$graphRole'")) {
        Write-Host 'No changes were made.' -ForegroundColor Yellow
        return
    }
    #endregion

    #region Step 1: Grant the signed-in user Site Collection Administrator
    Write-Step 'Step 1: Granting temporary Site Collection Administrator'
    if ($wasAlreadyAdmin) {
        Write-Host "$userUpn is already a Site Collection Administrator; skipping."
    }
    else {
        # Set the flag first so the finally block attempts removal even if this call fails part-way.
        $adminAddedByScript = $true
        Set-SPOUser -Site $siteUrl -LoginName $userUpn -IsSiteCollectionAdmin $true | Out-Null
        if (-not (Get-SiteAdminState -SiteUrl $siteUrl -LoginName $userUpn)) {
            throw "Set-SPOUser reported success but $userUpn is not a Site Collection Administrator."
        }
        Write-Host "Added $userUpn as Site Collection Administrator." -ForegroundColor Green
    }
    #endregion

    #region Step 2: Grant the Sites.Selected Graph application permission
    Write-Step 'Step 2: Granting Microsoft Graph Sites.Selected to the app'

    # 2a. Add Sites.Selected to the app's configured permissions (what the portal shows under API permissions).
    if ($application -and -not $alreadyConfigured) {
        $newAccess = @{ id = $sitesSelectedRole['id']; type = 'Role' }
        if ($graphAccessEntry) {
            $graphAccessEntry['resourceAccess'] = @($graphAccessEntry['resourceAccess']) + $newAccess
        }
        else {
            $requiredAccess += @{ resourceAppId = $GraphAppId; resourceAccess = @($newAccess) }
        }
        # PATCH replaces the whole list, so the existing entries are sent back unchanged alongside the new one.
        $body = @{ requiredResourceAccess = @($requiredAccess) } | ConvertTo-Json -Depth 10
        Invoke-MgGraphRequest -Method PATCH -Uri "v1.0/applications/$($application['id'])" -Body $body -ContentType 'application/json' | Out-Null
        Write-Host 'Added Sites.Selected to the app registration''s configured permissions.' -ForegroundColor Green
    }

    # 2b. Make sure the app has a service principal in this tenant.
    if (-not $appSp) {
        $body = @{ appId = $appGuid } | ConvertTo-Json
        $appSp = Invoke-MgGraphRequest -Method POST -Uri 'v1.0/servicePrincipals' -Body $body -ContentType 'application/json'
        Write-Host "Created service principal $($appSp['id'])." -ForegroundColor Green
    }

    # 2c. Grant admin consent: assign the Sites.Selected app role to the app's service principal.
    if ($alreadyConsented) {
        Write-Host 'Sites.Selected is already consented; skipping.'
    }
    else {
        $body = @{
            principalId = $appSp['id']
            resourceId  = $graphSp['id']
            appRoleId   = $sitesSelectedRole['id']
        } | ConvertTo-Json
        # Retried because a service principal created moments ago may not have replicated yet.
        Invoke-WithRetry {
            Invoke-MgGraphRequest -Method POST -Uri "v1.0/servicePrincipals/$($graphSp['id'])/appRoleAssignedTo" -Body $body -ContentType 'application/json'
        } | Out-Null
        Write-Host 'Granted admin consent for Sites.Selected.' -ForegroundColor Green
    }

    # Verify through a fresh read.
    $verified = Invoke-WithRetry {
        @(Invoke-GraphGetAll -Uri "v1.0/servicePrincipals/$($appSp['id'])/appRoleAssignments") |
            Where-Object { $_['resourceId'] -eq $graphSp['id'] -and $_['appRoleId'] -eq $sitesSelectedRole['id'] }
    }
    if (-not $verified) {
        throw 'Verification failed: the Sites.Selected app role assignment was not found after granting it.'
    }
    #endregion

    #region Step 3: Grant the app Read or Write on the site
    Write-Step "Step 3: Granting the app '$graphRole' on the site"
    if ($existingSitePerm -and $existingRoles.Count -eq 1 -and $existingRoles[0] -eq $graphRole) {
        Write-Host "The app already has '$graphRole' on this site; skipping."
        $sitePermId = $existingSitePerm['id']
    }
    elseif ($existingSitePerm) {
        $body = @{ roles = @($graphRole) } | ConvertTo-Json
        Invoke-MgGraphRequest -Method PATCH -Uri "v1.0/sites/$graphSiteId/permissions/$($existingSitePerm['id'])" -Body $body -ContentType 'application/json' | Out-Null
        $sitePermId = $existingSitePerm['id']
        Write-Host "Changed the app's site permission from '$($existingRoles -join ',')' to '$graphRole'." -ForegroundColor Green
    }
    else {
        $body = @{
            roles               = @($graphRole)
            grantedToIdentities = @(@{ application = @{ id = $appGuid; displayName = $appName } })
        } | ConvertTo-Json -Depth 5
        $created = Invoke-MgGraphRequest -Method POST -Uri "v1.0/sites/$graphSiteId/permissions" -Body $body -ContentType 'application/json'
        $sitePermId = $created['id']
        Write-Host "Granted '$graphRole' to the app on the site." -ForegroundColor Green
    }

    # Verify through a fresh read.
    $verifyPerm = Invoke-WithRetry {
        $p = Invoke-MgGraphRequest -Method GET -Uri "v1.0/sites/$graphSiteId/permissions/$sitePermId"
        if (@($p['roles']) -contains $graphRole) { $p }
    }
    if (-not $verifyPerm) {
        throw "Verification failed: the site permission does not show role '$graphRole' after granting it."
    }
    Write-Host "Verified: app $appGuid has '$graphRole' on $siteUrl." -ForegroundColor Green
    #endregion
}
catch {
    $exitCode = 1
    Write-Host ''
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "At: $($_.InvocationInfo.PositionMessage)" -ForegroundColor DarkGray
}
finally {
    #region Step 4: Remove the temporary Site Collection Administrator right (always attempted)
    if ($adminAddedByScript) {
        Write-Step 'Step 4: Removing temporary Site Collection Administrator'
        try {
            if (-not $spoConnected) { throw 'Not connected to SharePoint Online.' }
            Set-SPOUser -Site $siteUrl -LoginName $userUpn -IsSiteCollectionAdmin $false | Out-Null
            if (Get-SiteAdminState -SiteUrl $siteUrl -LoginName $userUpn) {
                throw 'Set-SPOUser reported success but the user is still a Site Collection Administrator.'
            }
            Write-Host "Removed $userUpn as Site Collection Administrator." -ForegroundColor Green
        }
        catch {
            $exitCode = 1
            Write-Host "MANUAL ACTION REQUIRED: could not remove $userUpn as Site Collection Administrator on $siteUrl." -ForegroundColor Red
            Write-Host "Reason: $($_.Exception.Message)" -ForegroundColor Red
            Write-Host "Run: Set-SPOUser -Site '$siteUrl' -LoginName '$userUpn' -IsSiteCollectionAdmin `$false" -ForegroundColor Red
        }
    }
    elseif ($siteUrl -and $spoConnected) {
        Write-Step 'Step 4: Site Collection Administrator'
        Write-Host 'No temporary admin right was added by this run; nothing to remove.'
    }
    #endregion

    # Disconnect only the sessions this script opened.
    if ($spoConnected)   { try { Disconnect-SPOService } catch { Write-Verbose $_.Exception.Message } }
    if ($graphConnected) { try { Disconnect-MgGraph | Out-Null } catch { Write-Verbose $_.Exception.Message } }

    if ($exitCode -eq 0) { Write-Host "`nCompleted successfully." -ForegroundColor Green }
    else { Write-Host "`nCompleted with errors. See the log." -ForegroundColor Red }
    if ($transcriptStarted) { Stop-Transcript | Out-Null; Write-Host "Log: $logFile" }
}

exit $exitCode
