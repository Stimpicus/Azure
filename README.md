# Azure

Microsoft 365 / Azure administration scripts.

## Grant-SitesSelectedPermission.ps1

Grants an Entra ID app registration access to **one** SharePoint site through the Microsoft Graph `Sites.Selected` permission.

| Step | What it does |
|---|---|
| 1 | Makes the signed-in admin a temporary Site Collection Administrator on the site |
| 2 | Adds Microsoft Graph `Sites.Selected` (Application) to the app and grants admin consent |
| 3 | Grants the app `read` or `write` on the site (`POST /sites/{id}/permissions`) |
| 4 | Removes the temporary Site Collection Administrator right. This runs in `finally`, so it's attempted on failure too |

### Inputs

| Variable | Meaning |
|---|---|
| `$appGuid` | App registration **Application (client) ID**. An Object ID is detected and converted |
| `$siteGuid` | SharePoint **site collection ID** |
| `$perm` | `Read` or `ReadWrite` (`Read/Write`, `RW` and `Write` are also accepted) |

Anything not passed as a parameter is prompted for. No credentials are captured. Sign-in happens in the browser (see below).

### Requirements

- Windows PowerShell 5.1 (recommended)
- `Install-Module Microsoft.Graph.Authentication -Scope CurrentUser`
- `Install-Module Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser`
- Global Administrator. Delegated Graph scopes `Application.ReadWrite.All`, `AppRoleAssignment.ReadWrite.All` and `Sites.FullControl.All` are consented at sign-in.

### Usage

```powershell
# Interactive: prompts for everything, shows a plan, asks to confirm
.\Grant-SitesSelectedPermission.ps1

# Dry run: reads and reports current state, changes nothing
.\Grant-SitesSelectedPermission.ps1 -appGuid <app-id> -siteGuid <site-id> -perm Read -WhatIf

# No confirmation prompt
.\Grant-SitesSelectedPermission.ps1 -appGuid <app-id> -siteGuid <site-id> -perm ReadWrite -Confirm:$false
```

### Safety behaviour

- All input is validated before the script connects to anything.
- The Graph sign-in must be a delegated user sign-in holding every required scope, or the script stops. That account is the one elevated and cleaned up.
- Locked sites (ReadOnly or NoAccess) are refused.
- A plan showing the current state must be confirmed before any change. `-WhatIf` gives a dry run.
- Idempotent: existing grants are detected and left alone. An existing app permission on the site is changed to the requested role, and the plan warns when that is a downgrade from `fullcontrol` or `owner`.
- Step 4 removes only an admin right that **this run added**. A pre-existing Site Collection Administrator assignment is never removed.
- The script verifies each grant with a fresh read after it writes.
- If the temporary admin right can't be removed, the script prints `MANUAL ACTION REQUIRED` and the exact command to run.
- The admin-state check fails closed. If SharePoint can't be read, the script aborts before step 1, or prints `MANUAL ACTION REQUIRED` at step 4. It never guesses.
- Each run is logged to a transcript in `.\logs\`.

### Testing status

The script's flow has been tested end-to-end against stubbed Graph and SharePoint cmdlets in five scenarios: a fresh grant, an idempotent re-run with a role change, failure at step 3 (the temporary admin right is still removed), SharePoint read failure (aborts before any change), and `-WhatIf`. It has **not** yet been run against a live tenant. Try it with `-WhatIf` on a test site first.

### Sign-in

There are two interactive browser sign-ins, both supporting MFA: Microsoft Graph first, then SharePoint Online admin. **Use the same account for both.** The script reads the Graph account and uses it as the admin to elevate. It can't read back which account signed in to SharePoint, so it can't check that the two match.

### Undo

```powershell
# Remove the app's site permission
Invoke-MgGraphRequest DELETE "v1.0/sites/{graph-site-id}/permissions/{permission-id}"
# Remove the Sites.Selected consent: delete the appRoleAssignment from the app's service principal
```
