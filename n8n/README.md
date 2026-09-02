# n8n workflow — Intune rename + ServiceNow

`Intune-rename-devices.workflow.json` is an importable n8n workflow that runs the whole loop:

```
Manual trigger
   → Collect devices      (Get-IntuneGroupDevices.ps1  → work\devices.json)
   → Rename devices        (Rename-IntuneDevice.ps1 -Rename -Force)
        ├─ output 0 (exit 0 = all clean)     → Close incident   (Close-ServiceNowIncident.ps1)
        └─ output 1 (exit 1/2 = attention)   → Create incident  (New-ServiceNowIncident.ps1)
                                                    → Send a message (Outlook notify)
```

The **create** and **close** steps share a `CorrelationId` (`intune-rename-devices`), so no ticket
sys_id has to pass between runs: a failing run opens (or reuses) one incident; a later clean run
resolves it automatically.

## Why the scripts are *called*, not pasted

The earlier version pasted each full `.ps1` into the node. That **cannot work**, because the
Intune scripts `dot-source IntuneGraphCommon.ps1` from their own folder (`$PSScriptRoot`), which
doesn't exist when the body is pasted inline. Each node here is a **thin caller** that runs the
real file from disk, so `$PSScriptRoot` resolves and the shared library loads.

## One-time host setup

1. **Copy all six files into one folder** on the n8n host, e.g. `C:\Scripts\RenameDevice\`:
   `IntuneGraphCommon.ps1`, `Get-IntuneGroupDevices.ps1`, `Rename-IntuneDevice.ps1`,
   `New-ServiceNowIncident.ps1`, `Close-ServiceNowIncident.ps1` (and this folder gets a `work\`
   subfolder automatically). If your folder differs, edit `$Base` at the top of each node.

2. **Set environment variables** for the account the n8n PowerShell node runs as
   (Machine or that user), so no secrets live in the workflow:

   | Variable | Purpose |
   |----------|---------|
   | `RENAMEDEVICE_TENANT_ID` | Entra tenant ID |
   | `RENAMEDEVICE_CLIENT_ID` | App registration (client) ID |
   | `RENAMEDEVICE_CLIENT_SECRET` | App client secret |
   | `SNOW_INSTANCE` | ServiceNow instance (`myorg` or `https://myorg.service-now.com`) |
   | `SNOW_USER` | ServiceNow integration user |
   | `SNOW_PASSWORD` | ServiceNow integration user password |

3. **Edit the "Collect devices" node**: set `$GroupId` to your group's object ID (and `$Base` if needed).

4. **Graph app permissions** (application, admin-consented): `GroupMember.Read.All`,
   `Device.Read.All`, `DeviceManagementManagedDevices.Read.All`,
   `DeviceManagementManagedDevices.PrivilegedOperations.All`.
   **ServiceNow**: the integration user needs to create/read/update `incident` (e.g. the `itil` role).

5. **PowerShell**: the nodes use `pwsh`/Windows PowerShell to run `.ps1` files; make sure the n8n
   PowerShell node is configured for the shell you have, and that running local scripts is allowed
   (`Unblock-File` the five files, or an appropriate execution policy).

## Test safely first

Before wiring `-Rename`, change the "Rename devices" node to a **dry run** by removing `-Rename -Force`.
It will preview names and write `work\rename-result.json` without changing anything. When the preview
looks right, put `-Rename -Force` back.

## Notes

- **Branching relies on the exit code** of the "Rename devices" node (0 = clean → *Close*; non-zero =
  attention → *Create*), surfaced through the node's success/error outputs (`onError:
  continueErrorOutput`). If your PowerShell node doesn't map exit codes to outputs, add an **IF** node
  that parses the node's stdout JSON (`countsByStatus`) instead.
- To run on a schedule, replace the Manual Trigger with a **Schedule Trigger**.
- The Outlook "Send a message" node is left as-is — configure its credential and recipient.
