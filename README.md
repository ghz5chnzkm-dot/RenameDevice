# Rename-IntuneDevice

Suggests and (optionally) applies a standardized device name to Intune-managed Windows
devices, based on the country of the device's **Windows Autopilot deployment profile** and
its hardware serial number.

```
Naming convention:  <CountryCode><CF><SerialNumber>
                     └─ ISO 3166-1 alpha-3, derived from the Autopilot profile
                                └─ literal "CF"
                                        └─ hardware serial number from Intune

Example:  profile "Autopilot User Driven profile MYS" + serial 0001234  ->  MYSCF0001234
```

The parts are **joined with no separator** (Windows computer names cannot contain `+`).

## What it does

For each Intune managed device ID in your CSV, the script:

1. Looks up the managed device in Microsoft Graph (serial number, current name, join type, OS).
2. Resolves the device's assigned **Autopilot deployment profile** and extracts the trailing
   ISO 3166-1 alpha-3 country code from its display name.
3. Builds and validates the suggested name against Windows computer-name rules.
4. **Dry-run (default):** prints the suggested name. **`-Rename`:** applies it via Graph.

If a device has **no Autopilot deployment profile**, it is reported as needing
**re-enrollment** and is never renamed.

### No reboot

`-Rename` calls the Intune `setDeviceName` action, which sets the name only. It does **not**
trigger "Restart after rename". The new name takes effect at the next user-initiated reboot.

## Requirements

- Windows Server (or any host) with **Windows PowerShell 5.1** or **PowerShell 7+**.
- Outbound HTTPS to `login.microsoftonline.com` and `graph.microsoft.com`.
- An **Entra app registration** (client credentials) with the following **application**
  Microsoft Graph permissions, **admin-consented**:

  | Permission                                               | Why |
  |----------------------------------------------------------|-----|
  | `DeviceManagementManagedDevices.Read.All`                | Read managed devices (name, serial, join type, OS). |
  | `DeviceManagementManagedDevices.PrivilegedOperations.All`| Perform the rename (`setDeviceName`). Only needed for `-Rename`. |
  | `DeviceManagementServiceConfig.Read.All`                 | Read Autopilot device identities and deployment profiles. |

  > For a dry-run-only service principal you can omit `...PrivilegedOperations.All`.

### App registration (quick steps)

1. Entra admin center → **App registrations** → **New registration**.
2. **Certificates & secrets** → **New client secret** → copy the value.
3. **API permissions** → **Add** → **Microsoft Graph** → **Application permissions** → add the
   three permissions above → **Grant admin consent**.
4. Note the **Directory (tenant) ID** and **Application (client) ID**.

## Input CSV

A CSV containing a column of **Intune managed device IDs** (GUIDs). The column is auto-detected
(`IntuneDeviceId`, `ManagedDeviceId`, `DeviceId`, or `Id`); otherwise pass `-IdColumnName`.
A single-column file is accepted as-is. See [`examples/devices.sample.csv`](examples/devices.sample.csv).

```csv
IntuneDeviceId
11111111-1111-1111-1111-111111111111
22222222-2222-2222-2222-222222222222
```

## Usage

**Preview (dry run) — no changes:**
```powershell
.\Rename-IntuneDevice.ps1 -CsvPath .\devices.csv `
    -TenantId <tenant-guid> -ClientId <client-guid> -ClientSecret <secret>
```

**Apply renames (interactive, prompts per device):**
```powershell
.\Rename-IntuneDevice.ps1 -CsvPath .\devices.csv -Rename `
    -TenantId <tenant-guid> -ClientId <client-guid> -ClientSecret <secret>
```

**Apply renames unattended (no prompts):** add `-Force`.

**Credentials via environment variables** (recommended — keeps the secret off the command line):
```powershell
$env:RENAMEDEVICE_TENANT_ID     = '<tenant-guid>'
$env:RENAMEDEVICE_CLIENT_ID     = '<client-guid>'
$env:RENAMEDEVICE_CLIENT_SECRET = '<secret>'
.\Rename-IntuneDevice.ps1 -CsvPath .\devices.csv -Rename -Force
```

### Key parameters

| Parameter | Default | Description |
|-----------|---------|-------------|
| `-CsvPath` | *(required)* | Path to the input CSV. |
| `-IdColumnName` | *(auto)* | CSV column holding the Intune device ID. |
| `-Delimiter` | `,` | CSV delimiter (use `;` for many EU exports). |
| `-DryRun` | *(default behavior)* | Preview only. Same as running with no mode switch. |
| `-Rename` | off | Apply the rename via Graph. Ignored if `-DryRun` is also set (safety). |
| `-Force` | off | Suppress the per-device confirmation prompt (required for unattended). |
| `-MaxNameLength` | `15` | Max name length. Longer names are flagged and skipped, never truncated. |
| `-ProfileCountryRegex` | trailing 3 letters | Regex (named group `cc`) to extract the code from the profile name. |
| `-SkipCountryCodeValidation` | off | Skip checking the code against the ISO 3166-1 alpha-3 list. |
| `-OutputDirectory` | `.\output` | Where the CSV report and JSON summary are written. |
| `-ReportPath` / `-JsonSummaryPath` | *(auto)* | Explicit output paths (override `-OutputDirectory`). |
| `-LogDirectory` | `.\logs` | Where the run log is written. |
| `-MaxRetries` | `5` | Retry attempts for transient Graph failures (429/5xx/network). |

Run `Get-Help .\Rename-IntuneDevice.ps1 -Full` for the complete reference.

## Output

- **Console** — live progress and a results table.
- **Log** — `logs\RenameIntuneDevice_<timestamp>.log`.
- **CSV report** — `output\RenameIntuneDevice_<timestamp>.csv`, one row per device.
- **JSON summary** — `output\RenameIntuneDevice_<timestamp>.json` (machine-readable; ideal for n8n).

### Per-device statuses

| Status | Meaning |
|--------|---------|
| `WouldRename` | Dry run: a rename is suggested. |
| `Renamed` | `-Rename`: rename submitted successfully (applies at next reboot). |
| `AlreadyCompliant` | Current name already matches the convention. |
| `Error-NoAutopilotProfile` | No deployment profile assigned — **device needs re-enrollment**. |
| `Error-CountryCodeParse` | Profile name did not yield a valid ISO 3166-1 alpha-3 code. |
| `Error-NoSerialNumber` | Intune has no serial number for the device. |
| `Error-NameTooLong` | Suggested name exceeds `-MaxNameLength` (default 15). |
| `Error-InvalidName` | Suggested name breaks Windows naming rules. |
| `Error-NotFound` | Device ID not found in Intune. |
| `Error-InvalidInput` | CSV value is not a valid GUID. |
| `Skipped-HybridJoined` | Entra hybrid-joined — rename via Intune is unsupported (rename in AD). |
| `Skipped-NonWindows` | Device OS is not Windows. |
| `RenameFailed` | The `setDeviceName` call failed (details in the message). |

### Exit codes

| Code | Meaning |
|------|---------|
| `0` | Completed; every device is clean (compliant / would-rename / renamed). |
| `1` | Completed, but one or more devices need attention (errors / skips / failures). |
| `2` | Fatal (bad arguments, unreadable/empty CSV, authentication failure). |

## Unattended execution with n8n (future)

The script is built for headless automation:

- Credentials come from `RENAMEDEVICE_TENANT_ID` / `RENAMEDEVICE_CLIENT_ID` /
  `RENAMEDEVICE_CLIENT_SECRET` environment variables — no secrets on the command line.
- `-Force` suppresses all prompts.
- The **exit code** signals success (`0`), attention (`1`), or fatal (`2`).
- The **JSON summary** is a stable, parseable result for downstream nodes.

**Execute Command** node example (secret supplied via the node's environment):
```bash
pwsh -NoProfile -File /opt/scripts/Rename-IntuneDevice.ps1 \
  -CsvPath /data/devices.csv \
  -Rename -Force \
  -JsonSummaryPath /data/rename-result.json
```

Then read `/data/rename-result.json` in a subsequent node, and branch on the exit code.
Recommended flow: run a **dry run first** (omit `-Rename`), inspect the JSON, then run again
with `-Rename -Force` once the suggested names look correct.

## Notes & limitations

- **15-character limit.** Windows NetBIOS/hostnames are capped at 15 characters. Names that
  would exceed this are flagged (`Error-NameTooLong`) and skipped, not truncated. Adjust with
  `-MaxNameLength` if your estate uses a different cap.
- **Hybrid-joined devices** cannot be renamed from Intune; rename them via on-premises AD.
- **Autopilot profile naming** drives the country code. The default regex takes the trailing
  3-letter token (e.g. `... MYS`). If your profiles encode the country differently, pass a
  custom `-ProfileCountryRegex` with a named `cc` group.
- Uses the Microsoft Graph **beta** endpoint, because `setDeviceName` is beta-only.
