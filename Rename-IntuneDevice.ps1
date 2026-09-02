#Requires -Version 5.1
<#
.SYNOPSIS
    Suggests and (optionally) applies standardized Intune device names based on the
    "<CountryCode-ISO3166-1-Alpha3>CF<SerialNumber>" naming convention.

.DESCRIPTION
    Input can be (exactly one of):
      * -CsvPath   : a CSV containing a column of Intune managed device IDs (GUIDs).
      * -InputJson : a JSON file (array of device IDs, or enriched device records).
      * -Stdin     : a JSON array piped in (e.g. from Get-IntuneGroupDevices.ps1).

    For each device the script:
      1. Obtains the managed-device details. If the input record is already enriched (has a
         SerialNumber, as produced by Get-IntuneGroupDevices.ps1) it is used as-is with NO Graph
         read; otherwise the device is fetched from Graph by ID.
      2. Derives the ISO 3166-1 alpha-3 country code from the enrollment profile name
         (managedDevice.enrollmentProfileName, e.g. "Autopilot User Driven profile MYS" -> "MYS");
         falls back to the assigned Autopilot deployment profile if that field is empty.
      3. Builds the suggested name:  <CountryCode> + "CF" + <SerialNumber>  (parts joined, no separators).
      4. Validates it against Windows computer-name rules (length + allowed characters).

    Modes:
      * Default / -DryRun : PREVIEW ONLY. No writes.
      * -Rename           : Applies the rename via the Graph 'setDeviceName' action. Sets the name
                            only; does NOT restart the device. The new name applies at next reboot.

    Requires IntuneGraphCommon.ps1 in the same folder (dot-sourced).

.PARAMETER CsvPath
    Input CSV of Intune managed device IDs. Column auto-detected (IntuneDeviceId/ManagedDeviceId/
    DeviceId/Id) unless -IdColumnName is given; a single-column file is used as-is.

.PARAMETER InputJson
    Path to a JSON file: an array of ID strings, or an array of device records (objects).

.PARAMETER Stdin
    Read the JSON array from standard input (for piping from Get-IntuneGroupDevices.ps1).

.PARAMETER IdColumnName
    Explicit CSV column header holding the Intune device ID (auto-detected otherwise).

.PARAMETER Delimiter
    Single-character CSV delimiter. Default ','. Use ';' for typical European-locale exports.

.PARAMETER AuthMode
    ClientSecret (default, app-only; best for unattended/n8n) or Interactive (browser sign-in as your admin).

.PARAMETER TenantId / ClientId / ClientSecret
    Credentials. Fall back to RENAMEDEVICE_TENANT_ID / _CLIENT_ID / _CLIENT_SECRET env vars.
    For -AuthMode Interactive, ClientId is optional (defaults to the Graph public client) and no secret is needed.

.PARAMETER DryRun
    Explicit preview mode (same as no mode switch). No writes.

.PARAMETER Rename
    Perform the rename via Graph. Ignored (safety wins) if -DryRun is also specified.

.PARAMETER Force
    Suppress the per-device confirmation prompt in -Rename mode. Required for unattended / n8n runs.

.PARAMETER MaxNameLength
    Max generated-name length. Default 15 (Windows NetBIOS limit). Longer names are flagged and skipped.

.PARAMETER ProfileCountryRegex
    Regex (named group 'cc') to extract the country code from the profile name. Default: trailing 3 letters.

.PARAMETER SkipCountryCodeValidation
    Skip validating the extracted code against the ISO 3166-1 alpha-3 list.

.PARAMETER PassThruJson
    Emit the run summary as JSON to STDOUT (logs go to STDERR, no console table). Ideal for n8n.

.PARAMETER GraphBaseUri / OutputDirectory / ReportPath / JsonSummaryPath / LogDirectory / MaxRetries
    See README. Defaults: beta endpoint; <script>\output and <script>\logs; 5 retries.

.EXAMPLE
    .\Rename-IntuneDevice.ps1 -CsvPath .\devices.csv -AuthMode Interactive -TenantId <tenant>

.EXAMPLE
    pwsh -File .\Get-IntuneGroupDevices.ps1 -GroupId <guid> | pwsh -File .\Rename-IntuneDevice.ps1 -Stdin -Rename -Force -PassThruJson

.NOTES
    Compatibility: Windows PowerShell 5.1 and PowerShell 7+.

    Required Microsoft Graph permissions (Application for ClientSecret; delegated for Interactive):
      * DeviceManagementManagedDevices.Read.All                  (read managed devices; only needed for non-enriched input)
      * DeviceManagementManagedDevices.PrivilegedOperations.All  (setDeviceName / rename)
      * DeviceManagementServiceConfig.Read.All                   (only for the Autopilot fallback)

    Exit codes (for orchestration / n8n):
      0 = completed, every device clean (compliant / would-rename / renamed)
      1 = completed, but one or more devices need attention (errors, skips, failed renames)
      2 = fatal (bad arguments, unreadable/empty input, authentication failure)
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [string] $CsvPath,
    [string] $InputJson,
    [switch] $Stdin,

    [string] $IdColumnName,

    [ValidatePattern('^.$')]
    [string] $Delimiter = ',',

    [ValidateSet('ClientSecret', 'Interactive')]
    [string] $AuthMode = 'ClientSecret',

    [string] $TenantId,
    [string] $ClientId,
    [string] $ClientSecret,

    [switch] $DryRun,
    [switch] $Rename,
    [switch] $Force,

    [ValidateRange(1, 63)]
    [int] $MaxNameLength = 15,

    [string] $ProfileCountryRegex = '(?<cc>[A-Za-z]{3})\s*$',
    [switch] $SkipCountryCodeValidation,

    [switch] $PassThruJson,

    [string] $GraphBaseUri = 'https://graph.microsoft.com/beta',
    [string] $OutputDirectory,
    [string] $ReportPath,
    [string] $JsonSummaryPath,
    [string] $LogDirectory,

    [ValidateRange(0, 10)]
    [int] $MaxRetries = 5
)

Set-StrictMode -Version 1.0
$ErrorActionPreference = 'Stop'

$ExitSuccess = 0; $ExitAttention = 1; $ExitFatal = 2

$scriptRoot = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
. (Join-Path $scriptRoot 'IntuneGraphCommon.ps1')

$runStamp = (Get-Date).ToString('yyyyMMdd_HHmmss')
if (-not $LogDirectory)    { $LogDirectory    = Join-Path $scriptRoot 'logs' }
if (-not $OutputDirectory) { $OutputDirectory = Join-Path $scriptRoot 'output' }
if (-not $ReportPath)      { $ReportPath      = Join-Path $OutputDirectory ("RenameIntuneDevice_{0}.csv"  -f $runStamp) }
if (-not $JsonSummaryPath) { $JsonSummaryPath = Join-Path $OutputDirectory ("RenameIntuneDevice_{0}.json" -f $runStamp) }
foreach ($d in @($LogDirectory, $OutputDirectory)) {
    if (-not (Test-Path -LiteralPath $d)) { New-Item -ItemType Directory -Path $d -Force | Out-Null }
}
$script:GraphLogFile     = Join-Path $LogDirectory ("RenameIntuneDevice_{0}.log" -f $runStamp)
$script:GraphLogToStderr = [bool]$PassThruJson   # keep STDOUT clean when emitting JSON

# ISO 3166-1 alpha-3 country codes (case-insensitive) for country-code validation.
$script:Iso3166Alpha3 = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
$isoAlpha3Raw = @'
ABW AFG AGO AIA ALA ALB AND ARE ARG ARM ASM ATA ATF ATG AUS AUT AZE BDI BEL BEN BES BFA BGD
BGR BHR BHS BIH BLM BLR BLZ BMU BOL BRA BRB BRN BTN BVT BWA CAF CAN CCK CHE CHL CHN CIV CMR COD COG
COK COL COM CPV CRI CUB CUW CXR CYM CYP CZE DEU DJI DMA DNK DOM DZA ECU EGY ERI ESH ESP EST ETH FIN
FJI FLK FRA FRO FSM GAB GBR GEO GGY GHA GIB GIN GLP GMB GNB GNQ GRC GRD GRL GTM GUF GUM GUY HKG HMD
HND HRV HTI HUN IDN IMN IND IOT IRL IRN IRQ ISL ISR ITA JAM JEY JOR JPN KAZ KEN KGZ KHM KIR KNA KOR
KWT LAO LBN LBR LBY LCA LIE LKA LSO LTU LUX LVA MAC MAF MAR MCO MDA MDG MDV MEX MHL MKD MLI MLT MMR
MNE MNG MNP MOZ MRT MSR MTQ MUS MWI MYS MYT NAM NCL NER NFK NGA NIC NIU NLD NOR NPL NRU NZL OMN PAK
PAN PCN PER PHL PLW PNG POL PRI PRK PRT PRY PSE PYF QAT REU ROU RUS RWA SAU SDN SEN SGP SGS SHN SJM
SLB SLE SLV SMR SOM SPM SRB SSD STP SUR SVK SVN SWE SWZ SXM SYC SYR TCA TCD TGO THA TJK TKL TKM TLS
TON TTO TUN TUR TUV TWN TZA UGA UKR UMI URY USA UZB VAT VCT VEN VGB VIR VNM VUT WLF WSM YEM ZAF ZMB ZWE
'@
foreach ($code in ($isoAlpha3Raw -split '\s+')) { if ($code) { $null = $script:Iso3166Alpha3.Add($code) } }

function Get-CountryCodeFromProfile {
    param([string] $ProfileName)
    if ([string]::IsNullOrWhiteSpace($ProfileName)) { return $null }
    $m = [regex]::Match($ProfileName, $ProfileCountryRegex)
    if (-not $m.Success -or -not $m.Groups['cc'].Success) { return $null }
    $cc = $m.Groups['cc'].Value.ToUpperInvariant()
    if (-not $SkipCountryCodeValidation -and -not $script:Iso3166Alpha3.Contains($cc)) { return $null }
    return $cc
}

function Test-WindowsComputerName {
    param([Parameter(Mandatory = $true)][AllowEmptyString()][string] $Name, [int] $MaxLength = 15)
    if ([string]::IsNullOrWhiteSpace($Name)) { return @{ Valid = $false; Reason = 'Generated name is empty.'; Status = 'Error-InvalidName' } }
    if ($Name.Length -gt $MaxLength) { return @{ Valid = $false; Reason = ("Name '{0}' is {1} chars; exceeds max {2}." -f $Name, $Name.Length, $MaxLength); Status = 'Error-NameTooLong' } }
    if ($Name -match '[^A-Za-z0-9-]') { return @{ Valid = $false; Reason = ("Name '{0}' contains characters other than A-Z, 0-9 and hyphen." -f $Name); Status = 'Error-InvalidName' } }
    if ($Name -match '^\d+$') { return @{ Valid = $false; Reason = ("Name '{0}' is all numeric, which Windows disallows." -f $Name); Status = 'Error-InvalidName' } }
    if ($Name.StartsWith('-') -or $Name.EndsWith('-')) { return @{ Valid = $false; Reason = ("Name '{0}' must not start or end with a hyphen." -f $Name); Status = 'Error-InvalidName' } }
    return @{ Valid = $true; Reason = ''; Status = 'OK' }
}

function Resolve-AutopilotProfileName {
    # Fallback country-code source: the assigned Autopilot deployment profile display name.
    param([Parameter(Mandatory = $true)] $ManagedDevice)
    $serial = ("$($ManagedDevice.serialNumber)").Trim()
    if ([string]::IsNullOrWhiteSpace($serial)) { return @{ Found = $false; ProfileName = $null; Source = $null } }
    $serialEsc = $serial -replace "'", "''"
    $filter    = "contains(serialNumber,'$serialEsc')"
    $uri       = "$script:GraphBaseUri/deviceManagement/windowsAutopilotDeviceIdentities" + '?$filter=' + [uri]::EscapeDataString($filter)
    $candidates = Get-GraphCollection -Uri $uri
    if ($candidates.Count -eq 0) { return @{ Found = $false; ProfileName = $null; Source = $null } }
    $exact = @($candidates | Where-Object { ("$($_.serialNumber)").Trim() -ieq $serial })
    if ($exact.Count -eq 0) { $exact = @($candidates) }
    $ap = $null
    if ($exact.Count -gt 1) {
        $ap = $exact | Where-Object { $_.managedDeviceId -and ($_.managedDeviceId -ieq $ManagedDevice.id) } | Select-Object -First 1
        if (-not $ap -and $ManagedDevice.azureADDeviceId) {
            $ap = $exact | Where-Object { $_.azureActiveDirectoryDeviceId -and ($_.azureActiveDirectoryDeviceId -ieq $ManagedDevice.azureADDeviceId) } | Select-Object -First 1
        }
        if (-not $ap) { Write-Log -Level WARN -Message ("Serial '{0}' matched {1} Autopilot identities; using the first." -f $serial, $exact.Count); $ap = $exact | Select-Object -First 1 }
    }
    else { $ap = $exact | Select-Object -First 1 }
    foreach ($nav in @('deploymentProfile', 'intendedDeploymentProfile')) {
        $res = Invoke-GraphApi -Method GET -Uri "$script:GraphBaseUri/deviceManagement/windowsAutopilotDeviceIdentities/$($ap.id)/$nav"
        if ($res.Success -and $res.Data -and $res.Data.displayName) { return @{ Found = $true; ProfileName = $res.Data.displayName; Source = $nav } }
    }
    return @{ Found = $false; ProfileName = $null; Source = $null }
}

function Get-PropValue {
    param($Object, [string[]] $Names)
    foreach ($nm in $Names) {
        $p = $Object.PSObject.Properties | Where-Object { $_.Name -ieq $nm } | Select-Object -First 1
        if ($p) { return $p.Value }
    }
    return $null
}

try {
    Write-Log -Level INFO -Message '=========================================================='
    Write-Log -Level INFO -Message 'Rename-IntuneDevice starting.'
    Write-Log -Level INFO -Message ("PowerShell {0} ({1})" -f $PSVersionTable.PSVersion, $PSVersionTable.PSEdition)

    if ($Rename -and $DryRun) { Write-Log -Level WARN -Message 'Both -Rename and -DryRun supplied; running in DRY-RUN for safety.' }
    $doRename = ($Rename.IsPresent -and -not $DryRun.IsPresent)
    $modeName = if ($doRename) { 'RENAME' } else { 'DRY-RUN' }
    Write-Log -Level INFO -Message ("Mode: {0}" -f $modeName)
    if ($doRename -and $Force) { $ConfirmPreference = 'None' }

    # ----- Resolve input source (exactly one) --------------------------------
    $sourceCount = @($CsvPath, $InputJson).Where({ $_ }).Count + [int]$Stdin.IsPresent
    if ($sourceCount -ne 1) {
        Write-Log -Level ERROR -Message 'Specify exactly one input: -CsvPath, -InputJson, or -Stdin.'
        exit $ExitFatal
    }

    # A normalized input item = @{ InputId = <string>; Md = <pscustomobject or $null> }
    $inputItems = New-Object System.Collections.Generic.List[object]
    $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $totalInput = 0

    function Add-RecordObject {
        param($Obj)
        $script:totalInput++
        $idVal = ("$(Get-PropValue -Object $Obj -Names @('IntuneDeviceId','ManagedDeviceId','DeviceId','id'))").Trim()
        # Enriched (from the collector) if it carries a SerialNumber property -> use as-is, no fetch.
        $hasSerialProp = [bool]($Obj.PSObject.Properties | Where-Object { $_.Name -ieq 'SerialNumber' })
        $md = $null
        if ($idVal -and $hasSerialProp) {
            $md = [pscustomobject]@{
                id                    = $idVal
                deviceName            = Get-PropValue -Object $Obj -Names @('DeviceName', 'deviceName')
                serialNumber          = Get-PropValue -Object $Obj -Names @('SerialNumber', 'serialNumber')
                operatingSystem       = Get-PropValue -Object $Obj -Names @('OperatingSystem', 'operatingSystem')
                joinType              = Get-PropValue -Object $Obj -Names @('JoinType', 'joinType')
                azureADDeviceId       = Get-PropValue -Object $Obj -Names @('AzureADDeviceId', 'azureADDeviceId')
                enrollmentProfileName = Get-PropValue -Object $Obj -Names @('EnrollmentProfileName', 'enrollmentProfileName')
            }
        }
        $key = if ($idVal) { $idVal } else { [guid]::NewGuid().ToString() }
        if ($seen.Add($key)) { $inputItems.Add(@{ InputId = $idVal; Md = $md }) }
    }

    if ($CsvPath) {
        if (-not (Test-Path -LiteralPath $CsvPath)) { Write-Log -Level ERROR -Message ("CSV not found: {0}" -f $CsvPath); exit $ExitFatal }
        $rows = @(Import-Csv -LiteralPath $CsvPath -Delimiter $Delimiter)
        if ($rows.Count -eq 0) { Write-Log -Level ERROR -Message 'CSV contains no data rows.'; exit $ExitFatal }
        $columns = @($rows[0].psobject.Properties.Name)
        if (-not $IdColumnName) {
            foreach ($c in 'IntuneDeviceId', 'ManagedDeviceId', 'DeviceId', 'Id', 'intune_device_id', 'managedDeviceId') {
                $hit = $columns | Where-Object { $_ -ieq $c } | Select-Object -First 1
                if ($hit) { $IdColumnName = $hit; break }
            }
            if (-not $IdColumnName -and $columns.Count -eq 1) { $IdColumnName = $columns[0] }
        }
        else {
            $hit = $columns | Where-Object { $_ -ieq $IdColumnName } | Select-Object -First 1
            if (-not $hit) { Write-Log -Level ERROR -Message ("Column '{0}' not found. Available: {1}" -f $IdColumnName, ($columns -join ', ')); exit $ExitFatal }
            $IdColumnName = $hit
        }
        if (-not $IdColumnName) { Write-Log -Level ERROR -Message ("Cannot determine the ID column. Use -IdColumnName. Available: {0}" -f ($columns -join ', ')); exit $ExitFatal }
        Write-Log -Level INFO -Message ("Using CSV ID column '{0}'." -f $IdColumnName)
        foreach ($r in $rows) {
            $totalInput++
            $val = ("$($r.$IdColumnName)").Trim()
            if ($val -and $seen.Add($val)) { $inputItems.Add(@{ InputId = $val; Md = $null }) }
        }
    }
    else {
        if ($Stdin) {
            $raw = [Console]::In.ReadToEnd()
        } else {
            if (-not (Test-Path -LiteralPath $InputJson)) { Write-Log -Level ERROR -Message ("JSON input not found: {0}" -f $InputJson); exit $ExitFatal }
            $raw = Get-Content -LiteralPath $InputJson -Raw
        }
        if ([string]::IsNullOrWhiteSpace($raw)) { Write-Log -Level ERROR -Message 'No JSON input received.'; exit $ExitFatal }
        try { $parsed = $raw | ConvertFrom-Json } catch { Write-Log -Level ERROR -Message ("Input is not valid JSON: {0}" -f $_.Exception.Message); exit $ExitFatal }
        foreach ($el in @($parsed)) {
            if ($null -eq $el) { continue }
            if ($el -is [string]) {
                $totalInput++
                $v = $el.Trim()
                if ($v -and $seen.Add($v)) { $inputItems.Add(@{ InputId = $v; Md = $null }) }
            }
            else { Add-RecordObject -Obj $el }
        }
    }

    Write-Log -Level INFO -Message ("{0} input record(s), {1} unique device(s) to process." -f $totalInput, $inputItems.Count)
    if ($inputItems.Count -eq 0) { Write-Log -Level ERROR -Message 'No devices to process.'; exit $ExitFatal }

    # ----- Authenticate ------------------------------------------------------
    $scopes = @('DeviceManagementManagedDevices.Read.All', 'DeviceManagementServiceConfig.Read.All')
    if ($doRename) { $scopes += 'DeviceManagementManagedDevices.PrivilegedOperations.All' }
    Connect-Graph -AuthMode $AuthMode -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret `
        -Scopes $scopes -MaxRetries $MaxRetries -GraphBaseUri $GraphBaseUri

    # ----- Process -----------------------------------------------------------
    $results  = New-Object System.Collections.Generic.List[object]
    $mdSelect = 'id,deviceName,serialNumber,operatingSystem,azureADDeviceId,joinType,enrollmentProfileName'
    $index = 0

    foreach ($item in $inputItems) {
        $index++
        $id = $item.InputId
        if (-not $PassThruJson) { Write-Progress -Activity 'Processing devices' -Status ("{0}/{1}: {2}" -f $index, $inputItems.Count, $id) -PercentComplete (($index / $inputItems.Count) * 100) }

        $rec = [ordered]@{
            InputId = $id; IntuneDeviceId = $null; CurrentDeviceName = $null; SerialNumber = $null
            OperatingSystem = $null; JoinType = $null; ProfileName = $null; ProfileSource = $null
            CountryCode = $null; SuggestedName = $null; Status = $null; Action = 'None'
            Message = $null; Timestamp = (Get-Date).ToString('o')
        }

        try {
            $guidRef = [ref]([guid]::Empty)
            if (-not [guid]::TryParse($id, $guidRef)) {
                $rec.Status = 'Error-InvalidInput'
                $rec.Message = if ($null -ne $item.Md) { 'Record has no valid Intune managed-device ID (device may not be enrolled).' } else { 'Value is not a valid GUID (expected an Intune managed device ID).' }
                Write-Log -Level ERROR -Message ("[{0}] {1}" -f $id, $rec.Message)
                $results.Add([pscustomobject]$rec); continue
            }

            # 1) Obtain managed-device details (from enriched input, or by fetching).
            if ($null -ne $item.Md) {
                $md = $item.Md
            }
            else {
                $mdRes = Invoke-GraphApi -Method GET -Uri ("$script:GraphBaseUri/deviceManagement/managedDevices/$id" + '?$select=' + $mdSelect)
                if (-not $mdRes.Success) {
                    if ($mdRes.NotFound) { $rec.Status = 'Error-NotFound'; $rec.Message = 'Device ID not found in Intune.' }
                    else { $rec.Status = 'Error-GraphLookup'; $rec.Message = ("Managed device lookup failed ({0}): {1}" -f $mdRes.StatusCode, $mdRes.Error) }
                    Write-Log -Level ERROR -Message ("[{0}] {1}" -f $id, $rec.Message)
                    $results.Add([pscustomobject]$rec); continue
                }
                $md = $mdRes.Data
            }

            $rec.IntuneDeviceId    = $md.id
            $rec.CurrentDeviceName = $md.deviceName
            $rec.SerialNumber      = ("$($md.serialNumber)").Trim()
            $rec.OperatingSystem   = $md.operatingSystem
            $rec.JoinType          = $md.joinType

            # 2) Guard rails.
            if ([string]::IsNullOrWhiteSpace($rec.SerialNumber)) {
                $rec.Status = 'Error-NoSerialNumber'; $rec.Message = 'Managed device has no serial number in Intune.'
                Write-Log -Level ERROR -Message ("[{0}] {1}" -f $id, $rec.Message); $results.Add([pscustomobject]$rec); continue
            }
            if ($md.joinType -and $md.joinType -ieq 'hybridAzureADJoined') {
                $rec.Status = 'Skipped-HybridJoined'; $rec.Message = 'Entra hybrid-joined devices cannot be renamed via Intune; rename on-premises (AD) instead.'
                Write-Log -Level WARN -Message ("[{0}] {1}" -f $id, $rec.Message); $results.Add([pscustomobject]$rec); continue
            }
            if ($md.operatingSystem -and $md.operatingSystem -notmatch '^(?i)windows') {
                $rec.Status = 'Skipped-NonWindows'; $rec.Message = ("OS '{0}' is not Windows; this naming convention targets Windows hosts." -f $md.operatingSystem)
                Write-Log -Level WARN -Message ("[{0}] {1}" -f $id, $rec.Message); $results.Add([pscustomobject]$rec); continue
            }

            # 3) Country code from enrollment profile (fallback: Autopilot deployment profile).
            $profileName = $null; $profileSource = $null
            if (-not [string]::IsNullOrWhiteSpace($md.enrollmentProfileName)) {
                $profileName = ("$($md.enrollmentProfileName)").Trim(); $profileSource = 'enrollmentProfileName'
            }
            else {
                try {
                    $apInfo = Resolve-AutopilotProfileName -ManagedDevice $md
                    if ($apInfo.Found) { $profileName = $apInfo.ProfileName; $profileSource = $apInfo.Source }
                } catch { Write-Log -Level WARN -Message ("[{0}] Autopilot profile fallback failed: {1}" -f $id, $_.Exception.Message) }
            }
            if ([string]::IsNullOrWhiteSpace($profileName)) {
                $rec.Status = 'Error-NoAutopilotProfile'; $rec.Message = 'No enrollment/Autopilot profile found; device needs re-enrollment.'
                Write-Log -Level ERROR -Message ("[{0}] {1}" -f $id, $rec.Message); $results.Add([pscustomobject]$rec); continue
            }
            $rec.ProfileName = $profileName; $rec.ProfileSource = $profileSource

            $cc = Get-CountryCodeFromProfile -ProfileName $profileName
            if (-not $cc) {
                $rec.Status = 'Error-CountryCodeParse'; $rec.Message = ("Could not derive a valid ISO 3166-1 alpha-3 country code from profile '{0}'." -f $profileName)
                Write-Log -Level ERROR -Message ("[{0}] {1}" -f $id, $rec.Message); $results.Add([pscustomobject]$rec); continue
            }
            $rec.CountryCode = $cc

            # 4) Build + validate the suggested name.
            $suggested = ($cc + 'CF' + $rec.SerialNumber).ToUpperInvariant()
            $rec.SuggestedName = $suggested
            $nameCheck = Test-WindowsComputerName -Name $suggested -MaxLength $MaxNameLength
            if (-not $nameCheck.Valid) {
                $rec.Status = $nameCheck.Status; $rec.Message = $nameCheck.Reason
                Write-Log -Level ERROR -Message ("[{0}] {1}" -f $id, $rec.Message); $results.Add([pscustomobject]$rec); continue
            }

            # 5) Already compliant?
            if ($rec.CurrentDeviceName -and ($rec.CurrentDeviceName -ieq $suggested)) {
                $rec.Status = 'AlreadyCompliant'; $rec.Message = 'Current name already matches the naming convention.'
                Write-Log -Level INFO -Message ("[{0}] Already compliant: {1}" -f $id, $suggested); $results.Add([pscustomobject]$rec); continue
            }

            # 6) Apply or preview.
            if ($doRename) {
                $target = "{0} ({1})" -f $rec.CurrentDeviceName, $md.id
                if ($PSCmdlet.ShouldProcess($target, ("Set device name to '{0}'" -f $suggested))) {
                    $renRes = Invoke-GraphApi -Method POST -Uri "$script:GraphBaseUri/deviceManagement/managedDevices/$($md.id)/setDeviceName" -Body @{ deviceName = $suggested }
                    if ($renRes.Success) {
                        $rec.Status = 'Renamed'; $rec.Action = 'Renamed'
                        $rec.Message = ("Rename requested: '{0}' -> '{1}'. Applies at next reboot; no restart was triggered." -f $rec.CurrentDeviceName, $suggested)
                        Write-Log -Level SUCCESS -Message ("[{0}] {1}" -f $id, $rec.Message)
                    }
                    else {
                        $rec.Status = 'RenameFailed'
                        $hint = if ($renRes.StatusCode -eq 403) { ' (check the app has DeviceManagementManagedDevices.PrivilegedOperations.All with admin consent)' } else { '' }
                        $rec.Message = ("setDeviceName failed ({0}): {1}{2}" -f $renRes.StatusCode, $renRes.Error, $hint)
                        Write-Log -Level ERROR -Message ("[{0}] {1}" -f $id, $rec.Message)
                    }
                }
                else {
                    $rec.Status = 'Skipped-NotConfirmed'; $rec.Message = 'Rename not confirmed (WhatIf/declined).'
                    Write-Log -Level WARN -Message ("[{0}] {1}" -f $id, $rec.Message)
                }
            }
            else {
                $rec.Status = 'WouldRename'; $rec.Action = 'WouldRename'
                $rec.Message = ("Suggested rename: '{0}' -> '{1}'." -f $rec.CurrentDeviceName, $suggested)
                Write-Log -Level INFO -Message ("[{0}] {1}" -f $id, $rec.Message)
            }
        }
        catch {
            $rec.Status = 'Error-Unhandled'; $rec.Message = $_.Exception.Message
            Write-Log -Level ERROR -Message ("[{0}] Unhandled error: {1}" -f $id, $_.Exception.Message)
        }
        $results.Add([pscustomobject]$rec)
    }
    if (-not $PassThruJson) { Write-Progress -Activity 'Processing devices' -Completed }

    # ----- Output ------------------------------------------------------------
    if (-not $PassThruJson) {
        Write-Host ''
        $results | Select-Object CurrentDeviceName, SerialNumber, CountryCode, SuggestedName, Status | Format-Table -AutoSize | Out-Host
    }

    $summary = $results | Group-Object Status | Sort-Object Name | ForEach-Object { '{0}={1}' -f $_.Name, $_.Count }
    Write-Log -Level INFO -Message ("Summary ({0}): {1}" -f $modeName, ($summary -join ', '))

    try { $results | Export-Csv -LiteralPath $ReportPath -NoTypeInformation -Encoding UTF8; Write-Log -Level SUCCESS -Message ("CSV report written: {0}" -f $ReportPath) }
    catch { Write-Log -Level WARN -Message ("Failed to write CSV report: {0}" -f $_.Exception.Message) }

    $countsByStatus = @{}
    foreach ($g in ($results | Group-Object Status)) { $countsByStatus[$g.Name] = $g.Count }
    $jsonObject = [ordered]@{
        startedAtUtc = (Get-Date).ToUniversalTime().ToString('o'); mode = $modeName; authMode = $AuthMode
        tenantId = $script:GraphTenantId; totalInput = $totalInput; totalProcessed = $results.Count
        countsByStatus = $countsByStatus; results = $results
    }
    try { $jsonObject | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $JsonSummaryPath -Encoding UTF8; Write-Log -Level SUCCESS -Message ("JSON summary written: {0}" -f $JsonSummaryPath) }
    catch { Write-Log -Level WARN -Message ("Failed to write JSON summary: {0}" -f $_.Exception.Message) }

    if ($PassThruJson) { Write-Output ($jsonObject | ConvertTo-Json -Depth 6) }

    # ----- Exit code ---------------------------------------------------------
    $cleanStatuses = @('AlreadyCompliant', 'WouldRename', 'Renamed')
    $attention = @($results | Where-Object { $cleanStatuses -notcontains $_.Status })
    if ($attention.Count -gt 0) { Write-Log -Level WARN -Message ("Completed with {0} device(s) needing attention." -f $attention.Count); exit $ExitAttention }
    Write-Log -Level SUCCESS -Message 'Completed successfully; no devices need attention.'
    exit $ExitSuccess
}
catch {
    Write-Log -Level ERROR -Message ("FATAL: {0}" -f $_.Exception.Message)
    Write-Log -Level DEBUG -Message ($_.ScriptStackTrace)
    exit $ExitFatal
}
finally { Clear-GraphState }
