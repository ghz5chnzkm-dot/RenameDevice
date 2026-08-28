#Requires -Version 5.1
<#
.SYNOPSIS
    Suggests and (optionally) applies standardized Intune device names based on the
    "<CountryCode-ISO3166-1-Alpha3>CF<SerialNumber>" naming convention.

.DESCRIPTION
    Reads a CSV of Intune managed device IDs (GUIDs), and for each device:
      1. Looks up the managed device in Microsoft Graph (serial number, current name, join type, OS).
      2. Resolves the assigned Windows Autopilot deployment profile to derive the ISO 3166-1
         alpha-3 country code (e.g. profile "Autopilot User Driven profile MYS" -> "MYS").
      3. Builds the suggested name:  <CountryCode> + "CF" + <SerialNumber>  (parts joined, no separators).
      4. Validates the name against Windows computer-name rules (length + allowed characters).

    Two operating modes (mutually safe):
      * Default / -DryRun : PREVIEW ONLY. Shows the suggested new name for every device. No writes.
      * -Rename           : Applies the rename via the Graph 'setDeviceName' action.
                            This sets the device name only; it does NOT restart the device
                            ("Restart after rename" is intentionally not triggered). The new name
                            takes effect at the next user-initiated reboot.

    If a device has no Autopilot deployment profile assigned, it is reported as requiring
    re-enrollment (no country code can be derived) and is never renamed.

    Authentication is app-only (client credentials) against Microsoft Graph, suitable for
    unattended execution on a Windows Server and for orchestration platforms such as n8n.

.PARAMETER CsvPath
    Path to the input CSV. Must contain a column of Intune managed device IDs (GUIDs).
    The column is auto-detected (IntuneDeviceId / ManagedDeviceId / DeviceId / Id) unless
    -IdColumnName is supplied. If the CSV has a single column, that column is used.

.PARAMETER IdColumnName
    Explicit CSV column header that holds the Intune device ID. Optional (auto-detected otherwise).

.PARAMETER Delimiter
    Single-character CSV delimiter. Default ','. Use ';' for typical European-locale exports.

.PARAMETER TenantId
    Entra tenant ID (GUID) or verified domain. Falls back to env var RENAMEDEVICE_TENANT_ID.

.PARAMETER ClientId
    App registration (client) ID. Falls back to env var RENAMEDEVICE_CLIENT_ID.

.PARAMETER ClientSecret
    App registration client secret. PREFER supplying this via env var RENAMEDEVICE_CLIENT_SECRET
    (avoids exposure in process listings / shell history). Used only to obtain the token.

.PARAMETER DryRun
    Explicit preview mode. Same behaviour as running with no mode switch. No writes are performed.

.PARAMETER Rename
    Perform the actual rename via Graph. Ignored (safety wins) if -DryRun is also specified.

.PARAMETER Force
    Suppress the per-device confirmation prompt in -Rename mode. Required for unattended / n8n runs.

.PARAMETER MaxNameLength
    Maximum allowed length of the generated name. Default 15 (Windows NetBIOS/hostname limit).
    Names exceeding this are flagged and skipped (never truncated). Range 1-63.

.PARAMETER ProfileCountryRegex
    Regex (with a named capture group 'cc') used to extract the country code from the deployment
    profile display name. Default extracts the trailing 3-letter token.

.PARAMETER SkipCountryCodeValidation
    Skip validating the extracted code against the ISO 3166-1 alpha-3 list (use if your codes are custom).

.PARAMETER GraphBaseUri
    Graph endpoint base. Default 'https://graph.microsoft.com/beta' (setDeviceName is beta-only).

.PARAMETER OutputDirectory
    Directory for the CSV report and JSON summary. Default: <script folder>\output.

.PARAMETER ReportPath
    Explicit path for the CSV result report (overrides OutputDirectory).

.PARAMETER JsonSummaryPath
    Explicit path for the machine-readable JSON summary (overrides OutputDirectory). Ideal for n8n.

.PARAMETER LogDirectory
    Directory for the run log. Default: <script folder>\logs.

.PARAMETER MaxRetries
    Max retry attempts for transient Graph failures (429/5xx/network). Default 5.

.EXAMPLE
    .\Rename-IntuneDevice.ps1 -CsvPath .\devices.csv -TenantId <guid> -ClientId <guid>
    Preview (dry run). Prints suggested names; writes report + JSON summary; makes no changes.

.EXAMPLE
    .\Rename-IntuneDevice.ps1 -CsvPath .\devices.csv -Rename -Force
    Applies renames unattended (credentials read from RENAMEDEVICE_* environment variables).

.NOTES
    Author       : Heidelberg Materials IT
    Created      : 2026-08-28
    Compatibility: Windows PowerShell 5.1 and PowerShell 7+
    Runs on      : Windows Server (admin/automation host), not on the managed endpoint.

    Required Microsoft Graph APPLICATION permissions (admin consent required):
      * DeviceManagementManagedDevices.Read.All                 (read managed devices)
      * DeviceManagementManagedDevices.PrivilegedOperations.All  (setDeviceName / rename)
      * DeviceManagementServiceConfig.Read.All                   (read Autopilot profiles)

    Exit codes (for orchestration / n8n):
      0 = completed, every device ended in a clean state (compliant / would-rename / renamed)
      1 = completed, but one or more devices need attention (errors, skips, failed renames)
      2 = fatal (bad arguments, unreadable/empty CSV, authentication failure)
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [Parameter(Mandatory = $true)]
    [string] $CsvPath,

    [string] $IdColumnName,

    [ValidatePattern('^.$')]
    [string] $Delimiter = ',',

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

    [string] $GraphBaseUri = 'https://graph.microsoft.com/beta',

    [string] $OutputDirectory,

    [string] $ReportPath,

    [string] $JsonSummaryPath,

    [string] $LogDirectory,

    [ValidateRange(0, 10)]
    [int] $MaxRetries = 5
)

# Version 1.0 (not Latest): still flags uninitialized variables, but lets absent
# object properties read as $null. Microsoft Graph omits null-valued properties,
# so stricter modes would throw on a device that simply has no serial number, etc.
Set-StrictMode -Version 1.0
$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# Constants / paths
# ---------------------------------------------------------------------------
$ExitSuccess   = 0
$ExitAttention = 1
$ExitFatal     = 2

$scriptRoot = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$runStamp   = (Get-Date).ToString('yyyyMMdd_HHmmss')

if (-not $LogDirectory)    { $LogDirectory    = Join-Path $scriptRoot 'logs' }
if (-not $OutputDirectory) { $OutputDirectory = Join-Path $scriptRoot 'output' }
if (-not $ReportPath)      { $ReportPath      = Join-Path $OutputDirectory ("RenameIntuneDevice_{0}.csv"  -f $runStamp) }
if (-not $JsonSummaryPath) { $JsonSummaryPath = Join-Path $OutputDirectory ("RenameIntuneDevice_{0}.json" -f $runStamp) }

foreach ($d in @($LogDirectory, $OutputDirectory)) {
    if (-not (Test-Path -LiteralPath $d)) { New-Item -ItemType Directory -Path $d -Force | Out-Null }
}
$script:LogFile = Join-Path $LogDirectory ("RenameIntuneDevice_{0}.log" -f $runStamp)

# Ensure TLS 1.2 (older Windows PowerShell defaults can break the token endpoint).
try {
    if (([Net.ServicePointManager]::SecurityProtocol -band [Net.SecurityProtocolType]::Tls12) -ne [Net.SecurityProtocolType]::Tls12) {
        [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    }
} catch { }

# ISO 3166-1 alpha-3 country codes (case-insensitive set) for country-code validation.
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
foreach ($code in ($isoAlpha3Raw -split '\s+')) {
    if ($code) { $null = $script:Iso3166Alpha3.Add($code) }
}

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------
function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string] $Message,
        [ValidateSet('INFO', 'WARN', 'ERROR', 'SUCCESS', 'DEBUG')][string] $Level = 'INFO'
    )
    $line = ('{0} [{1}] {2}' -f (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'), $Level, $Message)
    try { Add-Content -LiteralPath $script:LogFile -Value $line -Encoding UTF8 } catch { }
    switch ($Level) {
        'ERROR'   { Write-Host $line -ForegroundColor Red }
        'WARN'    { Write-Host $line -ForegroundColor Yellow }
        'SUCCESS' { Write-Host $line -ForegroundColor Green }
        'DEBUG'   { Write-Verbose $line }
        default   { Write-Host $line }
    }
}

# ---------------------------------------------------------------------------
# Graph error parsing (cross-version: Windows PowerShell 5.1 & PowerShell 7+)
# ---------------------------------------------------------------------------
function Get-GraphErrorDetail {
    param([Parameter(Mandatory = $true)] $ErrorRecord)

    $info = [ordered]@{ StatusCode = $null; RetryAfter = $null; Message = $null }
    $ex   = $ErrorRecord.Exception

    try { if ($null -ne $ex.Response) { $info.StatusCode = [int]$ex.Response.StatusCode } } catch { }

    $body = $null
    if ($ErrorRecord.ErrorDetails -and $ErrorRecord.ErrorDetails.Message) {
        $body = $ErrorRecord.ErrorDetails.Message
    }
    elseif ($null -ne $ex.Response) {
        try {
            $stream = $ex.Response.GetResponseStream()
            if ($stream) {
                $reader = New-Object System.IO.StreamReader($stream)
                $body = $reader.ReadToEnd()
                $reader.Close()
            }
        } catch { }
    }

    if ($body) {
        try {
            $parsed = $body | ConvertFrom-Json
            if ($parsed.error -and $parsed.error.message) { $info.Message = $parsed.error.message }
            else { $info.Message = $body }
        } catch { $info.Message = $body }
    }
    else { $info.Message = $ex.Message }

    try {
        if ($null -ne $ex.Response -and $null -ne $ex.Response.Headers) {
            $ra = $null
            try { $ra = $ex.Response.Headers['Retry-After'] } catch { }
            if (-not $ra) { try { $ra = $ex.Response.Headers.GetValues('Retry-After') | Select-Object -First 1 } catch { } }
            if (-not $ra) {
                try {
                    if ($ex.Response.Headers.RetryAfter -and $ex.Response.Headers.RetryAfter.Delta) {
                        $ra = [int]$ex.Response.Headers.RetryAfter.Delta.TotalSeconds
                    }
                } catch { }
            }
            if ($ra) { $info.RetryAfter = [int]$ra }
        }
    } catch { }

    return $info
}

# ---------------------------------------------------------------------------
# Token acquisition (client credentials) + caching / refresh
# ---------------------------------------------------------------------------
$script:TokenCache = $null   # @{ AccessToken = ''; ExpiresOn = [datetime] }

function Get-GraphToken {
    param(
        [Parameter(Mandatory = $true)][string] $TenantId,
        [Parameter(Mandatory = $true)][string] $ClientId,
        [Parameter(Mandatory = $true)][securestring] $ClientSecret
    )

    $bstr  = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($ClientSecret)
    try { $plain = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr) }
    finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }

    $body = @{
        client_id     = $ClientId
        scope         = 'https://graph.microsoft.com/.default'
        client_secret = $plain
        grant_type    = 'client_credentials'
    }
    $uri = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"

    try {
        $attempt = 0
        while ($true) {
            $attempt++
            try {
                $resp = Invoke-RestMethod -Method Post -Uri $uri -Body $body -ContentType 'application/x-www-form-urlencoded' -ErrorAction Stop
                break
            } catch {
                $det = Get-GraphErrorDetail -ErrorRecord $_
                $retryable = ($null -eq $det.StatusCode) -or ($det.StatusCode -in 429, 500, 502, 503, 504)
                if ($retryable -and $attempt -le $MaxRetries) {
                    $delay = if ($det.RetryAfter) { $det.RetryAfter } else { [int][Math]::Min(60, [Math]::Pow(2, $attempt)) }
                    Write-Log -Level WARN -Message ("Token request failed (attempt {0}/{1}, status {2}); retrying in {3}s." -f $attempt, $MaxRetries, $det.StatusCode, $delay)
                    Start-Sleep -Seconds $delay
                    continue
                }
                throw ("Failed to acquire Graph token: {0}" -f $det.Message)
            }
        }
    }
    finally {
        # Best-effort scrub of the plaintext secret from memory.
        if (Get-Variable -Name plain -Scope Local -ErrorAction SilentlyContinue) { $plain = $null }
    }

    $expiresIn = if ($resp.expires_in) { [int]$resp.expires_in } else { 3600 }
    return @{
        AccessToken = $resp.access_token
        # Refresh 2 minutes early to avoid mid-run expiry on long batches.
        ExpiresOn   = (Get-Date).ToUniversalTime().AddSeconds($expiresIn - 120)
    }
}

function Get-ValidToken {
    param([switch] $ForceRefresh)
    if ($ForceRefresh -or $null -eq $script:TokenCache -or (Get-Date).ToUniversalTime() -ge $script:TokenCache.ExpiresOn) {
        $script:TokenCache = Get-GraphToken -TenantId $script:TenantId -ClientId $script:ClientId -ClientSecret $script:ClientSecretSecure
    }
    return $script:TokenCache.AccessToken
}

# ---------------------------------------------------------------------------
# Graph request wrapper (retry, throttling, auth refresh)
# ---------------------------------------------------------------------------
function Invoke-GraphApi {
    param(
        [Parameter(Mandatory = $true)][ValidateSet('GET', 'POST', 'PATCH', 'DELETE')][string] $Method,
        [Parameter(Mandatory = $true)][string] $Uri,
        $Body
    )

    $jsonBody = $null
    if ($null -ne $Body) { $jsonBody = ($Body | ConvertTo-Json -Depth 6 -Compress) }

    $attempt      = 0
    $authRefreshed = $false

    while ($true) {
        $attempt++
        $token   = Get-ValidToken
        $headers = @{ Authorization = "Bearer $token"; Accept = 'application/json' }

        try {
            if ($null -ne $jsonBody) {
                $data = Invoke-RestMethod -Method $Method -Uri $Uri -Headers $headers -Body $jsonBody -ContentType 'application/json' -ErrorAction Stop
            } else {
                $data = Invoke-RestMethod -Method $Method -Uri $Uri -Headers $headers -ErrorAction Stop
            }
            return @{ Success = $true; Data = $data; StatusCode = 200 }
        }
        catch {
            $det  = Get-GraphErrorDetail -ErrorRecord $_
            $code = $det.StatusCode

            if ($code -eq 404) {
                return @{ Success = $false; NotFound = $true; StatusCode = 404; Error = $det.Message }
            }

            if ($code -eq 401 -and -not $authRefreshed) {
                Write-Log -Level DEBUG -Message 'Received 401; refreshing access token and retrying.'
                $null = Get-ValidToken -ForceRefresh
                $authRefreshed = $true
                continue
            }

            $retryable = ($null -eq $code) -or ($code -in 429, 500, 502, 503, 504)
            if ($retryable -and $attempt -le $MaxRetries) {
                $delay = if ($det.RetryAfter) { $det.RetryAfter } else { [int][Math]::Min(60, [Math]::Pow(2, $attempt)) }
                Write-Log -Level WARN -Message ("Graph {0} failed (attempt {1}/{2}, status {3}); retrying in {4}s." -f $Method, $attempt, $MaxRetries, $code, $delay)
                Start-Sleep -Seconds $delay
                continue
            }

            return @{ Success = $false; NotFound = $false; StatusCode = $code; Error = $det.Message }
        }
    }
}

function Get-GraphCollection {
    # GET a collection, following @odata.nextLink. Returns an array (may be empty).
    param([Parameter(Mandatory = $true)][string] $Uri)

    $items = New-Object System.Collections.Generic.List[object]
    $next  = $Uri
    while ($next) {
        $res = Invoke-GraphApi -Method GET -Uri $next
        if (-not $res.Success) {
            if ($res.NotFound) { break }
            throw ("Graph collection GET failed ({0}): {1}" -f $res.StatusCode, $res.Error)
        }
        if ($res.Data.value) { foreach ($v in $res.Data.value) { $items.Add($v) } }
        $next = $null
        if ($res.Data.PSObject.Properties.Name -contains '@odata.nextLink') { $next = $res.Data.'@odata.nextLink' }
    }
    return $items
}

# ---------------------------------------------------------------------------
# Domain helpers
# ---------------------------------------------------------------------------
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
    # Returns @{ Valid = [bool]; Reason = ''; Status = '' } based on Windows rename rules.
    param([Parameter(Mandatory = $true)][AllowEmptyString()][string] $Name, [int] $MaxLength = 15)

    if ([string]::IsNullOrWhiteSpace($Name)) {
        return @{ Valid = $false; Reason = 'Generated name is empty.'; Status = 'Error-InvalidName' }
    }
    if ($Name.Length -gt $MaxLength) {
        return @{ Valid = $false; Reason = ("Name '{0}' is {1} chars; exceeds max {2}." -f $Name, $Name.Length, $MaxLength); Status = 'Error-NameTooLong' }
    }
    if ($Name -match '[^A-Za-z0-9-]') {
        return @{ Valid = $false; Reason = ("Name '{0}' contains characters other than A-Z, 0-9 and hyphen." -f $Name); Status = 'Error-InvalidName' }
    }
    if ($Name -match '^\d+$') {
        return @{ Valid = $false; Reason = ("Name '{0}' is all numeric, which Windows disallows." -f $Name); Status = 'Error-InvalidName' }
    }
    if ($Name.StartsWith('-') -or $Name.EndsWith('-')) {
        return @{ Valid = $false; Reason = ("Name '{0}' must not start or end with a hyphen." -f $Name); Status = 'Error-InvalidName' }
    }
    return @{ Valid = $true; Reason = ''; Status = 'OK' }
}

function Resolve-AutopilotProfileName {
    # Given a managed device, find its Autopilot identity and return the assigned deployment profile display name.
    # Returns @{ ProfileName = ''; Source = 'deploymentProfile|intendedDeploymentProfile'; Found = [bool] }
    param([Parameter(Mandatory = $true)] $ManagedDevice)

    $serial = ("$($ManagedDevice.serialNumber)").Trim()
    if ([string]::IsNullOrWhiteSpace($serial)) { return @{ Found = $false; ProfileName = $null; Source = $null } }

    $serialEsc = $serial -replace "'", "''"
    $filter    = "contains(serialNumber,'$serialEsc')"
    $uri       = "$GraphBaseUri/deviceManagement/windowsAutopilotDeviceIdentities" + '?$filter=' + [uri]::EscapeDataString($filter)

    $candidates = Get-GraphCollection -Uri $uri
    if ($candidates.Count -eq 0) { return @{ Found = $false; ProfileName = $null; Source = $null } }

    # Prefer an exact serial match; disambiguate by managedDeviceId / azureADDeviceId when several share a serial.
    $exact = @($candidates | Where-Object { ("$($_.serialNumber)").Trim() -ieq $serial })
    if ($exact.Count -eq 0) { $exact = @($candidates) }

    $ap = $null
    if ($exact.Count -gt 1) {
        $ap = $exact | Where-Object { $_.managedDeviceId -and ($_.managedDeviceId -ieq $ManagedDevice.id) } | Select-Object -First 1
        if (-not $ap -and $ManagedDevice.azureADDeviceId) {
            $ap = $exact | Where-Object { $_.azureActiveDirectoryDeviceId -and ($_.azureActiveDirectoryDeviceId -ieq $ManagedDevice.azureADDeviceId) } | Select-Object -First 1
        }
        if (-not $ap) {
            Write-Log -Level WARN -Message ("Serial '{0}' matched {1} Autopilot identities; using the first. Verify the result." -f $serial, $exact.Count)
            $ap = $exact | Select-Object -First 1
        }
    }
    else { $ap = $exact | Select-Object -First 1 }

    foreach ($nav in @('deploymentProfile', 'intendedDeploymentProfile')) {
        $pUri = "$GraphBaseUri/deviceManagement/windowsAutopilotDeviceIdentities/$($ap.id)/$nav"
        $res  = Invoke-GraphApi -Method GET -Uri $pUri
        if ($res.Success -and $res.Data -and $res.Data.displayName) {
            return @{ Found = $true; ProfileName = $res.Data.displayName; Source = $nav }
        }
    }
    return @{ Found = $false; ProfileName = $null; Source = $null }
}

# ---------------------------------------------------------------------------
# Preflight: mode, credentials, input
# ---------------------------------------------------------------------------
try {
    Write-Log -Level INFO -Message '=========================================================='
    Write-Log -Level INFO -Message 'Rename-IntuneDevice starting.'
    Write-Log -Level INFO -Message ("PowerShell {0} ({1})" -f $PSVersionTable.PSVersion, $PSVersionTable.PSEdition)

    # Resolve mode.
    if ($Rename -and $DryRun) {
        Write-Log -Level WARN -Message 'Both -Rename and -DryRun supplied; running in DRY-RUN for safety.'
    }
    $doRename = ($Rename.IsPresent -and -not $DryRun.IsPresent)
    $modeName = if ($doRename) { 'RENAME' } else { 'DRY-RUN' }
    Write-Log -Level INFO -Message ("Mode: {0}" -f $modeName)

    if ($doRename -and $Force) { $ConfirmPreference = 'None' }

    # Resolve credentials (parameters override, else environment variables).
    if (-not $TenantId)     { $TenantId     = $env:RENAMEDEVICE_TENANT_ID }
    if (-not $ClientId)     { $ClientId     = $env:RENAMEDEVICE_CLIENT_ID }
    if (-not $ClientSecret) { $ClientSecret = $env:RENAMEDEVICE_CLIENT_SECRET }
    elseif ($PSBoundParameters.ContainsKey('ClientSecret')) {
        Write-Log -Level WARN -Message 'ClientSecret passed as a parameter; prefer the RENAMEDEVICE_CLIENT_SECRET environment variable for unattended runs.'
    }

    $missing = @()
    if ([string]::IsNullOrWhiteSpace($TenantId))     { $missing += 'TenantId (or RENAMEDEVICE_TENANT_ID)' }
    if ([string]::IsNullOrWhiteSpace($ClientId))     { $missing += 'ClientId (or RENAMEDEVICE_CLIENT_ID)' }
    if ([string]::IsNullOrWhiteSpace($ClientSecret)) { $missing += 'ClientSecret (or RENAMEDEVICE_CLIENT_SECRET)' }
    if ($missing.Count -gt 0) {
        Write-Log -Level ERROR -Message ("Missing required credential(s): {0}" -f ($missing -join ', '))
        exit $ExitFatal
    }

    $script:TenantId           = $TenantId
    $script:ClientId           = $ClientId
    $script:ClientSecretSecure = ConvertTo-SecureString -String $ClientSecret -AsPlainText -Force
    $ClientSecret = $null  # drop the plaintext copy

    # Load + validate CSV.
    if (-not (Test-Path -LiteralPath $CsvPath)) {
        Write-Log -Level ERROR -Message ("CSV not found: {0}" -f $CsvPath)
        exit $ExitFatal
    }
    $rows = @(Import-Csv -LiteralPath $CsvPath -Delimiter $Delimiter)
    if ($rows.Count -eq 0) {
        Write-Log -Level ERROR -Message ("CSV '{0}' contains no data rows." -f $CsvPath)
        exit $ExitFatal
    }

    $columns = @($rows[0].psobject.Properties.Name)
    if (-not $IdColumnName) {
        $candidates = 'IntuneDeviceId', 'ManagedDeviceId', 'DeviceId', 'Id', 'intune_device_id', 'managedDeviceId'
        foreach ($c in $candidates) {
            $hit = $columns | Where-Object { $_ -ieq $c } | Select-Object -First 1
            if ($hit) { $IdColumnName = $hit; break }
        }
        if (-not $IdColumnName -and $columns.Count -eq 1) { $IdColumnName = $columns[0] }
    }
    else {
        $hit = $columns | Where-Object { $_ -ieq $IdColumnName } | Select-Object -First 1
        if (-not $hit) {
            Write-Log -Level ERROR -Message ("Column '{0}' not found in CSV. Available: {1}" -f $IdColumnName, ($columns -join ', '))
            exit $ExitFatal
        }
        $IdColumnName = $hit
    }
    if (-not $IdColumnName) {
        Write-Log -Level ERROR -Message ("Cannot determine the ID column. Use -IdColumnName. Available: {0}" -f ($columns -join ', '))
        exit $ExitFatal
    }
    Write-Log -Level INFO -Message ("Using ID column '{0}'." -f $IdColumnName)

    # Collect distinct, non-empty IDs (preserve first-seen order).
    $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $ids  = New-Object System.Collections.Generic.List[string]
    foreach ($r in $rows) {
        $val = ("$($r.$IdColumnName)").Trim()
        if ($val -and $seen.Add($val)) { $ids.Add($val) }
    }
    Write-Log -Level INFO -Message ("{0} row(s) read, {1} unique device ID(s) to process." -f $rows.Count, $ids.Count)
    if ($ids.Count -eq 0) {
        Write-Log -Level ERROR -Message 'No device IDs to process after de-duplication.'
        exit $ExitFatal
    }

    # Validate connectivity / credentials up front (fail fast before the loop).
    $null = Get-ValidToken
    Write-Log -Level SUCCESS -Message 'Acquired Microsoft Graph access token.'

    # -----------------------------------------------------------------------
    # Main processing loop
    # -----------------------------------------------------------------------
    $results   = New-Object System.Collections.Generic.List[object]
    $mdSelect  = 'id,deviceName,serialNumber,operatingSystem,azureADDeviceId,joinType,manufacturer,model'
    $index     = 0

    foreach ($id in $ids) {
        $index++
        Write-Progress -Activity 'Processing devices' -Status ("{0}/{1}: {2}" -f $index, $ids.Count, $id) -PercentComplete (($index / $ids.Count) * 100)

        $rec = [ordered]@{
            InputId               = $id
            IntuneDeviceId        = $null
            CurrentDeviceName     = $null
            SerialNumber          = $null
            OperatingSystem       = $null
            JoinType              = $null
            DeploymentProfileName = $null
            CountryCode           = $null
            SuggestedName         = $null
            Status                = $null
            Action                = 'None'
            Message               = $null
            Timestamp             = (Get-Date).ToString('o')
        }

        try {
            # 0) Basic input sanity: must look like a GUID (Intune managedDevice id).
            $guidRef = [ref]([guid]::Empty)
            if (-not [guid]::TryParse($id, $guidRef)) {
                $rec.Status  = 'Error-InvalidInput'
                $rec.Message = 'Value is not a valid GUID (expected an Intune managed device ID).'
                Write-Log -Level ERROR -Message ("[{0}] {1}" -f $id, $rec.Message)
                $results.Add([pscustomobject]$rec); continue
            }

            # 1) Managed device lookup.
            $mdUri = "$GraphBaseUri/deviceManagement/managedDevices/$id" + '?$select=' + $mdSelect
            $mdRes = Invoke-GraphApi -Method GET -Uri $mdUri
            if (-not $mdRes.Success) {
                if ($mdRes.NotFound) {
                    $rec.Status  = 'Error-NotFound'
                    $rec.Message = 'Device ID not found in Intune.'
                } else {
                    $rec.Status  = 'Error-GraphLookup'
                    $rec.Message = ("Managed device lookup failed ({0}): {1}" -f $mdRes.StatusCode, $mdRes.Error)
                }
                Write-Log -Level ERROR -Message ("[{0}] {1}" -f $id, $rec.Message)
                $results.Add([pscustomobject]$rec); continue
            }
            $md = $mdRes.Data

            $rec.IntuneDeviceId    = $md.id
            $rec.CurrentDeviceName = $md.deviceName
            $rec.SerialNumber      = ("$($md.serialNumber)").Trim()
            $rec.OperatingSystem   = $md.operatingSystem
            $rec.JoinType          = $md.joinType

            # 2) Guard rails: serial, hybrid-join, non-Windows.
            if ([string]::IsNullOrWhiteSpace($rec.SerialNumber)) {
                $rec.Status  = 'Error-NoSerialNumber'
                $rec.Message = 'Managed device has no serial number in Intune.'
                Write-Log -Level ERROR -Message ("[{0}] {1}" -f $id, $rec.Message)
                $results.Add([pscustomobject]$rec); continue
            }
            if ($md.joinType -and $md.joinType -ieq 'hybridAzureADJoined') {
                $rec.Status  = 'Skipped-HybridJoined'
                $rec.Message = 'Entra hybrid-joined devices cannot be renamed via Intune; rename on-premises (AD) instead.'
                Write-Log -Level WARN -Message ("[{0}] {1}" -f $id, $rec.Message)
                $results.Add([pscustomobject]$rec); continue
            }
            if ($md.operatingSystem -and $md.operatingSystem -notmatch '^(?i)windows') {
                $rec.Status  = 'Skipped-NonWindows'
                $rec.Message = ("OS '{0}' is not Windows; this naming convention targets Windows hosts." -f $md.operatingSystem)
                Write-Log -Level WARN -Message ("[{0}] {1}" -f $id, $rec.Message)
                $results.Add([pscustomobject]$rec); continue
            }

            # 3) Country code from Autopilot deployment profile.
            $profileInfo = Resolve-AutopilotProfileName -ManagedDevice $md
            if (-not $profileInfo.Found) {
                $rec.Status  = 'Error-NoAutopilotProfile'
                $rec.Message = 'No Autopilot deployment profile is assigned; device needs re-enrollment.'
                Write-Log -Level ERROR -Message ("[{0}] {1}" -f $id, $rec.Message)
                $results.Add([pscustomobject]$rec); continue
            }
            $rec.DeploymentProfileName = $profileInfo.ProfileName

            $cc = Get-CountryCodeFromProfile -ProfileName $profileInfo.ProfileName
            if (-not $cc) {
                $rec.Status  = 'Error-CountryCodeParse'
                $rec.Message = ("Could not derive a valid ISO 3166-1 alpha-3 country code from profile '{0}'." -f $profileInfo.ProfileName)
                Write-Log -Level ERROR -Message ("[{0}] {1}" -f $id, $rec.Message)
                $results.Add([pscustomobject]$rec); continue
            }
            $rec.CountryCode = $cc

            # 4) Build + validate suggested name:  <CC> + CF + <Serial>  (uppercased, parts joined).
            $suggested       = ($cc + 'CF' + $rec.SerialNumber).ToUpperInvariant()
            $rec.SuggestedName = $suggested

            $nameCheck = Test-WindowsComputerName -Name $suggested -MaxLength $MaxNameLength
            if (-not $nameCheck.Valid) {
                $rec.Status  = $nameCheck.Status
                $rec.Message = $nameCheck.Reason
                Write-Log -Level ERROR -Message ("[{0}] {1}" -f $id, $rec.Message)
                $results.Add([pscustomobject]$rec); continue
            }

            # 5) Already compliant?
            if ($rec.CurrentDeviceName -and ($rec.CurrentDeviceName -ieq $suggested)) {
                $rec.Status  = 'AlreadyCompliant'
                $rec.Message = 'Current name already matches the naming convention.'
                Write-Log -Level INFO -Message ("[{0}] Already compliant: {1}" -f $id, $suggested)
                $results.Add([pscustomobject]$rec); continue
            }

            # 6) Apply or preview.
            if ($doRename) {
                $target = "{0} ({1})" -f $rec.CurrentDeviceName, $md.id
                if ($PSCmdlet.ShouldProcess($target, ("Set device name to '{0}'" -f $suggested))) {
                    $renUri = "$GraphBaseUri/deviceManagement/managedDevices/$($md.id)/setDeviceName"
                    $renRes = Invoke-GraphApi -Method POST -Uri $renUri -Body @{ deviceName = $suggested }
                    if ($renRes.Success) {
                        $rec.Status  = 'Renamed'
                        $rec.Action  = 'Renamed'
                        $rec.Message = ("Rename requested: '{0}' -> '{1}'. Applies at next reboot; no restart was triggered." -f $rec.CurrentDeviceName, $suggested)
                        Write-Log -Level SUCCESS -Message ("[{0}] {1}" -f $id, $rec.Message)
                    }
                    else {
                        $rec.Status  = 'RenameFailed'
                        $hint = ''
                        if ($renRes.StatusCode -eq 403) { $hint = ' (check the app has DeviceManagementManagedDevices.PrivilegedOperations.All with admin consent)' }
                        $rec.Message = ("setDeviceName failed ({0}): {1}{2}" -f $renRes.StatusCode, $renRes.Error, $hint)
                        Write-Log -Level ERROR -Message ("[{0}] {1}" -f $id, $rec.Message)
                    }
                }
                else {
                    $rec.Status  = 'Skipped-NotConfirmed'
                    $rec.Message = 'Rename not confirmed (WhatIf/declined).'
                    Write-Log -Level WARN -Message ("[{0}] {1}" -f $id, $rec.Message)
                }
            }
            else {
                $rec.Status  = 'WouldRename'
                $rec.Action  = 'WouldRename'
                $rec.Message = ("Suggested rename: '{0}' -> '{1}'." -f $rec.CurrentDeviceName, $suggested)
                Write-Log -Level INFO -Message ("[{0}] {1}" -f $id, $rec.Message)
            }
        }
        catch {
            $rec.Status  = 'Error-Unhandled'
            $rec.Message = $_.Exception.Message
            Write-Log -Level ERROR -Message ("[{0}] Unhandled error: {1}" -f $id, $_.Exception.Message)
        }

        $results.Add([pscustomobject]$rec)
    }
    Write-Progress -Activity 'Processing devices' -Completed

    # -----------------------------------------------------------------------
    # Output: console table, CSV report, JSON summary
    # -----------------------------------------------------------------------
    Write-Host ''
    $results |
        Select-Object CurrentDeviceName, SerialNumber, CountryCode, SuggestedName, Status |
        Format-Table -AutoSize | Out-Host

    $summary = $results | Group-Object Status | Sort-Object Name | ForEach-Object { '{0}={1}' -f $_.Name, $_.Count }
    Write-Log -Level INFO -Message ("Summary ({0}): {1}" -f $modeName, ($summary -join ', '))

    try {
        $results | Export-Csv -LiteralPath $ReportPath -NoTypeInformation -Encoding UTF8
        Write-Log -Level SUCCESS -Message ("CSV report written: {0}" -f $ReportPath)
    } catch { Write-Log -Level WARN -Message ("Failed to write CSV report: {0}" -f $_.Exception.Message) }

    $countsByStatus = @{}
    foreach ($g in ($results | Group-Object Status)) { $countsByStatus[$g.Name] = $g.Count }

    $jsonObject = [ordered]@{
        startedAtUtc   = (Get-Date).ToUniversalTime().ToString('o')
        mode           = $modeName
        tenantId       = $script:TenantId
        idColumn       = $IdColumnName
        totalInput     = $rows.Count
        totalProcessed = $results.Count
        countsByStatus = $countsByStatus
        results        = $results
    }
    try {
        $jsonObject | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $JsonSummaryPath -Encoding UTF8
        Write-Log -Level SUCCESS -Message ("JSON summary written: {0}" -f $JsonSummaryPath)
    } catch { Write-Log -Level WARN -Message ("Failed to write JSON summary: {0}" -f $_.Exception.Message) }

    # -----------------------------------------------------------------------
    # Exit code
    # -----------------------------------------------------------------------
    $cleanStatuses = @('AlreadyCompliant', 'WouldRename', 'Renamed')
    $attention = @($results | Where-Object { $cleanStatuses -notcontains $_.Status })
    if ($attention.Count -gt 0) {
        Write-Log -Level WARN -Message ("Completed with {0} device(s) needing attention." -f $attention.Count)
        exit $ExitAttention
    }
    Write-Log -Level SUCCESS -Message 'Completed successfully; no devices need attention.'
    exit $ExitSuccess
}
catch {
    Write-Log -Level ERROR -Message ("FATAL: {0}" -f $_.Exception.Message)
    Write-Log -Level DEBUG -Message ($_.ScriptStackTrace)
    exit $ExitFatal
}
finally {
    $script:ClientSecretSecure = $null
    $script:TokenCache = $null
    [System.GC]::Collect()
}
