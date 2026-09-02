#Requires -Version 5.1
<#
.SYNOPSIS
    Creates a ServiceNow incident via the Table API and returns its sys_id / number as JSON.

.DESCRIPTION
    Standalone (no shared library). Authenticates with Basic auth (an integration user), creates
    an incident, and writes { created, sys_id, number, url } to STDOUT. All logs go to STDERR so
    STDOUT stays clean for n8n / piping. Optionally builds the incident description from a
    Rename-IntuneDevice.ps1 JSON summary so the ticket lists exactly which devices need attention.

.PARAMETER Instance
    ServiceNow instance: short name ('dev12345') or full URL ('https://dev12345.service-now.com').
    Falls back to env var SNOW_INSTANCE.

.PARAMETER User / Password
    Basic-auth credentials (integration user). Fall back to env vars SNOW_USER / SNOW_PASSWORD.

.PARAMETER ShortDescription
    Incident short description (required).

.PARAMETER Description
    Incident description. If -InputJson is given and this is empty, a description is built from the
    rename summary (the devices whose status is not clean).

.PARAMETER InputJson
    Optional path to a Rename-IntuneDevice.ps1 JSON summary. Used to auto-build the description and,
    with -OnlyIfAttention, to skip creating a ticket when every device is clean.

.PARAMETER OnlyIfAttention
    With -InputJson: only create the incident if the summary contains devices needing attention
    (anything other than AlreadyCompliant / WouldRename / Renamed). Otherwise emit {created:false}.

.PARAMETER Urgency / Impact
    ServiceNow urgency / impact (1 = High, 2 = Medium, 3 = Low). Default 3.

.PARAMETER Category / AssignmentGroup / CallerId / ConfigurationItem
    Optional incident fields. AssignmentGroup / CallerId accept a sys_id or a name.

.PARAMETER CorrelationId
    Stable key stamped on the incident (correlation_id). If set, an existing ACTIVE incident with the
    same value is reused instead of creating a duplicate (unless -AllowDuplicate). Pair with the close
    step's -CorrelationId to run a stateless open/close loop across scheduled runs.

.PARAMETER AllowDuplicate
    Create a new incident even if an active one with the same -CorrelationId already exists.

.PARAMETER OutputPath
    Optional file to also write the { created, sys_id, number, url } result to (used by the close step).

.PARAMETER MaxRetries
    Retry attempts for transient failures (429/5xx/network). Default 3.

.EXAMPLE
    .\New-ServiceNowIncident.ps1 -ShortDescription 'Intune rename: devices need manual action' `
        -InputJson .\work\rename-result.json -OnlyIfAttention -OutputPath .\work\incident.json

.NOTES
    Exit codes: 0 = created (or skipped cleanly with -OnlyIfAttention); 2 = fatal (bad args/auth/API).
#>

[CmdletBinding()]
param(
    [string] $Instance,
    [string] $User,
    [string] $Password,

    [Parameter(Mandatory = $true)]
    [string] $ShortDescription,

    [string] $Description,
    [string] $InputJson,
    [switch] $OnlyIfAttention,

    [ValidateRange(1, 3)][int] $Urgency = 3,
    [ValidateRange(1, 3)][int] $Impact = 3,

    [string] $Category,
    [string] $AssignmentGroup,
    [string] $CallerId,
    [string] $ConfigurationItem,

    [string] $CorrelationId,
    [switch] $AllowDuplicate,

    [string] $OutputPath,
    [ValidateRange(0, 8)][int] $MaxRetries = 3
)

Set-StrictMode -Version 1.0
$ErrorActionPreference = 'Stop'
try { [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12 } catch { }

function Write-Log { param([string]$Message, [ValidateSet('INFO','WARN','ERROR','SUCCESS')][string]$Level='INFO')
    [Console]::Error.WriteLine(('{0} [{1}] {2}' -f (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'), $Level, $Message)) }

function Get-SnowBaseUri {
    param([string]$Instance)
    if ([string]::IsNullOrWhiteSpace($Instance)) { throw 'ServiceNow instance not provided (-Instance or SNOW_INSTANCE).' }
    if ($Instance -match '^https?://') { return $Instance.TrimEnd('/') }
    return "https://$Instance.service-now.com"
}

function Invoke-Snow {
    param([ValidateSet('GET','POST','PATCH')][string]$Method, [string]$Uri, $Body, [hashtable]$Headers, [int]$MaxRetries = 3)
    $json = if ($null -ne $Body) { $Body | ConvertTo-Json -Depth 6 } else { $null }
    $attempt = 0
    while ($true) {
        $attempt++
        try {
            if ($json) { return Invoke-RestMethod -Method $Method -Uri $Uri -Headers $Headers -Body $json -ContentType 'application/json' -ErrorAction Stop }
            return Invoke-RestMethod -Method $Method -Uri $Uri -Headers $Headers -ErrorAction Stop
        }
        catch {
            $code = $null; try { if ($_.Exception.Response) { $code = [int]$_.Exception.Response.StatusCode } } catch { }
            $retryable = ($null -eq $code) -or ($code -in 429, 500, 502, 503, 504)
            if ($retryable -and $attempt -le $MaxRetries) {
                $delay = [int][Math]::Min(30, [Math]::Pow(2, $attempt))
                Write-Log -Level WARN -Message ("ServiceNow {0} failed (attempt {1}/{2}, status {3}); retrying in {4}s." -f $Method, $attempt, $MaxRetries, $code, $delay)
                Start-Sleep -Seconds $delay; continue
            }
            $detail = $_.Exception.Message
            try { if ($_.ErrorDetails -and $_.ErrorDetails.Message) { $p = $_.ErrorDetails.Message | ConvertFrom-Json; if ($p.error -and $p.error.message) { $detail = $p.error.message } } } catch { }
            throw ("ServiceNow {0} {1} failed ({2}): {3}" -f $Method, $Uri, $code, $detail)
        }
    }
}

try {
    if (-not $Instance) { $Instance = $env:SNOW_INSTANCE }
    if (-not $User)     { $User     = $env:SNOW_USER }
    if (-not $Password) { $Password = $env:SNOW_PASSWORD }

    $missing = @()
    if ([string]::IsNullOrWhiteSpace($Instance)) { $missing += 'Instance (or SNOW_INSTANCE)' }
    if ([string]::IsNullOrWhiteSpace($User))     { $missing += 'User (or SNOW_USER)' }
    if ([string]::IsNullOrWhiteSpace($Password)) { $missing += 'Password (or SNOW_PASSWORD)' }
    if ($missing.Count -gt 0) { Write-Log -Level ERROR -Message ("Missing: {0}" -f ($missing -join ', ')); exit 2 }

    $baseUri = Get-SnowBaseUri -Instance $Instance
    $auth = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("{0}:{1}" -f $User, $Password)))
    $headers = @{ Authorization = "Basic $auth"; Accept = 'application/json' }

    # De-duplicate: if an active incident already carries this correlation_id, reuse it.
    if ($CorrelationId -and -not $AllowDuplicate) {
        $q = "$baseUri/api/now/table/incident?sysparm_query=active=true^correlation_id=$([uri]::EscapeDataString($CorrelationId))&sysparm_limit=1&sysparm_fields=sys_id,number"
        $existing = Invoke-Snow -Method GET -Uri $q -Headers $headers -MaxRetries $MaxRetries
        if ($existing.result -and @($existing.result).Count -gt 0) {
            $ex = @($existing.result)[0]
            Write-Log -Level INFO -Message ("Active incident {0} already exists for correlation '{1}'; reusing." -f $ex.number, $CorrelationId)
            $out = [ordered]@{ created = $false; reason = 'exists'; sys_id = $ex.sys_id; number = $ex.number; url = "$baseUri/nav_to.do?uri=incident.do?sys_id=$($ex.sys_id)" }
            if ($OutputPath) { try { $dir = Split-Path -Parent $OutputPath; if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }; ($out | ConvertTo-Json) | Set-Content -LiteralPath $OutputPath -Encoding UTF8 } catch { } }
            Write-Output ($out | ConvertTo-Json)
            exit 0
        }
    }

    # Optionally derive the description from a rename summary (and honor -OnlyIfAttention).
    $attentionCount = $null
    if ($InputJson) {
        if (-not (Test-Path -LiteralPath $InputJson)) { Write-Log -Level ERROR -Message ("InputJson not found: {0}" -f $InputJson); exit 2 }
        try { $summary = Get-Content -LiteralPath $InputJson -Raw | ConvertFrom-Json } catch { Write-Log -Level ERROR -Message ("InputJson is not valid JSON: {0}" -f $_.Exception.Message); exit 2 }
        $clean = @('AlreadyCompliant', 'WouldRename', 'Renamed')
        $attention = @()
        if ($summary.results) { $attention = @($summary.results | Where-Object { $clean -notcontains $_.Status }) }
        $attentionCount = $attention.Count

        if ($OnlyIfAttention -and $attentionCount -eq 0) {
            Write-Log -Level INFO -Message 'No devices need attention; skipping incident creation.'
            $skip = [ordered]@{ created = $false; reason = 'no-attention'; sys_id = $null; number = $null; url = $null }
            if ($OutputPath) { try { $dir = Split-Path -Parent $OutputPath; if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }; ($skip | ConvertTo-Json) | Set-Content -LiteralPath $OutputPath -Encoding UTF8 } catch { } }
            Write-Output ($skip | ConvertTo-Json)
            exit 0
        }
        if (-not $Description) {
            $lines = New-Object System.Collections.Generic.List[string]
            $lines.Add(("Automated Intune device rename summary (mode: {0})." -f $summary.mode))
            $lines.Add(("Devices needing attention: {0} of {1} processed." -f $attentionCount, $summary.totalProcessed))
            $lines.Add('')
            foreach ($r in $attention) {
                $lines.Add(("- {0} (serial {1}) [{2}]: {3}" -f $r.CurrentDeviceName, $r.SerialNumber, $r.Status, $r.Message))
            }
            $Description = ($lines -join "`n")
            if ($Description.Length -gt 3900) { $Description = $Description.Substring(0, 3900) + "`n... (truncated)" }
        }
    }

    $body = @{ short_description = $ShortDescription; urgency = "$Urgency"; impact = "$Impact" }
    if ($Description)       { $body.description      = $Description }
    if ($Category)          { $body.category         = $Category }
    if ($AssignmentGroup)   { $body.assignment_group = $AssignmentGroup }
    if ($CallerId)          { $body.caller_id        = $CallerId }
    if ($ConfigurationItem) { $body.cmdb_ci          = $ConfigurationItem }
    if ($CorrelationId)     { $body.correlation_id   = $CorrelationId }

    Write-Log -Level INFO -Message ("Creating incident on {0}..." -f $baseUri)
    $resp = Invoke-Snow -Method POST -Uri "$baseUri/api/now/table/incident" -Body $body -Headers $headers -MaxRetries $MaxRetries
    $sysId  = $resp.result.sys_id
    $number = $resp.result.number
    $url    = "$baseUri/nav_to.do?uri=incident.do?sys_id=$sysId"
    Write-Log -Level SUCCESS -Message ("Created incident {0} (sys_id {1})." -f $number, $sysId)

    $out = [ordered]@{ created = $true; sys_id = $sysId; number = $number; url = $url; attentionCount = $attentionCount }
    if ($OutputPath) {
        try {
            $dir = Split-Path -Parent $OutputPath
            if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
            ($out | ConvertTo-Json) | Set-Content -LiteralPath $OutputPath -Encoding UTF8
            Write-Log -Level SUCCESS -Message ("Wrote incident reference to {0}" -f $OutputPath)
        } catch { Write-Log -Level WARN -Message ("Failed to write OutputPath: {0}" -f $_.Exception.Message) }
    }
    Write-Output ($out | ConvertTo-Json)
    exit 0
}
catch {
    Write-Log -Level ERROR -Message ("FATAL: {0}" -f $_.Exception.Message)
    exit 2
}
