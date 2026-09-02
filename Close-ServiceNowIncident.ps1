#Requires -Version 5.1
<#
.SYNOPSIS
    Resolves or closes a ServiceNow incident via the Table API.

.DESCRIPTION
    Standalone (no shared library). Basic auth. Targets the incident by -SysId, by -Number, or by
    -InputPath (a JSON file written by New-ServiceNowIncident.ps1 containing sys_id/number). Sets the
    incident state to resolved (6) or closed (7) with a close code and close notes. Logs to STDERR;
    the JSON result goes to STDOUT.

.PARAMETER Instance / User / Password
    ServiceNow instance and Basic-auth credentials. Fall back to SNOW_INSTANCE / SNOW_USER / SNOW_PASSWORD.

.PARAMETER SysId
    Incident sys_id to close. One of -SysId / -Number / -InputPath is required.

.PARAMETER Number
    Incident number (e.g. INC0012345). Resolved to sys_id via a query.

.PARAMETER InputPath
    Path to a JSON file with a sys_id (or number) field — e.g. the -OutputPath from New-ServiceNowIncident.ps1.
    If it marks created=false, this script no-ops (nothing to close).

.PARAMETER CorrelationId
    Close every ACTIVE incident carrying this correlation_id (the key set by New-ServiceNowIncident.ps1).
    Enables a stateless loop: no sys_id needs to pass between runs. No-ops if none are open.

.PARAMETER State
    Target state: 'resolved' (6, default) or 'closed' (7).

.PARAMETER CloseCode
    ServiceNow close code. Default 'Solved (Permanently)'. Must match a value configured in your instance.

.PARAMETER CloseNotes
    Close notes. If -InputJson is given and this is empty, notes are built from the rename summary counts.

.PARAMETER InputJson
    Optional Rename-IntuneDevice.ps1 JSON summary, used to auto-build close notes.

.PARAMETER MaxRetries
    Retry attempts for transient failures. Default 3.

.EXAMPLE
    .\Close-ServiceNowIncident.ps1 -InputPath .\work\incident.json -InputJson .\work\rename-result.json

.NOTES
    Exit codes: 0 = closed (or nothing to close); 2 = fatal (bad args/auth/API/not found).
#>

[CmdletBinding()]
param(
    [string] $Instance,
    [string] $User,
    [string] $Password,

    [string] $SysId,
    [string] $Number,
    [string] $InputPath,
    [string] $CorrelationId,

    [ValidateSet('resolved', 'closed')]
    [string] $State = 'resolved',

    [string] $CloseCode = 'Solved (Permanently)',
    [string] $CloseNotes,
    [string] $InputJson,

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

    # Resolve target from -InputPath if provided.
    if ($InputPath) {
        if (-not (Test-Path -LiteralPath $InputPath)) { Write-Log -Level ERROR -Message ("InputPath not found: {0}" -f $InputPath); exit 2 }
        try { $ref = Get-Content -LiteralPath $InputPath -Raw | ConvertFrom-Json } catch { Write-Log -Level ERROR -Message ("InputPath is not valid JSON: {0}" -f $_.Exception.Message); exit 2 }
        if ($ref.PSObject.Properties.Name -contains 'created' -and -not $ref.created) {
            Write-Log -Level INFO -Message 'Incident reference marks created=false; nothing to close.'
            Write-Output (@{ closed = $false; reason = 'no-incident' } | ConvertTo-Json)
            exit 0
        }
        if (-not $SysId -and $ref.sys_id)  { $SysId  = $ref.sys_id }
        if (-not $Number -and $ref.number) { $Number = $ref.number }
    }

    $missing = @()
    if ([string]::IsNullOrWhiteSpace($Instance)) { $missing += 'Instance (or SNOW_INSTANCE)' }
    if ([string]::IsNullOrWhiteSpace($User))     { $missing += 'User (or SNOW_USER)' }
    if ([string]::IsNullOrWhiteSpace($Password)) { $missing += 'Password (or SNOW_PASSWORD)' }
    if (-not $SysId -and -not $Number -and -not $CorrelationId) { $missing += 'SysId, Number, CorrelationId, or -InputPath' }
    if ($missing.Count -gt 0) { Write-Log -Level ERROR -Message ("Missing: {0}" -f ($missing -join ', ')); exit 2 }

    $baseUri = Get-SnowBaseUri -Instance $Instance
    $auth = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("{0}:{1}" -f $User, $Password)))
    $headers = @{ Authorization = "Basic $auth"; Accept = 'application/json' }

    # Build the list of target sys_ids.
    $targets = New-Object System.Collections.Generic.List[string]
    if ($SysId) { $targets.Add($SysId) }
    elseif ($Number) {
        Write-Log -Level INFO -Message ("Looking up incident {0}..." -f $Number)
        $q = "$baseUri/api/now/table/incident?sysparm_query=number=$([uri]::EscapeDataString($Number))&sysparm_limit=1&sysparm_fields=sys_id"
        $look = Invoke-Snow -Method GET -Uri $q -Headers $headers -MaxRetries $MaxRetries
        if (-not $look.result -or @($look.result).Count -eq 0) { Write-Log -Level ERROR -Message ("Incident {0} not found." -f $Number); exit 2 }
        $targets.Add(@($look.result)[0].sys_id)
    }
    else {
        # By correlation_id: close every active incident that carries it.
        $q = "$baseUri/api/now/table/incident?sysparm_query=active=true^correlation_id=$([uri]::EscapeDataString($CorrelationId))&sysparm_fields=sys_id,number"
        $look = Invoke-Snow -Method GET -Uri $q -Headers $headers -MaxRetries $MaxRetries
        foreach ($r in @($look.result)) { $targets.Add($r.sys_id) }
        if ($targets.Count -eq 0) {
            Write-Log -Level INFO -Message ("No active incident with correlation '{0}'; nothing to close." -f $CorrelationId)
            Write-Output (@{ closed = $false; reason = 'no-incident'; count = 0 } | ConvertTo-Json)
            exit 0
        }
    }

    # Build close notes from a rename summary if requested.
    if (-not $CloseNotes -and $InputJson -and (Test-Path -LiteralPath $InputJson)) {
        try {
            $summary = Get-Content -LiteralPath $InputJson -Raw | ConvertFrom-Json
            $parts = @()
            if ($summary.countsByStatus) { foreach ($p in $summary.countsByStatus.PSObject.Properties) { $parts += ('{0}={1}' -f $p.Name, $p.Value) } }
            $CloseNotes = ("Automated Intune rename completed ({0}). {1}" -f $summary.mode, ($parts -join ', '))
        } catch { }
    }
    if (-not $CloseNotes) { $CloseNotes = 'Closed by automation.' }

    $stateValue = if ($State -eq 'closed') { '7' } else { '6' }
    $body = @{ state = $stateValue; close_code = $CloseCode; close_notes = $CloseNotes }

    $closed = New-Object System.Collections.Generic.List[object]
    foreach ($sid in $targets) {
        Write-Log -Level INFO -Message ("Setting incident {0} to state {1}..." -f $sid, $State)
        $resp = Invoke-Snow -Method PATCH -Uri "$baseUri/api/now/table/incident/$sid" -Body $body -Headers $headers -MaxRetries $MaxRetries
        $num = if ($resp.result -and $resp.result.number) { $resp.result.number } else { $Number }
        $closed.Add([ordered]@{ sys_id = $sid; number = $num })
        Write-Log -Level SUCCESS -Message ("Incident {0} set to {1}." -f $num, $State)
    }

    Write-Output (@{ closed = $true; state = $State; count = $closed.Count; incidents = $closed } | ConvertTo-Json -Depth 5)
    exit 0
}
catch {
    Write-Log -Level ERROR -Message ("FATAL: {0}" -f $_.Exception.Message)
    exit 2
}
