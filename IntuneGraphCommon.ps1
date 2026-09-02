<#
.SYNOPSIS
    Shared Microsoft Graph helpers for the Intune device tooling (auth, HTTP, retry, batching,
    logging). Dot-sourced by Rename-IntuneDevice.ps1 and Get-IntuneGroupDevices.ps1.

.DESCRIPTION
    Not meant to be run directly. Dot-source it, set $script:GraphLogFile (and optionally
    $script:GraphLogToStderr), then call Connect-Graph once, followed by Invoke-GraphApi /
    Get-GraphCollection / Invoke-GraphBatch.

    Configuration lives in $script:Graph* variables populated by Connect-Graph:
      GraphAuthMode, GraphTenantId, GraphClientId, GraphClientSecret (secure),
      GraphScopes, GraphMaxRetries, GraphBaseUri, GraphTokenCache.

.NOTES
    Compatibility: Windows PowerShell 5.1 and PowerShell 7+.
#>

# Public client "Microsoft Graph Command Line Tools" (present in every tenant); default client
# for interactive sign-in so no app registration is required to test.
$script:GraphPublicClientId = '14d82eec-204b-4c2f-b7e8-296a70dab67e'
if (-not (Get-Variable -Name GraphTokenCache -Scope Script -ErrorAction SilentlyContinue)) { $script:GraphTokenCache = $null }
if (-not (Get-Variable -Name GraphLogToStderr -Scope Script -ErrorAction SilentlyContinue)) { $script:GraphLogToStderr = $false }

# Ensure TLS 1.2 (older Windows PowerShell defaults can break the token endpoint).
try {
    if (([Net.ServicePointManager]::SecurityProtocol -band [Net.SecurityProtocolType]::Tls12) -ne [Net.SecurityProtocolType]::Tls12) {
        [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    }
} catch { }

function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string] $Message,
        [ValidateSet('INFO', 'WARN', 'ERROR', 'SUCCESS', 'DEBUG')][string] $Level = 'INFO'
    )
    $line = ('{0} [{1}] {2}' -f (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'), $Level, $Message)
    if ($script:GraphLogFile) {
        try { Add-Content -LiteralPath $script:GraphLogFile -Value $line -Encoding UTF8 } catch { }
    }
    if ($Level -eq 'DEBUG') { Write-Verbose $line; return }
    # When logs share stdout with machine-readable output, send them to stderr instead.
    if ($script:GraphLogToStderr) { [Console]::Error.WriteLine($line); return }
    switch ($Level) {
        'ERROR'   { Write-Host $line -ForegroundColor Red }
        'WARN'    { Write-Host $line -ForegroundColor Yellow }
        'SUCCESS' { Write-Host $line -ForegroundColor Green }
        default   { Write-Host $line }
    }
}

function Get-GraphErrorDetail {
    param([Parameter(Mandatory = $true)] $ErrorRecord)

    $info = [ordered]@{ StatusCode = $null; RetryAfter = $null; Message = $null; Raw = $null }
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
        $info.Raw = $body
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

function New-TokenCache {
    param([Parameter(Mandatory = $true)] $Response)
    $expiresIn = if ($Response.expires_in) { [int]$Response.expires_in } else { 3600 }
    return @{
        AccessToken  = $Response.access_token
        ExpiresOn    = (Get-Date).ToUniversalTime().AddSeconds($expiresIn - 120)  # refresh 2 min early
        RefreshToken = $Response.refresh_token
    }
}

function Get-ScopeString {
    param([string[]] $Scopes)
    return ((($Scopes | ForEach-Object { "https://graph.microsoft.com/$_" }) -join ' ') + ' offline_access')
}

function ConvertTo-Base64Url {
    param([Parameter(Mandatory = $true)][byte[]] $Bytes)
    return ([Convert]::ToBase64String($Bytes)).TrimEnd('=').Replace('+', '-').Replace('/', '_')
}

function ConvertFrom-HttpQueryString {
    param([string] $Query)
    $result = @{}
    if ([string]::IsNullOrEmpty($Query)) { return $result }
    foreach ($pair in ($Query -split '&')) {
        if (-not $pair) { continue }
        $kv = $pair -split '=', 2
        $key = [uri]::UnescapeDataString($kv[0])
        $val = if ($kv.Count -gt 1) { [uri]::UnescapeDataString($kv[1]) } else { '' }
        $result[$key] = $val
    }
    return $result
}

function Get-GraphTokenClientSecret {
    param(
        [Parameter(Mandatory = $true)][string] $TenantId,
        [Parameter(Mandatory = $true)][string] $ClientId,
        [Parameter(Mandatory = $true)][securestring] $ClientSecret
    )
    $bstr  = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($ClientSecret)
    try { $plain = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr) }
    finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }

    $body = @{ client_id = $ClientId; scope = 'https://graph.microsoft.com/.default'; client_secret = $plain; grant_type = 'client_credentials' }
    $uri  = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"

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
                if ($retryable -and $attempt -le $script:GraphMaxRetries) {
                    $delay = if ($det.RetryAfter) { $det.RetryAfter } else { [int][Math]::Min(60, [Math]::Pow(2, $attempt)) }
                    Write-Log -Level WARN -Message ("Token request failed (attempt {0}/{1}, status {2}); retrying in {3}s." -f $attempt, $script:GraphMaxRetries, $det.StatusCode, $delay)
                    Start-Sleep -Seconds $delay
                    continue
                }
                throw ("Failed to acquire Graph token: {0}" -f $det.Message)
            }
        }
    }
    finally { if (Get-Variable -Name plain -Scope Local -ErrorAction SilentlyContinue) { $plain = $null } }

    return (New-TokenCache -Response $resp)
}

function Get-GraphTokenInteractive {
    # OAuth authorization-code flow with PKCE and a loopback (http://localhost) redirect.
    param(
        [Parameter(Mandatory = $true)][string] $TenantId,
        [Parameter(Mandatory = $true)][string] $ClientId,
        [Parameter(Mandatory = $true)][string[]] $Scopes,
        [int] $TimeoutSeconds = 300
    )
    $rng = [System.Security.Cryptography.RandomNumberGenerator]::Create()
    try {
        $vbytes = New-Object byte[] 32; $rng.GetBytes($vbytes)
        $verifier = ConvertTo-Base64Url -Bytes $vbytes
        $sha = [System.Security.Cryptography.SHA256]::Create()
        try { $challenge = ConvertTo-Base64Url -Bytes ($sha.ComputeHash([System.Text.Encoding]::ASCII.GetBytes($verifier))) }
        finally { $sha.Dispose() }
        $sbytes = New-Object byte[] 16; $rng.GetBytes($sbytes)
        $state = ConvertTo-Base64Url -Bytes $sbytes
    }
    finally { $rng.Dispose() }

    $probe = New-Object System.Net.Sockets.TcpListener ([System.Net.IPAddress]::Loopback, 0)
    $probe.Start(); $port = ([System.Net.IPEndPoint]$probe.LocalEndpoint).Port; $probe.Stop()
    $redirectUri = "http://localhost:$port"
    $scopeStr    = Get-ScopeString -Scopes $Scopes

    $authUrl = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/authorize" +
        "?client_id=$([uri]::EscapeDataString($ClientId))" +
        "&response_type=code&redirect_uri=$([uri]::EscapeDataString($redirectUri))" +
        "&response_mode=query&scope=$([uri]::EscapeDataString($scopeStr))" +
        "&code_challenge=$challenge&code_challenge_method=S256&state=$state&prompt=select_account"

    $listener = New-Object System.Net.Sockets.TcpListener ([System.Net.IPAddress]::Loopback, $port)
    $listener.Start()
    try {
        Write-Log -Level INFO -Message 'A browser window will open for sign-in. If it does not, open the URL printed below.'
        [Console]::Error.WriteLine('')
        [Console]::Error.WriteLine("  $authUrl")
        [Console]::Error.WriteLine('')
        try { Start-Process $authUrl | Out-Null }
        catch { Write-Log -Level WARN -Message 'Could not launch a browser automatically; open the URL above manually.' }
        Write-Log -Level INFO -Message 'Waiting for interactive browser sign-in to complete...'

        $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
        while (-not $listener.Pending()) {
            if ((Get-Date) -ge $deadline) { throw 'Timed out waiting for interactive sign-in.' }
            Start-Sleep -Milliseconds 300
        }
        $client = $listener.AcceptTcpClient()
        try {
            $stream = $client.GetStream()
            $reader = New-Object System.IO.StreamReader($stream)
            $requestLine = $reader.ReadLine()
            $html = '<html><body style="font-family:Segoe UI,Arial,sans-serif"><h3>Sign-in complete.</h3><p>You can close this window and return to PowerShell.</p></body></html>'
            $response = "HTTP/1.1 200 OK`r`nContent-Type: text/html; charset=utf-8`r`nContent-Length: $([System.Text.Encoding]::UTF8.GetByteCount($html))`r`nConnection: close`r`n`r`n$html"
            $rbytes = [System.Text.Encoding]::UTF8.GetBytes($response)
            $stream.Write($rbytes, 0, $rbytes.Length); $stream.Flush()
        }
        finally { $client.Close() }
    }
    finally { $listener.Stop() }

    if ([string]::IsNullOrWhiteSpace($requestLine)) { throw 'No redirect was received from the browser.' }
    $pathPart = ($requestLine -split ' ')[1]
    $query = ''; $qi = $pathPart.IndexOf('?'); if ($qi -ge 0) { $query = $pathPart.Substring($qi + 1) }
    $qp = ConvertFrom-HttpQueryString -Query $query
    if ($qp['error']) { throw ("Interactive sign-in failed: {0} - {1}" -f $qp['error'], $qp['error_description']) }
    if (-not $qp['code']) { throw 'No authorization code was returned from sign-in.' }
    if ($qp['state'] -ne $state) { throw 'State mismatch during interactive sign-in; aborting for safety.' }

    $tokenUri = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
    $resp = Invoke-RestMethod -Method Post -Uri $tokenUri -ContentType 'application/x-www-form-urlencoded' `
        -Body @{ client_id = $ClientId; grant_type = 'authorization_code'; code = $qp['code']; redirect_uri = $redirectUri; code_verifier = $verifier; scope = $scopeStr } -ErrorAction Stop
    Write-Log -Level SUCCESS -Message 'Interactive sign-in completed.'
    return (New-TokenCache -Response $resp)
}

function Get-GraphTokenByRefresh {
    param(
        [Parameter(Mandatory = $true)][string] $TenantId,
        [Parameter(Mandatory = $true)][string] $ClientId,
        [Parameter(Mandatory = $true)][string] $RefreshToken,
        [Parameter(Mandatory = $true)][string[]] $Scopes
    )
    $uri = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
    $resp = Invoke-RestMethod -Method Post -Uri $uri -ContentType 'application/x-www-form-urlencoded' `
        -Body @{ grant_type = 'refresh_token'; client_id = $ClientId; refresh_token = $RefreshToken; scope = (Get-ScopeString -Scopes $Scopes) } -ErrorAction Stop
    return (New-TokenCache -Response $resp)
}

function Get-GraphAccessToken {
    param([switch] $ForceRefresh)
    if (-not $ForceRefresh -and $null -ne $script:GraphTokenCache -and (Get-Date).ToUniversalTime() -lt $script:GraphTokenCache.ExpiresOn) {
        return $script:GraphTokenCache.AccessToken
    }
    if ($script:GraphAuthMode -eq 'Interactive') {
        if ($null -ne $script:GraphTokenCache -and $script:GraphTokenCache.RefreshToken) {
            try {
                $script:GraphTokenCache = Get-GraphTokenByRefresh -TenantId $script:GraphTenantId -ClientId $script:GraphClientId -RefreshToken $script:GraphTokenCache.RefreshToken -Scopes $script:GraphScopes
                return $script:GraphTokenCache.AccessToken
            } catch { Write-Log -Level WARN -Message 'Silent token refresh failed; requesting interactive sign-in again.' }
        }
        $script:GraphTokenCache = Get-GraphTokenInteractive -TenantId $script:GraphTenantId -ClientId $script:GraphClientId -Scopes $script:GraphScopes
        return $script:GraphTokenCache.AccessToken
    }
    $script:GraphTokenCache = Get-GraphTokenClientSecret -TenantId $script:GraphTenantId -ClientId $script:GraphClientId -ClientSecret $script:GraphClientSecret
    return $script:GraphTokenCache.AccessToken
}

function Connect-Graph {
    <#
      Resolves credentials (parameters, then RENAMEDEVICE_* env vars), validates them, stores the
      Graph config in $script:Graph* variables, and acquires an access token up front (interactive
      sign-in happens here when AuthMode is Interactive).
    #>
    param(
        [ValidateSet('ClientSecret', 'Interactive')][string] $AuthMode = 'ClientSecret',
        [string] $TenantId,
        [string] $ClientId,
        [string] $ClientSecret,
        [Parameter(Mandatory = $true)][string[]] $Scopes,
        [int] $MaxRetries = 5,
        [string] $GraphBaseUri = 'https://graph.microsoft.com/beta'
    )

    if (-not $TenantId) { $TenantId = $env:RENAMEDEVICE_TENANT_ID }
    if (-not $ClientId) { $ClientId = $env:RENAMEDEVICE_CLIENT_ID }

    $missing = @()
    if ([string]::IsNullOrWhiteSpace($TenantId)) { $missing += 'TenantId (or RENAMEDEVICE_TENANT_ID)' }

    if ($AuthMode -eq 'ClientSecret') {
        if (-not $ClientSecret) { $ClientSecret = $env:RENAMEDEVICE_CLIENT_SECRET }
        if ([string]::IsNullOrWhiteSpace($ClientId))     { $missing += 'ClientId (or RENAMEDEVICE_CLIENT_ID)' }
        if ([string]::IsNullOrWhiteSpace($ClientSecret)) { $missing += 'ClientSecret (or RENAMEDEVICE_CLIENT_SECRET)' }
    }
    else {
        if ([string]::IsNullOrWhiteSpace($ClientId)) {
            $ClientId = $script:GraphPublicClientId
            Write-Log -Level INFO -Message 'Using the default Microsoft Graph public client for interactive sign-in.'
        }
    }

    if ($missing.Count -gt 0) {
        $msg = "Missing required credential(s): {0}" -f ($missing -join ', ')
        if ($AuthMode -eq 'ClientSecret') {
            $msg += '. Supply the secret via -ClientSecret or $env:RENAMEDEVICE_CLIENT_SECRET, OR use -AuthMode Interactive (no secret, omit -ClientId).'
        }
        throw $msg
    }

    $script:GraphAuthMode   = $AuthMode
    $script:GraphTenantId   = $TenantId
    $script:GraphClientId   = $ClientId
    $script:GraphScopes     = $Scopes
    $script:GraphMaxRetries = $MaxRetries
    $script:GraphBaseUri    = $GraphBaseUri.TrimEnd('/')
    $script:GraphTokenCache = $null
    if ($AuthMode -eq 'ClientSecret') {
        $script:GraphClientSecret = ConvertTo-SecureString -String $ClientSecret -AsPlainText -Force
    }

    Write-Log -Level INFO -Message ("Auth mode: {0}" -f $AuthMode)
    if ($AuthMode -eq 'Interactive') {
        Write-Log -Level INFO -Message 'Interactive sign-in required. A browser will open for you to authenticate.'
    }
    $null = Get-GraphAccessToken
    Write-Log -Level SUCCESS -Message 'Acquired Microsoft Graph access token.'
}

function Invoke-GraphApi {
    param(
        [Parameter(Mandatory = $true)][ValidateSet('GET', 'POST', 'PATCH', 'DELETE')][string] $Method,
        [Parameter(Mandatory = $true)][string] $Uri,
        $Body
    )
    $jsonBody = $null
    if ($null -ne $Body) { $jsonBody = ($Body | ConvertTo-Json -Depth 8 -Compress) }

    $attempt = 0
    $authRefreshed = $false
    while ($true) {
        $attempt++
        $token   = Get-GraphAccessToken
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
            if ($code -eq 404) { return @{ Success = $false; NotFound = $true; StatusCode = 404; Error = $det.Message } }
            if ($code -eq 401 -and -not $authRefreshed) {
                Write-Log -Level DEBUG -Message 'Received 401; refreshing access token and retrying.'
                $null = Get-GraphAccessToken -ForceRefresh
                $authRefreshed = $true
                continue
            }
            $retryable = ($null -eq $code) -or ($code -in 429, 500, 502, 503, 504)
            if ($retryable -and $attempt -le $script:GraphMaxRetries) {
                $delay = if ($det.RetryAfter) { $det.RetryAfter } else { [int][Math]::Min(60, [Math]::Pow(2, $attempt)) }
                Write-Log -Level WARN -Message ("Graph {0} failed (attempt {1}/{2}, status {3}); retrying in {4}s." -f $Method, $attempt, $script:GraphMaxRetries, $code, $delay)
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

function Invoke-GraphBatch {
    <#
      Runs many Graph requests via the $batch endpoint (max 20 per call), with throttling-aware
      retry of individual sub-requests. Returns a hashtable keyed by each request id:
        @{ status = [int]; body = <parsed object or $null> }
      Requests: array of @{ id = '<unique>'; method = 'GET'|'POST'|...; url = '/relative/graph/url'; body = <obj, optional> }
    #>
    param([Parameter(Mandatory = $true)][object[]] $Requests)

    $results  = @{}
    $batchUri = "$script:GraphBaseUri/`$batch"

    for ($i = 0; $i -lt $Requests.Count; $i += 20) {
        $end     = [Math]::Min($i + 19, $Requests.Count - 1)
        $chunk   = @($Requests[$i..$end])
        $pending = $chunk
        $attempt = 0
        while ($pending.Count -gt 0) {
            $attempt++
            $reqList = @($pending | ForEach-Object {
                $r = @{ id = [string]$_.id; method = $_.method; url = $_.url }
                if ($null -ne $_.body) { $r.body = $_.body; $r.headers = @{ 'Content-Type' = 'application/json' } }
                $r
            })
            $resp = Invoke-GraphApi -Method POST -Uri $batchUri -Body @{ requests = $reqList }
            if (-not $resp.Success) { throw ("Batch request failed ({0}): {1}" -f $resp.StatusCode, $resp.Error) }

            $retry = New-Object System.Collections.Generic.List[object]
            $maxRetryAfter = 0
            foreach ($sub in $resp.Data.responses) {
                $sc = [int]$sub.status
                if ($sc -in 429, 503, 504) {
                    try { if ($sub.headers -and $sub.headers.'Retry-After') { $ra = [int]$sub.headers.'Retry-After'; if ($ra -gt $maxRetryAfter) { $maxRetryAfter = $ra } } } catch { }
                    $orig = $chunk | Where-Object { [string]$_.id -eq [string]$sub.id } | Select-Object -First 1
                    if ($orig) { $retry.Add($orig) }
                }
                else { $results[[string]$sub.id] = @{ status = $sc; body = $sub.body } }
            }

            if ($retry.Count -gt 0 -and $attempt -le $script:GraphMaxRetries) {
                $delay = if ($maxRetryAfter -gt 0) { $maxRetryAfter } else { [int][Math]::Min(60, [Math]::Pow(2, $attempt)) }
                Write-Log -Level WARN -Message ("Batch throttled on {0} request(s); retrying in {1}s." -f $retry.Count, $delay)
                Start-Sleep -Seconds $delay
                $pending = $retry.ToArray()
            }
            else {
                foreach ($o in $retry) { $results[[string]$o.id] = @{ status = 429; body = $null } }
                $pending = @()
            }
        }
    }
    return $results
}

function Clear-GraphState {
    $script:GraphClientSecret = $null
    $script:GraphTokenCache = $null
    [System.GC]::Collect()
}
