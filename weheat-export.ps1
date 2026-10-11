<#
.SYNOPSIS
    Export Weheat heat pump data to CSV. Plain PowerShell, nothing to install.
    Works in Windows PowerShell 5.1 and PowerShell 7+.

    Same API as https://github.com/wefabricate/wh-python (Keycloak password
    login + the third_party REST API), without Python.

.EXAMPLE
    # Credentials in weheat-credentials.txt next to the script:
    #   ClientId     = ...
    #   ClientSecret = ...
    #   Username     = me@mail.nl
    #   Password     = ...        (leave out to be asked)
    #   HeatPumpId   = ...        (optional)
    .\weheat-export.ps1 -Data Logs -Interval Hour -OutFile logs.csv

.EXAMPLE
    # Same, with the file somewhere else
    .\weheat-export.ps1 -CredentialFile C:\secrets\weheat.txt -Data Latest

.EXAMPLE
    # List your heat pumps (ids, names, models)
    .\weheat-export.ps1 -ClientId x -ClientSecret y -Username me@mail.nl -Data HeatPumps

.EXAMPLE
    # 15-minute logs for the last 7 days
    .\weheat-export.ps1 -ClientId x -ClientSecret y -Username me@mail.nl `
        -Data Logs -Interval FifteenMinute -Start (Get-Date).AddDays(-7) -OutFile logs.csv

.EXAMPLE
    # Daily energy for 2025, password from an env var
    $env:WEHEAT_PASSWORD = '...'
    .\weheat-export.ps1 -ClientId x -ClientSecret y -Username me@mail.nl `
        -Data Energy -Interval Day -Start 2025-01-01 -End 2026-01-01

.NOTES
    If PowerShell refuses to run the script, start it with:
    powershell -ExecutionPolicy Bypass -File .\weheat-export.ps1 ...
#>
[CmdletBinding()]
param(
    # Client ID and secret provided by Weheat
    [string]$ClientId,
    [string]$ClientSecret,
    # Login for the Weheat online portal; password is prompted if not given
    [string]$Username,
    [string]$Password,

    # Heat pump UUID; if empty, the first active heat pump on the account is used
    [string]$HeatPumpId,

    # Text file with "Key = value" lines (ClientId, ClientSecret, Username,
    # Password, HeatPumpId). Default: weheat-credentials.txt next to this script.
    # Command-line parameters win over the file, the file over WEHEAT_* env vars.
    [string]$CredentialFile,

    # What to export:
    #   Logs        aggregated logs per -Interval (/logs)
    #   RawLogs     every raw log record (/logs/raw)
    #   Latest      the latest log record (/logs/latest)
    #   Energy      energy per -Interval (/energy-logs)
    #   EnergyTotal lifetime energy totals (/energy-logs/total)
    #   HeatPumps   heat pumps on the account (/heat-pumps)
    [ValidateSet('Logs', 'RawLogs', 'Latest', 'Energy', 'EnergyTotal', 'HeatPumps')]
    [string]$Data = 'Logs',

    [ValidateSet('Minute', 'FiveMinute', 'FifteenMinute', 'Hour', 'Day', 'Week', 'Month', 'Year')]
    [string]$Interval = 'Hour',

    # Time range (local time unless you pass a UTC DateTime); default: last 24 hours
    [datetime]$Start = (Get-Date).AddDays(-1),
    [datetime]$End = (Get-Date),

    # CSV file to write; if empty, CSV goes to the console
    [string]$OutFile,
    # CSV delimiter, e.g. ';' for Dutch Excel
    [char]$Delimiter = ',',

    [string]$AuthUrl = 'https://auth.weheat.nl/auth/',
    [string]$Realm = 'Weheat',
    [string]$ApiUrl = 'https://api.weheat.nl/third_party'
)

$ErrorActionPreference = 'Stop'
# Windows PowerShell 5.1 doesn't enable TLS 1.2 by default
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

# --- Credentials --------------------------------------------------------------

$credKeys = [ordered]@{
    ClientId = 'WEHEAT_CLIENT_ID'; ClientSecret = 'WEHEAT_CLIENT_SECRET'
    Username = 'WEHEAT_USERNAME'; Password = 'WEHEAT_PASSWORD'; HeatPumpId = 'WEHEAT_HEAT_PUMP_ID'
}

$explicitFile = [bool]$CredentialFile
if (-not $CredentialFile) { $CredentialFile = Join-Path $PSScriptRoot 'weheat-credentials.txt' }
$fromFile = @{}
if (Test-Path -LiteralPath $CredentialFile) {
    Write-Host "Reading credentials from $CredentialFile" -ForegroundColor Cyan
    foreach ($line in Get-Content -LiteralPath $CredentialFile) {
        $line = $line.Trim()
        if (-not $line -or $line.StartsWith('#')) { continue }
        # Split on the first '=' or ':' so values may contain them
        if ($line -notmatch '^\s*([^=:]+?)\s*[=:]\s*(.*)$') { Write-Warning "Skipping line: $line"; continue }
        $key = $Matches[1] -replace '[\s_-]', ''
        $value = $Matches[2].Trim().Trim('"', "'")
        $match = @($credKeys.Keys) | Where-Object { $_ -eq $key }
        if ($match) { $fromFile[$match] = $value } else { Write-Warning "Unknown key '$($Matches[1])' in $CredentialFile" }
    }
} elseif ($explicitFile) {
    throw "Credential file not found: $CredentialFile"
}

foreach ($name in $credKeys.Keys) {
    if (Get-Variable $name -ValueOnly) { continue }
    $value = if ($fromFile[$name]) { $fromFile[$name] } else { [Environment]::GetEnvironmentVariable($credKeys[$name]) }
    Set-Variable $name $value
}

foreach ($name in 'ClientId', 'ClientSecret', 'Username') {
    if (-not (Get-Variable $name -ValueOnly)) { throw "Missing $name. Put it in $CredentialFile, pass -$name, or set $($credKeys[$name])." }
}
if (-not $Password) {
    $secure = Read-Host "Weheat password for $Username" -AsSecureString
    $Password = [Runtime.InteropServices.Marshal]::PtrToStringBSTR(
        [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure))
}

$oidc = $AuthUrl.TrimEnd('/') + "/realms/$Realm/protocol/openid-connect"
$ApiUrl = $ApiUrl.TrimEnd('/')

# --- Auth -------------------------------------------------------------------

function Get-WeheatToken {
    $body = @{
        grant_type    = 'password'
        client_id     = $ClientId
        client_secret = $ClientSecret
        username      = $Username
        password      = $Password
        scope         = 'openid'
    }
    Invoke-RestMethod -Method Post -Uri "$oidc/token" -Body $body -ContentType 'application/x-www-form-urlencoded'
}

function Close-WeheatSession($refreshToken) {
    # Same as the Python example: the access token stays valid until it expires
    try {
        Invoke-RestMethod -Method Post -Uri "$oidc/logout" -ContentType 'application/x-www-form-urlencoded' -Body @{
            client_id = $ClientId; client_secret = $ClientSecret; refresh_token = $refreshToken
        } | Out-Null
    } catch { Write-Verbose "Logout failed: $_" }
}

# --- API --------------------------------------------------------------------

function Invoke-Weheat([string]$Path, [hashtable]$Query = @{}) {
    $qs = ($Query.GetEnumerator() | ForEach-Object {
        '{0}={1}' -f $_.Key, [uri]::EscapeDataString([string]$_.Value)
    }) -join '&'
    $uri = "$ApiUrl$Path" + $(if ($qs) { "?$qs" } else { '' })
    Write-Verbose "GET $uri"
    # Fetch the raw body and parse it ourselves: Invoke-RestMethod in 5.1
    # wraps JSON arrays in a single object when piped.
    $resp = Invoke-WebRequest -Uri $uri -UseBasicParsing -Headers @{
        Authorization = "Bearer $script:AccessToken"
        Accept        = 'application/json'
    }
    $text = if ($resp.Content -is [byte[]]) { [Text.Encoding]::UTF8.GetString($resp.Content) } else { $resp.Content }
    if (-not $text) { return }
    $json = ConvertFrom-Json $text
    # Arrays come back as one object in 5.1; unroll them
    if ($json -is [array]) { $json | ForEach-Object { $_ } } else { $json }
}

function Format-UtcTime([datetime]$t) {
    $t.ToUniversalTime().ToString("yyyy-MM-dd'T'HH:mm:ss'Z'", [Globalization.CultureInfo]::InvariantCulture)
}

# Max time span per request for each interval (from the API docs)
$maxSpan = @{
    Minute = [timespan]::FromDays(2);   FiveMinute = [timespan]::FromDays(7)
    FifteenMinute = [timespan]::FromDays(28); Hour = [timespan]::FromDays(28)
    Day = [timespan]::FromDays(365);    Week = [timespan]::FromDays(730)
    Month = [timespan]::FromDays(1825); Year = [timespan]::FromDays(36500)
    Raw = [timespan]::FromDays(1)
}

# Split [Start, End) into chunks the API accepts and fetch each
function Get-Chunked([string]$Path, [timespan]$Span, [hashtable]$Extra = @{}) {
    $from = $Start
    while ($from -lt $End) {
        $to = $from + $Span
        if ($to -gt $End) { $to = $End }
        Write-Host ("  {0:yyyy-MM-dd HH:mm} -> {1:yyyy-MM-dd HH:mm}" -f $from, $to) -ForegroundColor DarkGray
        $q = @{ startTime = (Format-UtcTime $from); endTime = (Format-UtcTime $to) } + $Extra
        Invoke-Weheat $Path $q
        $from = $to
    }
}

function Get-ActiveHeatPumpId {
    # state=3 is "active", same filter the Python discovery uses
    $page = Invoke-Weheat '/api/v1/heat-pumps' @{ page = 1; pageSize = 1000; state = 3 }
    $pumps = @($page.data)
    if (-not $pumps) { throw 'No active heat pumps found on this account. Pass -HeatPumpId.' }
    if ($pumps.Count -gt 1) {
        Write-Warning ("Found {0} heat pumps, using '{1}'. Pass -HeatPumpId to pick another." -f $pumps.Count, $pumps[0].name)
    }
    $pumps[0].id
}

# --- CSV --------------------------------------------------------------------

# Flatten nested objects into "parent.child" columns, arrays into JSON text
function ConvertTo-FlatRow($obj, [string]$prefix = '', [System.Collections.Specialized.OrderedDictionary]$row = $null) {
    if ($null -eq $row) { $row = [ordered]@{} }
    foreach ($p in $obj.PSObject.Properties) {
        $key = if ($prefix) { "$prefix.$($p.Name)" } else { $p.Name }
        $v = $p.Value
        if ($v -is [System.Management.Automation.PSCustomObject]) {
            ConvertTo-FlatRow $v $key $row | Out-Null
        } elseif ($v -is [array]) {
            $row[$key] = ConvertTo-Json $v -Compress -Depth 10
        } elseif ($v -is [datetime]) {
            # 5.1 parses ISO dates into local DateTime; keep them as UTC ISO text
            $row[$key] = $v.ToUniversalTime().ToString('o')
        } else {
            $row[$key] = $v
        }
    }
    $row
}

function Export-Rows($items) {
    $rows = @($items | Where-Object { $null -ne $_ } | ForEach-Object { ConvertTo-FlatRow $_ })
    if (-not $rows) { Write-Warning 'The API returned no data for this request.'; return }

    # Union of all columns (Export-Csv only uses the first row's columns)
    $cols = New-Object System.Collections.Generic.List[string]
    foreach ($r in $rows) { foreach ($k in $r.Keys) { if (-not $cols.Contains($k)) { $cols.Add($k) } } }
    $objs = foreach ($r in $rows) {
        $o = [ordered]@{}
        foreach ($c in $cols) { $o[$c] = $r[$c] }
        [pscustomobject]$o
    }

    if ($OutFile) {
        $objs | Export-Csv -Path $OutFile -NoTypeInformation -Delimiter $Delimiter -Encoding UTF8
        Write-Host "Wrote $($rows.Count) rows, $($cols.Count) columns to $OutFile" -ForegroundColor Green
    } else {
        $objs | ConvertTo-Csv -NoTypeInformation -Delimiter $Delimiter
    }
}

# --- Main -------------------------------------------------------------------

Write-Host "Logging in as $Username..." -ForegroundColor Cyan
$token = Get-WeheatToken
$script:AccessToken = $token.access_token
Close-WeheatSession $token.refresh_token

if ($Data -ne 'HeatPumps' -and -not $HeatPumpId) {
    $HeatPumpId = Get-ActiveHeatPumpId
    Write-Host "Using heat pump $HeatPumpId" -ForegroundColor Cyan
}
if ($End -le $Start -and $Data -in 'Logs', 'RawLogs', 'Energy') { throw '-End must be after -Start.' }

Write-Host "Fetching $Data..." -ForegroundColor Cyan
$items = switch ($Data) {
    'HeatPumps'   { (Invoke-Weheat '/api/v1/heat-pumps' @{ page = 1; pageSize = 1000 }).data }
    'Latest'      { Invoke-Weheat "/api/v1/heat-pumps/$HeatPumpId/logs/latest" }
    'EnergyTotal' { Invoke-Weheat "/api/v1/energy-logs/$HeatPumpId/total" }
    'Logs'        { Get-Chunked "/api/v1/heat-pumps/$HeatPumpId/logs" $maxSpan[$Interval] @{ interval = $Interval } }
    'Energy'      { Get-Chunked "/api/v1/energy-logs/$HeatPumpId" $maxSpan[$Interval] @{ interval = $Interval } }
    'RawLogs'     { Get-Chunked "/api/v1/heat-pumps/$HeatPumpId/logs/raw" $maxSpan['Raw'] }
}

Export-Rows $items
