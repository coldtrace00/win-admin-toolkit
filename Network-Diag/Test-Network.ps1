#Requires -Version 5.1
<#
.SYNOPSIS
    Діагностика мережі: покроково знаходить, де саме проблема.

.DESCRIPTION
    Перевіряє по черзі: мережевий адаптер (кабель / Wi-Fi) -> IP-адресу (DHCP) -> шлюз (роутер)
    -> інтернет за IP -> DNS -> відкриття сайтів (HTTP/HTTPS), а також:
    пінг із втратами та джитером, трасування маршруту, якість для IP-телефонії,
    доступність серверів і портів, швидкість, проксі, файл hosts, точність годинника.

    Результат: висновок «де зламалося» + HTML-звіт на робочий стіл.
    Права адміністратора не потрібні.

.PARAMETER Quick
    Швидкий режим: без трасування і тесту швидкості.

.EXAMPLE
    .\Test-Network.ps1
.EXAMPLE
    .\Test-Network.ps1 -Quick
#>
[CmdletBinding()]
param(
    [switch]$Quick,
    [switch]$NoOpen
)

# =====================================================================
#  НАЛАШТУВАННЯ — змінюйте тут
# =====================================================================

# Ваші сервери та сервіси для перевірки (TCP-порт). Розкоментуйте й змініть під себе.
$Targets = @(
    # @{ Name = 'Файловий сервер';        Host = 'fileserver';       Port = 445  }
    # @{ Name = '1С / база даних';        Host = '192.168.1.10';     Port = 1541 }
    # @{ Name = 'Мережевий принтер';      Host = '192.168.1.50';     Port = 9100 }
    # @{ Name = 'АТС / SIP-сервер (TCP)'; Host = 'pbx.example.com';  Port = 5060 }
    @{ Name = 'Google (HTTPS)';          Host = 'www.google.com';   Port = 443 }
    @{ Name = 'Microsoft 365 (HTTPS)';   Host = 'outlook.office.com'; Port = 443 }
)
$InternetHosts       = @('8.8.8.8', '1.1.1.1')      # пінг «в інтернет» за IP
$DnsTestName         = 'www.google.com'             # що резолвимо для перевірки DNS
$TraceTarget         = '8.8.8.8'                    # куди робимо трасування
$SpeedTestUrl        = 'https://speed.cloudflare.com/__down?bytes=50000000'
$SpeedTestSeconds    = 8                            # скільки секунд міряти швидкість

# =====================================================================
$ErrorActionPreference = 'Continue'
$ProgressPreference    = 'SilentlyContinue'
try { [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12 } catch { }
try { $Host.UI.RawUI.WindowTitle = 'Діагностика мережі' } catch { }

# =====================================================================
#  СПІЛЬНІ ФУНКЦІЇ HTML-ЗВІТУ
# =====================================================================
$Desktop = [Environment]::GetFolderPath('Desktop')
$Stamp   = Get-Date -Format 'yyyy-MM-dd_HH-mm'
$Utf8Bom = New-Object System.Text.UTF8Encoding $true
$script:TblCounter = 0

function Enc([object]$v) {
    if ($null -eq $v) { return '' }
    [System.Net.WebUtility]::HtmlEncode([string]$v)
}

function New-Tag([string]$Status, [string]$Text) {
    $cls = 'info'
    switch ($Status) { 'OK' { $cls = 'ok' } 'WARN' { $cls = 'warn' } 'FAIL' { $cls = 'bad' } }
    if (-not $Text) {
        switch ($Status) {
            'OK'    { $Text = 'OK' }
            'WARN'  { $Text = 'Увага' }
            'FAIL'  { $Text = 'Помилка' }
            default { $Text = 'Інфо' }
        }
    }
    "<span class=""tag $cls"">$(Enc $Text)</span>"
}

function ConvertTo-HtmlTable {
    param(
        [object[]]$Rows,
        [string[]]$RawColumns = @(),
        [string]$Name = 'export',
        [switch]$NoToolbar
    )
    $Rows = @($Rows | Where-Object { $null -ne $_ })
    if ($Rows.Count -eq 0) { return '<p class="muted">Немає даних</p>' }
    $script:TblCounter++
    $id = "t$($script:TblCounter)"
    $cols = @($Rows[0].PSObject.Properties.Name)
    $sb = New-Object System.Text.StringBuilder
    if (-not $NoToolbar) {
        [void]$sb.Append('<div class="toolbar">')
        if ($Rows.Count -gt 8) {
            [void]$sb.Append("<input class=""search"" type=""search"" placeholder=""Пошук..."" oninput=""filterTable(this,'$id')"">")
            [void]$sb.Append("<span class=""muted"">Показано: <b id=""$id-count"">$($Rows.Count)</b> з $($Rows.Count)</span>")
        }
        [void]$sb.Append("<button class=""btn"" type=""button"" data-table=""$id"" data-name=""$(Enc $Name)"" onclick=""exportCsv(this)"">Експорт у CSV</button></div>")
    }
    [void]$sb.Append("<div class=""tbl""><table class=""sortable"" id=""$id""><thead><tr>")
    foreach ($c in $cols) { [void]$sb.Append("<th>$(Enc $c)</th>") }
    [void]$sb.Append('</tr></thead><tbody>')
    foreach ($r in $Rows) {
        [void]$sb.Append('<tr>')
        foreach ($c in $cols) {
            if ($RawColumns -contains $c) { [void]$sb.Append("<td>$($r.$c)</td>") }
            else                          { [void]$sb.Append("<td>$(Enc $r.$c)</td>") }
        }
        [void]$sb.Append('</tr>')
    }
    [void]$sb.Append('</tbody></table></div>')
    $sb.ToString()
}

function ConvertTo-KvTable([System.Collections.IDictionary]$Data) {
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('<div class="tbl"><table class="kv"><tbody>')
    foreach ($k in $Data.Keys) {
        [void]$sb.Append("<tr><th>$(Enc $k)</th><td>$(Enc $Data[$k])</td></tr>")
    }
    [void]$sb.Append('</tbody></table></div>')
    $sb.ToString()
}

function New-Cards([System.Collections.IDictionary]$Data) {
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('<div class="cards">')
    foreach ($k in $Data.Keys) {
        [void]$sb.Append("<div class=""card""><div class=""k"">$(Enc $k)</div><div class=""v"">$(Enc $Data[$k])</div></div>")
    }
    [void]$sb.Append('</div>')
    $sb.ToString()
}

function New-Section([string]$Title, [string]$Html, [string]$Description, [string]$Anchor) {
    $a = ''
    if ($Anchor) { $a = " id=""$Anchor""" }
    $d = ''
    if ($Description) { $d = "<p class=""muted"">$(Enc $Description)</p>" }
    "<section$a><h2>$(Enc $Title)</h2>$d$Html</section>"
}

function New-HtmlPage([string]$Title, [string]$Subtitle, [string]$Body, [string]$Tool) {
    $css = @'
:root{--bg:#f4f5f7;--card:#fff;--text:#1d2330;--muted:#6b7280;--line:#e5e7eb;--head:#f9fafb;--accent:#2563eb;--ok:#16a34a;--warn:#d97706;--bad:#dc2626;--info:#64748b}
@media (prefers-color-scheme:dark){:root{--bg:#0f1115;--card:#171a21;--text:#e5e7eb;--muted:#9ca3af;--line:#2a2f3a;--head:#1d212a;--accent:#60a5fa}}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--text);font:14px/1.5 "Segoe UI",system-ui,-apple-system,sans-serif}
.wrap{max-width:1200px;margin:0 auto;padding:24px 16px 40px}
header h1{margin:0;font-size:26px}
header p{margin:4px 0 0;color:var(--muted)}
.cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:12px;margin:20px 0 4px}
.card{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:12px 16px}
.card .k{color:var(--muted);font-size:12px;text-transform:uppercase;letter-spacing:.04em}
.card .v{font-size:16px;font-weight:600;margin-top:2px;word-break:break-word}
section{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:16px 18px;margin-top:16px}
section h2{margin:0 0 8px;font-size:17px}
section h3{margin:16px 0 8px;font-size:14px;color:var(--muted)}
.tbl{overflow-x:auto}
table{width:100%;border-collapse:collapse}
th,td{text-align:left;padding:7px 10px;border-bottom:1px solid var(--line);vertical-align:top}
table.sortable thead th{background:var(--head);font-weight:600;white-space:nowrap;cursor:pointer;user-select:none}
table.sortable thead th:hover{color:var(--accent)}
table.kv th{width:35%;color:var(--muted);font-weight:500}
tbody tr:hover{background:var(--head)}
.tag{display:inline-block;padding:0 8px;border-radius:999px;font-size:12px;font-weight:600;color:#fff;background:var(--info);white-space:nowrap}
.tag.ok{background:var(--ok)}.tag.warn{background:var(--warn)}.tag.bad{background:var(--bad)}
.toolbar{display:flex;gap:10px;align-items:center;flex-wrap:wrap;margin:6px 0 10px}
.search{flex:1 1 220px;max-width:340px;padding:7px 10px;border:1px solid var(--line);border-radius:8px;background:var(--bg);color:var(--text);font:inherit}
.btn{padding:6px 12px;border:1px solid var(--line);border-radius:8px;background:var(--head);color:var(--text);font:inherit;cursor:pointer;margin-left:auto}
.btn:hover{border-color:var(--accent);color:var(--accent)}
.verdict{margin-top:20px;padding:16px 20px;border-radius:10px;border:1px solid var(--line);border-left:6px solid var(--info);background:var(--card)}
.verdict.ok{border-left-color:var(--ok)}.verdict.warn{border-left-color:var(--warn)}.verdict.bad{border-left-color:var(--bad)}
.verdict h2{margin:0;font-size:20px}.verdict p{margin:6px 0 0}
.toc{display:flex;flex-wrap:wrap;gap:6px 14px;margin-top:14px}
.muted{color:var(--muted)}
a{color:var(--accent)}
code{background:var(--head);border:1px solid var(--line);border-radius:4px;padding:0 4px;font-family:Consolas,monospace}
details summary{cursor:pointer;font-weight:600}
footer{margin-top:24px;color:var(--muted);font-size:12px;text-align:center}
@media print{body{background:#fff;color:#000}.toolbar{display:none}section,.card{break-inside:avoid;border-color:#ccc}}
'@
    $js = @'
function filterTable(input,id){var q=input.value.toLowerCase(),n=0;
document.querySelectorAll('#'+id+' tbody tr').forEach(function(r){var m=r.textContent.toLowerCase().indexOf(q)>-1;r.style.display=m?'':'none';if(m)n++;});
var c=document.getElementById(id+'-count');if(c)c.textContent=n;}
function exportCsv(btn){var t=document.getElementById(btn.getAttribute('data-table'));if(!t)return;
var esc=function(s){s=s.replace(/\s+/g,' ').trim();return /[";]/.test(s)?'"'+s.replace(/"/g,'""')+'"':s;};
var out=[Array.prototype.map.call(t.tHead.rows[0].cells,function(c){return esc(c.textContent);}).join(';')];
Array.prototype.forEach.call(t.tBodies[0].rows,function(r){if(r.style.display==='none')return;
out.push(Array.prototype.map.call(r.cells,function(c){return esc(c.textContent);}).join(';'));});
var blob=new Blob(['﻿'+out.join('\r\n')],{type:'text/csv;charset=utf-8'});
var a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download=(btn.getAttribute('data-name')||'export')+'.csv';
document.body.appendChild(a);a.click();setTimeout(function(){URL.revokeObjectURL(a.href);a.remove();},500);}
document.addEventListener('click',function(e){var th=e.target.closest('table.sortable thead th');if(!th)return;
var table=th.closest('table'),tb=table.tBodies[0],i=Array.prototype.indexOf.call(th.parentNode.children,th);
var asc=th.getAttribute('data-dir')!=='asc';th.parentNode.querySelectorAll('th').forEach(function(x){x.removeAttribute('data-dir');});th.setAttribute('data-dir',asc?'asc':'desc');
var rows=Array.prototype.slice.call(tb.rows);rows.sort(function(a,b){var x=a.cells[i].textContent.trim(),y=b.cells[i].textContent.trim();
var nx=parseFloat(x.replace(/\s/g,'').replace(',','.')),ny=parseFloat(y.replace(/\s/g,'').replace(',','.'));
var r=(!isNaN(nx)&&!isNaN(ny))?nx-ny:x.localeCompare(y,'uk',{numeric:true});return asc?r:-r;});
rows.forEach(function(r){tb.appendChild(r);});});
'@
    $gen = Get-Date -Format 'dd.MM.yyyy HH:mm'
    @"
<!DOCTYPE html>
<html lang="uk">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>$(Enc $Title)</title>
<style>$css</style>
</head>
<body>
<div class="wrap">
<header><h1>$(Enc $Title)</h1><p>$(Enc $Subtitle)</p></header>
$Body
<footer>Згенеровано $gen на $(Enc $env:COMPUTERNAME) · $(Enc $Tool)</footer>
</div>
<script>$js</script>
</body>
</html>
"@
}

function Save-Html([string]$Html, [string]$FileName) {
    $path = Join-Path $Desktop $FileName
    [System.IO.File]::WriteAllText($path, $Html, $Utf8Bom)
    $path
}

$Checks = New-Object System.Collections.Generic.List[object]
function Add-Check([string]$Status, [string]$Name, [string]$Detail) {
    $Checks.Add([pscustomobject]@{ Status = $Status; Name = $Name; Detail = $Detail })
    $color = @{ OK = 'Green'; WARN = 'Yellow'; FAIL = 'Red'; INFO = 'Cyan' }[$Status]
    Write-Host ('[{0,-4}] ' -f $Status) -ForegroundColor $color -NoNewline
    Write-Host $Name -NoNewline
    if ($Detail) { Write-Host "  $Detail" -ForegroundColor DarkGray } else { Write-Host '' }
}

function Test-PingStats([string]$Target, [int]$Count = 10, [int]$Timeout = 1000) {
    $p = New-Object System.Net.NetworkInformation.Ping
    $times = New-Object System.Collections.Generic.List[double]
    for ($i = 0; $i -lt $Count; $i++) {
        try {
            $r = $p.Send($Target, $Timeout)
            if ($r.Status -eq 'Success') { $times.Add([double]$r.RoundtripTime) }
        } catch { }
        if ($i -lt $Count - 1) { Start-Sleep -Milliseconds 150 }
    }
    $p.Dispose()
    $avg = $null; $min = $null; $max = $null; $jit = $null
    if ($times.Count -gt 0) {
        $m = $times | Measure-Object -Average -Minimum -Maximum
        $avg = [math]::Round($m.Average, 1); $min = $m.Minimum; $max = $m.Maximum
        if ($times.Count -gt 1) {
            $d = for ($j = 1; $j -lt $times.Count; $j++) { [math]::Abs($times[$j] - $times[$j - 1]) }
            $jit = [math]::Round(($d | Measure-Object -Average).Average, 1)
        } else { $jit = 0 }
    }
    [pscustomobject]@{
        Target = $Target; Sent = $Count; Received = $times.Count
        Loss = [math]::Round(($Count - $times.Count) / $Count * 100)
        Avg = $avg; Min = $min; Max = $max; Jitter = $jit
    }
}

function Test-TcpPort([string]$HostName, [int]$Port, [int]$Timeout = 2500) {
    $c = New-Object System.Net.Sockets.TcpClient
    $sw = [Diagnostics.Stopwatch]::StartNew()
    try {
        $t = $c.ConnectAsync($HostName, $Port)
        if ($t.Wait($Timeout) -and $c.Connected) {
            return [pscustomobject]@{ Ok = $true; Ms = $sw.ElapsedMilliseconds; Error = '' }
        }
        return [pscustomobject]@{ Ok = $false; Ms = $null; Error = 'Немає відповіді (таймаут)' }
    } catch {
        $e = $_.Exception
        while ($e.InnerException) { $e = $e.InnerException }
        return [pscustomobject]@{ Ok = $false; Ms = $null; Error = $e.Message }
    } finally {
        $c.Close()
    }
}

function Test-Dns([string]$Name, [string]$Server) {
    $sw = [Diagnostics.Stopwatch]::StartNew()
    try {
        $p = @{ Name = $Name; Type = 'A'; DnsOnly = $true; NoHostsFile = $true; QuickTimeout = $true; ErrorAction = 'Stop' }
        if ($Server) { $p.Server = $Server }
        $r = @(Resolve-DnsName @p | Where-Object { $_.Type -eq 'A' })
        $sw.Stop()
        if ($r.Count -eq 0) { return [pscustomobject]@{ Ok = $false; Ms = $null; Result = 'Порожня відповідь' } }
        [pscustomobject]@{ Ok = $true; Ms = $sw.ElapsedMilliseconds; Result = (($r | Select-Object -First 3).IPAddress -join ', ') }
    } catch {
        [pscustomobject]@{ Ok = $false; Ms = $null; Result = $_.Exception.Message }
    }
}

function Format-Ms($v) { if ($null -eq $v) { '-' } else { "$v мс" } }

$cs = Get-CimInstance Win32_ComputerSystem
Write-Host ''
Write-Host '========================================================' -ForegroundColor Cyan
Write-Host "  ДІАГНОСТИКА МЕРЕЖІ: $env:COMPUTERNAME" -ForegroundColor Cyan
Write-Host '========================================================' -ForegroundColor Cyan

# =====================================================================
#  1. АДАПТЕР
# =====================================================================
$adapters = @(Get-NetAdapter -ErrorAction SilentlyContinue | Where-Object { $_.Status -ne 'Not Present' })
$up       = @($adapters | Where-Object { $_.Status -eq 'Up' })
$upPhys   = @($up | Where-Object { $_.HardwareInterface })
$wifiUp   = @($upPhys | Where-Object { $_.NdisPhysicalMedium -eq 9 -or $_.InterfaceDescription -match 'Wi-?Fi|Wireless|WLAN|802\.11' })

$adapterRows = foreach ($a in ($adapters | Sort-Object { $_.Status -ne 'Up' }, Name)) {
    $st = 'INFO'; $stText = [string]$a.Status
    if ($a.Status -eq 'Up') { $st = 'OK'; $stText = 'Підключено' }
    elseif ($a.Status -eq 'Disconnected') { $stText = 'Не підключено' }
    elseif ($a.Status -eq 'Disabled') { $stText = 'Вимкнено' }
    [pscustomobject]@{
        'Стан'      = (New-Tag $st $stText)
        'Адаптер'   = $a.Name
        'Опис'      = $a.InterfaceDescription
        'Швидкість' = $a.LinkSpeed
        'MAC'       = $a.MacAddress
        'Фізичний'  = $(if ($a.HardwareInterface) { 'Так' } else { 'Ні (віртуальний/VPN)' })
    }
}

$wifiInfo = $null
if ($wifiUp.Count -gt 0) {
    $wl = @(netsh wlan show interfaces 2>$null)
    $ssid = ($wl | Where-Object { $_ -match '^\s*SSID\s*:' } | Select-Object -First 1) -replace '^\s*SSID\s*:\s*', ''
    $sig  = ($wl | Select-String -Pattern ':\s*(\d{1,3})\s*%' | Select-Object -First 1)
    $sigVal = $null
    if ($sig) { $sigVal = [int]$sig.Matches[0].Groups[1].Value }
    $wifiInfo = [pscustomobject]@{ SSID = $ssid.Trim(); Signal = $sigVal }
}

if ($upPhys.Count -eq 0) {
    Add-Check 'FAIL' 'Мережевий адаптер' 'Немає підключення: кабель не вставлений, Wi-Fi не підключено або адаптер вимкнено'
} else {
    foreach ($a in $upPhys) {
        $isWifi = $wifiUp -contains $a
        if ($isWifi -and $wifiInfo) {
            $st = 'OK'; $hint = ''
            if ($wifiInfo.Signal -ne $null -and $wifiInfo.Signal -lt 50) { $st = 'WARN'; $hint = ' - слабкий сигнал, підійдіть ближче до роутера' }
            Add-Check $st "Wi-Fi: $($a.Name)" "мережа «$($wifiInfo.SSID)», сигнал $($wifiInfo.Signal)%, $($a.LinkSpeed)$hint"
        } else {
            $st = 'OK'; $hint = ''
            if ($a.Speed -and $a.Speed -le 10000000) { $st = 'WARN'; $hint = ' - дуже низька швидкість лінку, перевірте кабель/розетку' }
            elseif ($a.Speed -and $a.Speed -le 100000000) { $hint = ' (100 Мбіт: старий комутатор або пошкоджений кабель)' }
            Add-Check $st "Кабель: $($a.Name)" "$($a.LinkSpeed)$hint"
        }
    }
}

# =====================================================================
#  2. IP-АДРЕСА
# =====================================================================
$defRoute = Get-NetRoute -DestinationPrefix '0.0.0.0/0' -ErrorAction SilentlyContinue |
            Sort-Object { [int]$_.RouteMetric + [int]$_.InterfaceMetric } | Select-Object -First 1
$gateway  = $null
$mainIdx  = $null
if ($defRoute) { $gateway = $defRoute.NextHop; $mainIdx = $defRoute.InterfaceIndex }

$ipConfigs = @(Get-NetIPConfiguration -ErrorAction SilentlyContinue | Where-Object { $_.NetAdapter.Status -eq 'Up' })
if (-not $mainIdx -and $ipConfigs.Count) { $mainIdx = $ipConfigs[0].InterfaceIndex }

$ipRows = foreach ($c in $ipConfigs) {
    $dhcp = (Get-NetIPInterface -InterfaceIndex $c.InterfaceIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue).Dhcp
    [pscustomobject]@{
        'Інтерфейс' = $c.InterfaceAlias + $(if ($c.InterfaceIndex -eq $mainIdx) { ' (основний)' } else { '' })
        'IPv4'      = (@($c.IPv4Address) | ForEach-Object { "$($_.IPAddress)/$($_.PrefixLength)" }) -join ', '
        'Шлюз'      = (@($c.IPv4DefaultGateway) | ForEach-Object { $_.NextHop }) -join ', '
        'DNS'       = (@($c.DNSServer) | Where-Object { $_.AddressFamily -eq 2 } | ForEach-Object { $_.ServerAddresses }) -join ', '
        'DHCP'      = $(if ($dhcp -eq 'Enabled') { 'Так' } else { 'Ні (статична)' })
    }
}

$mainIp = $null
$mainCfg = $ipConfigs | Where-Object { $_.InterfaceIndex -eq $mainIdx } | Select-Object -First 1
if ($mainCfg) { $mainIp = @($mainCfg.IPv4Address)[0] }
$dnsServers = @()
if ($mainIdx) {
    $dnsServers = @((Get-DnsClientServerAddress -InterfaceIndex $mainIdx -AddressFamily IPv4 -ErrorAction SilentlyContinue).ServerAddresses)
}

$ipOk = $false
if (-not $mainIp) {
    Add-Check 'FAIL' 'IP-адреса' 'Немає IPv4-адреси'
} elseif ($mainIp.IPAddress -like '169.254.*') {
    Add-Check 'FAIL' 'IP-адреса' "$($mainIp.IPAddress) - ПК не отримав адресу від DHCP (роутер/DHCP-сервер не відповідає)"
} else {
    $ipOk = $true
    Add-Check 'OK' 'IP-адреса' "$($mainIp.IPAddress)/$($mainIp.PrefixLength) на «$($mainCfg.InterfaceAlias)»"
}

# =====================================================================
#  3. ШЛЮЗ
# =====================================================================
$gwStats = $null; $gwOk = $false; $gwArp = $false
if (-not $gateway) {
    Add-Check 'FAIL' 'Шлюз (роутер)' 'Шлюз не налаштований'
} else {
    $gwStats = Test-PingStats $gateway 10
    $nb = Get-NetNeighbor -IPAddress $gateway -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($nb -and $nb.State -in 'Reachable', 'Stale', 'Delay', 'Probe', 'Permanent') { $gwArp = $true }
    if ($gwStats.Received -gt 0) {
        $gwOk = $true
        $st = 'OK'; $hint = ''
        if ($gwStats.Loss -gt 0) { $st = 'WARN'; $hint = ' - є втрати в локальній мережі (кабель, Wi-Fi, комутатор)' }
        elseif ($gwStats.Avg -gt 30) { $st = 'WARN'; $hint = ' - висока затримка до роутера' }
        Add-Check $st 'Шлюз (роутер)' "$gateway : $($gwStats.Avg) мс, втрати $($gwStats.Loss)%$hint"
    } elseif ($gwArp) {
        Add-Check 'WARN' 'Шлюз (роутер)' "$gateway не відповідає на ping, але присутній у мережі (можливо, ping заблоковано)"
    } else {
        Add-Check 'FAIL' 'Шлюз (роутер)' "$gateway недоступний"
    }
}

# =====================================================================
#  4. ІНТЕРНЕТ ЗА IP
# =====================================================================
$inetStats = @()
foreach ($h in $InternetHosts) { $inetStats += Test-PingStats $h 10 }
$inetOk = @($inetStats | Where-Object { $_.Received -gt 0 }).Count -gt 0
$tcpInet = $null
if (-not $inetOk) {
    $tcpInet = Test-TcpPort '1.1.1.1' 443 3000
    if ($tcpInet.Ok) { $inetOk = $true }
}
$best = $inetStats | Where-Object { $_.Received -gt 0 } | Sort-Object Avg | Select-Object -First 1
if ($inetOk -and $best) {
    $st = 'OK'; $hint = ''
    if ($best.Loss -gt 2) { $st = 'WARN'; $hint = ' - втрати пакетів, нестабільний інтернет' }
    elseif ($best.Avg -gt 100) { $st = 'WARN'; $hint = ' - висока затримка' }
    Add-Check $st 'Інтернет (за IP)' "$($best.Target): $($best.Avg) мс, втрати $($best.Loss)%, джитер $($best.Jitter) мс$hint"
} elseif ($inetOk) {
    Add-Check 'WARN' 'Інтернет (за IP)' 'Ping заблоковано, але TCP 443 до 1.1.1.1 працює'
} else {
    Add-Check 'FAIL' 'Інтернет (за IP)' "Немає відповіді від $($InternetHosts -join ', ')"
}

# =====================================================================
#  5. DNS
# =====================================================================
$dnsRows = @()
$sysDns = Test-Dns $DnsTestName $null
$dnsRows += [pscustomobject]@{ 'DNS-сервер' = 'Системний (як у браузера)'; 'Результат' = (New-Tag $(if ($sysDns.Ok) { 'OK' } else { 'FAIL' })); 'Час' = (Format-Ms $sysDns.Ms); 'Відповідь' = $sysDns.Result }
$cfgOk = 0
foreach ($s in $dnsServers) {
    $r = Test-Dns $DnsTestName $s
    if ($r.Ok) { $cfgOk++ }
    $dnsRows += [pscustomobject]@{ 'DNS-сервер' = "$s (з налаштувань)"; 'Результат' = (New-Tag $(if ($r.Ok) { 'OK' } else { 'FAIL' })); 'Час' = (Format-Ms $r.Ms); 'Відповідь' = $r.Result }
}
$pubDns = Test-Dns $DnsTestName '8.8.8.8'
$dnsRows += [pscustomobject]@{ 'DNS-сервер' = '8.8.8.8 (Google, для порівняння)'; 'Результат' = (New-Tag $(if ($pubDns.Ok) { 'OK' } else { 'FAIL' })); 'Час' = (Format-Ms $pubDns.Ms); 'Відповідь' = $pubDns.Result }

$dnsOk = $sysDns.Ok
if ($sysDns.Ok -and $cfgOk -eq $dnsServers.Count) {
    Add-Check 'OK' 'DNS' "$DnsTestName -> $($sysDns.Result) ($($sysDns.Ms) мс)"
} elseif ($sysDns.Ok) {
    Add-Check 'WARN' 'DNS' "Працює, але не всі DNS-сервери відповідають ($cfgOk з $($dnsServers.Count))"
} elseif ($pubDns.Ok) {
    Add-Check 'FAIL' 'DNS' "DNS-сервер із налаштувань ($($dnsServers -join ', ')) не працює, а 8.8.8.8 працює"
} else {
    Add-Check 'FAIL' 'DNS' 'Імена сайтів не перетворюються на адреси'
}

$domainName = $null
if ($cs.PartOfDomain) {
    $domainName = $cs.Domain
    $dd = Test-Dns $domainName $null
    $dnsRows += [pscustomobject]@{ 'DNS-сервер' = "Домен $domainName"; 'Результат' = (New-Tag $(if ($dd.Ok) { 'OK' } else { 'FAIL' })); 'Час' = (Format-Ms $dd.Ms); 'Відповідь' = $dd.Result }
    if ($dd.Ok) { Add-Check 'OK' 'DNS домену' "$domainName -> $($dd.Result)" }
    else { Add-Check 'FAIL' 'DNS домену' "Не знаходиться $domainName - вхід у домен і мережеві диски можуть не працювати. Як DNS має стояти контролер домену" }
}

# =====================================================================
#  6. САЙТИ (HTTP / HTTPS) + ГОДИННИК
# =====================================================================
$httpOk = $false; $httpsOk = $false; $serverDate = $null
try {
    $r = Invoke-WebRequest 'http://www.msftconnecttest.com/connecttest.txt' -UseBasicParsing -TimeoutSec 8 -ErrorAction Stop
    if ($r.Content -match 'Microsoft Connect Test') { $httpOk = $true; Add-Check 'OK' 'HTTP' 'Перевірка Windows (msftconnecttest) пройдена' }
    else { Add-Check 'WARN' 'HTTP' 'Відповідь підмінена: можливо, сторінка авторизації (гостьовий Wi-Fi) або проксі' }
} catch {
    Add-Check 'FAIL' 'HTTP' $_.Exception.Message
}
try {
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $r = Invoke-WebRequest 'https://www.google.com/generate_204' -UseBasicParsing -TimeoutSec 10 -ErrorAction Stop
    $sw.Stop()
    $httpsOk = $true
    $serverDate = $r.Headers['Date']
    Add-Check 'OK' 'HTTPS' "google.com відкривається ($($sw.ElapsedMilliseconds) мс)"
} catch {
    Add-Check 'FAIL' 'HTTPS' "Сайти не відкриваються: $($_.Exception.Message)"
}

$timeInfo = ''
if ($serverDate) {
    try {
        $srv = [datetime]::ParseExact($serverDate, 'r', [Globalization.CultureInfo]::InvariantCulture)
        $diff = [math]::Round(([DateTime]::UtcNow - $srv).TotalSeconds)
        $timeInfo = "Різниця з точним часом: $diff с"
        if ([math]::Abs($diff) -gt 120) { Add-Check 'WARN' 'Годинник ПК' ("Відстає/поспішає на {0:N0} хв - це ламає HTTPS і вхід у домен" -f ([math]::Abs($diff) / 60)) }
        else { Add-Check 'OK' 'Годинник ПК' "Точний (різниця $diff с)" }
    } catch { }
}

# =====================================================================
#  7. ПРОКСІ, HOSTS, ПРОФІЛЬ МЕРЕЖІ
# =====================================================================
$ie = Get-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings' -ErrorAction SilentlyContinue
$proxyText = 'Вимкнено'
if ($ie.ProxyEnable -eq 1) { $proxyText = "Увімкнено: $($ie.ProxyServer)" }
if ($ie.AutoConfigURL)     { $proxyText += "; скрипт автоналаштування: $($ie.AutoConfigURL)" }
$winhttp = ((netsh winhttp show proxy 2>$null) | Where-Object { $_.Trim() } | Select-Object -Skip 1 | ForEach-Object { $_.Trim() }) -join ' '
if ($ie.ProxyEnable -eq 1 -or $ie.AutoConfigURL) { Add-Check 'INFO' 'Проксі' $proxyText }

$hostsPath = Join-Path $env:SystemRoot 'System32\drivers\etc\hosts'
$hostsLines = @(Get-Content $hostsPath -ErrorAction SilentlyContinue | Where-Object { $_.Trim() -and $_.Trim() -notmatch '^#' })
if ($hostsLines.Count -gt 0) { Add-Check 'WARN' 'Файл hosts' "Є записи ($($hostsLines.Count)) - вони перекривають DNS, див. звіт" }

$profiles = @(Get-NetConnectionProfile -ErrorAction SilentlyContinue)
$profRows = foreach ($p in $profiles) {
    $cat = switch ([string]$p.NetworkCategory) { 'Public' { 'Загальнодоступна' } 'Private' { 'Приватна' } 'DomainAuthenticated' { 'Доменна' } default { [string]$p.NetworkCategory } }
    [pscustomobject]@{
        'Мережа'     = $p.Name
        'Адаптер'    = $p.InterfaceAlias
        'Тип'        = $cat
        'IPv4'       = [string]$p.IPv4Connectivity
    }
    if ($p.NetworkCategory -eq 'Public' -and $p.InterfaceIndex -eq $mainIdx) {
        Add-Check 'WARN' 'Тип мережі' "«$($p.Name)» позначена як загальнодоступна - спільні папки й принтери в офісі можуть не працювати"
    }
}

# =====================================================================
#  8. СЕРВЕРИ ТА ПОРТИ
# =====================================================================
$allTargets = New-Object System.Collections.Generic.List[object]
if ($cs.PartOfDomain) {
    $dc = $null
    try {
        Add-Type -AssemblyName System.DirectoryServices -ErrorAction Stop
        $dc = ([System.DirectoryServices.ActiveDirectory.Domain]::GetComputerDomain().FindDomainController()).Name
    } catch { }
    if (-not $dc -and $env:LOGONSERVER -and ($env:LOGONSERVER -replace '\\', '') -ne $env:COMPUTERNAME) { $dc = $env:LOGONSERVER -replace '\\', '' }
    if ($dc) {
        foreach ($pp in @(@{ P = 53; N = 'DNS' }, @{ P = 88; N = 'Kerberos' }, @{ P = 389; N = 'LDAP' }, @{ P = 445; N = 'SMB' })) {
            $allTargets.Add(@{ Name = "Контролер домену ($($pp.N))"; Host = $dc; Port = $pp.P })
        }
    }
}
foreach ($t in $Targets) { $allTargets.Add($t) }

$portRows = @()
$portFails = 0
foreach ($t in $allTargets) {
    $r = Test-TcpPort $t.Host $t.Port
    if (-not $r.Ok) { $portFails++ }
    $portRows += [pscustomobject]@{
        'Стан'   = (New-Tag $(if ($r.Ok) { 'OK' } else { 'FAIL' }) $(if ($r.Ok) { 'Доступно' } else { 'Недоступно' }))
        'Сервіс' = $t.Name
        'Адреса' = "$($t.Host):$($t.Port)"
        'Час'    = (Format-Ms $r.Ms)
        'Помилка'= $r.Error
    }
    if ($r.Ok) { Add-Check 'OK' $t.Name "$($t.Host):$($t.Port) ($($r.Ms) мс)" }
    else       { Add-Check 'FAIL' $t.Name "$($t.Host):$($t.Port) - $($r.Error)" }
}

# =====================================================================
#  9. ЯКІСТЬ ДЛЯ IP-ТЕЛЕФОНІЇ
# =====================================================================
if ($best) {
    $voipProblems = @()
    if ($best.Avg -gt 150)   { $voipProblems += "затримка $($best.Avg) мс (норма < 150)" }
    if ($best.Jitter -gt 30) { $voipProblems += "джитер $($best.Jitter) мс (норма < 30)" }
    if ($best.Loss -gt 1)    { $voipProblems += "втрати $($best.Loss)% (норма < 1%)" }
    if ($voipProblems.Count) { Add-Check 'WARN' 'Якість для IP-телефонії' ("Можливі обриви та «робот» у розмові: " + ($voipProblems -join ', ')) }
    else { Add-Check 'OK' 'Якість для IP-телефонії' "Затримка $($best.Avg) мс, джитер $($best.Jitter) мс, втрати $($best.Loss)%" }
}

# =====================================================================
#  10. ПУБЛІЧНА IP, ТРАСУВАННЯ, ШВИДКІСТЬ
# =====================================================================
$publicIp = ''
if ($httpsOk) {
    try { $publicIp = (Invoke-RestMethod 'https://api.ipify.org' -TimeoutSec 6 -ErrorAction Stop).ToString().Trim() } catch { }
    if ($publicIp) { Add-Check 'INFO' 'Зовнішня IP-адреса' $publicIp }
}

$traceRows = @()
if (-not $Quick -and $ipOk) {
    Write-Host ''
    Write-Host "Трасування до $TraceTarget ..." -ForegroundColor Cyan
    $p = New-Object System.Net.NetworkInformation.Ping
    $buf = New-Object byte[] 32
    $silent = 0
    for ($ttl = 1; $ttl -le 20; $ttl++) {
        $opt = New-Object System.Net.NetworkInformation.PingOptions($ttl, $true)
        $times = @(); $addr = $null; $reached = $false
        for ($k = 0; $k -lt 2; $k++) {
            $sw = [Diagnostics.Stopwatch]::StartNew()
            $r = $null
            try { $r = $p.Send($TraceTarget, 1000, $buf, $opt) } catch { }
            $sw.Stop()
            if ($r -and ($r.Status -eq 'TtlExpired' -or $r.Status -eq 'Success')) {
                $addr = $r.Address.ToString(); $times += $sw.ElapsedMilliseconds
                if ($r.Status -eq 'Success') { $reached = $true }
            }
        }
        $name = ''
        if ($addr) {
            $silent = 0
            try { $task = [System.Net.Dns]::GetHostEntryAsync($addr); if ($task.Wait(600)) { $name = $task.Result.HostName } } catch { }
            if ($name -eq $addr) { $name = '' }
        } else { $silent++ }
        $tText = '*'
        if ($times.Count) { $tText = ($times | ForEach-Object { "$_ мс" }) -join ' / ' }
        $traceRows += [pscustomobject]@{ 'Хоп' = $ttl; 'Адреса' = $(if ($addr) { $addr } else { '*' }); 'Ім''я' = $name; 'Час' = $tText }
        Write-Host ('  {0,2}  {1,-16} {2,-18} {3}' -f $ttl, $(if ($addr) { $addr } else { '*' }), $tText, $name) -ForegroundColor DarkGray
        if ($reached -or $silent -ge 5) { break }
    }
    $p.Dispose()
}

$speed = $null
if (-not $Quick -and $httpsOk) {
    Write-Host ''
    Write-Host "Тест швидкості завантаження (~$SpeedTestSeconds с)..." -ForegroundColor Cyan
    try {
        $req = [System.Net.HttpWebRequest]::Create($SpeedTestUrl)
        $req.Timeout = 10000; $req.ReadWriteTimeout = 10000
        $req.UserAgent = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) NetworkDiag'
        $resp = $req.GetResponse(); $s = $resp.GetResponseStream()
        $buf2 = New-Object byte[] 131072; $total = 0
        $sw = [Diagnostics.Stopwatch]::StartNew()
        while ($sw.Elapsed.TotalSeconds -lt $SpeedTestSeconds) {
            $n = $s.Read($buf2, 0, $buf2.Length)
            if ($n -le 0) { break }
            $total += $n
        }
        $sw.Stop(); $s.Close(); $resp.Close()
        $speed = [math]::Round($total * 8 / $sw.Elapsed.TotalSeconds / 1e6, 1)
        if ($speed -lt 5) { Add-Check 'WARN' 'Швидкість завантаження' "$speed Мбіт/с - дуже повільно" }
        else { Add-Check 'OK' 'Швидкість завантаження' "$speed Мбіт/с" }
    } catch {
        Add-Check 'INFO' 'Швидкість завантаження' "Не вдалося виміряти: $($_.Exception.Message)"
    }
}

# =====================================================================
#  ВИСНОВОК
# =====================================================================
$warnCount = @($Checks | Where-Object { $_.Status -eq 'WARN' }).Count
$failCount = @($Checks | Where-Object { $_.Status -eq 'FAIL' }).Count

if ($upPhys.Count -eq 0) {
    $v = @('bad', 'Немає підключення до мережі', 'Перевірте, чи вставлений кабель і чи світиться індикатор на мережевій карті, або підключіться до Wi-Fi. Перевірте, чи не вимкнено адаптер.')
} elseif (-not $ipOk) {
    $v = @('bad', 'ПК не отримав IP-адресу', 'Роутер або DHCP-сервер не видав адресу. Перезавантажте роутер, перевірте кабель. Спробуйте: ipconfig /release, потім ipconfig /renew.')
} elseif (-not $gateway) {
    $v = @('bad', 'Не налаштований шлюз', 'У налаштуваннях мережі немає основного шлюзу. Якщо адреса статична, вкажіть шлюз вручну або увімкніть DHCP.')
} elseif (-not $gwOk -and -not $gwArp -and -not $inetOk) {
    $v = @('bad', 'Роутер (шлюз) не відповідає', 'Проблема між ПК і роутером: кабель, мережева розетка, комутатор або сам роутер. Перевірте, чи працює мережа на сусідніх ПК.')
} elseif (-not $inetOk) {
    $v = @('bad', 'Локальна мережа працює, а інтернету немає', 'Роутер доступний, але далі зв''язку немає. Ймовірно, проблема в роутері або в провайдера. Перевірте інші ПК; якщо в них теж немає, телефонуйте провайдеру.')
} elseif (-not $dnsOk) {
    $v = @('bad', 'Інтернет є, але не працює DNS', 'Сайти не відкриваються за назвою. Виконайте ipconfig /flushdns. Якщо не допоможе, перевірте DNS-сервер у налаштуваннях (у домені це контролер домену; без домену можна тимчасово вказати 8.8.8.8).')
} elseif (-not $httpsOk) {
    $v = @('bad', 'Інтернет і DNS працюють, але сайти не відкриваються', 'Ймовірні причини: проксі, антивірус або брандмауер блокує браузер, неправильний час на ПК, сторінка авторизації гостьового Wi-Fi.')
} elseif ($failCount -gt 0 -or $warnCount -gt 0) {
    $v = @('warn', 'Інтернет працює, але є зауваження', "Помилок: $failCount, попереджень: $warnCount. Див. пункти нижче.")
} else {
    $v = @('ok', 'Мережа працює нормально', 'Усі перевірки пройдені.')
}

Write-Host ''
Write-Host '========================================================' -ForegroundColor Cyan
$vColor = @{ ok = 'Green'; warn = 'Yellow'; bad = 'Red' }[$v[0]]
Write-Host "  ВИСНОВОК: $($v[1])" -ForegroundColor $vColor
Write-Host "  $($v[2])"
Write-Host '========================================================' -ForegroundColor Cyan

# =====================================================================
#  HTML
# =====================================================================
$checkRows = foreach ($c in $Checks) {
    [pscustomobject]@{ 'Статус' = (New-Tag $c.Status); 'Перевірка' = $c.Name; 'Деталі' = $c.Detail }
}

$pingRows = @()
if ($gwStats) { $pingRows += $gwStats | Select-Object @{ n = 'Куди'; e = { "$($_.Target) (шлюз)" } }, @{ n = 'Отримано'; e = { "$($_.Received) з $($_.Sent)" } }, @{ n = 'Втрати, %'; e = { $_.Loss } }, @{ n = 'Мін, мс'; e = { $_.Min } }, @{ n = 'Сер, мс'; e = { $_.Avg } }, @{ n = 'Макс, мс'; e = { $_.Max } }, @{ n = 'Джитер, мс'; e = { $_.Jitter } } }
foreach ($st in $inetStats) { $pingRows += $st | Select-Object @{ n = 'Куди'; e = { $_.Target } }, @{ n = 'Отримано'; e = { "$($_.Received) з $($_.Sent)" } }, @{ n = 'Втрати, %'; e = { $_.Loss } }, @{ n = 'Мін, мс'; e = { $_.Min } }, @{ n = 'Сер, мс'; e = { $_.Avg } }, @{ n = 'Макс, мс'; e = { $_.Max } }, @{ n = 'Джитер, мс'; e = { $_.Jitter } } }

$cards = [ordered]@{
    'Комп''ютер'    = $env:COMPUTERNAME
    'Локальна IP'   = $(if ($mainIp) { $mainIp.IPAddress } else { '-' })
    'Шлюз'          = $(if ($gateway) { $gateway } else { '-' })
    'Зовнішня IP'   = $(if ($publicIp) { $publicIp } else { '-' })
    'Пінг до інтернету' = $(if ($best) { "$($best.Avg) мс" } else { '-' })
    'Швидкість'     = $(if ($speed) { "$speed Мбіт/с" } else { '-' })
}

$other = [ordered]@{
    'Проксі (браузер)'  = $proxyText
    'Проксі (WinHTTP)'  = $winhttp
    'Годинник'          = $timeInfo
    'Домен'             = $(if ($domainName) { $domainName } else { 'Не в домені' })
}

$cmds = @'
<table class="kv"><tbody>
<tr><th><code>ipconfig /flushdns</code></th><td>Очистити кеш DNS (сайт переїхав / відкривається старий)</td></tr>
<tr><th><code>ipconfig /release</code> + <code>ipconfig /renew</code></th><td>Отримати нову IP-адресу від DHCP</td></tr>
<tr><th><code>netsh winsock reset</code></th><td>Скинути мережевий стек (від адміністратора, потім перезавантаження)</td></tr>
<tr><th><code>netsh int ip reset</code></th><td>Скинути налаштування TCP/IP (від адміністратора, потім перезавантаження)</td></tr>
<tr><th>Параметри > Мережа та Інтернет > Скидання мережі</th><td>Повне скидання всіх адаптерів (крайній захід)</td></tr>
</tbody></table>
'@

$body  = "<div class=""verdict $($v[0])""><h2>$(Enc $v[1])</h2><p>$(Enc $v[2])</p></div>"
$body += New-Cards $cards
$body += New-Section 'Покрокові перевірки' (ConvertTo-HtmlTable $checkRows -RawColumns 'Статус' -Name 'checks')
$body += New-Section 'Мережеві адаптери' (ConvertTo-HtmlTable $adapterRows -RawColumns 'Стан' -Name 'adapters')
$body += New-Section 'IP-налаштування' (ConvertTo-HtmlTable $ipRows -Name 'ip')
if ($profRows) { $body += New-Section 'Тип мережі' (ConvertTo-HtmlTable $profRows -Name 'profiles') 'Для офісної мережі має бути «Приватна» або «Доменна». «Загальнодоступна» блокує спільні папки й принтери.' }
$body += New-Section 'Пінг' (ConvertTo-HtmlTable $pingRows -Name 'ping') 'Джитер - наскільки «скаче» затримка. Для дзвінків (MicroSIP) важливо: затримка < 150 мс, джитер < 30 мс, втрати < 1%.'
$body += New-Section 'DNS' (ConvertTo-HtmlTable $dnsRows -RawColumns 'Результат' -Name 'dns')
if ($traceRows) { $body += New-Section "Трасування до $TraceTarget" (ConvertTo-HtmlTable $traceRows -Name 'traceroute') 'Зірочка (*) - вузол не відповідає. Це нормально для окремих вузлів провайдера, якщо далі маршрут продовжується.' }
if ($portRows) { $body += New-Section 'Сервери та порти' (ConvertTo-HtmlTable $portRows -RawColumns 'Стан' -Name 'ports') 'Свої сервери додаються в розділі НАЛАШТУВАННЯ на початку скрипта.' }
$body += New-Section 'Проксі, годинник, домен' (ConvertTo-KvTable $other)
if ($hostsLines.Count) { $body += New-Section 'Записи у файлі hosts' ("<pre>$(Enc ($hostsLines -join "`n"))</pre>") }
$body += New-Section 'Корисні команди' $cmds

$html = New-HtmlPage "Діагностика мережі: $env:COMPUTERNAME" "$($v[1]) · $(Get-Date -Format 'dd.MM.yyyy HH:mm')" $body 'Test-Network'
$path = Save-Html $html "NetworkDiag_${env:COMPUTERNAME}_$Stamp.html"
Write-Host "Звіт: $path" -ForegroundColor Green
if (-not $NoOpen) { Start-Process $path }
