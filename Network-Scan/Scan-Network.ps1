#Requires -Version 5.1
<#
.SYNOPSIS
    Сканер мережі з акцентом на безпеку: що є в мережі, що відкрито і як це закрити.

.DESCRIPTION
    Працює однаково локально і всередині RDP-сеансу (тоді сканує мережі віддаленого ПК).
    Без установки, без прав адміністратора.

    - Знаходить пристрої: ping, ARP, TCP та NetBIOS (знаходить і ті, що блокують ping).
    - Визначає ім'я, MAC, виробника, тип (ПК, принтер, камера, роутер, NAS...).
    - Перевіряє ~35 важливих портів і пояснює ризик та як закрити.
    - Показує, що видно з інтернету (Shodan InternetDB) і що прокинув роутер (UPnP).
    - Перевіряє сам цей ПК: RDP, NLA, брандмауер, SMBv1, які порти він слухає.
    - Звіряє з Active Directory: пристрої поза доменом і ПК домену, що не відповідають.
    - Пам'ятає попередні сканування: нові та зниклі пристрої.
    - Показує вільні IP-адреси.

.PARAMETER Subnet
    Свої мережі для сканування, напр. -Subnet 192.168.1.0/24,10.0.5.0/24.
    Без параметра скануються мережі всіх активних адаптерів (включно з VPN).

.PARAMETER NoExternal
    Не перевіряти, що видно з інтернету (не звертатися до api.ipify.org та internetdb.shodan.io).

.EXAMPLE
    .\Scan-Network.ps1
.EXAMPLE
    .\Scan-Network.ps1 -Subnet 192.168.10.0/24 -NoExternal
#>
[CmdletBinding()]
param(
    [string[]]$Subnet,
    [switch]$NoExternal,
    [switch]$NoOpen
)

# =====================================================================
#  НАЛАШТУВАННЯ
# =====================================================================
$MaxHosts       = 4094     # більші мережі урізаються до /24 навколо вашої адреси
$PingTimeoutMs  = 900
$PortTimeoutMs  = 900
$BatchSize      = 400      # одночасних TCP-з'єднань
$DataDir        = Join-Path $env:ProgramData 'NetScan'   # історія сканувань і база виробників

# Порти: Sev = HIGH (критично) / MED (увага) / LOW (до відома) / INFO (норма)
$PortInfo = @{
    21    = @{ N = 'FTP';            Sev = 'MED';  What = 'Передача файлів без шифрування: логіни й паролі йдуть відкритим текстом.'; Fix = 'Вимкніть FTP або замініть на SFTP/FTPS. Ніколи не відкривайте в інтернет.' }
    22    = @{ N = 'SSH';            Sev = 'INFO'; What = 'Віддалене керування (Linux, роутери, NAS).'; Fix = 'Вхід лише за ключем або складним паролем; в інтернет - тільки за потреби й з обмеженням IP.' }
    23    = @{ N = 'Telnet';         Sev = 'HIGH'; What = 'Віддалене керування без шифрування, пароль перехоплюється в мережі.'; Fix = 'Вимкніть Telnet у налаштуваннях пристрою, використовуйте SSH або веб-панель з HTTPS.' }
    25    = @{ N = 'SMTP';           Sev = 'INFO'; What = 'Поштовий сервер.'; Fix = 'Переконайтеся, що сервер не є відкритим ретранслятором (open relay).' }
    53    = @{ N = 'DNS';            Sev = 'INFO'; What = 'DNS-сервер.'; Fix = '' }
    80    = @{ N = 'HTTP';           Sev = 'LOW';  What = 'Веб-інтерфейс без шифрування.'; Fix = 'Змініть пароль за замовчуванням; якщо пристрій уміє, увімкніть HTTPS і вимкніть HTTP.' }
    88    = @{ N = 'Kerberos';       Sev = 'INFO'; What = 'Автентифікація домену (контролер домену).'; Fix = '' }
    110   = @{ N = 'POP3';           Sev = 'LOW';  What = 'Пошта без шифрування.'; Fix = 'Використовуйте POP3S (995).' }
    135   = @{ N = 'RPC';            Sev = 'INFO'; What = 'Служби Windows (керування, WMI).'; Fix = 'Норма для Windows у локальній мережі; ніколи не відкривайте в інтернет.' }
    139   = @{ N = 'NetBIOS';        Sev = 'LOW';  What = 'Застарілий протокол спільного доступу Windows.'; Fix = 'Якщо немає старих пристроїв, вимкніть NetBIOS over TCP/IP у властивостях адаптера.' }
    143   = @{ N = 'IMAP';           Sev = 'LOW';  What = 'Пошта без шифрування.'; Fix = 'Використовуйте IMAPS (993).' }
    389   = @{ N = 'LDAP';           Sev = 'INFO'; What = 'Каталог (контролер домену).'; Fix = '' }
    443   = @{ N = 'HTTPS';          Sev = 'INFO'; What = 'Веб-інтерфейс з шифруванням.'; Fix = '' }
    445   = @{ N = 'SMB';            Sev = 'INFO'; What = 'Спільні папки та принтери Windows.'; Fix = 'Норма в локальній мережі. Ніколи не відкривайте в інтернет; вимкніть SMBv1.' }
    515   = @{ N = 'LPD';            Sev = 'INFO'; What = 'Друк (старий протокол).'; Fix = '' }
    554   = @{ N = 'RTSP';           Sev = 'MED';  What = 'Відеопотік камери/реєстратора.'; Fix = 'Змініть пароль за замовчуванням, оновіть прошивку, винесіть камери в окрему мережу (VLAN) без доступу в інтернет.' }
    631   = @{ N = 'IPP';            Sev = 'INFO'; What = 'Мережевий друк.'; Fix = '' }
    1433  = @{ N = 'MS SQL';         Sev = 'MED';  What = 'База даних доступна з мережі.'; Fix = 'Дозвольте підключення лише з потрібних серверів (правило брандмауера), вимкніть обліковий запис sa.' }
    1540  = @{ N = '1С/BAS агент';   Sev = 'INFO'; What = 'Сервер 1С/BAS.'; Fix = 'Обмежте доступ брандмауером лише для робочих ПК.' }
    1541  = @{ N = '1С/BAS кластер'; Sev = 'INFO'; What = 'Сервер 1С/BAS.'; Fix = 'Обмежте доступ брандмауером лише для робочих ПК.' }
    3306  = @{ N = 'MySQL';          Sev = 'MED';  What = 'База даних доступна з мережі.'; Fix = 'Прив''яжіть до 127.0.0.1 або дозвольте лише потрібні IP.' }
    3389  = @{ N = 'RDP';            Sev = 'MED';  What = 'Віддалений робочий стіл.'; Fix = 'Тільки з NLA і складними паролями; ззовні - лише через VPN. Ніколи не прокидайте 3389 на роутері: це головний шлях шифрувальників.' }
    5000  = @{ N = 'NAS/веб';        Sev = 'LOW';  What = 'Веб-панель (часто Synology NAS).'; Fix = 'Не відкривайте в інтернет, увімкніть 2FA й автооновлення.' }
    5001  = @{ N = 'NAS/веб HTTPS';  Sev = 'INFO'; What = 'Веб-панель NAS з шифруванням.'; Fix = 'Не відкривайте в інтернет без потреби, увімкніть 2FA.' }
    5432  = @{ N = 'PostgreSQL';     Sev = 'MED';  What = 'База даних доступна з мережі.'; Fix = 'Обмежте доступ у pg_hba.conf і брандмауером.' }
    5900  = @{ N = 'VNC';            Sev = 'HIGH'; What = 'Віддалений доступ, часто без шифрування і зі слабким паролем.'; Fix = 'Вимкніть VNC або обмежте доступ брандмауером; краще RDP через VPN чи AnyDesk з паролем.' }
    5985  = @{ N = 'WinRM';          Sev = 'INFO'; What = 'Віддалене керування PowerShell.'; Fix = 'Норма в домені; не відкривайте назовні.' }
    7070  = @{ N = 'AnyDesk';        Sev = 'INFO'; What = 'Прямі підключення AnyDesk.'; Fix = 'Пароль для безконтрольного доступу + білий список ID в налаштуваннях AnyDesk.' }
    8000  = @{ N = 'Камера/веб';     Sev = 'LOW';  What = 'Часто SDK-порт камер Hikvision або веб-сервіс.'; Fix = 'Якщо це камера: пароль, оновлення прошивки, окрема мережа.' }
    8080  = @{ N = 'HTTP-alt';       Sev = 'LOW';  What = 'Додатковий веб-інтерфейс (адмінка, проксі, сервіс).'; Fix = 'Перевірте, що це і чи має пароль; закрийте, якщо не потрібен.' }
    8291  = @{ N = 'MikroTik Winbox';Sev = 'MED';  What = 'Керування роутером MikroTik.'; Fix = 'Дозвольте лише з адмінських IP (IP > Services), оновіть RouterOS.' }
    8443  = @{ N = 'HTTPS-alt';      Sev = 'INFO'; What = 'Додатковий веб-інтерфейс з шифруванням.'; Fix = '' }
    9100  = @{ N = 'Принтер RAW';    Sev = 'INFO'; What = 'Прямий друк на принтер.'; Fix = '' }
    37777 = @{ N = 'Dahua';          Sev = 'MED';  What = 'Службовий порт відеореєстраторів Dahua.'; Fix = 'Пароль, оновлення прошивки, окрема мережа; ніколи не прокидайте в інтернет.' }
}
$ScanPorts      = @($PortInfo.Keys | Sort-Object)
$DiscoveryPorts = @(445, 3389, 80, 443, 22, 9100, 554, 135)
$ExposedHigh    = @(21, 23, 135, 139, 445, 1433, 3306, 3389, 5432, 5900, 8291, 37777)

# =====================================================================
$ErrorActionPreference = 'Continue'
$ProgressPreference    = 'SilentlyContinue'
try { [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12 } catch { }
try { $Host.UI.RawUI.WindowTitle = 'Сканер мережі' } catch { }
Add-Type -AssemblyName System.Net.Http -ErrorAction SilentlyContinue
if (-not ('NetScanTrustAll' -as [type])) {
    Add-Type -TypeDefinition @'
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
public static class NetScanTrustAll {
    public static bool Accept(object s, X509Certificate c, X509Chain ch, SslPolicyErrors e) { return true; }
    public static void Enable()  { ServicePointManager.ServerCertificateValidationCallback = Accept; }
    public static void Disable() { ServicePointManager.ServerCertificateValidationCallback = null; }
}
'@
}

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

New-Item -ItemType Directory -Force -Path (Join-Path $DataDir 'history') -ErrorAction SilentlyContinue | Out-Null

function Write-Phase([string]$Text) { Write-Host "`n>> $Text" -ForegroundColor Cyan }
function Write-Info([string]$Text)  { Write-Host "   $Text" -ForegroundColor DarkGray }

# =====================================================================
#  IP-АРИФМЕТИКА
# =====================================================================
function ConvertTo-UInt([string]$Ip) {
    $b = ([Net.IPAddress]::Parse($Ip)).GetAddressBytes(); [Array]::Reverse($b); [BitConverter]::ToUInt32($b, 0)
}
function ConvertTo-Ip([uint32]$N) {
    $b = [BitConverter]::GetBytes($N); [Array]::Reverse($b); (New-Object Net.IPAddress (, $b)).ToString()
}
function Get-NetworkInfo([string]$Ip, [int]$Prefix) {
    $mask = [uint32](([long]4294967295 -shl (32 - $Prefix)) -band [long]4294967295)
    if ($Prefix -eq 0) { $mask = [uint32]0 }
    $net  = [uint32]((ConvertTo-UInt $Ip) -band $mask)
    $bc   = [uint32]($net + ([long]4294967295 - $mask))
    [pscustomobject]@{ Network = (ConvertTo-Ip $net); Prefix = $Prefix; First = $net + 1; Last = $bc - 1; Count = [long]$bc - $net - 1 }
}
function Get-IpSortKey([string]$Ip) { try { ConvertTo-UInt $Ip } catch { 0 } }

# =====================================================================
#  АСИНХРОННІ ПЕРЕВІРКИ
# =====================================================================
function Wait-Tasks($Tasks, [int]$TimeoutMs) {
    $arr = [System.Threading.Tasks.Task[]]@($Tasks | Where-Object { $_ })
    if ($arr.Count -eq 0) { return }
    try { [void][System.Threading.Tasks.Task]::WaitAll($arr, $TimeoutMs) } catch { }
}

function Invoke-PingSweep([string[]]$Ips) {
    $alive = @{}
    for ($i = 0; $i -lt $Ips.Count; $i += 512) {
        $chunk = $Ips[$i..([math]::Min($i + 511, $Ips.Count - 1))]
        $jobs = foreach ($ip in $chunk) {
            $p = New-Object System.Net.NetworkInformation.Ping
            try { [pscustomobject]@{ Ip = $ip; P = $p; T = $p.SendPingAsync($ip, $PingTimeoutMs) } } catch { $p.Dispose() }
        }
        Wait-Tasks $jobs.T ($PingTimeoutMs + 1500)
        foreach ($j in $jobs) {
            if ($j.T.Status -eq 'RanToCompletion' -and $j.T.Result.Status -eq 'Success') { $alive[$j.Ip] = [int]$j.T.Result.RoundtripTime }
            $j.P.Dispose()
        }
    }
    $alive
}

function Invoke-PortScan([string[]]$Ips, [int[]]$Ports) {
    $open = @{}
    $pairs = foreach ($ip in $Ips) { foreach ($port in $Ports) { ,@($ip, $port) } }
    $pairs = @($pairs)
    for ($i = 0; $i -lt $pairs.Count; $i += $BatchSize) {
        $chunk = $pairs[$i..([math]::Min($i + $BatchSize - 1, $pairs.Count - 1))]
        $jobs = foreach ($pp in $chunk) {
            $c = New-Object System.Net.Sockets.TcpClient
            try { [pscustomobject]@{ Ip = $pp[0]; Port = [int]$pp[1]; C = $c; T = $c.ConnectAsync($pp[0], [int]$pp[1]) } } catch { $c.Close() }
        }
        Wait-Tasks $jobs.T $PortTimeoutMs
        foreach ($j in $jobs) {
            if ($j.T.Status -eq 'RanToCompletion' -and $j.C.Connected) {
                if (-not $open.ContainsKey($j.Ip)) { $open[$j.Ip] = New-Object System.Collections.Generic.List[int] }
                $open[$j.Ip].Add($j.Port)
            }
            try { $j.C.Close() } catch { }
        }
    }
    $open
}

function Invoke-ReverseDns([string[]]$Ips) {
    $names = @{}
    $jobs = foreach ($ip in $Ips) { try { [pscustomobject]@{ Ip = $ip; T = [System.Net.Dns]::GetHostEntryAsync($ip) } } catch { } }
    Wait-Tasks $jobs.T 2500
    foreach ($j in $jobs) {
        if ($j.T.Status -eq 'RanToCompletion' -and $j.T.Result.HostName -and $j.T.Result.HostName -ne $j.Ip) { $names[$j.Ip] = $j.T.Result.HostName }
    }
    $names
}

# NetBIOS (UDP 137): ім'я ПК, робоча група і MAC - працює навіть через VPN, де немає ARP
$NbstatRequest = [byte[]](@(0x13, 0x37, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0x20, 0x43, 0x4B) + @(@(0x41) * 30) + @(0, 0, 0x21, 0, 1))
function ConvertFrom-Nbstat([byte[]]$D) {
    if (-not $D -or $D.Length -lt 57) { return $null }
    $n = $D[56]; $off = 57; $names = @()
    for ($i = 0; $i -lt $n -and ($off + 18) -le $D.Length; $i++) {
        $names += [pscustomobject]@{
            N = ([Text.Encoding]::ASCII.GetString($D, $off, 15)).Trim([char]0, ' ')
            S = $D[$off + 15]
            G = (($D[$off + 16] -band 0x80) -ne 0)
        }
        $off += 18
    }
    $mac = $null
    if (($off + 6) -le $D.Length) {
        $mac = (($D[$off..($off + 5)] | ForEach-Object { '{0:X2}' -f $_ }) -join '-')
        if ($mac -eq '00-00-00-00-00-00') { $mac = $null }
    }
    [pscustomobject]@{
        Name  = ($names | Where-Object { $_.S -eq 0 -and -not $_.G } | Select-Object -First 1).N
        Group = ($names | Where-Object { $_.S -eq 0 -and $_.G } | Select-Object -First 1).N
        Mac   = $mac
    }
}
function Invoke-NetBios([string[]]$Ips) {
    $res = @{}
    for ($i = 0; $i -lt $Ips.Count; $i += 512) {
    $chunk = $Ips[$i..([math]::Min($i + 511, $Ips.Count - 1))]
    $jobs = foreach ($ip in $chunk) {
        $u = New-Object System.Net.Sockets.UdpClient
        try { [void]$u.Send($NbstatRequest, $NbstatRequest.Length, $ip, 137); [pscustomobject]@{ Ip = $ip; U = $u; T = $u.ReceiveAsync() } } catch { $u.Close() }
    }
    Wait-Tasks $jobs.T 1500
    foreach ($j in $jobs) {
        if ($j.T.Status -eq 'RanToCompletion') {
            $info = ConvertFrom-Nbstat $j.T.Result.Buffer
            if ($info -and ($info.Name -or $info.Mac)) { $res[$j.Ip] = $info }
        }
        try { $j.U.Close() } catch { }
    }
    }
    $res
}

function Invoke-WebProbe($Targets) {
    # $Targets: список @{Ip; Url}
    $res = @{}
    if (-not $Targets) { return $res }
    [NetScanTrustAll]::Enable()
    $handler = New-Object System.Net.Http.HttpClientHandler
    $handler.UseProxy = $false
    $handler.AllowAutoRedirect = $true
    $client = New-Object System.Net.Http.HttpClient($handler)
    $client.Timeout = [TimeSpan]::FromSeconds(5)
    [void]$client.DefaultRequestHeaders.UserAgent.TryParseAdd('Mozilla/5.0 (Windows NT 10.0; Win64; x64) NetScan')
    $jobs = foreach ($t in $Targets) { try { [pscustomobject]@{ Ip = $t.Ip; Url = $t.Url; T = $client.GetAsync($t.Url) } } catch { } }
    Wait-Tasks $jobs.T 6000
    foreach ($j in $jobs) {
        if ($j.T.Status -ne 'RanToCompletion') { continue }
        try {
            $r = $j.T.Result
            $server = (@($r.Headers.Server) | ForEach-Object { $_.ToString() }) -join ' '
            $body = ''
            $ct = $r.Content.ReadAsStringAsync()
            if ($ct.Wait(1500)) { $body = $ct.Result }
            $title = ''
            if ($body -match '(?is)<title[^>]*>\s*(.*?)\s*</title>') { $title = [System.Net.WebUtility]::HtmlDecode($Matches[1]) -replace '\s+', ' ' }
            if ($title.Length -gt 80) { $title = $title.Substring(0, 80) + '...' }
            if (-not $res.ContainsKey($j.Ip)) { $res[$j.Ip] = [pscustomobject]@{ Title = $title; Server = $server; Url = $j.Url } }
        } catch { }
    }
    $client.Dispose()
    [NetScanTrustAll]::Disable()
    $res
}

# =====================================================================
#  БАЗА ВИРОБНИКІВ (IEEE OUI)
# =====================================================================
function Get-OuiTable {
    $file = Join-Path $DataDir 'oui.csv'
    $fresh = (Test-Path $file) -and ((Get-Item $file).LastWriteTime -gt (Get-Date).AddDays(-90))
    if (-not $fresh) {
        Write-Info 'Завантажую базу виробників мережевих карт (IEEE, ~5 МБ, раз на 90 днів)...'
        try {
            Invoke-WebRequest 'https://standards-oui.ieee.org/oui/oui.csv' -OutFile "$file.tmp" -UseBasicParsing -TimeoutSec 60 `
                -UserAgent 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0 Safari/537.36' -ErrorAction Stop
            Move-Item "$file.tmp" $file -Force
        } catch {
            Remove-Item "$file.tmp" -ErrorAction SilentlyContinue
            if (-not (Test-Path $file)) { Write-Info 'Не вдалося завантажити (немає інтернету?) - виробник не визначатиметься.'; return @{} }
        }
    }
    $t = @{}
    foreach ($line in [System.IO.File]::ReadLines($file)) {
        if ($line -match '^MA-L,([0-9A-F]{6}),("([^"]*)"|([^,]*))') {
            $name = $Matches[3]; if (-not $name) { $name = $Matches[4] }
            $t[$Matches[1]] = $name.Trim()
        }
    }
    $t
}

function Test-RandomMac([string]$Mac) {
    # другий шістнадцятковий символ 2/6/A/E = локально призначена (випадкова) адреса телефонів і ноутбуків
    $Mac -match '^.[26AEae]'
}

# =====================================================================
#  ТИП ПРИСТРОЮ І РИЗИКИ
# =====================================================================
function Get-DeviceType($H) {
    $p = $H.Ports
    $txt = "$($H.Vendor) $($H.WebTitle) $($H.WebServer) $($H.Name)"
    if ($H.IsSelf)    { return 'Цей комп''ютер' }
    if ($H.IsGateway) { return 'Роутер / шлюз' }
    if ($p -contains 88 -and $p -contains 389) { return 'Контролер домену' }
    if ($p -contains 9100 -or $p -contains 515 -or $p -contains 631 -or $txt -match 'printer|laserjet|officejet|deskjet|epson|canon|brother|kyocera|xerox|ricoh|konica|lexmark|pantum|sharp|oki data') { return 'Принтер / МФУ' }
    if ($p -contains 554 -or $p -contains 37777 -or $txt -match 'hikvision|dahua|uniview|ezviz|imou|axis comm|tiandy|camera|nvr|dvr|webs\b|ipcam') { return 'Камера / відеореєстратор' }
    if ($txt -match 'synology|qnap|diskstation|asustor|terramaster|western digital|my cloud') { return 'NAS (мережеве сховище)' }
    if ($p -contains 8291 -or $txt -match 'mikrotik|routeros|ubiquiti|unifi|tp-?link|keenetic|zyxel|d-link|cisco|netgear|aruba|juniper|ruijie|tenda|mercusys|grandstream|yealink|fanvil') {
        if ($txt -match 'grandstream|yealink|fanvil|snom|polycom') { return 'IP-телефон' }
        return 'Мережеве обладнання'
    }
    if ($p -contains 3389 -or $p -contains 445 -or $p -contains 135) {
        if ($p -contains 1433 -or $p -contains 1541 -or $p -contains 3306 -or $p -contains 5432 -or $H.Name -match 'srv|server|^dc|^fs|^sql|^ts|^rds') { return 'Windows-сервер' }
        return 'Windows-ПК'
    }
    if ($txt -match 'apple|samsung|xiaomi|oneplus|oppo|vivo|realme|huawei|honor|motorola|google') { return 'Смартфон / планшет / Mac' }
    if ($H.RandomMac) { return 'Телефон / ноутбук (приватна MAC)' }
    if ($p -contains 22) { return 'Linux / пристрій з SSH' }
    if ($p.Count -gt 0) { return 'Пристрій з відкритими портами' }
    'Невідомо'
}

$SevOrder = @{ HIGH = 3; MED = 2; LOW = 1; INFO = 0 }
$Findings = New-Object System.Collections.Generic.List[object]
function Add-Finding([string]$Sev, [string]$Where, [string]$What, [string]$Detail, [string]$Fix) {
    $Findings.Add([pscustomobject]@{ Sev = $Sev; Where = $Where; What = $What; Detail = $Detail; Fix = $Fix })
}
function New-SevTag([string]$Sev) {
    switch ($Sev) { 'HIGH' { New-Tag 'FAIL' 'Критично' } 'MED' { New-Tag 'WARN' 'Увага' } 'LOW' { New-Tag 'INFO' 'До відома' } default { New-Tag 'OK' 'Норма' } }
}

# =====================================================================
#  1. ДЕ МИ І ЩО СКАНУВАТИ
# =====================================================================
$cs = Get-CimInstance Win32_ComputerSystem -ErrorAction SilentlyContinue
$InRdp = [string]$env:SESSIONNAME -match '^RDP-'
Write-Host ''
Write-Host '========================================================' -ForegroundColor Cyan
Write-Host "  СКАНЕР МЕРЕЖІ - $env:COMPUTERNAME" -ForegroundColor Cyan
if ($InRdp) { Write-Host '  Ви в RDP-сеансі: скануються мережі ЦЬОГО (віддаленого) комп''ютера' -ForegroundColor Yellow }
Write-Host '========================================================' -ForegroundColor Cyan

$upAdapters = @{}
foreach ($a in @(Get-NetAdapter -ErrorAction SilentlyContinue | Where-Object { $_.Status -eq 'Up' })) { $upAdapters[[int]$a.InterfaceIndex] = $a }
$gateways = @(Get-NetRoute -DestinationPrefix '0.0.0.0/0' -ErrorAction SilentlyContinue | ForEach-Object { $_.NextHop } | Where-Object { $_ -and $_ -ne '0.0.0.0' } | Select-Object -Unique)
$myIps = @{}
$networks = New-Object System.Collections.Generic.List[object]

foreach ($addr in @(Get-NetIPAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue)) {
    $ad = $upAdapters[[int]$addr.InterfaceIndex]
    if (-not $ad -or $addr.IPAddress -like '169.254.*' -or $addr.IPAddress -like '127.*') { continue }
    $myIps[$addr.IPAddress] = $ad
    if ($Subnet) { continue }
    if ($ad.InterfaceDescription -match 'Hyper-V|VMware|VirtualBox|WSL|Loopback|Npcap') { continue }
    $prefix = [int]$addr.PrefixLength
    $note = ''
    if ($prefix -lt 20 -or (Get-NetworkInfo $addr.IPAddress $prefix).Count -gt $MaxHosts) { $note = "мережа /$prefix завелика, скануємо лише /24 навколо $($addr.IPAddress)"; $prefix = 24 }
    $ni = Get-NetworkInfo $addr.IPAddress $prefix
    $kind = 'локальна'
    if (-not $ad.HardwareInterface -or $ad.InterfaceDescription -match 'VPN|WireGuard|OpenVPN|TAP-|Wintun|Fortinet|Cisco AnyConnect|PANGP|SSTP|L2TP|PPP') { $kind = 'VPN / віртуальна' }
    $networks.Add([pscustomobject]@{ Cidr = "$($ni.Network)/$prefix"; Info = $ni; Adapter = $ad.Name; MyIp = $addr.IPAddress; Kind = $kind; Note = $note })
}
foreach ($s in $Subnet) {
    if ($s -match '^(\d{1,3}(\.\d{1,3}){3})(/(\d{1,2}))?$') {
        $prefix = 32; if ($Matches[4]) { $prefix = [int]$Matches[4] }
        $ni = Get-NetworkInfo $Matches[1] $prefix
        if ($prefix -eq 32) { $ni = [pscustomobject]@{ Network = $Matches[1]; Prefix = 32; First = (ConvertTo-UInt $Matches[1]); Last = (ConvertTo-UInt $Matches[1]); Count = 1 } }
        if ($ni.Count -gt $MaxHosts) { Write-Host "  $s завелика (> $MaxHosts адрес), пропускаю" -ForegroundColor Yellow; continue }
        $networks.Add([pscustomobject]@{ Cidr = "$($ni.Network)/$prefix"; Info = $ni; Adapter = 'задано вручну'; MyIp = ''; Kind = 'задано вручну'; Note = '' })
    } else { Write-Host "  Не розумію мережу «$s», формат: 192.168.1.0/24" -ForegroundColor Yellow }
}
if ($networks.Count -eq 0) { Write-Host 'Немає мереж для сканування.' -ForegroundColor Red; exit 1 }

$allIps = New-Object System.Collections.Generic.List[string]
$ipNet  = @{}
foreach ($n in $networks) {
    Write-Info ("Мережа {0} ({1}, адаптер «{2}», {3} адрес) {4}" -f $n.Cidr, $n.Kind, $n.Adapter, $n.Info.Count, $n.Note)
    for ($x = [long]$n.Info.First; $x -le [long]$n.Info.Last; $x++) {
        $ip = ConvertTo-Ip ([uint32]$x)
        if (-not $ipNet.ContainsKey($ip)) { $allIps.Add($ip); $ipNet[$ip] = $n.Cidr }
    }
}
$scanStart = Get-Date

# =====================================================================
#  2. ПОШУК ПРИСТРОЇВ
# =====================================================================
Write-Phase "Пошук пристроїв ($($allIps.Count) адрес)..."
$ping = Invoke-PingSweep $allIps.ToArray()
Write-Info "відповіли на ping: $($ping.Count)"

$arp = @{}
foreach ($nb in @(Get-NetNeighbor -AddressFamily IPv4 -ErrorAction SilentlyContinue)) {
    if ($nb.State -notin 'Reachable', 'Stale', 'Delay', 'Probe', 'Permanent') { continue }
    $mac = [string]$nb.LinkLayerAddress
    if (-not $mac -or $mac -match '^(00-){5}00$|^(FF-){5}FF$|^01-00-5E') { continue }
    if ($ipNet.ContainsKey($nb.IPAddress)) { $arp[$nb.IPAddress] = $mac.ToUpper() }
}
Write-Info "знайдено в ARP-таблиці: $($arp.Count)"

$silent = @($allIps | Where-Object { -not $ping.ContainsKey($_) -and -not $arp.ContainsKey($_) })
$tcpFound = @{}
if ($silent.Count -gt 0) {
    Write-Info "перевіряю TCP-портами ті, що мовчать ($($silent.Count))..."
    $tcpFound = Invoke-PortScan $silent $DiscoveryPorts
    Write-Info "знайдено через TCP: $($tcpFound.Count)"
}
$nbAll = Invoke-NetBios $allIps.ToArray()

$liveSet = @{}
foreach ($k in $ping.Keys)     { $liveSet[$k] = 1 }
foreach ($k in $arp.Keys)      { $liveSet[$k] = 1 }
foreach ($k in $tcpFound.Keys) { $liveSet[$k] = 1 }
foreach ($k in $nbAll.Keys)    { $liveSet[$k] = 1 }
foreach ($k in $myIps.Keys)    { if ($ipNet.ContainsKey($k)) { $liveSet[$k] = 1 } }
$live = @($liveSet.Keys | Sort-Object { Get-IpSortKey $_ })
Write-Host "   Знайдено пристроїв: $($live.Count)" -ForegroundColor Green

# =====================================================================
#  3. ПОРТИ, ІМЕНА, ВЕБ-ІНТЕРФЕЙСИ
# =====================================================================
Write-Phase "Перевірка портів ($($ScanPorts.Count) портів x $($live.Count) пристроїв)..."
$open = Invoke-PortScan $live $ScanPorts
Write-Phase 'Імена, виробники, веб-інтерфейси...'
$dns = Invoke-ReverseDns $live
$oui = Get-OuiTable
$webTargets = foreach ($ip in $live) {
    if (-not $open.ContainsKey($ip)) { continue }
    $pl = $open[$ip]
    foreach ($wp in @(80, 443, 8080, 8443, 5000, 5001, 8000)) {
        if ($pl -contains $wp) {
            $scheme = 'http'; if ($wp -in 443, 8443, 5001) { $scheme = 'https' }
            $url = "${scheme}://${ip}"; if ($wp -notin 80, 443) { $url += ":$wp" }
            @{ Ip = $ip; Url = $url }
            break
        }
    }
}
$web = Invoke-WebProbe @($webTargets)

$Hosts = foreach ($ip in $live) {
    $mac = $arp[$ip]
    if (-not $mac -and $nbAll[$ip]) { $mac = $nbAll[$ip].Mac }
    if (-not $mac -and $myIps.ContainsKey($ip)) { $mac = ([string]$myIps[$ip].MacAddress).ToUpper() }
    $vendor = ''
    $macHex = ([string]$mac) -replace '[-:]', ''
    if ($macHex.Length -ge 6) { $vendor = $oui[$macHex.Substring(0, 6).ToUpper()] }
    $name = $dns[$ip]
    if (-not $name -and $nbAll[$ip]) { $name = $nbAll[$ip].Name }
    $ports = @(); if ($open.ContainsKey($ip)) { $ports = @($open[$ip] | Sort-Object -Unique) }
    $h = [pscustomobject]@{
        Ip = $ip; Mac = $mac; Vendor = $vendor; Name = $name; Ports = $ports
        Workgroup = $(if ($nbAll[$ip]) { $nbAll[$ip].Group } else { '' })
        WebTitle = $(if ($web[$ip]) { $web[$ip].Title } else { '' })
        WebServer = $(if ($web[$ip]) { $web[$ip].Server } else { '' })
        WebUrl = $(if ($web[$ip]) { $web[$ip].Url } else { '' })
        Ping = $ping[$ip]
        IsSelf = $myIps.ContainsKey($ip); IsGateway = ($gateways -contains $ip)
        RandomMac = ($mac -and (Test-RandomMac $mac))
        Network = $ipNet[$ip]
        Type = ''; Risk = 'INFO'; InAd = $null; Status = ''; FirstSeen = $null
    }
    $h.Type = Get-DeviceType $h
    $h
}
$Hosts = @($Hosts)

# =====================================================================
#  4. РИЗИКИ ПО ПРИСТРОЯХ
# =====================================================================
foreach ($h in $Hosts) {
    $label = $h.Ip; if ($h.Name) { $label = "$($h.Ip) ($($h.Name))" }
    foreach ($p in $h.Ports) {
        $pi = $PortInfo[$p]; if (-not $pi) { continue }
        $sev = $pi.Sev; $what = "$($pi.N) (порт $p)"; $detail = $pi.What
        if ($h.IsGateway -and $p -in 3389, 445, 135, 139, 5900) { $sev = 'HIGH'; $detail = "На роутері/шлюзі відкрито $($pi.N) - так бути не повинно. " + $pi.What }
        if ($h.IsGateway -and $p -in 80, 8080) { $detail = 'Веб-панель роутера без шифрування. Переконайтеся, що вона недоступна з інтернету і пароль змінено.' }
        if ($h.Type -eq 'Камера / відеореєстратор' -and $p -in 80, 8000, 8080) { $sev = 'MED'; $detail = 'Веб-панель камери: камери часто мають заводські паролі й старі прошивки з відомими вразливостями.'; $pi = $PortInfo[554] }
        if ($SevOrder[$sev] -gt $SevOrder[$h.Risk]) { $h.Risk = $sev }
        if ($sev -ne 'INFO') { Add-Finding $sev $label $what $detail $pi.Fix }
    }
}

# =====================================================================
#  5. ЦЕЙ КОМП'ЮТЕР
# =====================================================================
Write-Phase 'Перевірка цього комп''ютера...'
$selfLabel = "Цей ПК ($env:COMPUTERNAME)"
$selfRows = [ordered]@{}
$ts = Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server' -ErrorAction SilentlyContinue
$rdpTcp = Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp' -ErrorAction SilentlyContinue
if ($ts) {
    $rdpOn = $ts.fDenyTSConnections -eq 0
    $nla = $rdpTcp.UserAuthentication -eq 1
    $selfRows['Віддалений робочий стіл (RDP)'] = $(if ($rdpOn) { "Увімкнено, порт $($rdpTcp.PortNumber)" } else { 'Вимкнено' })
    if ($rdpOn) {
        $selfRows['NLA (перевірка до входу)'] = $(if ($nla) { 'Увімкнено' } else { 'ВИМКНЕНО' })
        if (-not $nla) { Add-Finding 'HIGH' $selfLabel 'RDP без NLA' 'Вікно входу доступне ще до перевірки пароля: легше підбирати паролі й атакувати вразливості RDP.' 'Властивості системи > Віддалений доступ > «Дозволяти підключення лише з автентифікацією на рівні мережі».' }
    }
}
$fw = @(Get-NetFirewallProfile -ErrorAction SilentlyContinue)
if ($fw.Count) {
    $selfRows['Брандмауер'] = ($fw | ForEach-Object { "$($_.Name): $(if ($_.Enabled) { 'увімк.' } else { 'ВИМК.' })" }) -join ', '
    foreach ($f in ($fw | Where-Object { -not $_.Enabled })) {
        Add-Finding 'HIGH' $selfLabel "Брандмауер вимкнено (профіль $($f.Name))" 'Усі служби ПК доступні будь-кому в мережі.' 'Увімкніть брандмауер: Безпека Windows > Брандмауер і захист мережі.'
    }
}
$smb1 = Get-Service -Name mrxsmb10 -ErrorAction SilentlyContinue
if ($smb1 -and $smb1.StartType -ne 'Disabled') {
    $selfRows['SMBv1'] = 'Встановлено'
    Add-Finding 'MED' $selfLabel 'Увімкнено SMBv1' 'Застарілий протокол, через який поширювались WannaCry і NotPetya.' 'Компоненти Windows > зніміть «Підтримка загального доступу до файлів SMB 1.0/CIFS» (якщо немає дуже старих пристроїв).'
} elseif ($fw.Count) { $selfRows['SMBv1'] = 'Не встановлено' }

$listenRows = @()
$svcByPid = @{}
foreach ($s in @(Get-CimInstance Win32_Service -Filter "State='Running'" -ErrorAction SilentlyContinue)) {
    if (-not $svcByPid.ContainsKey([int]$s.ProcessId)) { $svcByPid[[int]$s.ProcessId] = @() }
    $svcByPid[[int]$s.ProcessId] += $s.Name
}
$listen = @(Get-NetTCPConnection -State Listen -ErrorAction SilentlyContinue | Where-Object { $_.LocalAddress -notmatch '^(127\.|::1$)' })
foreach ($g in ($listen | Group-Object LocalPort | Sort-Object { [int]$_.Name })) {
    $port = [int]$g.Name
    $c = $g.Group[0]
    if ($port -ge 49152 -and -not $PortInfo[$port]) { continue }   # динамічні порти RPC
    $proc = (Get-Process -Id $c.OwningProcess -ErrorAction SilentlyContinue).ProcessName
    $svc = ''; if ($svcByPid.ContainsKey([int]$c.OwningProcess)) { $svc = ($svcByPid[[int]$c.OwningProcess] | Select-Object -First 3) -join ', ' }
    $pi = $PortInfo[$port]
    $listenRows += [pscustomobject]@{
        'Порт'      = $port
        'Що це'     = $(if ($pi) { $pi.N } else { '' })
        'Програма'  = $proc
        'Служба'    = $svc
        'Адреса'    = (($g.Group | ForEach-Object { $_.LocalAddress } | Select-Object -Unique) -join ', ')
        'Ризик'     = $(if ($pi) { New-SevTag $pi.Sev } else { '' })
    }
}

# =====================================================================
#  6. ЩО ВИДНО З ІНТЕРНЕТУ
# =====================================================================
$publicIp = ''; $shodan = $null; $extRows = @(); $upnpRows = @()
Write-Phase 'Що відкрито назовні (UPnP на роутері, Shodan)...'
try {
    $nat = New-Object -ComObject HNetCfg.NATUPnP -ErrorAction Stop
    $col = $nat.StaticPortMappingCollection
    if ($col) {
        foreach ($m in $col) {
            $upnpRows += [pscustomobject]@{
                'Ззовні'     = "$($m.Protocol) $($m.ExternalPort)"
                'Куди'       = "$($m.InternalClient):$($m.InternalPort)"
                'Увімкнено'  = $(if ($m.Enabled) { 'Так' } else { 'Ні' })
                'Опис'       = $m.Description
            }
            $sev = 'MED'; if ([int]$m.InternalPort -in $ExposedHigh) { $sev = 'HIGH' }
            Add-Finding $sev "Роутер (UPnP)" "Прокинуто порт $($m.ExternalPort) -> $($m.InternalClient):$($m.InternalPort)" "Програма «$($m.Description)» сама відкрила порт в інтернет через UPnP." 'Перевірте, чи це потрібно. Вимкніть UPnP на роутері, а потрібні порти прокидайте вручну й свідомо.'
        }
    }
} catch { }

if (-not $NoExternal) {
    try { $publicIp = ([string](Invoke-RestMethod 'https://api.ipify.org' -TimeoutSec 8 -ErrorAction Stop)).Trim() } catch { }
    if ($publicIp) {
        try {
            $shodan = Invoke-RestMethod "https://internetdb.shodan.io/$publicIp" -TimeoutSec 10 -ErrorAction Stop
        } catch {
            $code = $null; try { $code = [int]$_.Exception.Response.StatusCode } catch { }
            if ($code -eq 404) { $shodan = [pscustomobject]@{ ports = @(); vulns = @(); hostnames = @(); empty = $true } }
        }
        if ($shodan) {
            foreach ($p in @($shodan.ports)) {
                $pi = $PortInfo[[int]$p]
                $sev = 'MED'; if ([int]$p -in $ExposedHigh) { $sev = 'HIGH' } elseif ([int]$p -in 443, 8443) { $sev = 'LOW' }
                $extRows += [pscustomobject]@{ 'Порт' = $p; 'Що це' = $(if ($pi) { $pi.N } else { '?' }); 'Ризик' = (New-SevTag $sev) }
                $fix = 'Закрийте прокидання порту на роутері; для віддаленого доступу використовуйте VPN.'
                if ($pi -and $pi.Fix) { $fix = "$fix $($pi.Fix)" }
                Add-Finding $sev "Інтернет ($publicIp)" "З інтернету видно порт $p $(if ($pi) { $pi.N })" 'Цей порт бачать сканери по всьому світу, і його постійно атакують.' $fix
            }
            if (@($shodan.vulns).Count) {
                Add-Finding 'HIGH' "Інтернет ($publicIp)" "Відомі вразливості: $((@($shodan.vulns) | Select-Object -First 8) -join ', ')" 'Shodan бачить на вашій зовнішній адресі ПЗ з відомими вразливостями.' 'Оновіть прошивку роутера/сервера, що відповідає на цих портах, або закрийте доступ ззовні.'
            }
        }
    }
}

# =====================================================================
#  7. ДОМЕН
# =====================================================================
$adMissing = @()
if ($cs.PartOfDomain) {
    Write-Phase 'Звірка з Active Directory...'
    try {
        Add-Type -AssemblyName System.DirectoryServices -ErrorAction Stop
        $s = New-Object System.DirectoryServices.DirectorySearcher
        $s.Filter = '(&(objectCategory=computer)(!(userAccountControl:1.2.840.113556.1.4.803:=2)))'
        $s.PageSize = 1000
        [void]$s.PropertiesToLoad.AddRange([string[]]@('name', 'dnshostname', 'operatingsystem'))
        $adComps = @($s.FindAll() | ForEach-Object {
            [pscustomobject]@{ Name = [string]$_.Properties['name'][0]; Dns = [string]$_.Properties['dnshostname'][0]; OS = [string]$_.Properties['operatingsystem'][0] }
        })
        $jobs = foreach ($c in $adComps) { if ($c.Dns) { try { [pscustomobject]@{ C = $c; T = [System.Net.Dns]::GetHostAddressesAsync($c.Dns) } } catch { } } }
        Wait-Tasks $jobs.T 4000
        $adByIp = @{}
        foreach ($j in $jobs) {
            if ($j.T.Status -ne 'RanToCompletion') { continue }
            foreach ($a in $j.T.Result) { if ($a.AddressFamily -eq 'InterNetwork') { $adByIp[$a.ToString()] = $j.C } }
        }
        $adNames = @{}; foreach ($c in $adComps) { $adNames[$c.Name.ToUpper()] = $c }
        foreach ($h in $Hosts) {
            $short = ''; if ($h.Name) { $short = ($h.Name -split '\.')[0].ToUpper() }
            $h.InAd = ($adByIp.ContainsKey($h.Ip) -or ($short -and $adNames.ContainsKey($short)))
            if (-not $h.InAd -and $h.Type -in 'Windows-ПК', 'Windows-сервер') {
                Add-Finding 'MED' "$($h.Ip) $(if ($h.Name) { "($($h.Name))" })" 'Windows-комп''ютер поза доменом' 'На нього не діють групові політики, оновлення й контроль доступу домену. Це може бути особистий ноутбук або забутий ПК.' 'З''ясуйте, чий це пристрій: введіть у домен або перенесіть у гостьову мережу.'
            }
        }
        $liveIdx = @{}; foreach ($h in $Hosts) { $liveIdx[$h.Ip] = 1 }
        foreach ($ip in $adByIp.Keys) {
            if ($ipNet.ContainsKey($ip) -and -not $liveIdx.ContainsKey($ip)) {
                $c = $adByIp[$ip]
                $adMissing += [pscustomobject]@{ 'Комп''ютер' = $c.Name; 'IP з DNS' = $ip; 'ОС' = $c.OS }
            }
        }
        Write-Info "ПК у домені: $($adComps.Count); у цій мережі не відповідають: $($adMissing.Count)"
    } catch { Write-Info "Не вдалося прочитати AD: $($_.Exception.Message)" }
}

# =====================================================================
#  8. ІСТОРІЯ: НОВІ ТА ЗНИКЛІ
# =====================================================================
$histKey  = (($networks | ForEach-Object { $_.Cidr }) -join '_') -replace '[/\\:*?"<>|]', '-'
$histFile = Join-Path $DataDir "history\$env:COMPUTERNAME`_$histKey.json"
$prev = @()
if (Test-Path $histFile) {
    try { foreach ($x in (ConvertFrom-Json (Get-Content $histFile -Raw -Encoding UTF8))) { $prev += $x } } catch { }
}
$hadHistory = $prev.Count -gt 0
$prevByKey = @{}; foreach ($p in $prev) { $prevByKey[$p.Key] = $p }
function Get-HostKey($H) { if ($H.Mac -and -not $H.RandomMac) { "MAC:$($H.Mac)" } else { "IP:$($H.Ip)" } }

$nowStr = (Get-Date).ToString('yyyy-MM-dd HH:mm')
$newCount = 0
$seenKeys = @{}
foreach ($h in $Hosts) {
    $k = Get-HostKey $h; $seenKeys[$k] = 1
    if ($prevByKey.ContainsKey($k)) {
        $h.FirstSeen = $prevByKey[$k].FirstSeen
        if ($prevByKey[$k].Ip -ne $h.Ip) { $h.Status = "IP змінено (було $($prevByKey[$k].Ip))" }
    } else {
        $h.FirstSeen = $nowStr
        if ($hadHistory) { $h.Status = 'Новий'; $newCount++ }
    }
}
$gone = @($prev | Where-Object { -not $seenKeys.ContainsKey($_.Key) -and $_.LastSeen -gt (Get-Date).AddDays(-60).ToString('yyyy-MM-dd') } |
          Sort-Object LastSeen -Descending | ForEach-Object {
            [pscustomobject]@{ 'IP' = $_.Ip; 'Ім''я' = $_.Name; 'Тип' = $_.Type; 'Виробник' = $_.Vendor; 'MAC' = $_.Mac; 'Востаннє бачили' = $_.LastSeen }
          })
$save = @($Hosts | ForEach-Object { [pscustomobject]@{ Key = (Get-HostKey $_); Ip = $_.Ip; Mac = $_.Mac; Name = $_.Name; Vendor = $_.Vendor; Type = $_.Type; FirstSeen = $_.FirstSeen; LastSeen = $nowStr } })
$save += @($prev | Where-Object { -not $seenKeys.ContainsKey($_.Key) -and $_.LastSeen -gt (Get-Date).AddDays(-180).ToString('yyyy-MM-dd') })
try { [System.IO.File]::WriteAllText($histFile, (ConvertTo-Json @($save) -Depth 3), $Utf8Bom) } catch { }

# =====================================================================
#  9. ВІЛЬНІ IP І КОНФЛІКТИ
# =====================================================================
$freeRows = foreach ($n in $networks) {
    if ($n.Info.Count -gt 1024) { continue }
    $ranges = @(); $start = $null; $prevN = $null; $cnt = 0
    for ($x = [long]$n.Info.First; $x -le [long]$n.Info.Last + 1; $x++) {
        $ip = $null; $isFree = $false
        if ($x -le [long]$n.Info.Last) { $ip = ConvertTo-Ip ([uint32]$x); $isFree = -not $liveSet.ContainsKey($ip) }
        if ($isFree) { if ($null -eq $start) { $start = $x }; $prevN = $x; $cnt++ }
        elseif ($null -ne $start) {
            if ($start -eq $prevN) { $ranges += (ConvertTo-Ip ([uint32]$start)) } else { $ranges += "$(ConvertTo-Ip ([uint32]$start)) - $((ConvertTo-Ip ([uint32]$prevN)).Split('.')[-1])" }
            $start = $null
        }
    }
    [pscustomobject]@{ 'Мережа' = $n.Cidr; 'Вільних' = $cnt; 'Діапазони' = ($ranges -join ', ') }
}
$conflicts = @(Get-WinEvent -FilterHashtable @{ LogName = 'System'; ProviderName = 'Microsoft-Windows-Tcpip'; Id = 4199; StartTime = (Get-Date).AddDays(-30) } -MaxEvents 20 -ErrorAction SilentlyContinue |
               ForEach-Object { [pscustomobject]@{ 'Коли' = $_.TimeCreated.ToString('dd.MM.yyyy HH:mm'); 'Подія' = ($_.Message -split "`n")[0] } })
if ($conflicts.Count) { Add-Finding 'MED' $selfLabel "Конфлікти IP-адрес: $($conflicts.Count) за 30 днів" 'Інший пристрій мав ту саму IP-адресу: обриви зв''язку, «зникає мережа».' 'Знайдіть пристрій зі статичною IP всередині пулу DHCP і перенесіть його за межі пулу або зарезервуйте адресу на DHCP.' }

# =====================================================================
#  ЗВІТ
# =====================================================================
$elapsed = [int]((Get-Date) - $scanStart).TotalSeconds
# однакові проблеми на різних пристроях - в один рядок
$Grouped = @($Findings | Group-Object { "$($_.Sev)|$($_.What)|$($_.Detail)|$($_.Fix)" } | ForEach-Object {
    $f = $_.Group[0]
    $wh = @($_.Group | ForEach-Object { $_.Where } | Select-Object -Unique)
    $whereHtml = Enc ($wh -join ', ')
    if ($wh.Count -gt 3) { $whereHtml = "<details><summary>$($wh.Count) пристроїв</summary>$(Enc ($wh -join ', '))</details>" }
    [pscustomobject]@{ Sev = $f.Sev; WhereHtml = $whereHtml; What = $f.What; Detail = $f.Detail; Fix = $f.Fix; N = $wh.Count }
} | Sort-Object @{ e = { $SevOrder[$_.Sev] }; Descending = $true }, @{ e = { $_.N }; Descending = $true })
$high = @($Grouped | Where-Object { $_.Sev -eq 'HIGH' }).Count
$med  = @($Grouped | Where-Object { $_.Sev -eq 'MED' }).Count
$outsideAd = @($Hosts | Where-Object { $_.InAd -eq $false }).Count

Write-Host ''
Write-Host '========================================================' -ForegroundColor Cyan
Write-Host ("  Пристроїв: {0}   Нових: {1}   Критичних: {2}   Увага: {3}   ({4} с)" -f $Hosts.Count, $newCount, $high, $med, $elapsed) -ForegroundColor $(if ($high) { 'Red' } elseif ($med) { 'Yellow' } else { 'Green' })
Write-Host '========================================================' -ForegroundColor Cyan

foreach ($f in ($Grouped | Where-Object { $_.Sev -in 'HIGH', 'MED' } | Select-Object -First 12)) {
    $c = 'Yellow'; if ($f.Sev -eq 'HIGH') { $c = 'Red' }
    $n = ''; if ($f.N -gt 1) { $n = " [$($f.N) пристр.]" }
    Write-Host ("  [{0,-4}] {1}{2}" -f $f.Sev, $f.What, $n) -ForegroundColor $c
}
if ($high) { $v = @('bad', "Знайдено критичних проблем: $high", 'Почніть з розділу «Що закрити в першу чергу» - там для кожної проблеми написано, як її усунути.') }
elseif ($med) { $v = @('warn', "Критичних проблем немає, зауважень: $med", 'Перегляньте розділ «Що закрити в першу чергу».') }
else { $v = @('ok', 'Серйозних проблем не знайдено', 'Відкриті порти відповідають звичайній офісній мережі.') }

$context = "Сканування виконано з комп'ютера $env:COMPUTERNAME"
if ($InRdp) { $context += ' через RDP-сеанс: показано мережі цього віддаленого комп''ютера, а не вашого локального.' } else { $context += ' (локально).' }
$context += ' Мережі: ' + (($networks | ForEach-Object { "$($_.Cidr) [$($_.Kind)]" }) -join ', ') + '.'

$cards = [ordered]@{
    'Пристроїв'           = $Hosts.Count
    'Нових'               = $(if ($hadHistory) { $newCount } else { 'перше сканування' })
    'Критично'            = $high
    'Увага'               = $med
    'Зовнішня IP'         = $(if ($publicIp) { $publicIp } else { '-' })
    'Поза доменом'        = $(if ($cs.PartOfDomain) { $outsideAd } else { 'не домен' })
}

$findRows = foreach ($f in ($Grouped | Where-Object { $_.Sev -in 'HIGH', 'MED' })) {
    [pscustomobject]@{ 'Рівень' = (New-SevTag $f.Sev); 'Де' = $f.WhereHtml; 'Проблема' = $f.What; 'Чому це погано' = $f.Detail; 'Як закрити' = $f.Fix }
}
$lowRows = foreach ($f in ($Grouped | Where-Object { $_.Sev -eq 'LOW' })) {
    [pscustomobject]@{ 'Де' = $f.WhereHtml; 'Що' = $f.What; 'Пояснення' = $f.Detail; 'Рекомендація' = $f.Fix }
}

$hostRows = foreach ($h in $Hosts) {
    $ipCell = Enc $h.Ip
    if ($h.WebUrl) { $ipCell = "<a href=""$(Enc $h.WebUrl)"" target=""_blank"">$(Enc $h.Ip)</a>" }
    $st = ''
    if ($h.Status -eq 'Новий') { $st = New-Tag 'WARN' 'Новий' } elseif ($h.Status) { $st = New-Tag 'INFO' $h.Status }
    $adCell = ''
    if ($h.InAd -eq $true) { $adCell = New-Tag 'OK' 'У домені' } elseif ($h.InAd -eq $false) { $adCell = New-Tag 'INFO' 'Поза доменом' }
    $vend = $h.Vendor; if (-not $vend -and $h.RandomMac) { $vend = '(приватна MAC)' }
    [pscustomobject]@{
        'Ризик'           = (New-SevTag $h.Risk)
        'IP'              = $ipCell
        'Ім''я'           = $h.Name
        'Тип'             = $h.Type
        'Виробник'        = $vend
        'MAC'             = $h.Mac
        'Відкриті порти'  = (($h.Ports | ForEach-Object { if ($PortInfo[$_]) { "$_ $($PortInfo[$_].N)" } else { "$_" } }) -join ', ')
        'Веб-сторінка'    = (@($h.WebTitle, $h.WebServer) | Where-Object { $_ }) -join ' · '
        'Домен'           = $adCell
        'Статус'          = $st
        'Ping'            = $(if ($null -ne $h.Ping) { "$($h.Ping) мс" } else { 'не відповідає' })
    }
}

$typeSum = @($Hosts | Group-Object Type | Sort-Object Count -Descending | ForEach-Object { [pscustomobject]@{ 'Тип' = $_.Name; 'Кількість' = $_.Count } })

$body  = "<div class=""verdict $($v[0])""><h2>$(Enc $v[1])</h2><p>$(Enc $v[2])</p><p class=""muted"">$(Enc $context)</p></div>"
$body += New-Cards $cards
$body += New-Section 'Що закрити в першу чергу' $(if ($findRows) { ConvertTo-HtmlTable @($findRows) -RawColumns 'Рівень', 'Де' -Name 'findings' } else { "<p>$(New-Tag 'OK' 'Чисто') Критичних проблем і зауважень не знайдено.</p>" }) 'Відсортовано від найважливішого. «Критично» - закрити сьогодні, «Увага» - запланувати.'
$extHtml = ''
if ($NoExternal) { $extHtml += '<p class="muted">Перевірку з інтернету вимкнено параметром -NoExternal.</p>' }
elseif (-not $publicIp) { $extHtml += '<p class="muted">Не вдалося визначити зовнішню IP (немає інтернету або доступ закрито).</p>' }
elseif ($shodan -and $shodan.empty) { $extHtml += "<p>$(New-Tag 'OK' 'Чисто') Shodan не бачить відкритих портів на $(Enc $publicIp).</p>" }
elseif ($shodan) {
    $extHtml += "<p>Зовнішня адреса <b>$(Enc $publicIp)</b>. Дані Shodan можуть бути давністю до кількох тижнів.</p>"
    $extHtml += ConvertTo-HtmlTable $extRows -RawColumns 'Ризик' -Name 'internet' -NoToolbar
    if (@($shodan.vulns).Count) { $extHtml += "<p>$(New-Tag 'FAIL' 'Вразливості') $(Enc ((@($shodan.vulns)) -join ', '))</p>" }
} else { $extHtml += '<p class="muted">Сервіс Shodan InternetDB не відповів.</p>' }
$extHtml += '<h3>Порти, які роутер прокинув автоматично (UPnP)</h3>'
if ($upnpRows) { $extHtml += ConvertTo-HtmlTable $upnpRows -Name 'upnp' -NoToolbar } else { $extHtml += '<p class="muted">Немає або UPnP на роутері вимкнено (це добре).</p>' }
$extHtml += '<p class="muted">Порти, прокинуті на роутері вручну, звідси не видно: перевірте розділ Port Forwarding / Virtual Server у веб-панелі роутера.</p>'
$body += New-Section 'Що видно з інтернету' $extHtml
$body += New-Section 'Пристрої в мережі' ((ConvertTo-HtmlTable $typeSum -Name 'types' -NoToolbar) + '<h3>Усі пристрої</h3>' + (ConvertTo-HtmlTable @($hostRows) -RawColumns 'Ризик', 'IP', 'Домен', 'Статус' -Name 'devices')) 'IP-адреса з посиланням відкриває веб-інтерфейс пристрою.'
$selfHtml = ConvertTo-KvTable $selfRows
if ($listenRows) { $selfHtml += '<h3>Які порти слухає цей комп''ютер</h3>' + (ConvertTo-HtmlTable $listenRows -RawColumns 'Ризик' -Name 'listening') }
$body += New-Section "Цей комп'ютер ($env:COMPUTERNAME)" $selfHtml 'Корисно в RDP-сеансі на сервері: видно, що саме цей сервер відкриває в мережу і яка програма за цим стоїть.'
if ($hadHistory -or $gone) {
    $histHtml = ''
    $newRows = @($hostRows | Where-Object { $_.'Статус' -match 'Новий' })
    $histHtml += '<h3>Нові з минулого сканування</h3>' + (ConvertTo-HtmlTable $newRows -RawColumns 'Ризик', 'IP', 'Домен', 'Статус' -Name 'new' -NoToolbar)
    $histHtml += '<h3>Зникли (не відповідають зараз)</h3>' + (ConvertTo-HtmlTable $gone -Name 'gone' -NoToolbar)
    $body += New-Section 'Зміни в мережі' $histHtml 'Історія зберігається в C:\ProgramData\NetScan\history окремо для кожного ПК і мережі.'
}
if ($adMissing) { $body += New-Section 'У домені, але не відповідають' (ConvertTo-HtmlTable $adMissing -Name 'ad-missing') 'ПК, чия адреса з DNS потрапляє в цю мережу, але зараз не відповідає: вимкнені, списані або в DNS застарілий запис.' }
if ($freeRows) { $body += New-Section 'Вільні IP-адреси' (ConvertTo-HtmlTable @($freeRows) -Name 'free-ip' -NoToolbar) 'Не відповіли на жодну перевірку. Перед тим як прописати статичну IP, переконайтеся, що адреса поза пулом DHCP роутера.' }
if ($conflicts) { $body += New-Section 'Конфлікти IP (журнал Windows, 30 днів)' (ConvertTo-HtmlTable $conflicts -Name 'conflicts' -NoToolbar) }
if ($lowRows) { $body += New-Section 'До відома' (ConvertTo-HtmlTable @($lowRows) -RawColumns 'Де' -Name 'info') 'Не проблема саме по собі, але варто знати.' }

$html = New-HtmlPage "Сканер мережі: $env:COMPUTERNAME" "$($v[1]) · $(Get-Date -Format 'dd.MM.yyyy HH:mm') · $elapsed с" $body 'Scan-Network'
$path = Save-Html $html "NetScan_${env:COMPUTERNAME}_$Stamp.html"
Write-Host "Звіт: $path" -ForegroundColor Green
if (-not $NoOpen) { Start-Process $path }
