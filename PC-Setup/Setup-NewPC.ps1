#Requires -Version 5.1
<#
.SYNOPSIS
    Розгортання нового ПК: встановлення програм через winget і базові налаштування Windows 10/11.

.DESCRIPTION
    1. Перевіряє winget і, якщо його немає (часто на Windows 10), встановлює автоматично.
    2. Показує меню програм: позначені ставляться, решту можна увімкнути номером.
    3. Встановлює програми тихо, без кліків. Вже встановлені пропускає.
    4. Застосовує налаштування системи (розділ НАЛАШТУВАННЯ нижче).
    5. Зберігає HTML-звіт на робочий стіл і журнал у C:\ProgramData\PC-Setup.

    Скрипт сам попросить права адміністратора (UAC).

.PARAMETER Auto
    Без меню: ставить усе, що позначено Enabled = $true. Зручно для масового розгортання.

.PARAMETER ComputerName
    Нове ім'я ПК (застосується після перезавантаження).

.EXAMPLE
    .\Setup-NewPC.ps1
.EXAMPLE
    .\Setup-NewPC.ps1 -Auto -ComputerName BUH-03
#>
[CmdletBinding()]
param(
    [switch]$Auto,
    [string]$ComputerName,
    [switch]$NoOpen
)

# =====================================================================
#  НАЛАШТУВАННЯ — змінюйте тут
# =====================================================================

# Програми. Enabled = $true — позначено за замовчуванням.
# Id — ідентифікатор winget. Знайти будь-яку програму: winget search <назва>
# UserScope = $true — програма ставиться лише для одного користувача (див. примітку в кінці).
$Apps = @(
    @{ Name = 'Google Chrome';                    Id = 'Google.Chrome';                     Enabled = $true  }
    @{ Name = 'AnyDesk';                          Id = 'AnyDesk.AnyDesk';                   Enabled = $true  }
    @{ Name = 'MicroSIP (IP-телефонія)';          Id = 'MicroSIP.MicroSIP';                 Enabled = $true;  UserScope = $true }
    @{ Name = '7-Zip (архіватор)';                Id = '7zip.7zip';                         Enabled = $true  }
    @{ Name = 'Adobe Acrobat Reader (PDF)';       Id = 'Adobe.Acrobat.Reader.64-bit';       Enabled = $true  }
    @{ Name = 'Visual C++ Runtime x64';           Id = 'Microsoft.VCRedist.2015+.x64';      Enabled = $true  }
    @{ Name = 'Visual C++ Runtime x86';           Id = 'Microsoft.VCRedist.2015+.x86';      Enabled = $true  }
    @{ Name = 'Notepad++';                        Id = 'Notepad++.Notepad++';               Enabled = $true  }
    @{ Name = 'VLC (відео/аудіо)';                Id = 'VideoLAN.VLC';                      Enabled = $true  }
    @{ Name = 'LibreOffice (безкоштовний офіс)';  Id = 'TheDocumentFoundation.LibreOffice'; Enabled = $false }
    @{ Name = 'Microsoft 365 / Office (ліцензія)';Id = 'Microsoft.Office';                  Enabled = $false }
    @{ Name = 'Mozilla Firefox';                  Id = 'Mozilla.Firefox';                   Enabled = $false }
    @{ Name = 'Telegram';                         Id = 'Telegram.TelegramDesktop';          Enabled = $false; UserScope = $true }
    @{ Name = 'Viber';                            Id = 'Rakuten.Viber';                     Enabled = $false }
    @{ Name = 'Zoom';                             Id = 'Zoom.Zoom';                         Enabled = $false }
    @{ Name = 'Microsoft Teams';                  Id = 'Microsoft.Teams';                   Enabled = $false }
    @{ Name = 'TeamViewer';                       Id = 'TeamViewer.TeamViewer';             Enabled = $false }
)

# Налаштування Windows. $true — виконати, $false — пропустити.
$Settings = [ordered]@{
    RestorePoint       = $true                 # точка відновлення перед змінами
    TimeZone           = 'FLE Standard Time'   # Київ. '' — не змінювати
    PowerNoSleepAC     = $true                 # від мережі не засинати, монітор гасне через 20 хв
    DisableFastStartup = $true                 # вимкнути «Швидкий запуск» (менше проблем з оновленнями й мережею)
    ShowFileExtensions = $true                 # показувати розширення (захист від "рахунок.pdf.exe")
    ExplorerThisPC     = $true                 # Провідник відкривається на «Цей ПК»
    DesktopThisPC      = $true                 # значок «Цей ПК» на робочому столі
    DisableAdsTips     = $true                 # без реклами, порад і автоустановки «рекомендованих» програм
    EnableRDP          = $false                # дозволити віддалений робочий стіл (RDP)
    UpgradeAll         = $false                # оновити всі вже встановлені програми
    AskComputerName    = $true                 # запитати нове ім'я ПК
}

# =====================================================================
#  ПРАВА АДМІНІСТРАТОРА
# =====================================================================
$identity  = [Security.Principal.WindowsIdentity]::GetCurrent()
$principal = New-Object Security.Principal.WindowsPrincipal($identity)
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host 'Потрібні права адміністратора, зараз з''явиться запит UAC...' -ForegroundColor Yellow
    $argList = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$PSCommandPath`"")
    if ($Auto)         { $argList += '-Auto' }
    if ($NoOpen)       { $argList += '-NoOpen' }
    if ($ComputerName) { $argList += @('-ComputerName', $ComputerName) }
    try {
        $proc = Start-Process powershell.exe -Verb RunAs -ArgumentList $argList -Wait -PassThru -ErrorAction Stop
    } catch {
        Write-Host 'Запуск скасовано: без прав адміністратора встановлення неможливе.' -ForegroundColor Red
        Read-Host 'Натисніть Enter, щоб закрити' | Out-Null
        exit 1
    }
    exit $proc.ExitCode
}

$ErrorActionPreference = 'Continue'
$ProgressPreference    = 'SilentlyContinue'
try { [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12 } catch { }
try { $Host.UI.RawUI.WindowTitle = 'Налаштування нового ПК' } catch { }

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

$LogDir = Join-Path $env:ProgramData 'PC-Setup'
New-Item -ItemType Directory -Force -Path $LogDir | Out-Null
$LogFile = Join-Path $LogDir "Setup_$Stamp.log"
try { Start-Transcript -Path $LogFile -Append | Out-Null } catch { }

$Results       = New-Object System.Collections.Generic.List[object]
$NeedReboot    = $false

function Add-Result([string]$Group, [string]$Name, [string]$Status, [string]$Detail) {
    $Results.Add([pscustomobject]@{ Group = $Group; Name = $Name; Status = $Status; Detail = $Detail })
    $color = @{ OK = 'Green'; WARN = 'Yellow'; FAIL = 'Red'; INFO = 'Cyan' }[$Status]
    Write-Host ('  [{0,-4}] ' -f $Status) -ForegroundColor $color -NoNewline
    Write-Host $Name -NoNewline
    if ($Detail) { Write-Host " - $Detail" -ForegroundColor DarkGray } else { Write-Host '' }
}

function Write-Title([string]$Text) {
    Write-Host ''
    Write-Host "=== $Text ===" -ForegroundColor Cyan
}

# Користувач, що зараз сидить за ПК (може відрізнятися від адміна, який ввів пароль у UAC)
$CurrentSid  = $identity.User.Value
$ConsoleUser = (Get-CimInstance Win32_ComputerSystem).UserName
$ConsoleSid  = $null
if ($ConsoleUser) {
    try { $ConsoleSid = (New-Object Security.Principal.NTAccount($ConsoleUser)).Translate([Security.Principal.SecurityIdentifier]).Value } catch { }
}
$OtherConsoleUser = ($ConsoleSid -and $ConsoleSid -ne $CurrentSid)

# =====================================================================
#  МЕНЮ ВИБОРУ ПРОГРАМ
# =====================================================================
function Show-AppMenu {
    while ($true) {
        Clear-Host
        Write-Host '========================================================' -ForegroundColor Cyan
        Write-Host '   НАЛАШТУВАННЯ НОВОГО ПК' -ForegroundColor Cyan
        Write-Host '========================================================' -ForegroundColor Cyan
        Write-Host ''
        Write-Host ' Програми для встановлення:' -ForegroundColor White
        for ($i = 0; $i -lt $Apps.Count; $i++) {
            $a = $Apps[$i]
            if ($a.Enabled) { $mark = '[x]'; $col = 'White' } else { $mark = '[ ]'; $col = 'DarkGray' }
            Write-Host ('  {0,2}. {1} {2}' -f ($i + 1), $mark, $a.Name) -ForegroundColor $col
        }
        Write-Host ''
        Write-Host ' Налаштування Windows (змінюються на початку скрипта):' -ForegroundColor White
        $labels = @{
            RestorePoint = 'точка відновлення'; TimeZone = 'часовий пояс'; PowerNoSleepAC = 'без сну від мережі'
            DisableFastStartup = 'без швидкого запуску'; ShowFileExtensions = 'розширення файлів'
            ExplorerThisPC = 'Провідник -> Цей ПК'; DesktopThisPC = 'Цей ПК на столі'; DisableAdsTips = 'без реклами'
            EnableRDP = 'RDP'; UpgradeAll = 'оновити всі програми'; AskComputerName = 'ім''я ПК'
        }
        $on = @($Settings.Keys | Where-Object { $Settings[$_] -and $Settings[$_] -ne '' } | ForEach-Object { $labels[$_] })
        Write-Host ('  ' + ($on -join ', ')) -ForegroundColor DarkGray
        Write-Host ''
        Write-Host '  Номери через кому - позначити/зняти (напр. 10,13)' -ForegroundColor Yellow
        Write-Host '  A - усі   N - жодної   Enter - ПОЧАТИ   Q - вихід' -ForegroundColor Yellow
        $in = Read-Host ' >'
        if ([string]::IsNullOrWhiteSpace($in)) { return $true }
        $in = $in.Trim()
        if     ($in -match '^[qQйЙ]$') { return $false }
        elseif ($in -match '^[aAфФ]$') { foreach ($a in $Apps) { $a.Enabled = $true } }
        elseif ($in -match '^[nNтТ]$') { foreach ($a in $Apps) { $a.Enabled = $false } }
        else {
            foreach ($tok in ($in -split '[,;\s]+')) {
                $n = 0
                if ([int]::TryParse($tok, [ref]$n) -and $n -ge 1 -and $n -le $Apps.Count) {
                    $Apps[$n - 1].Enabled = -not $Apps[$n - 1].Enabled
                }
            }
        }
    }
}

if (-not $Auto) {
    if (-not (Show-AppMenu)) { try { Stop-Transcript | Out-Null } catch { }; exit 0 }
    if ($Settings.AskComputerName -and -not $ComputerName) {
        Write-Host ''
        $ComputerName = Read-Host " Нове ім'я ПК (Enter - залишити $env:COMPUTERNAME)"
    }
}
$Selected = @($Apps | Where-Object { $_.Enabled })
Clear-Host
Write-Host "Початок: $(Get-Date -Format 'dd.MM.yyyy HH:mm'). Журнал: $LogFile" -ForegroundColor DarkGray

# =====================================================================
#  WINGET
# =====================================================================
function Get-WingetPath {
    $cmd = Get-Command winget.exe -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }
    $pkg = Get-AppxPackage -Name Microsoft.DesktopAppInstaller -AllUsers -ErrorAction SilentlyContinue |
           Sort-Object Version -Descending | Select-Object -First 1
    if ($pkg -and $pkg.InstallLocation) {
        $p = Join-Path $pkg.InstallLocation 'winget.exe'
        if (Test-Path $p) { return $p }
    }
    $null
}

function Install-Winget {
    Write-Host '  winget не знайдено. Пробую зареєструвати App Installer...' -ForegroundColor Yellow
    try { Add-AppxPackage -RegisterByFamilyName -MainPackage Microsoft.DesktopAppInstaller_8wekyb3d8bbwe -ErrorAction Stop; Start-Sleep 3 } catch { }
    $p = Get-WingetPath
    if ($p) { return $p }

    Write-Host '  Встановлюю winget офіційним способом Microsoft (модуль Microsoft.WinGet.Client)...' -ForegroundColor Yellow
    try {
        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope AllUsers -ErrorAction Stop | Out-Null
        Install-Module -Name Microsoft.WinGet.Client -Repository PSGallery -Force -AllowClobber -Scope AllUsers -ErrorAction Stop
        Import-Module Microsoft.WinGet.Client -Force -ErrorAction Stop
        Repair-WinGetPackageManager -AllUsers -Force -Latest -ErrorAction Stop | Out-Null
        Start-Sleep 3
    } catch {
        Write-Host "  Помилка: $($_.Exception.Message)" -ForegroundColor Red
    }
    Get-WingetPath
}

# коди завершення winget (HRESULT як Int32)
$WG_ALREADY_INSTALLED = -1978335135   # 0x8A150061
$WG_UPDATE_NOT_APPL   = -1978335189   # 0x8A15002B
$WG_NOT_FOUND         = -1978335212   # 0x8A150014
$WG_NO_INSTALLER      = -1978335216   # 0x8A150010 (немає інсталятора для такого scope/архітектури)
$WG_REBOOT            = @(-1978334967, 3010, 1641)

function ConvertTo-Int32([long]$Code) {
    if ($Code -gt [int]::MaxValue) { return [int]($Code - 4294967296) }
    [int]$Code
}

function Get-InstallStatus([int]$Code) {
    if ($Code -eq 0)                                                    { return @('OK',   'Встановлено') }
    if ($Code -eq $WG_ALREADY_INSTALLED -or $Code -eq $WG_UPDATE_NOT_APPL) { return @('OK',   'Вже встановлено') }
    if ($WG_REBOOT -contains $Code)                                     { $script:NeedReboot = $true; return @('OK', 'Встановлено, потрібне перезавантаження') }
    if ($Code -eq $WG_NOT_FOUND)                                        { return @('FAIL', 'Не знайдено в winget, перевірте Id') }
    @('FAIL', ('Помилка, код 0x{0:X8}' -f $Code))
}

# Встановлення від імені користувача за ПК (для програм, що ставляться лише в профіль)
function Install-AsConsoleUser($App) {
    $task = 'PCSetup_' + ($App.Id -replace '[^\w]', '_')
    $cmd  = "winget install --id $($App.Id) --exact --silent --accept-package-agreements --accept-source-agreements"
    try {
        $action   = New-ScheduledTaskAction -Execute 'cmd.exe' -Argument "/c $cmd"
        $prin     = New-ScheduledTaskPrincipal -UserId $ConsoleUser -LogonType Interactive -RunLevel Limited
        $set      = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Minutes 30)
        Register-ScheduledTask -TaskName $task -Action $action -Principal $prin -Settings $set -Force -ErrorAction Stop | Out-Null
        Start-ScheduledTask -TaskName $task -ErrorAction Stop
        Start-Sleep -Seconds 5
        $deadline = (Get-Date).AddMinutes(20)
        while ((Get-ScheduledTask -TaskName $task).State -eq 'Running' -and (Get-Date) -lt $deadline) { Start-Sleep -Seconds 3 }
        $code = ConvertTo-Int32 (Get-ScheduledTaskInfo -TaskName $task).LastTaskResult
        return $code
    } catch {
        Write-Host "  Не вдалося запустити від імені $ConsoleUser : $($_.Exception.Message)" -ForegroundColor Red
        return -1
    } finally {
        Unregister-ScheduledTask -TaskName $task -Confirm:$false -ErrorAction SilentlyContinue
    }
}

$Winget = $null
if ($Selected.Count -gt 0 -or $Settings.UpgradeAll) {
    Write-Title 'Перевірка winget'
    $Winget = Get-WingetPath
    if (-not $Winget) { $Winget = Install-Winget }
    if ($Winget) {
        $ver = (& $Winget --version 2>$null | Select-Object -First 1)
        Add-Result 'winget' 'Менеджер пакетів winget' 'OK' "$ver"
    } else {
        Add-Result 'winget' 'Менеджер пакетів winget' 'FAIL' 'Не вдалося встановити. Поставте «App Installer» з Microsoft Store (apps.microsoft.com/detail/9NBLGGH4NNS1) і запустіть скрипт знову'
    }
}

# =====================================================================
#  ВСТАНОВЛЕННЯ ПРОГРАМ
# =====================================================================
if ($Winget -and $Selected.Count -gt 0) {
    Write-Title "Встановлення програм ($($Selected.Count))"
    $n = 0
    foreach ($app in $Selected) {
        $n++
        Write-Host ''
        Write-Host ("[{0}/{1}] {2}" -f $n, $Selected.Count, $app.Name) -ForegroundColor White

        $code = $null
        if ($app.UserScope -and $OtherConsoleUser) {
            Write-Host "  Ставлю в профіль користувача $ConsoleUser (вікно з'явиться в його сесії)..." -ForegroundColor DarkGray
            $code = Install-AsConsoleUser $app
        } else {
            & $Winget list --id $app.Id --exact --accept-source-agreements *> $null
            if ($LASTEXITCODE -eq 0) {
                $code = $WG_ALREADY_INSTALLED
            } else {
                $common = @('install', '--id', $app.Id, '--exact', '--silent', '--accept-package-agreements', '--accept-source-agreements')
                if ($app.UserScope) {
                    & $Winget @common
                    $code = ConvertTo-Int32 $LASTEXITCODE
                } else {
                    & $Winget @common --scope machine
                    $code = ConvertTo-Int32 $LASTEXITCODE
                    if ($code -eq $WG_NO_INSTALLER) {
                        Write-Host '  Немає версії «для всіх користувачів», ставлю звичайну...' -ForegroundColor DarkGray
                        & $Winget @common
                        $code = ConvertTo-Int32 $LASTEXITCODE
                    }
                }
            }
        }
        $st = Get-InstallStatus $code
        Add-Result 'Програми' $app.Name $st[0] $st[1]
    }
}

if ($Winget -and $Settings.UpgradeAll) {
    Write-Title 'Оновлення встановлених програм'
    & $Winget upgrade --all --silent --accept-package-agreements --accept-source-agreements
    $c = ConvertTo-Int32 $LASTEXITCODE
    if ($c -eq 0 -or $c -eq $WG_UPDATE_NOT_APPL) { Add-Result 'Програми' 'Оновлення всіх програм' 'OK' '' }
    else { Add-Result 'Програми' 'Оновлення всіх програм' 'WARN' ('Частина не оновилась, код 0x{0:X8}' -f $c) }
}

# =====================================================================
#  НАЛАШТУВАННЯ WINDOWS
# =====================================================================
Write-Title 'Налаштування Windows'

# --- куди писати налаштування користувача ---
# поточний (адмін) + користувач за ПК + профіль за замовчуванням (для всіх нових користувачів)
$UserHives = New-Object System.Collections.Generic.List[string]
$UserHives.Add('Registry::HKEY_CURRENT_USER')
if ($OtherConsoleUser -and (Test-Path "Registry::HKEY_USERS\$ConsoleSid")) { $UserHives.Add("Registry::HKEY_USERS\$ConsoleSid") }
$DefaultLoaded = $false
$needUserSettings = $Settings.ShowFileExtensions -or $Settings.ExplorerThisPC -or $Settings.DesktopThisPC -or $Settings.DisableAdsTips
if ($needUserSettings) {
    try {
        $defProfile = [Environment]::ExpandEnvironmentVariables((Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList').Default)
        $ntuser = Join-Path $defProfile 'NTUSER.DAT'
        if (Test-Path $ntuser) {
            & reg.exe load 'HKU\PCSetupDefault' "$ntuser" *> $null
            if ($LASTEXITCODE -eq 0) { $DefaultLoaded = $true; $UserHives.Add('Registry::HKEY_USERS\PCSetupDefault') }
        }
    } catch { }
}

function Set-UserReg([string]$SubKey, [string]$Name, $Value, [string]$Type = 'DWord') {
    foreach ($root in $UserHives) {
        $path = "$root\$SubKey"
        try {
            if (-not (Test-Path $path)) { New-Item -Path $path -Force -ErrorAction Stop | Out-Null }
            New-ItemProperty -Path $path -Name $Name -Value $Value -PropertyType $Type -Force -ErrorAction Stop | Out-Null
        } catch { Write-Verbose "$path\$Name : $($_.Exception.Message)" }
    }
}

function Set-MachineReg([string]$Path, [string]$Name, $Value, [string]$Type = 'DWord') {
    if (-not (Test-Path $Path)) { New-Item -Path $Path -Force | Out-Null }
    New-ItemProperty -Path $Path -Name $Name -Value $Value -PropertyType $Type -Force -ErrorAction Stop | Out-Null
}

function Invoke-Setting([string]$Name, [scriptblock]$Action) {
    try {
        $detail = & $Action
        Add-Result 'Налаштування' $Name 'OK' ([string]$detail)
    } catch {
        Add-Result 'Налаштування' $Name 'WARN' $_.Exception.Message
    }
}

if ($Settings.RestorePoint) {
    Invoke-Setting 'Точка відновлення' {
        Enable-ComputerRestore -Drive "$env:SystemDrive\" -ErrorAction Stop
        Set-MachineReg 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SystemRestore' 'SystemRestorePointCreationFrequency' 0
        Checkpoint-Computer -Description 'Before PC-Setup' -RestorePointType MODIFY_SETTINGS -ErrorAction Stop
        'Створено'
    }
}

if ($Settings.TimeZone) {
    Invoke-Setting 'Часовий пояс' {
        Set-TimeZone -Id $Settings.TimeZone -ErrorAction Stop
        Set-Service w32time -StartupType Automatic -ErrorAction SilentlyContinue
        Start-Service w32time -ErrorAction SilentlyContinue
        & w32tm.exe /resync /force *> $null
        (Get-TimeZone).DisplayName
    }
}

if ($Settings.PowerNoSleepAC) {
    Invoke-Setting 'Електроживлення' {
        & powercfg.exe /change standby-timeout-ac 0
        & powercfg.exe /change hibernate-timeout-ac 0
        & powercfg.exe /change monitor-timeout-ac 20
        'Від мережі: без сну, монітор - 20 хв'
    }
}

if ($Settings.DisableFastStartup) {
    Invoke-Setting 'Вимкнути «Швидкий запуск»' {
        Set-MachineReg 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Power' 'HiberbootEnabled' 0
        'Вимкнено'
    }
}

if ($Settings.ShowFileExtensions) {
    Invoke-Setting 'Показувати розширення файлів' {
        Set-UserReg 'Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced' 'HideFileExt' 0
        'Увімкнено'
    }
}

if ($Settings.ExplorerThisPC) {
    Invoke-Setting 'Провідник відкривається на «Цей ПК»' {
        Set-UserReg 'Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced' 'LaunchTo' 1
        'Увімкнено'
    }
}

if ($Settings.DesktopThisPC) {
    Invoke-Setting 'Значок «Цей ПК» на робочому столі' {
        Set-UserReg 'Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel' '{20D04FE0-3AEA-1069-A2D8-08002B30309D}' 0
        'Увімкнено'
    }
}

if ($Settings.DisableAdsTips) {
    Invoke-Setting 'Вимкнути рекламу та «рекомендовані» програми' {
        $cdm = 'Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager'
        foreach ($v in 'SilentInstalledAppsEnabled', 'SystemPaneSuggestionsEnabled', 'SoftLandingEnabled',
                       'PreInstalledAppsEnabled', 'OemPreInstalledAppsEnabled', 'SubscribedContent-338388Enabled',
                       'SubscribedContent-338389Enabled', 'SubscribedContent-353694Enabled', 'SubscribedContent-353696Enabled') {
            Set-UserReg $cdm $v 0
        }
        Set-UserReg 'Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced' 'ShowSyncProviderNotifications' 0
        Set-MachineReg 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent' 'DisableWindowsConsumerFeatures' 1
        'Вимкнено'
    }
}

if ($Settings.EnableRDP) {
    Invoke-Setting 'Віддалений робочий стіл (RDP)' {
        Set-MachineReg 'HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server' 'fDenyTSConnections' 0
        Set-MachineReg 'HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp' 'UserAuthentication' 1
        Enable-NetFirewallRule -Group '@FirewallAPI.dll,-28752' -ErrorAction Stop
        'Увімкнено (з NLA), правило брандмауера відкрито'
    }
}

if ($DefaultLoaded) {
    [GC]::Collect(); [GC]::WaitForPendingFinalizers(); Start-Sleep -Milliseconds 500
    & reg.exe unload 'HKU\PCSetupDefault' *> $null
}

if ($ComputerName -and $ComputerName -ne $env:COMPUTERNAME) {
    if ($ComputerName -match '^(?=.*[A-Za-z-])[A-Za-z0-9-]{1,15}$') {
        Invoke-Setting "Перейменування ПК на $ComputerName" {
            Rename-Computer -NewName $ComputerName -Force -ErrorAction Stop -WarningAction SilentlyContinue
            $script:NeedReboot = $true
            'Застосується після перезавантаження'
        }
    } else {
        Add-Result 'Налаштування' "Перейменування ПК на $ComputerName" 'WARN' 'Некоректне ім''я: до 15 символів, латиниця, цифри, дефіс'
    }
}

# =====================================================================
#  ЗВІТ
# =====================================================================
$os = Get-CimInstance Win32_OperatingSystem
$cv = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion'
$appRes = @($Results | Where-Object { $_.Group -eq 'Програми' })
$okApps = @($appRes | Where-Object { $_.Status -eq 'OK' }).Count
$fails  = @($Results | Where-Object { $_.Status -ne 'OK' -and $_.Status -ne 'INFO' })

$cards = [ordered]@{
    'Комп''ютер'       = $env:COMPUTERNAME
    'Windows'          = "$(([string]$os.Caption) -replace '^Microsoft\s+','') $($cv.DisplayVersion)"
    'Програми'         = "$okApps з $($appRes.Count) успішно"
    'Помилки/увага'    = $fails.Count
    'Перезавантаження' = $(if ($NeedReboot) { 'Потрібне' } else { 'Не потрібне' })
}

function ConvertTo-ResultRows($items) {
    foreach ($r in $items) {
        [pscustomobject]@{ 'Статус' = (New-Tag $r.Status); 'Що' = $r.Name; 'Результат' = $r.Detail }
    }
}

$body  = New-Cards $cards
$body += New-Section 'Програми' (ConvertTo-HtmlTable (ConvertTo-ResultRows $appRes) -RawColumns 'Статус' -Name 'programs')
$body += New-Section 'Налаштування Windows' (ConvertTo-HtmlTable (ConvertTo-ResultRows ($Results | Where-Object { $_.Group -ne 'Програми' })) -RawColumns 'Статус' -Name 'settings')
$note = 'Налаштування Провідника застосовано для поточного користувача, користувача за ПК та профілю за замовчуванням (усі нові користувачі). Вони набудуть чинності після повторного входу.'
if ($OtherConsoleUser) { $note += " Програми «для одного користувача» встановлено в профіль $ConsoleUser." }
$body += New-Section 'Примітки' "<p>$(Enc $note)</p><p class=""muted"">Журнал: <code>$(Enc $LogFile)</code></p>"

$html = New-HtmlPage "Налаштування ПК: $env:COMPUTERNAME" "Виконав: $($identity.Name) · $(Get-Date -Format 'dd.MM.yyyy HH:mm')" $body 'Setup-NewPC'
$reportPath = Save-Html $html "Setup_${env:COMPUTERNAME}_$Stamp.html"

Write-Host ''
Write-Host '========================================================' -ForegroundColor Cyan
Write-Host (" Готово. Програми: {0} з {1}. Помилок/зауважень: {2}" -f $okApps, $appRes.Count, $fails.Count) -ForegroundColor $(if ($fails.Count) { 'Yellow' } else { 'Green' })
Write-Host " Звіт: $reportPath"
Write-Host '========================================================' -ForegroundColor Cyan
try { Stop-Transcript | Out-Null } catch { }
if (-not $NoOpen) { Start-Process $reportPath }

if (-not $Auto) {
    if ($NeedReboot) {
        $a = Read-Host ' Потрібне перезавантаження. Перезавантажити зараз? (Y/N)'
        if ($a -match '^[yYнН]') { Restart-Computer -Force }
    } else {
        Read-Host ' Натисніть Enter, щоб закрити' | Out-Null
    }
}
