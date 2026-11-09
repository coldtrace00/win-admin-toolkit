#Requires -Version 5.1
<#
.SYNOPSIS
    Інвентаризація ПК: процесор, пам'ять, диски, Windows, серійний номер, встановлене ПЗ.

.DESCRIPTION
    Збирає дані про локальний комп'ютер і зберігає самодостатній HTML-звіт на робочий стіл.

    Працює без прав адміністратора, на Windows 10/11 з вбудованим PowerShell 5.1.

.EXAMPLE
    .\Get-PCInventory.ps1
    Звіт по цьому ПК на робочий стіл.

#>
[CmdletBinding()]
param(
    [switch]$NoOpen
)

$ErrorActionPreference = 'Continue'
$ScriptVersion = '1.0'
$Desktop  = [Environment]::GetFolderPath('Desktop')   # враховує перенаправлення в OneDrive
$Stamp    = Get-Date -Format 'yyyy-MM-dd_HH-mm'
$Utf8Bom  = New-Object System.Text.UTF8Encoding $true
$Inv      = [Globalization.CultureInfo]::InvariantCulture

# =====================================================================
#  ДОПОМІЖНІ ФУНКЦІЇ
# =====================================================================
function Enc([object]$v) {
    if ($null -eq $v) { return '' }
    [System.Net.WebUtility]::HtmlEncode([string]$v)
}

function Safe([scriptblock]$Block) {
    try { & $Block } catch { Write-Verbose $_.Exception.Message; $null }
}

function Format-Size([double]$Bytes) {
    if     ($Bytes -ge 1TB) { '{0:N2} ТБ' -f ($Bytes / 1TB) }
    elseif ($Bytes -ge 1GB) { '{0:N0} ГБ' -f ($Bytes / 1GB) }
    elseif ($Bytes -ge 1MB) { '{0:N0} МБ' -f ($Bytes / 1MB) }
    elseif ($Bytes -gt 0)   { '{0:N0} Б'  -f $Bytes }
    else { '' }
}

function Format-Date($dt) {
    if ($dt -is [datetime]) { $dt.ToString('dd.MM.yyyy') } else { '' }
}

function ConvertTo-HtmlTable {
    param(
        [object[]]$Rows,
        [string]$TableId = '',
        [string[]]$RawColumns = @()
    )
    $Rows = @($Rows | Where-Object { $_ })
    if ($Rows.Count -eq 0) { return '<p class="muted">Немає даних</p>' }

    $cols = $Rows[0].PSObject.Properties.Name
    $sb = New-Object System.Text.StringBuilder
    $idAttr = ''
    if ($TableId) { $idAttr = " id=""$TableId""" }
    [void]$sb.Append("<div class=""tbl""><table class=""sortable""$idAttr><thead><tr>")
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

function New-SearchBox([string]$TableId, [int]$Total) {
    "<input class=""search"" type=""search"" placeholder=""Пошук..."" oninput=""filterTable(this,'$TableId')"">" +
    " <span class=""muted"">Показано: <b id=""$TableId-count"">$Total</b> з $Total</span>"
}

function New-HtmlPage([string]$Title, [string]$Subtitle, [string]$Body) {
    $css = @'
:root{--bg:#f4f5f7;--card:#fff;--text:#1d2330;--muted:#6b7280;--line:#e5e7eb;--head:#f9fafb;--accent:#2563eb;--ok:#16a34a;--warn:#d97706;--bad:#dc2626}
@media (prefers-color-scheme:dark){:root{--bg:#0f1115;--card:#171a21;--text:#e5e7eb;--muted:#9ca3af;--line:#2a2f3a;--head:#1d212a;--accent:#60a5fa}}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--text);font:14px/1.5 "Segoe UI",system-ui,-apple-system,sans-serif}
.wrap{max-width:1200px;margin:0 auto;padding:24px 16px 40px}
header h1{margin:0;font-size:26px}
header p{margin:4px 0 0;color:var(--muted)}
.cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(210px,1fr));gap:12px;margin:20px 0 4px}
.card{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:12px 16px}
.card .k{color:var(--muted);font-size:12px;text-transform:uppercase;letter-spacing:.04em}
.card .v{font-size:16px;font-weight:600;margin-top:2px;word-break:break-word}
section{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:16px 18px;margin-top:16px}
section h2{margin:0 0 12px;font-size:17px}
section h3{margin:16px 0 8px;font-size:14px;color:var(--muted)}
.tbl{overflow-x:auto}
table{width:100%;border-collapse:collapse}
th,td{text-align:left;padding:7px 10px;border-bottom:1px solid var(--line);vertical-align:top}
table.sortable thead th{background:var(--head);font-weight:600;white-space:nowrap;cursor:pointer;user-select:none}
table.sortable thead th:hover{color:var(--accent)}
table.kv th{width:35%;color:var(--muted);font-weight:500}
tbody tr:hover{background:var(--head)}
.bar{display:inline-block;width:110px;height:8px;background:var(--line);border-radius:4px;overflow:hidden;vertical-align:middle;margin-right:6px}
.bar span{display:block;height:100%;background:var(--ok)}
.bar.warn span{background:var(--warn)}.bar.bad span{background:var(--bad)}
.tag{display:inline-block;padding:0 8px;border-radius:999px;font-size:12px;font-weight:600;color:#fff;background:var(--ok)}
.tag.warn{background:var(--warn)}.tag.bad{background:var(--bad)}
.search{width:100%;max-width:340px;padding:8px 10px;border:1px solid var(--line);border-radius:8px;background:var(--bg);color:var(--text);margin-bottom:10px;font:inherit}
.muted{color:var(--muted)}
a{color:var(--accent)}
details summary{cursor:pointer;font-weight:600}
details[open] summary{margin-bottom:4px}
footer{margin-top:24px;color:var(--muted);font-size:12px;text-align:center}
@media print{body{background:#fff;color:#000}.search{display:none}section,.card{break-inside:avoid;border-color:#ccc}}
'@
    $js = @'
function filterTable(input,id){var q=input.value.toLowerCase(),n=0;
document.querySelectorAll('#'+id+' tbody tr').forEach(function(r){var m=r.textContent.toLowerCase().indexOf(q)>-1;r.style.display=m?'':'none';if(m)n++;});
var c=document.getElementById(id+'-count');if(c)c.textContent=n;}
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
<footer>Згенеровано $gen · Get-PCInventory v$ScriptVersion</footer>
</div>
<script>$js</script>
</body>
</html>
"@
}

# =====================================================================
#  ЗБІР ДАНИХ
# =====================================================================
Write-Host 'Збираю дані про комп''ютер...' -ForegroundColor Cyan

$cs      = Safe { Get-CimInstance Win32_ComputerSystem }
$os      = Safe { Get-CimInstance Win32_OperatingSystem }
$bios    = Safe { Get-CimInstance Win32_BIOS }
$board   = Safe { Get-CimInstance Win32_BaseBoard }
$cpu     = @(Safe { Get-CimInstance Win32_Processor })
$mem     = @(Safe { Get-CimInstance Win32_PhysicalMemory })
$memArr  = @(Safe { Get-CimInstance Win32_PhysicalMemoryArray })
$disks   = @(Safe { Get-CimInstance Win32_DiskDrive } | Sort-Object Index)
$vols    = @(Safe { Get-CimInstance Win32_LogicalDisk -Filter 'DriveType=3' })
$gpus    = @(Safe { Get-CimInstance Win32_VideoController })
$nets    = @(Safe { Get-CimInstance Win32_NetworkAdapterConfiguration -Filter 'IPEnabled=True' })
$cv      = Safe { Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' }

$pcName  = $env:COMPUTERNAME
$user    = if ($cs.UserName) { $cs.UserName } else { "$env:USERDOMAIN\$env:USERNAME" }

# --- Windows ---
$winVer  = if ($cv.DisplayVersion) { $cv.DisplayVersion } else { $cv.ReleaseId }
$winName = ([string]$os.Caption) -replace '^Microsoft\s+', ''
$build   = if ($cv.UBR -ne $null) { "$($os.BuildNumber).$($cv.UBR)" } else { [string]$os.BuildNumber }
$uptime  = ''
if ($os.LastBootUpTime) {
    $up = (Get-Date) - $os.LastBootUpTime
    $uptime = '{0} д {1} год {2} хв' -f $up.Days, $up.Hours, $up.Minutes
}

# --- серійні номери ---
$junk = @('', '0', 'None', 'Default string', 'To be filled by O.E.M.', 'System Serial Number', 'Not Specified', 'Not Applicable')
$serialBios  = ([string]$bios.SerialNumber).Trim()
$serialBoard = ([string]$board.SerialNumber).Trim()
$serialMain  = if ($junk -notcontains $serialBios) { $serialBios }
               elseif ($junk -notcontains $serialBoard) { "$serialBoard (плата)" }
               else { 'не задано виробником' }

# --- процесор ---
$cpuRows = foreach ($c in $cpu) {
    [pscustomobject]@{
        'Модель'          = (([string]$c.Name).Trim() -replace '\s+', ' ')
        'Ядра'            = $c.NumberOfCores
        'Потоки'          = $c.NumberOfLogicalProcessors
        'Базова частота'  = ('{0:N2} ГГц' -f ($c.MaxClockSpeed / 1000))
        'Сокет'           = $c.SocketDesignation
        'Віртуалізація'   = $(if ($c.VirtualizationFirmwareEnabled) { 'Увімкнена' } else { 'Вимкнена / н/д' })
    }
}
$cpuName  = ($cpuRows | ForEach-Object { $_.'Модель' } | Select-Object -Unique) -join ' + '
$cpuCores = ($cpu | Measure-Object NumberOfCores -Sum).Sum
$cpuThr   = ($cpu | Measure-Object NumberOfLogicalProcessors -Sum).Sum

# --- пам'ять ---
$ramBytes = ($mem | Measure-Object Capacity -Sum).Sum
if (-not $ramBytes) { $ramBytes = $cs.TotalPhysicalMemory }
$ramGB      = [math]::Round($ramBytes / 1GB)
$slotsTotal = ($memArr | Measure-Object MemoryDevices -Sum).Sum
$slotsUsed  = $mem.Count
$memRows = foreach ($m in $mem) {
    [pscustomobject]@{
        'Слот'          = $m.DeviceLocator
        'Обсяг'         = (Format-Size $m.Capacity)
        'Частота, МГц'  = $(if ($m.ConfiguredClockSpeed) { $m.ConfiguredClockSpeed } else { $m.Speed })
        'Виробник'      = ([string]$m.Manufacturer).Trim()
        'Part Number'   = ([string]$m.PartNumber).Trim()
    }
}
$ramText = "$ramGB ГБ"
if ($slotsTotal) { $ramText += " · слотів $slotsUsed з $slotsTotal" }

# --- фізичні диски ---
$phys = @{}
Safe { Get-PhysicalDisk -ErrorAction Stop | ForEach-Object { $phys[[string]$_.DeviceId] = $_ } } | Out-Null
$diskRows = foreach ($d in $disks) {
    $p = $phys[[string]$d.Index]
    $type = 'Невідомо'
    $bus  = [string]$d.InterfaceType
    $health = ''
    if ($p) {
        $bus = [string]$p.BusType
        $health = [string]$p.HealthStatus
        if     ($bus -eq 'NVMe')               { $type = 'NVMe SSD' }
        elseif ([string]$p.MediaType -eq 'SSD') { $type = 'SSD' }
        elseif ([string]$p.MediaType -eq 'HDD') { $type = 'HDD' }
    }
    [pscustomobject]@{
        'Модель'          = ([string]$d.Model).Trim()
        'Тип'             = $type
        'Інтерфейс'       = $bus
        'Обсяг'           = (Format-Size $d.Size)
        'Стан'            = $health
        'Серійний номер'  = ([string]$d.SerialNumber).Trim()
    }
}
$diskSummary = ($diskRows | ForEach-Object { "$($_.'Обсяг') $($_.'Тип')" }) -join ' + '

# --- розділи ---
$minFree = $null
$volRows = foreach ($v in ($vols | Sort-Object DeviceID)) {
    $pct = 0
    if ($v.Size) { $pct = [math]::Round(($v.Size - $v.FreeSpace) / $v.Size * 100) }
    $free = 100 - $pct
    if ($null -eq $minFree -or $free -lt $minFree) { $minFree = $free }
    $cls = 'ok'
    if ($pct -ge 90) { $cls = 'bad' } elseif ($pct -ge 75) { $cls = 'warn' }
    [pscustomobject]@{
        'Диск'       = $v.DeviceID
        'Мітка'      = $v.VolumeName
        'ФС'         = $v.FileSystem
        'Обсяг'      = (Format-Size $v.Size)
        'Вільно'     = (Format-Size $v.FreeSpace)
        'Заповнено'  = "<span class=""bar $cls""><span style=""width:$pct%""></span></span>$pct%"
    }
}
$volSummary = ($vols | Sort-Object DeviceID | ForEach-Object {
    '{0} {1} (вільно {2})' -f $_.DeviceID, (Format-Size $_.Size), (Format-Size $_.FreeSpace)
}) -join '; '

# --- відео ---
$gpuRows = foreach ($g in $gpus) {
    $res = ''
    if ($g.CurrentHorizontalResolution) { $res = "$($g.CurrentHorizontalResolution)×$($g.CurrentVerticalResolution)" }
    [pscustomobject]@{
        'Відеоадаптер'       = $g.Name
        'Версія драйвера'    = $g.DriverVersion
        'Дата драйвера'      = (Format-Date $g.DriverDate)
        'Роздільна здатність' = $res
    }
}

# --- мережа ---
$netRows = foreach ($n in ($nets | Sort-Object { -not $_.DefaultIPGateway })) {
    [pscustomobject]@{
        'Адаптер' = $n.Description
        'IPv4'    = (@($n.IPAddress) | Where-Object { $_ -match '^\d+\.\d+\.\d+\.\d+$' }) -join ', '
        'MAC'     = $n.MACAddress
        'Шлюз'    = (@($n.DefaultIPGateway) | Where-Object { $_ }) -join ', '
        'DNS'     = (@($n.DNSServerSearchOrder) | Where-Object { $_ }) -join ', '
        'DHCP'    = $(if ($n.DHCPEnabled) { 'Так' } else { 'Ні' })
    }
}
$mainNet = $netRows | Select-Object -First 1

# --- встановлене ПЗ ---
Write-Host 'Збираю список програм...' -ForegroundColor Cyan
$regPaths = @(
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*',
    'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*'
)
$software = @(Get-ItemProperty $regPaths -ErrorAction SilentlyContinue |
    Where-Object {
        $_.DisplayName -and
        $_.SystemComponent -ne 1 -and
        -not $_.ParentKeyName -and
        $_.ReleaseType -notin @('Update', 'Hotfix', 'Security Update') -and
        $_.DisplayName -notmatch '^(Security Update|Update for|Hotfix for)'
    } |
    ForEach-Object {
        $d = [string]$_.InstallDate
        if ($d -match '^\d{8}$') {
            try { $d = [datetime]::ParseExact($d, 'yyyyMMdd', $Inv).ToString('dd.MM.yyyy') } catch { }
        }
        [pscustomobject]@{
            'Програма'          = ([string]$_.DisplayName).Trim()
            'Версія'            = ([string]$_.DisplayVersion).Trim()
            'Видавець'          = ([string]$_.Publisher).Trim()
            'Дата встановлення' = $d
        }
    } |
    Sort-Object 'Програма', 'Версія' -Unique)

# =====================================================================
#  HTML
# =====================================================================
$winFull = "$winName $winVer".Trim()
$model   = "$(([string]$cs.Manufacturer).Trim()) $(([string]$cs.Model).Trim())".Trim()

$cards = [ordered]@{
    'Модель'          = $model
    'Процесор'        = $cpuName
    'Оперативна пам''ять' = $ramText
    'Диски'           = $diskSummary
    'Windows'         = "$winFull ($build)"
    'Серійний номер'  = $serialMain
}

$system = [ordered]@{
    'Ім''я комп''ютера'      = $pcName
    'Домен / робоча група'   = $(if ($cs.PartOfDomain) { "$($cs.Domain) (домен)" } else { "$($cs.Workgroup) (робоча група)" })
    'Користувач'             = $user
    'Виробник'               = $cs.Manufacturer
    'Модель'                 = $cs.Model
    'Серійний номер (BIOS)'  = $serialBios
    'Материнська плата'      = "$($board.Manufacturer) $($board.Product)".Trim()
    'Серійний номер плати'   = $serialBoard
    'BIOS'                   = "$($bios.SMBIOSBIOSVersion) від $(Format-Date $bios.ReleaseDate)"
    'Windows'                = $winFull
    'Збірка'                 = $build
    'Розрядність'            = $os.OSArchitecture
    'ОС встановлена'         = (Format-Date $os.InstallDate)
    'Останнє завантаження'   = $(if ($os.LastBootUpTime) { $os.LastBootUpTime.ToString('dd.MM.yyyy HH:mm') } else { '' })
    'Час роботи'             = $uptime
}

$body  = New-Cards $cards
$body += '<section><h2>Система</h2>' + (ConvertTo-KvTable $system) + '</section>'
$body += '<section><h2>Процесор</h2>' + (ConvertTo-HtmlTable $cpuRows) + '</section>'
$body += "<section><h2>Оперативна пам'ять — $(Enc $ramText)</h2>" + (ConvertTo-HtmlTable $memRows) + '</section>'
$body += '<section><h2>Диски</h2><h3>Фізичні накопичувачі</h3>' + (ConvertTo-HtmlTable $diskRows) +
         '<h3>Розділи</h3>' + (ConvertTo-HtmlTable $volRows -RawColumns @('Заповнено')) + '</section>'
$body += '<section><h2>Відеоадаптери</h2>' + (ConvertTo-HtmlTable $gpuRows) + '</section>'
$body += '<section><h2>Мережа</h2>' + (ConvertTo-HtmlTable $netRows) + '</section>'
$body += "<section><h2>Встановлене ПЗ — $($software.Count)</h2>" + (New-SearchBox 'sw' $software.Count) +
         (ConvertTo-HtmlTable $software 'sw') + '</section>'

$html = New-HtmlPage "Інвентаризація: $pcName" "$model · $user · $(Get-Date -Format 'dd.MM.yyyy HH:mm')" $body

$htmlPath = Join-Path $Desktop "Inventory_${pcName}_$Stamp.html"
[System.IO.File]::WriteAllText($htmlPath, $html, $Utf8Bom)
Write-Host "Звіт збережено: $htmlPath" -ForegroundColor Green

if (-not $NoOpen) { Start-Process $htmlPath }
