#Requires -Version 5.1
<#
.SYNOPSIS
    Аптечка ПК: за хвилину показує, чому комп'ютер гальмує або глючить, і що з цим робити.

.DESCRIPTION
    Перевіряє: час без перезавантаження, оновлення, що чекають перезавантаження, місце і стан дисків (SMART),
    HDD під системою, пам'ять і процесор (хто їх їсть), сині екрани, раптові вимкнення, збої програм,
    помилки диска в журналі, давність оновлень Windows, антивірус, пристрої без драйверів, важливі служби,
    автозавантаження, тимчасові файли, мережу, батарею ноутбука, схему живлення.

    Три способи запуску:
      - на самому ПК (Enter на запит імені);
      - віддалено з вашого ПК: ввести ім'я ПК (потрібні домен і WinRM);
      - у RDP-сеансі на сервері.

    Без прав адміністратора працює, але частина перевірок (SMART-лічильники, знос батареї) буде пропущена,
    тому скрипт пропонує підняти права.

.EXAMPLE
    .\Test-PCHealth.ps1
.EXAMPLE
    .\Test-PCHealth.ps1 -ComputerName BUH-03
#>
[CmdletBinding()]
param(
    [string]$ComputerName,
    [switch]$Local,
    [switch]$NoElevate,
    [switch]$NoOpen
)

$ErrorActionPreference = 'Continue'
$ProgressPreference    = 'SilentlyContinue'
try { $Host.UI.RawUI.WindowTitle = 'Аптечка ПК' } catch { }

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

# =====================================================================
#  ЗБІР ДАНИХ (виконується на перевірюваному ПК - локально або через WinRM)
# =====================================================================
$Collector = {
    $ErrorActionPreference = 'SilentlyContinue'
    $ProgressPreference = 'SilentlyContinue'
    $checks = New-Object System.Collections.Generic.List[object]
    function Add-C($S, $Area, $Name, $Detail, $Fix) {
        $checks.Add([pscustomobject]@{ Status = $S; Area = $Area; Name = $Name; Detail = [string]$Detail; Fix = [string]$Fix })
    }
    function Get-Ev($Log, $Ids, $Days, $Max = 300) {
        try { @(Get-WinEvent -FilterHashtable @{ LogName = $Log; Id = $Ids; StartTime = (Get-Date).AddDays(-$Days) } -MaxEvents $Max -ErrorAction Stop) } catch { @() }
    }
    function Get-SizeMB($Path) {
        if (-not (Test-Path $Path)) { return 0 }
        [math]::Round(((Get-ChildItem $Path -Recurse -File -Force -ErrorAction SilentlyContinue | Measure-Object Length -Sum).Sum) / 1MB)
    }
    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    $os  = Get-CimInstance Win32_OperatingSystem
    $cs  = Get-CimInstance Win32_ComputerSystem
    $cpu = @(Get-CimInstance Win32_Processor)
    $cv  = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion'
    $cores = [int](($cpu | Measure-Object NumberOfLogicalProcessors -Sum).Sum); if ($cores -lt 1) { $cores = 1 }
    $isLaptop = [bool](Get-CimInstance Win32_Battery)

    # --- перезавантаження ---
    $upDays = $null
    if ($os.LastBootUpTime) {
        $up = (Get-Date) - $os.LastBootUpTime
        $upDays = [math]::Floor($up.TotalDays)
        $upText = '{0} дн. {1} год.' -f $upDays, $up.Hours
        $fix = 'Перезавантажте ПК. «Завершення роботи» при ввімкненому швидкому запуску не очищує систему - потрібне саме «Перезавантаження».'
        if ($upDays -ge 30)     { Add-C 'FAIL' 'Система' "Без перезавантаження $upText" 'Накопичуються витоки пам''яті, зависають служби, не застосовуються оновлення.' $fix }
        elseif ($upDays -ge 14) { Add-C 'WARN' 'Система' "Без перезавантаження $upText" 'Довга робота без перезавантаження - часта причина гальм і дивних збоїв.' $fix }
        else                    { Add-C 'OK'   'Система' "Без перезавантаження $upText" '' '' }
    }
    $pend = @()
    if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending') { $pend += 'компоненти Windows' }
    if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired') { $pend += 'оновлення Windows' }
    $cn1 = (Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\ComputerName\ActiveComputerName').ComputerName
    $cn2 = (Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\ComputerName\ComputerName').ComputerName
    if ($cn1 -and $cn2 -and $cn1 -ne $cn2) { $pend += 'перейменування ПК' }
    if ($pend.Count) { Add-C 'WARN' 'Система' 'Чекає перезавантаження' ("Не завершено: " + ($pend -join ', ')) 'Перезавантажте ПК, щоб завершити встановлення.' }

    # --- пам'ять ---
    $ramGB = [math]::Round($os.TotalVisibleMemorySize / 1MB, 1)
    $ramUsed = $null
    if ($os.TotalVisibleMemorySize) { $ramUsed = [math]::Round(100 - $os.FreePhysicalMemory / $os.TotalVisibleMemorySize * 100) }
    if ($ramGB -gt 0 -and $ramGB -lt 3.8)      { Add-C 'FAIL' 'Пам''ять' "Встановлено лише $ramGB ГБ" 'Для Windows 10/11 з браузером і офісом цього критично мало.' 'Додайте пам''ять: мінімум 8 ГБ, краще 16 ГБ.' }
    elseif ($ramGB -gt 0 -and $ramGB -lt 7.5)  { Add-C 'WARN' 'Пам''ять' "Встановлено $ramGB ГБ" 'Сучасний браузер з десятком вкладок займає 2-4 ГБ.' 'Додайте пам''ять до 8-16 ГБ - це дешево і помітно прискорює роботу.' }
    if ($null -ne $ramUsed) {
        if ($ramUsed -ge 90)     { Add-C 'FAIL' 'Пам''ять' "Зайнято $ramUsed% з $ramGB ГБ" 'Пам''яті не вистачає, Windows вивантажує дані на диск - звідси гальма.' 'Подивіться таблицю «Хто їсть пам''ять» і закрийте зайве; якщо так постійно - додайте пам''ять.' }
        elseif ($ramUsed -ge 80) { Add-C 'WARN' 'Пам''ять' "Зайнято $ramUsed% з $ramGB ГБ" 'Запас пам''яті малий.' 'Перевірте таблицю «Хто їсть пам''ять».' }
        else                     { Add-C 'OK'   'Пам''ять' "Зайнято $ramUsed% з $ramGB ГБ" '' '' }
    }
    $topMem = @(Get-Process | Group-Object ProcessName | ForEach-Object {
        [pscustomobject]@{ Name = $_.Name; Count = $_.Count; MB = [math]::Round((($_.Group | Measure-Object WorkingSet64 -Sum).Sum) / 1MB); Desc = ($_.Group | Select-Object -First 1).Description }
    } | Sort-Object MB -Descending | Select-Object -First 8)

    # --- процесор ---
    $loads = @()
    for ($i = 0; $i -lt 3; $i++) {
        $p = Get-CimInstance Win32_PerfFormattedData_PerfOS_Processor -Filter "Name='_Total'"
        if ($p) { $loads += [int]$p.PercentProcessorTime }
        Start-Sleep -Milliseconds 700
    }
    $cpuLoad = $null
    if ($loads.Count) { $cpuLoad = [math]::Round(($loads | Measure-Object -Average).Average) }
    $topCpu = @(Get-CimInstance Win32_PerfFormattedData_PerfProc_Process | Where-Object { $_.Name -notin '_Total', 'Idle' } |
        ForEach-Object { [pscustomobject]@{ Name = ($_.Name -replace '#\d+$', ''); Pct = [double]$_.PercentProcessorTime / $cores } } |
        Group-Object Name | ForEach-Object { [pscustomobject]@{ Name = $_.Name; Pct = [math]::Round(($_.Group | Measure-Object Pct -Sum).Sum, 1) } } |
        Where-Object { $_.Pct -gt 0 } | Sort-Object Pct -Descending | Select-Object -First 6)
    if ($null -ne $cpuLoad) {
        $who = ''; if ($topCpu.Count) { $who = " Найбільше: $($topCpu[0].Name) ($($topCpu[0].Pct)%)." }
        if ($cpuLoad -ge 85)     { Add-C 'FAIL' 'Процесор' "Навантаження $cpuLoad%" "Процесор завантажений майже повністю.$who" 'Перевірте процес-лідер: зависла програма, оновлення Windows, сканування антивірусом або майнер.' }
        elseif ($cpuLoad -ge 60) { Add-C 'WARN' 'Процесор' "Навантаження $cpuLoad%" "Процесор помітно зайнятий.$who" 'Якщо триває довго - перевірте процес-лідер.' }
        else                     { Add-C 'OK'   'Процесор' "Навантаження $cpuLoad%" '' '' }
    }

    # --- місце на дисках ---
    $disks = @()
    foreach ($d in @(Get-CimInstance Win32_LogicalDisk -Filter 'DriveType=3')) {
        if (-not $d.Size) { continue }
        $freeGB = [math]::Round($d.FreeSpace / 1GB, 1); $sizeGB = [math]::Round($d.Size / 1GB, 1)
        $freePct = [math]::Round($d.FreeSpace / $d.Size * 100)
        $disks += [pscustomobject]@{ Drive = $d.DeviceID; Label = $d.VolumeName; SizeGB = $sizeGB; FreeGB = $freeGB; FreePct = $freePct }
        $isSys = $d.DeviceID -eq $env:SystemDrive
        $fix = 'Очистіть тимчасові файли (аптечка запропонує в кінці), «Очищення диска» > «Очистити системні файли», перенесіть великі файли на інший диск.'
        if (($freePct -lt 10) -or ($isSys -and $freeGB -lt 5))       { Add-C 'FAIL' 'Диски' "Диск $($d.DeviceID) вільно $freeGB ГБ ($freePct%)" 'Коли диск майже повний, Windows гальмує, не ставить оновлення, програми падають.' $fix }
        elseif (($freePct -lt 20) -or ($isSys -and $freeGB -lt 15))  { Add-C 'WARN' 'Диски' "Диск $($d.DeviceID) вільно $freeGB ГБ ($freePct%)" 'Місце закінчується.' $fix }
        else                                                          { Add-C 'OK'   'Диски' "Диск $($d.DeviceID) вільно $freeGB ГБ ($freePct%)" '' '' }
    }

    # --- стан дисків ---
    $phys = @()
    $sysDiskNum = $null
    try { $sysDiskNum = (Get-Partition -DriveLetter ($env:SystemDrive.TrimEnd(':')) -ErrorAction Stop | Select-Object -First 1).DiskNumber } catch { }
    foreach ($p in @(Get-PhysicalDisk)) {
        $type = 'Невідомо'
        if ([string]$p.BusType -eq 'NVMe') { $type = 'NVMe SSD' } elseif ([string]$p.MediaType -eq 'SSD') { $type = 'SSD' } elseif ([string]$p.MediaType -eq 'HDD') { $type = 'HDD' }
        $rel = $null
        if ($isAdmin) { $rel = $p | Get-StorageReliabilityCounter }
        $phys += [pscustomobject]@{
            Model = ([string]$p.FriendlyName).Trim(); Type = $type; SizeGB = [math]::Round($p.Size / 1GB)
            Health = [string]$p.HealthStatus; Wear = $rel.Wear; Temp = $rel.Temperature
            ReadErr = $rel.ReadErrorsUncorrected; Hours = $rel.PowerOnHours; IsSystem = ($null -ne $sysDiskNum -and [string]$p.DeviceId -eq [string]$sysDiskNum)
        }
        $nm = "$($p.FriendlyName) ($type)"
        $backup = 'Негайно зробіть резервну копію даних і плануйте заміну диска.'
        if ([string]$p.HealthStatus -eq 'Unhealthy')    { Add-C 'FAIL' 'Диски' "Диск несправний: $nm" 'Windows позначила диск як несправний.' $backup }
        elseif ([string]$p.HealthStatus -eq 'Warning')  { Add-C 'WARN' 'Диски' "Попередження про стан: $nm" 'Диск повідомляє про проблеми.' $backup }
        if ($null -ne $rel.Wear -and $rel.Wear -ge 90)  { Add-C 'FAIL' 'Диски' "SSD зношений на $($rel.Wear)%: $nm" 'Ресурс запису майже вичерпано.' $backup }
        elseif ($null -ne $rel.Wear -and $rel.Wear -ge 70) { Add-C 'WARN' 'Диски' "SSD зношений на $($rel.Wear)%: $nm" 'Ресурс запису добігає кінця.' 'Плануйте заміну, тримайте свіжу резервну копію.' }
        if ($rel.ReadErrorsUncorrected -gt 0)           { Add-C 'WARN' 'Диски' "Невиправлені помилки читання: $($rel.ReadErrorsUncorrected), $nm" 'Диск не зміг прочитати частину даних.' $backup }
        if ($type -eq 'HDD' -and $phys[-1].IsSystem)   { Add-C 'WARN' 'Диски' 'Windows стоїть на HDD' 'Механічний диск - найчастіша причина повільного завантаження і «все думає».' 'Замініть системний диск на SSD: найдешевший спосіб прискорити старий ПК у 3-5 разів.' }
    }
    if ($isAdmin) {
        $pred = @(Get-CimInstance -Namespace root\wmi -ClassName MSStorageDriver_FailurePredictStatus | Where-Object { $_.PredictFailure })
        if ($pred.Count) { Add-C 'FAIL' 'Диски' 'SMART передбачає відмову диска' 'Диск сам повідомляє, що скоро вийде з ладу.' 'Терміново скопіюйте дані й замініть диск.' }
    }
    $healthyDisks = @($phys | Where-Object { $_.Health -eq 'Healthy' }).Count
    if ($phys.Count -and $healthyDisks -eq $phys.Count) { Add-C 'OK' 'Диски' "Стан дисків: справні ($($phys.Count))" '' '' }

    # --- журнали: сині екрани, вимкнення, збої ---
    $bsod = @(Get-Ev 'System' 1001 30 | Where-Object { $_.ProviderName -match 'WER-SystemErrorReporting|BugCheck' })
    $crashes = @($bsod | Select-Object -First 10 | ForEach-Object {
        $code = ''; if ($_.Message -match '0x[0-9a-fA-F]{8}') { $code = $Matches[0] }
        [pscustomobject]@{ When = $_.TimeCreated; Kind = 'Синій екран (BSOD)'; Detail = $code }
    })
    if ($bsod.Count) { Add-C 'FAIL' 'Стабільність' "Сині екрани за 30 днів: $($bsod.Count)" 'Система аварійно перезавантажувалась.' 'Найчастіше винен драйвер (відео, мережа), перегрів або пам''ять. Оновіть драйвери; якщо повторюється - перевірте пам''ять (mdsched.exe).' }
    else { Add-C 'OK' 'Стабільність' 'Синіх екранів за 30 днів немає' '' '' }

    $kp = @(Get-Ev 'System' 41 30 | Where-Object { $_.ProviderName -match 'Kernel-Power' })
    $crashes += @($kp | Select-Object -First 10 | ForEach-Object { [pscustomobject]@{ When = $_.TimeCreated; Kind = 'Раптове вимкнення'; Detail = 'Живлення зникло, зависання або утримання кнопки' } })
    if ($kp.Count -ge 3)    { Add-C 'WARN' 'Стабільність' "Раптових вимкнень за 30 днів: $($kp.Count)" 'ПК вимикався без завершення роботи: відключення світла, зависання або кнопка.' 'Якщо це відключення світла - поставте ДБЖ; якщо зависання - дивіться сині екрани та температуру.' }
    elseif ($kp.Count -gt 0) { Add-C 'OK' 'Стабільність' "Раптових вимкнень за 30 днів: $($kp.Count)" '' '' }

    $appErr = @(Get-Ev 'Application' @(1000, 1002) 14 | Where-Object { $_.ProviderName -match 'Application (Error|Hang)' })
    $appCrashes = @($appErr | ForEach-Object {
        $n = ''; try { $n = [string]$_.Properties[0].Value } catch { }
        [pscustomobject]@{ App = $n; Kind = $(if ($_.Id -eq 1002) { 'Зависання' } else { 'Аварійне закриття' }); When = $_.TimeCreated }
    } | Where-Object { $_.App } | Group-Object App, Kind | ForEach-Object {
        [pscustomobject]@{ App = $_.Group[0].App; Kind = $_.Group[0].Kind; Count = $_.Count; Last = ($_.Group | Sort-Object When -Descending | Select-Object -First 1).When }
    } | Sort-Object Count -Descending | Select-Object -First 10)
    $badApp = @($appCrashes | Where-Object { $_.Count -ge 5 })
    if ($badApp.Count) { Add-C 'WARN' 'Стабільність' "Програми часто падають: $(($badApp | ForEach-Object { "$($_.App) ($($_.Count))" }) -join ', ')" 'За 14 днів програма закривалась із помилкою або зависала 5+ разів.' 'Оновіть або перевстановіть програму; для Office - «Відновлення» в «Програми та компоненти».' }

    $diskEv = @(Get-Ev 'System' @(7, 51, 55, 129, 153) 30 | Where-Object { $_.ProviderName -match 'disk|Ntfs|stor|nvme|iaStor' })
    $crashes += @($diskEv | Select-Object -First 10 | ForEach-Object { [pscustomobject]@{ When = $_.TimeCreated; Kind = "Помилка диска (подія $($_.Id))"; Detail = ([string]$_.Message -split "`n")[0] } })
    if (@($diskEv | Where-Object { $_.Id -in 7, 55 }).Count) { Add-C 'FAIL' 'Диски' "Помилки диска в журналі: $($diskEv.Count) за 30 днів" 'Пошкоджені сектори або пошкоджена файлова система.' 'Зробіть резервну копію; запустіть chkdsk C: /scan; якщо помилки тривають - замініть диск.' }
    elseif ($diskEv.Count) { Add-C 'WARN' 'Диски' "Збої обміну з диском: $($diskEv.Count) за 30 днів" 'Диск або кабель/контролер не відповідали вчасно.' 'Перевірте кабель SATA, оновіть драйвер контролера; слідкуйте за станом диска.' }

    # --- оновлення Windows ---
    $lastUpd = $null
    try {
        $sess = New-Object -ComObject Microsoft.Update.Session
        $srch = $sess.CreateUpdateSearcher()
        $cnt = $srch.GetTotalHistoryCount()
        if ($cnt -gt 0) { $lastUpd = ($srch.QueryHistory(0, [math]::Min($cnt, 100)) | Where-Object { $_.ResultCode -eq 2 } | Sort-Object Date -Descending | Select-Object -First 1).Date }
    } catch { }
    if (-not $lastUpd) { $lastUpd = (Get-HotFix | Where-Object { $_.InstalledOn } | Sort-Object InstalledOn -Descending | Select-Object -First 1).InstalledOn }
    if ($lastUpd) {
        $ud = [math]::Floor(((Get-Date) - [datetime]$lastUpd).TotalDays)
        if ($ud -gt 60) { Add-C 'WARN' 'Безпека' "Останнє оновлення Windows $ud дн. тому" 'ПК не отримує виправлень безпеки.' 'Параметри > Windows Update > Перевірити наявність оновлень; якщо помилка - перевірте місце на диску та службу оновлення.' }
        else { Add-C 'OK' 'Безпека' "Останнє оновлення Windows $ud дн. тому" '' '' }
    }

    # --- антивірус ---
    $avList = @()
    foreach ($a in @(Get-CimInstance -Namespace root/SecurityCenter2 -ClassName AntiVirusProduct)) {
        $hex = '{0:X6}' -f [int]$a.productState
        $avList += [pscustomobject]@{ Name = $a.displayName; On = ($hex.Substring(2, 2) -in '10', '11'); Fresh = ($hex.Substring(4, 2) -eq '00') }
    }
    if ($avList.Count) {
        $on = @($avList | Where-Object { $_.On })
        if ($on.Count -eq 0) { Add-C 'FAIL' 'Безпека' 'Антивірус вимкнено' ("Встановлено: " + (($avList | ForEach-Object { $_.Name }) -join ', ')) 'Увімкніть захист у реальному часі; якщо закінчилась ліцензія - продовжіть або поверніть Microsoft Defender.' }
        elseif (@($on | Where-Object { -not $_.Fresh }).Count) { Add-C 'WARN' 'Безпека' "Застарілі бази антивіруса: $(($on | ForEach-Object { $_.Name }) -join ', ')" 'Антивірус не знає свіжих загроз.' 'Оновіть бази вручну і перевірте, чи є інтернет і ліцензія.' }
        else { Add-C 'OK' 'Безпека' "Антивірус працює: $(($on | ForEach-Object { $_.Name }) -join ', ')" '' '' }
    } else {
        $mp = Get-MpComputerStatus
        if ($mp) {
            if (-not $mp.RealTimeProtectionEnabled) { Add-C 'FAIL' 'Безпека' 'Microsoft Defender: захист у реальному часі вимкнено' '' 'Увімкніть захист у «Безпека Windows».' }
            elseif ($mp.AntivirusSignatureAge -gt 7) { Add-C 'WARN' 'Безпека' "Бази Defender застарілі ($($mp.AntivirusSignatureAge) дн.)" '' 'Оновіть бази: «Безпека Windows» > «Оновлення захисту».' }
            else { Add-C 'OK' 'Безпека' 'Microsoft Defender працює' '' '' }
        }
    }

    # --- пристрої з помилками ---
    $codeText = @{ 1 = 'не налаштовано'; 3 = 'драйвер пошкоджено'; 10 = 'не запускається'; 12 = 'конфлікт ресурсів'; 18 = 'перевстановіть драйвер'; 19 = 'помилка реєстру'; 21 = 'видаляється'; 24 = 'не знайдено/проблема'; 28 = 'немає драйвера'; 31 = 'драйвер не завантажився'; 39 = 'драйвер пошкоджено'; 43 = 'пристрій повідомив про збій'; 52 = 'драйвер без підпису' }
    $badDev = @(Get-CimInstance Win32_PnPEntity | Where-Object { $_.ConfigManagerErrorCode -and $_.ConfigManagerErrorCode -notin 22, 45 } | ForEach-Object {
        $c = [int]$_.ConfigManagerErrorCode
        $t = $codeText[$c]; if (-not $t) { $t = "код $c" }
        [pscustomobject]@{ Name = $(if ($_.Name) { $_.Name } else { $_.DeviceID }); Problem = $t }
    })
    if ($badDev.Count) { Add-C 'WARN' 'Обладнання' "Пристрої з помилками: $($badDev.Count)" (($badDev | Select-Object -First 3 | ForEach-Object { "$($_.Name) - $($_.Problem)" }) -join '; ') 'Встановіть драйвер із сайту виробника ноутбука/материнської плати або через Windows Update (Додаткові оновлення).' }
    else { Add-C 'OK' 'Обладнання' 'Усі пристрої працюють' '' '' }

    # --- важливі служби ---
    $critical = [ordered]@{ Spooler = 'Друк'; Dnscache = 'DNS-клієнт'; Dhcp = 'DHCP-клієнт'; LanmanWorkstation = 'Мережеві папки'; EventLog = 'Журнал подій'; Winmgmt = 'WMI'; BFE = 'Фільтрація (брандмауер)'; mpssvc = 'Брандмауер'; Audiosrv = 'Звук'; CryptSvc = 'Криптографія'; nsi = 'Мережа (NSI)'; Schedule = 'Планувальник' }
    $svcRows = @()
    foreach ($k in $critical.Keys) {
        $s = Get-CimInstance Win32_Service -Filter "Name='$k'"
        if (-not $s) { continue }
        $svcRows += [pscustomobject]@{ Name = $k; Title = $critical[$k]; State = [string]$s.State; Mode = [string]$s.StartMode }
        if ($s.State -ne 'Running' -and $s.StartMode -ne 'Disabled') {
            Add-C 'WARN' 'Служби' "Не працює служба «$($critical[$k])» ($k)" "Стан: $($s.State), запуск: $($s.StartMode)." "Запустіть службу (services.msc > $k > Запустити). Якщо падає знову - дивіться журнал подій."
        }
    }

    # --- автозавантаження ---
    $startup = @(Get-CimInstance Win32_StartupCommand | ForEach-Object { [pscustomobject]@{ Name = $_.Name; Command = $_.Command; Where = $_.Location; User = $_.User } })
    if ($startup.Count -gt 12) { Add-C 'WARN' 'Швидкодія' "Програм в автозавантаженні: $($startup.Count)" 'Кожна з них уповільнює старт і їсть пам''ять у фоні.' 'Диспетчер завдань > Автозавантаження: вимкніть усе, що не потрібно одразу після входу.' }
    else { Add-C 'OK' 'Швидкодія' "Програм в автозавантаженні: $($startup.Count)" '' '' }

    # --- тимчасові файли ---
    $tempUser = Get-SizeMB $env:TEMP
    $tempWin  = Get-SizeMB (Join-Path $env:SystemRoot 'Temp')
    $tempWu   = Get-SizeMB (Join-Path $env:SystemRoot 'SoftwareDistribution\Download')
    $tempTotal = $tempUser + $tempWin + $tempWu
    if ($tempTotal -gt 3072) { Add-C 'WARN' 'Швидкодія' ("Тимчасові файли: {0:N1} ГБ" -f ($tempTotal / 1024)) "TEMP користувача $tempUser МБ, Windows\Temp $tempWin МБ, кеш оновлень $tempWu МБ." 'Очистіть (аптечка запропонує в кінці) або «Очищення диска».' }

    # --- схема живлення ---
    $scheme = [string](& powercfg.exe /getactivescheme 2>$null)
    if ($scheme -match 'a1841308-3541-4fab-bc81-f71556f20b4a' -and -not $isLaptop) { Add-C 'WARN' 'Швидкодія' 'Схема живлення «Економія енергії»' 'На настільному ПК це штучно знижує частоту процесора.' 'Панель керування > Електроживлення > «Збалансована» або «Висока продуктивність».' }

    # --- батарея ---
    $battery = $null
    if ($isLaptop) {
        $b = Get-CimInstance Win32_Battery | Select-Object -First 1
        $wear = $null
        if ($isAdmin) {
            $full = (Get-CimInstance -Namespace root\wmi -ClassName BatteryFullChargedCapacity | Select-Object -First 1).FullChargedCapacity
            $design = (Get-CimInstance -Namespace root\wmi -ClassName BatteryStaticData | Select-Object -First 1).DesignedCapacity
            if ($full -and $design) { $wear = [math]::Round(100 - $full / $design * 100) }
        }
        $battery = [pscustomobject]@{ Charge = $b.EstimatedChargeRemaining; Wear = $wear }
        if ($null -ne $wear -and $wear -ge 40) { Add-C 'WARN' 'Обладнання' "Батарея зношена на $wear%" 'Ноутбук тримає заряд значно менше, ніж новий.' 'Плануйте заміну батареї, особливо якщо ноутбук потрібен під час відключень світла.' }
        elseif ($null -ne $wear) { Add-C 'OK' 'Обладнання' "Батарея: знос $wear%" '' '' }
    }

    # --- мережа (швидко) ---
    $gw = (Get-CimInstance Win32_NetworkAdapterConfiguration -Filter 'IPEnabled=True' | ForEach-Object { $_.DefaultIPGateway } | Where-Object { $_ -match '^\d+\.' } | Select-Object -First 1)
    $png = New-Object System.Net.NetworkInformation.Ping
    $gwOk = $false; $inetOk = $false; $dnsOk = $false
    if ($gw) { try { $gwOk = $png.Send($gw, 1000).Status -eq 'Success' } catch { } }
    foreach ($h in '8.8.8.8', '1.1.1.1') { try { if ($png.Send($h, 1500).Status -eq 'Success') { $inetOk = $true; break } } catch { } }
    try { $dnsOk = @([System.Net.Dns]::GetHostAddresses('www.google.com')).Count -gt 0 } catch { }
    if (-not $gw) { Add-C 'FAIL' 'Мережа' 'Немає підключення до мережі (немає шлюзу)' '' 'Запустіть «Діагностику мережі» з меню інструментів.' }
    elseif (-not $inetOk -and -not $dnsOk) { Add-C 'FAIL' 'Мережа' 'Немає інтернету' "Шлюз $gw $(if ($gwOk) { 'відповідає' } else { 'не відповідає' })." 'Запустіть «Діагностику мережі» з меню інструментів.' }
    elseif (-not $dnsOk) { Add-C 'WARN' 'Мережа' 'Не працює DNS' 'Інтернет за IP є, а імена сайтів не визначаються.' 'ipconfig /flushdns; детальніше - «Діагностика мережі».' }
    else { Add-C 'OK' 'Мережа' 'Мережа та інтернет працюють' '' '' }

    # --- загальна інформація ---
    $info = @(
        [pscustomobject]@{ K = 'Комп''ютер'; V = $env:COMPUTERNAME }
        [pscustomobject]@{ K = 'Користувач'; V = $cs.UserName }
        [pscustomobject]@{ K = 'Модель'; V = "$($cs.Manufacturer) $($cs.Model)".Trim() }
        [pscustomobject]@{ K = 'Windows'; V = "$(([string]$os.Caption) -replace '^Microsoft\s+', '') $($cv.DisplayVersion) (збірка $($os.BuildNumber).$($cv.UBR))" }
        [pscustomobject]@{ K = 'Процесор'; V = (($cpu | ForEach-Object { ([string]$_.Name).Trim() -replace '\s+', ' ' }) | Select-Object -Unique) -join ' + ' }
        [pscustomobject]@{ K = 'Оперативна пам''ять'; V = "$ramGB ГБ" }
        [pscustomobject]@{ K = 'Останнє ввімкнення'; V = $(if ($os.LastBootUpTime) { $os.LastBootUpTime.ToString('dd.MM.yyyy HH:mm') }) }
        [pscustomobject]@{ K = 'Перевірка з правами адміна'; V = $(if ($isAdmin) { 'Так' } else { 'Ні (SMART-лічильники та знос батареї пропущено)' }) }
    )

    [pscustomobject]@{
        Computer = $env:COMPUTERNAME; IsAdmin = $isAdmin; Info = $info; Checks = $checks.ToArray()
        UpDays = $upDays; RamGB = $ramGB; RamUsed = $ramUsed; CpuLoad = $cpuLoad
        SysFree = ($disks | Where-Object { $_.Drive -eq $env:SystemDrive } | Select-Object -First 1)
        Disks = $disks; Phys = $phys; TopMem = $topMem; TopCpu = $topCpu; Crashes = @($crashes | Sort-Object When -Descending)
        AppCrashes = $appCrashes; BadDevices = $badDev; Services = $svcRows; Startup = $startup
        TempMB = $tempTotal; TempUser = $tempUser; TempWin = $tempWin; Battery = $battery
        SpoolerDown = @($svcRows | Where-Object { $_.Name -eq 'Spooler' -and $_.State -ne 'Running' -and $_.Mode -ne 'Disabled' }).Count -gt 0
    }
}

# =====================================================================
#  ЗАПУСК: ДЕ ПЕРЕВІРЯЄМО
# =====================================================================
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $ComputerName -and -not $Local) {
    Write-Host ''
    Write-Host '========================================================' -ForegroundColor Cyan
    Write-Host '   АПТЕЧКА ПК' -ForegroundColor Cyan
    Write-Host '========================================================' -ForegroundColor Cyan
    $ComputerName = (Read-Host "  Ім'я або IP комп'ютера для перевірки (Enter - цей ПК, $env:COMPUTERNAME)").Trim()
}
$Remote = $ComputerName -and $ComputerName -notin '.', 'localhost', $env:COMPUTERNAME

if (-not $Remote -and -not $isAdmin -and -not $NoElevate) {
    Write-Host '  Для повної перевірки (стан дисків SMART, знос батареї) потрібні права адміністратора.' -ForegroundColor Yellow
    try {
        $argList = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$PSCommandPath`"", '-Local', '-NoElevate')
        if ($NoOpen) { $argList += '-NoOpen' }
        $proc = Start-Process powershell.exe -Verb RunAs -ArgumentList $argList -Wait -PassThru -ErrorAction Stop
        exit $proc.ExitCode
    } catch {
        Write-Host '  Права не надано - продовжую без них.' -ForegroundColor DarkGray
    }
}

Write-Host ''
if ($Remote) {
    Write-Host "  Перевіряю $ComputerName віддалено (WinRM)... ~20 с" -ForegroundColor Cyan
    try {
        $HD = Invoke-Command -ComputerName $ComputerName -ScriptBlock $Collector -ErrorAction Stop
    } catch {
        Write-Host ''
        Write-Host "  Не вдалося підключитися до $ComputerName :" -ForegroundColor Red
        Write-Host "  $($_.Exception.Message)" -ForegroundColor DarkGray
        Write-Host ''
        Write-Host '  Що робити:' -ForegroundColor Yellow
        Write-Host '   - перевірте, що ПК увімкнений і ім''я правильне;'
        Write-Host '   - потрібен домен і права адміністратора на тому ПК;'
        Write-Host '   - на тому ПК має бути ввімкнений WinRM: Enable-PSRemoting -Force (від адміна)'
        Write-Host '     або для всіх ПК одразу через групову політику;'
        Write-Host '   - або просто запустіть аптечку на самому ПК (через AnyDesk).'
        Read-Host '  Натисніть Enter' | Out-Null
        exit 1
    }
} else {
    Write-Host '  Перевіряю цей комп''ютер... ~20 с' -ForegroundColor Cyan
    $HD = & $Collector
}

# =====================================================================
#  ЗВІТ
# =====================================================================
$order = @{ FAIL = 0; WARN = 1; OK = 2; INFO = 3 }
$checks = @($HD.Checks | Sort-Object @{ e = { $order[[string]$_.Status] } }, Area)
$fails = @($checks | Where-Object { $_.Status -eq 'FAIL' })
$warns = @($checks | Where-Object { $_.Status -eq 'WARN' })
$oks   = @($checks | Where-Object { $_.Status -eq 'OK' })

Write-Host ''
Write-Host '========================================================' -ForegroundColor Cyan
Write-Host "  РЕЗУЛЬТАТ: $($HD.Computer)" -ForegroundColor Cyan
Write-Host '========================================================' -ForegroundColor Cyan
foreach ($c in ($fails + $warns)) {
    $col = 'Yellow'; if ($c.Status -eq 'FAIL') { $col = 'Red' }
    Write-Host ('  [{0,-4}] {1}' -f $c.Status, $c.Name) -ForegroundColor $col
    if ($c.Fix) { Write-Host "         -> $($c.Fix)" -ForegroundColor DarkGray }
}
if (-not ($fails.Count + $warns.Count)) { Write-Host '  Проблем не знайдено.' -ForegroundColor Green }
Write-Host ("  У нормі: {0} перевірок" -f $oks.Count) -ForegroundColor Green

if ($fails.Count)     { $v = @('bad',  "Знайдено серйозних проблем: $($fails.Count)", 'Нижче - що саме не так і що з цим робити, від найважливішого.') }
elseif ($warns.Count) { $v = @('warn', "Є зауваження: $($warns.Count)", 'Серйозних проблем немає, але ось що варто виправити.') }
else                  { $v = @('ok',   'Комп''ютер у нормі', 'Усі перевірки пройдено. Якщо скарга залишається - проблема, ймовірно, в конкретній програмі або мережі.') }

$sysFree = '-'
if ($HD.SysFree) { $sysFree = "$($HD.SysFree.FreeGB) ГБ ($($HD.SysFree.FreePct)%)" }
$cards = [ordered]@{
    'Без перезавантаження' = $(if ($null -ne $HD.UpDays) { "$($HD.UpDays) дн." } else { '-' })
    'Процесор зараз'       = $(if ($null -ne $HD.CpuLoad) { "$($HD.CpuLoad)%" } else { '-' })
    'Пам''ять зайнято'     = $(if ($null -ne $HD.RamUsed) { "$($HD.RamUsed)% з $($HD.RamGB) ГБ" } else { '-' })
    'Вільно на системному' = $sysFree
    'Проблем / зауважень'  = "$($fails.Count) / $($warns.Count)"
}

$probRows = foreach ($c in ($fails + $warns)) {
    [pscustomobject]@{ 'Стан' = (New-Tag $c.Status); 'Розділ' = $c.Area; 'Що не так' = $c.Name; 'Деталі' = $c.Detail; 'Що робити' = $c.Fix }
}
$okHtml = '<ul>' + (($oks | ForEach-Object { "<li>$(Enc $_.Name)</li>" }) -join '') + '</ul>'

$memRows = foreach ($m in $HD.TopMem) { [pscustomobject]@{ 'Програма' = $m.Name; 'Опис' = $m.Desc; 'Процесів' = $m.Count; 'Пам''ять, МБ' = $m.MB } }
$cpuRows = foreach ($c in $HD.TopCpu) { [pscustomobject]@{ 'Процес' = $c.Name; 'ЦП, %' = $c.Pct } }
$diskRows = foreach ($d in $HD.Disks) {
    $st = 'OK'; if ($d.FreePct -lt 10) { $st = 'FAIL' } elseif ($d.FreePct -lt 20) { $st = 'WARN' }
    [pscustomobject]@{ 'Диск' = $d.Drive; 'Мітка' = $d.Label; 'Обсяг, ГБ' = $d.SizeGB; 'Вільно, ГБ' = $d.FreeGB; 'Вільно' = (New-Tag $st "$($d.FreePct)%") }
}
$physRows = foreach ($p in $HD.Phys) {
    $st = 'OK'; if ($p.Health -eq 'Unhealthy') { $st = 'FAIL' } elseif ($p.Health -and $p.Health -ne 'Healthy') { $st = 'WARN' }
    [pscustomobject]@{
        'Модель' = $p.Model; 'Тип' = $p.Type; 'Обсяг, ГБ' = $p.SizeGB; 'Стан' = (New-Tag $st $(if ($p.Health) { $p.Health } else { 'н/д' }))
        'Знос, %' = $p.Wear; 'Темп., °C' = $p.Temp; 'Помилки читання' = $p.ReadErr; 'Годин роботи' = $p.Hours; 'Системний' = $(if ($p.IsSystem) { 'Так' } else { '' })
    }
}
$crashRows = foreach ($c in $HD.Crashes) { [pscustomobject]@{ 'Коли' = ([datetime]$c.When).ToString('dd.MM.yyyy HH:mm'); 'Що сталося' = $c.Kind; 'Деталі' = $c.Detail } }
$appRows   = foreach ($a in $HD.AppCrashes) { [pscustomobject]@{ 'Програма' = $a.App; 'Що' = $a.Kind; 'Разів за 14 дн.' = $a.Count; 'Востаннє' = ([datetime]$a.Last).ToString('dd.MM.yyyy HH:mm') } }
$devRows   = foreach ($b in $HD.BadDevices) { [pscustomobject]@{ 'Пристрій' = $b.Name; 'Проблема' = $b.Problem } }
$svcHtmlRows = foreach ($s in $HD.Services) {
    $st = 'OK'; if ($s.State -ne 'Running') { if ($s.Mode -eq 'Disabled') { $st = 'INFO' } else { $st = 'WARN' } }
    [pscustomobject]@{ 'Служба' = $s.Title; 'Ім''я' = $s.Name; 'Стан' = (New-Tag $st $s.State); 'Запуск' = $s.Mode }
}
$startRows = foreach ($s in $HD.Startup) { [pscustomobject]@{ 'Назва' = $s.Name; 'Команда' = $s.Command; 'Звідки' = $s.Where; 'Для кого' = $s.User } }

$body  = "<div class=""verdict $($v[0])""><h2>$(Enc $v[1])</h2><p>$(Enc $v[2])</p></div>"
$body += New-Cards $cards
if ($probRows) { $body += New-Section 'Що не так і що робити' (ConvertTo-HtmlTable @($probRows) -RawColumns 'Стан' -Name 'problems' -NoToolbar) }
$body += New-Section "У нормі ($($oks.Count))" "<details><summary>Показати</summary>$okHtml</details>"
$body += New-Section 'Хто їсть пам''ять' (ConvertTo-HtmlTable @($memRows) -Name 'memory' -NoToolbar) 'Сумарно по всіх процесах програми (у браузерів їх десятки).'
if ($cpuRows) { $body += New-Section 'Хто навантажує процесор зараз' (ConvertTo-HtmlTable @($cpuRows) -Name 'cpu' -NoToolbar) }
$body += New-Section 'Диски' ((ConvertTo-HtmlTable @($diskRows) -RawColumns 'Вільно' -Name 'volumes' -NoToolbar) + '<h3>Фізичні диски</h3>' + (ConvertTo-HtmlTable @($physRows) -RawColumns 'Стан' -Name 'physical' -NoToolbar)) $(if (-not $HD.IsAdmin) { 'Знос і помилки SSD видно лише при запуску з правами адміністратора.' })
if ($crashRows) { $body += New-Section 'Збої системи за 30 днів' (ConvertTo-HtmlTable @($crashRows) -Name 'crashes') }
if ($appRows)   { $body += New-Section 'Збої програм за 14 днів' (ConvertTo-HtmlTable @($appRows) -Name 'app-crashes' -NoToolbar) }
if ($devRows)   { $body += New-Section 'Пристрої з помилками' (ConvertTo-HtmlTable @($devRows) -Name 'devices' -NoToolbar) 'Те саме, що жовті знаки оклику в Диспетчері пристроїв.' }
$body += New-Section 'Важливі служби' (ConvertTo-HtmlTable @($svcHtmlRows) -RawColumns 'Стан' -Name 'services' -NoToolbar)
$body += New-Section "Автозавантаження ($(@($HD.Startup).Count))" (ConvertTo-HtmlTable @($startRows) -Name 'startup')
$infoDict = [ordered]@{}; foreach ($i in $HD.Info) { $infoDict[$i.K] = $i.V }
$body += New-Section 'Про комп''ютер' (ConvertTo-KvTable $infoDict)

$mode = 'локально'; if ($Remote) { $mode = "віддалено з $env:COMPUTERNAME" }
$html = New-HtmlPage "Аптечка: $($HD.Computer)" "$($v[1]) · перевірено $mode · $(Get-Date -Format 'dd.MM.yyyy HH:mm')" $body 'Test-PCHealth'
$path = Save-Html $html "Health_$($HD.Computer)_$Stamp.html"
Write-Host ''
Write-Host "  Звіт: $path" -ForegroundColor Green
if (-not $NoOpen) { Start-Process $path }

# =====================================================================
#  ШВИДКІ ВИПРАВЛЕННЯ (лише на цьому ПК)
# =====================================================================
if (-not $Remote) {
    if ($HD.TempMB -gt 500) {
        Write-Host ''
        $a = Read-Host ("  Очистити тимчасові файли (~{0:N1} ГБ)? Видаляються лише файли, старші за добу. (Y/N)" -f ($HD.TempMB / 1024))
        if ($a -match '^[yYнН]') {
            $before = 0; $after = 0
            $paths = @($env:TEMP)
            if ($isAdmin -or $HD.IsAdmin) { $paths += (Join-Path $env:SystemRoot 'Temp') }
            foreach ($p in $paths) {
                $files = @(Get-ChildItem $p -Recurse -File -Force -ErrorAction SilentlyContinue | Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-1) })
                $before += ($files | Measure-Object Length -Sum).Sum
                $files | Remove-Item -Force -ErrorAction SilentlyContinue
                $after += (@($files | Where-Object { Test-Path -LiteralPath $_.FullName }) | Measure-Object Length -Sum).Sum
                Get-ChildItem $p -Recurse -Directory -Force -ErrorAction SilentlyContinue | Sort-Object { $_.FullName.Length } -Descending |
                    Where-Object { -not (Get-ChildItem $_.FullName -Force -ErrorAction SilentlyContinue) } | Remove-Item -Force -ErrorAction SilentlyContinue
            }
            Write-Host ("  Звільнено {0:N0} МБ (зайняті програмами файли пропущено)." -f (($before - $after) / 1MB)) -ForegroundColor Green
            if ($HD.TempWin -gt 500 -and -not ($isAdmin -or $HD.IsAdmin)) { Write-Host '  Windows\Temp можна очистити лише з правами адміністратора.' -ForegroundColor DarkGray }
        }
    }
    if ($HD.SpoolerDown) {
        $a = Read-Host '  Служба друку зупинена. Запустити? (Y/N)'
        if ($a -match '^[yYнН]') {
            try { Start-Service Spooler -ErrorAction Stop; Write-Host '  Службу друку запущено.' -ForegroundColor Green } catch { Write-Host "  Не вдалося: $($_.Exception.Message)" -ForegroundColor Red }
        }
    }
    Write-Host ''
    Read-Host '  Натисніть Enter, щоб закрити' | Out-Null
}
