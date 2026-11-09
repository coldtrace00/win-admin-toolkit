#Requires -Version 5.1
<#
.SYNOPSIS
    Звіти Active Directory з меню: неактивні облікові записи, паролі, блокування, групи, адміни, ОС.

.DESCRIPTION
    Запускається на будь-якому ПК у домені (або на контролері домену).
    Потрібен модуль ActiveDirectory (RSAT). Якщо його немає, скрипт запропонує встановити.
    Для звітів достатньо звичайного доменного користувача; для розблокування потрібні права адміна/хелпдеску.

    Звіт: HTML на робочий стіл, у кожній таблиці пошук, сортування і кнопка «Експорт у CSV».

.PARAMETER Report
    Номери звітів без меню, напр. -Report 1,3,7 або -Report All.

.EXAMPLE
    .\Get-ADReports.ps1
.EXAMPLE
    .\Get-ADReports.ps1 -Report All -InactiveDays 60
#>
[CmdletBinding()]
param(
    [string[]]$Report,
    [int]$InactiveDays = 90,     # скільки днів без входу вважається «неактивний»
    [int]$PasswordDays = 14,     # за скільки днів попереджати про пароль, що спливає
    [switch]$NoOpen
)

$ErrorActionPreference = 'Continue'
try { $Host.UI.RawUI.WindowTitle = 'Звіти Active Directory' } catch { }

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
#  МОДУЛЬ ACTIVE DIRECTORY
# =====================================================================
if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
    Write-Host ''
    Write-Host 'Не знайдено модуль PowerShell для Active Directory (входить до RSAT).' -ForegroundColor Yellow
    $isServer = (Get-CimInstance Win32_OperatingSystem).ProductType -ne 1
    if ($isServer) { $cmd = 'Install-WindowsFeature RSAT-AD-PowerShell' }
    else           { $cmd = "Add-WindowsCapability -Online -Name 'Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0'" }
    $a = Read-Host 'Встановити зараз? Потрібні права адміністратора та інтернет (або WSUS), 1-5 хв. (Y/N)'
    if ($a -match '^[yYнН]') {
        Write-Host 'Встановлюю, зачекайте...' -ForegroundColor Cyan
        try { Start-Process powershell.exe -Verb RunAs -Wait -ArgumentList @('-NoProfile', '-Command', $cmd) -ErrorAction Stop } catch { }
    }
    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        Write-Host ''
        Write-Host 'Модуль не встановлено. Встановіть вручну:' -ForegroundColor Red
        Write-Host '  Параметри > Система > Додаткові компоненти > Додати компонент >'
        Write-Host '  «RSAT: Active Directory Domain Services and Lightweight Directory Services Tools»'
        Write-Host "  або в PowerShell від адміністратора: $cmd"
        Read-Host 'Натисніть Enter, щоб закрити' | Out-Null
        exit 1
    }
}

$env:ADPS_LoadDefaultDrive = 0
Import-Module ActiveDirectory -ErrorAction Stop -WarningAction SilentlyContinue
try {
    $Domain = Get-ADDomain -ErrorAction Stop
} catch {
    Write-Host "Не вдалося підключитися до домену: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host 'Переконайтеся, що ПК у домені й ви увійшли доменним обліковим записом.'
    Read-Host 'Натисніть Enter, щоб закрити' | Out-Null
    exit 1
}

# =====================================================================
#  ДОПОМІЖНІ
# =====================================================================
function Get-Cn([string]$Dn) {
    if (-not $Dn) { return '' }
    ((($Dn -split '(?<!\\),')[0]) -replace '^(CN|OU)=', '') -replace '\\(.)', '$1'
}

function Get-Container([string]$Dn) {
    $parts = @(($Dn -split '(?<!\\),') | Select-Object -Skip 1 | Where-Object { $_ -match '^(OU|CN)=' } |
               ForEach-Object { ($_ -replace '^(OU|CN)=', '') -replace '\\(.)', '$1' })
    if ($parts.Count -eq 0) { return '(корінь домену)' }
    [array]::Reverse($parts)
    $parts -join ' / '
}

function Fmt($Dt) { if ($Dt) { ([datetime]$Dt).ToString('dd.MM.yyyy HH:mm') } else { '' } }
function Get-DaysAgo($Dt) { if ($Dt) { [int][math]::Floor(((Get-Date) - [datetime]$Dt).TotalDays) } else { $null } }
function Get-DisplayName($Obj) { if ($Obj.DisplayName) { $Obj.DisplayName } else { $Obj.Name } }
function Test-BuiltinAccount($Obj) { [string]$Obj.SID -match '-50[1-4]$' }   # Гість, krbtgt, DefaultAccount, WDAG

# =====================================================================
#  ЗВІТИ
# =====================================================================
function Get-RptInactiveUsers {
    $cut = (Get-Date).AddDays(-$InactiveDays)
    $rows = Get-ADUser -Filter 'Enabled -eq $true' -Properties DisplayName, LastLogonDate, whenCreated, Department, Title |
        Where-Object { -not (Test-BuiltinAccount $_) -and (($_.LastLogonDate -and $_.LastLogonDate -lt $cut) -or (-not $_.LastLogonDate -and $_.whenCreated -lt $cut)) } |
        Sort-Object @{ e = { if ($_.LastLogonDate) { $_.LastLogonDate } else { [datetime]::MinValue } } } |
        ForEach-Object {
            [pscustomobject]@{
                'Логін'          = $_.SamAccountName
                'ПІБ'            = (Get-DisplayName $_)
                'Відділ'         = $_.Department
                'Посада'         = $_.Title
                'Останній вхід'  = $(if ($_.LastLogonDate) { Fmt $_.LastLogonDate } else { 'ніколи' })
                'Днів без входу' = (Get-DaysAgo $_.LastLogonDate)
                'Створено'       = (Fmt $_.whenCreated)
                'Розташування'   = (Get-Container $_.DistinguishedName)
            }
        }
    @{ Rows = @($rows); Desc = "Увімкнені облікові записи без входу понад $InactiveDays днів. Кандидати на вимкнення (звільнені, тестові, забуті). Дата входу в AD оновлюється із затримкою до 14 днів." }
}

function Get-RptInactiveComputers {
    $cut = (Get-Date).AddDays(-$InactiveDays)
    $rows = Get-ADComputer -Filter 'Enabled -eq $true' -Properties LastLogonDate, OperatingSystem, whenCreated, Description |
        Where-Object { ($_.LastLogonDate -and $_.LastLogonDate -lt $cut) -or (-not $_.LastLogonDate -and $_.whenCreated -lt $cut) } |
        Sort-Object @{ e = { if ($_.LastLogonDate) { $_.LastLogonDate } else { [datetime]::MinValue } } } |
        ForEach-Object {
            [pscustomobject]@{
                'Комп''ютер'       = $_.Name
                'ОС'               = $_.OperatingSystem
                'Останній вхід'    = $(if ($_.LastLogonDate) { Fmt $_.LastLogonDate } else { 'ніколи' })
                'Днів не в мережі' = (Get-DaysAgo $_.LastLogonDate)
                'Створено'         = (Fmt $_.whenCreated)
                'Опис'             = $_.Description
                'Розташування'     = (Get-Container $_.DistinguishedName)
            }
        }
    @{ Rows = @($rows); Desc = "Увімкнені комп'ютери, що не з'являлися в домені понад $InactiveDays днів: списані, переустановлені під іншим ім'ям або давно вимкнені." }
}

function Get-RptPasswordExpiry {
    $now = Get-Date
    $users = Get-ADUser -Filter 'Enabled -eq $true -and PasswordNeverExpires -eq $false' -Properties DisplayName, EmailAddress, Department, PasswordLastSet, 'msDS-UserPasswordExpiryTimeComputed'
    $rows = foreach ($u in $users) {
        if (Test-BuiltinAccount $u) { continue }
        $raw = $u.'msDS-UserPasswordExpiryTimeComputed'
        $exp = $null
        if (-not $u.PasswordLastSet) {
            $left = -1; $st = 'FAIL'; $txt = 'Змінити при вході'
        } elseif (-not $raw -or $raw -ge [long]::MaxValue) {
            continue
        } else {
            $exp  = [datetime]::FromFileTime($raw)
            $left = [int][math]::Floor(($exp - $now).TotalDays)
            if ($left -gt $PasswordDays) { continue }
            if     ($left -lt 0) { $st = 'FAIL'; $txt = 'Прострочено' }
            elseif ($left -le 3) { $st = 'FAIL'; $txt = "$left дн." }
            elseif ($left -le 7) { $st = 'WARN'; $txt = "$left дн." }
            else                 { $st = 'INFO'; $txt = "$left дн." }
        }
        [pscustomobject]@{
            'Статус'          = (New-Tag $st $txt)
            'Логін'           = $u.SamAccountName
            'ПІБ'             = (Get-DisplayName $u)
            'Email'           = $u.EmailAddress
            'Відділ'          = $u.Department
            'Пароль змінено'  = (Fmt $u.PasswordLastSet)
            'Спливає'         = (Fmt $exp)
            'Залишилось днів' = $left
        }
    }
    @{ Rows = @($rows | Sort-Object 'Залишилось днів'); Raw = @('Статус'); Desc = "Паролі, що спливають протягом $PasswordDays днів або вже прострочені. Варто попередити людей заздалегідь, особливо тих, хто працює віддалено через VPN." }
}

function Get-LockedUsers {
    @(Search-ADAccount -LockedOut -UsersOnly -ErrorAction Stop | ForEach-Object {
        Get-ADUser -Identity $_.DistinguishedName -Properties DisplayName, Department, LockoutTime, BadLogonCount, LastBadPasswordAttempt
    })
}

function Get-RptLocked {
    $rows = foreach ($u in (Get-LockedUsers)) {
        $lt = $null
        if ($u.LockoutTime) { try { $lt = [datetime]::FromFileTime([long]$u.LockoutTime) } catch { } }
        [pscustomobject]@{
            'Логін'                 = $u.SamAccountName
            'ПІБ'                   = (Get-DisplayName $u)
            'Відділ'                = $u.Department
            'Заблоковано'           = (Fmt $lt)
            'Невдалих спроб'        = $u.BadLogonCount
            'Остання невдала спроба'= (Fmt $u.LastBadPasswordAttempt)
            'Розташування'          = (Get-Container $u.DistinguishedName)
        }
    }
    @{ Rows = @($rows); Desc = 'Заблоковані через неправильний пароль. Розблокувати: пункт U у меню. Якщо обліковий запис блокується постійно, шукайте старий пароль у телефоні (пошта), підключених дисках чи збережених облікових даних.' }
}

function Get-RptGroups {
    $empty = 0
    $rows = foreach ($g in (Get-ADGroup -Filter * -Properties Members, Description | Sort-Object Name)) {
        $m = @($g.Members)
        if ($m.Count -eq 0) { $empty++; continue }
        $names = @($m | ForEach-Object { Get-Cn $_ } | Sort-Object)
        [pscustomobject]@{
            'Група'        = $g.Name
            'Тип'          = $(if ([string]$g.GroupCategory -eq 'Security') { 'Безпеки' } else { 'Розсилки' })
            'Область'      = [string]$g.GroupScope
            'Учасники'     = "<details><summary>$($m.Count)</summary>$(Enc ($names -join ', '))</details>"
            'Опис'         = $g.Description
            'Розташування' = (Get-Container $g.DistinguishedName)
        }
    }
    @{ Rows = @($rows); Raw = @('Учасники'); Desc = "Прямі учасники кожної групи (натисніть на число). Порожніх груп приховано: $empty. Основна група (зазвичай «Domain Users») тут не показується - так влаштовано AD." }
}

function Get-RptUserGroups {
    $rows = foreach ($u in (Get-ADUser -Filter 'Enabled -eq $true' -Properties DisplayName, Department, MemberOf, PrimaryGroup | Sort-Object SamAccountName)) {
        if (Test-BuiltinAccount $u) { continue }
        $gr = @(@($u.MemberOf) + @($u.PrimaryGroup) | Where-Object { $_ } | ForEach-Object { Get-Cn $_ } | Sort-Object -Unique)
        [pscustomobject]@{
            'Логін'   = $u.SamAccountName
            'ПІБ'     = (Get-DisplayName $u)
            'Відділ'  = $u.Department
            'Груп'    = $gr.Count
            'Групи'   = ($gr -join ', ')
        }
    }
    @{ Rows = @($rows); Desc = 'Увімкнені користувачі та групи, в яких вони прямо перебувають. Пошук працює за назвою групи: введіть її, щоб побачити всіх учасників.' }
}

function Get-RptPrivileged {
    $dsid = $Domain.DomainSID.Value
    $rootSid = $dsid
    try { $rootSid = (Get-ADDomain -Identity (Get-ADForest).RootDomain -ErrorAction Stop).DomainSID.Value } catch { }
    $ids = @("$dsid-512", "$rootSid-519", "$rootSid-518", 'S-1-5-32-544', 'S-1-5-32-548', 'S-1-5-32-549', 'S-1-5-32-551', 'S-1-5-32-550', "$dsid-520", 'DnsAdmins')
    $cutLogon = (Get-Date).AddDays(-$InactiveDays)
    $cutPwd   = (Get-Date).AddDays(-365)
    $rows = foreach ($id in $ids) {
        $g = $null
        try { $g = Get-ADGroup -Identity $id -ErrorAction Stop } catch { continue }
        $members = @()
        try {
            $members = @(Get-ADGroupMember -Identity $g -Recursive -ErrorAction Stop)
        } catch {
            $members = @((Get-ADGroup $g -Properties Members).Members | ForEach-Object {
                try { Get-ADObject -Identity $_ -Properties SamAccountName -ErrorAction Stop } catch { }
            })
        }
        foreach ($m in $members) {
            $notes = @(); $login = $m.SamAccountName; $name = $m.Name; $type = [string]$m.objectClass
            $enabled = ''; $lastLogon = ''; $pwdSet = ''
            if ($type -eq 'user') {
                $type = 'Користувач'
                try {
                    $u = Get-ADUser -Identity $m.DistinguishedName -Properties DisplayName, LastLogonDate, PasswordLastSet, PasswordNeverExpires -ErrorAction Stop
                    $name = Get-DisplayName $u
                    $enabled = $(if ($u.Enabled) { 'Так' } else { 'Ні' })
                    $lastLogon = Fmt $u.LastLogonDate
                    $pwdSet = Fmt $u.PasswordLastSet
                    if (-not $u.Enabled)                                         { $notes += New-Tag 'INFO' 'Вимкнений' }
                    if ($u.Enabled -and $u.PasswordNeverExpires)                 { $notes += New-Tag 'WARN' 'Пароль не спливає' }
                    if ($u.Enabled -and $u.PasswordLastSet -and $u.PasswordLastSet -lt $cutPwd) { $notes += New-Tag 'WARN' 'Пароль старший за рік' }
                    if ($u.Enabled -and (-not $u.LastLogonDate -or $u.LastLogonDate -lt $cutLogon)) { $notes += New-Tag 'WARN' 'Давно не входив' }
                } catch { }
            } elseif ($type -eq 'computer') { $type = 'Комп''ютер'; $notes += New-Tag 'WARN' 'Комп''ютер у групі адмінів' }
            elseif ($type -eq 'group')      { $type = 'Група' }
            if ($notes.Count -eq 0) { $notes += New-Tag 'OK' 'OK' }
            [pscustomobject]@{
                'Група'           = $g.Name
                'Логін'           = $login
                'Ім''я'           = $name
                'Тип'             = $type
                'Увімкнений'      = $enabled
                'Останній вхід'   = $lastLogon
                'Пароль змінено'  = $pwdSet
                'Зауваження'      = ($notes -join ' ')
            }
        }
    }
    @{ Rows = @($rows); Raw = @('Зауваження'); Desc = 'Хто має адміністративні права в домені (вкладені групи розгорнуто). Кожен зайвий обліковий запис тут - пряма загроза: саме їх атакують першими. Адмінам варто мати окремий обліковий запис для щоденної роботи.' }
}

function Get-RptNeverExpires {
    $rows = Get-ADUser -Filter 'Enabled -eq $true -and PasswordNeverExpires -eq $true' -Properties DisplayName, PasswordLastSet, LastLogonDate, Description |
        Where-Object { -not (Test-BuiltinAccount $_) } | Sort-Object PasswordLastSet |
        ForEach-Object {
            [pscustomobject]@{
                'Логін'           = $_.SamAccountName
                'ПІБ'             = (Get-DisplayName $_)
                'Пароль змінено'  = (Fmt $_.PasswordLastSet)
                'Вік пароля, дн.' = (Get-DaysAgo $_.PasswordLastSet)
                'Останній вхід'   = (Fmt $_.LastLogonDate)
                'Опис'            = $_.Description
                'Розташування'    = (Get-Container $_.DistinguishedName)
            }
        }
    @{ Rows = @($rows); Desc = 'Облікові записи з паролем «без терміну дії». Допустимо для службових облікових записів, але не для звичайних людей і тим паче не для адмінів.' }
}

function Get-RptDisabled {
    $rows = Get-ADUser -Filter 'Enabled -eq $false' -Properties DisplayName, whenChanged, LastLogonDate, Description, MemberOf |
        Where-Object { -not (Test-BuiltinAccount $_) } | Sort-Object whenChanged -Descending |
        ForEach-Object {
            [pscustomobject]@{
                'Логін'                = $_.SamAccountName
                'ПІБ'                  = (Get-DisplayName $_)
                'Змінено (≈вимкнено)'  = (Fmt $_.whenChanged)
                'Останній вхід'        = (Fmt $_.LastLogonDate)
                'Груп'                 = @($_.MemberOf).Count
                'Опис'                 = $_.Description
                'Розташування'         = (Get-Container $_.DistinguishedName)
            }
        }
    @{ Rows = @($rows); Desc = 'Вимкнені облікові записи. Добра практика: вимкнути при звільненні, прибрати з груп, а через 30-90 днів видалити.' }
}

function Get-WinRelease([string]$Os, [string]$Ver) {
    if ($Os -match 'Server' -or $Ver -notmatch '\((\d+)\)') { return '' }
    $map = @{ 10240 = '1507'; 10586 = '1511'; 14393 = '1607'; 15063 = '1703'; 16299 = '1709'; 17134 = '1803'; 17763 = '1809'
              18362 = '1903'; 18363 = '1909'; 19041 = '2004'; 19042 = '20H2'; 19043 = '21H1'; 19044 = '21H2'; 19045 = '22H2'
              22000 = '21H2'; 22621 = '22H2'; 22631 = '23H2'; 26100 = '24H2'; 26200 = '25H2' }
    [string]$map[[int]$Matches[1]]
}

function Get-RptComputers {
    $cut = (Get-Date).AddDays(-$InactiveDays)
    $all = @(Get-ADComputer -Filter * -Properties OperatingSystem, OperatingSystemVersion, LastLogonDate, whenCreated, Description | Sort-Object Name)
    $rows = foreach ($c in $all) {
        $os = [string]$c.OperatingSystem
        $rel = Get-WinRelease $os ([string]$c.OperatingSystemVersion)
        $notes = @()
        if (-not $c.Enabled) { $notes += New-Tag 'INFO' 'Вимкнений' }
        if ($os -match 'Windows (XP|Vista|7|8(\.1)?)(\s|$)|Server 20(03|08|12)') { $notes += New-Tag 'FAIL' 'Застаріла ОС' }
        elseif ($os -match '^Windows 10' -and $os -notmatch 'LTSC|LTSB') { $notes += New-Tag 'WARN' 'Win10: підтримку завершено 14.10.2025' }
        if ($c.Enabled -and $c.LastLogonDate -and $c.LastLogonDate -lt $cut) { $notes += New-Tag 'WARN' 'Давно не в мережі' }
        if ($notes.Count -eq 0) { $notes += New-Tag 'OK' 'OK' }
        [pscustomobject]@{
            'Комп''ютер'    = $c.Name
            'ОС'            = $os
            'Версія'        = $rel
            'Збірка'        = $c.OperatingSystemVersion
            'Останній вхід' = (Fmt $c.LastLogonDate)
            'Створено'      = (Fmt $c.whenCreated)
            'Зауваження'    = ($notes -join ' ')
            'Розташування'  = (Get-Container $c.DistinguishedName)
        }
    }
    $osSum = @($all | Where-Object { $_.Enabled } | Group-Object { $r = Get-WinRelease ([string]$_.OperatingSystem) ([string]$_.OperatingSystemVersion); ("$($_.OperatingSystem) $r").Trim() } |
               Sort-Object Count -Descending | ForEach-Object { [pscustomobject]@{ 'ОС' = $(if ($_.Name) { $_.Name } else { '(не вказано)' }); 'Кількість' = $_.Count } })
    $extra = '<h3>Розподіл ОС (увімкнені)</h3>' + (ConvertTo-HtmlTable $osSum -Name 'os-summary' -NoToolbar) + '<h3>Усі комп''ютери</h3>'
    @{ Rows = @($rows); Raw = @('Зауваження'); Extra = $extra; Desc = 'Усі облікові записи комп''ютерів домену з версією Windows. Застарілі ОС не отримують оновлень безпеки.' }
}

# --- реєстр звітів (номер -> назва, функція) ---
$Reports = [ordered]@{
    '1'  = @{ Title = 'Неактивні користувачі';           Fn = { Get-RptInactiveUsers };     File = 'inactive-users' }
    '2'  = @{ Title = 'Неактивні комп''ютери';           Fn = { Get-RptInactiveComputers }; File = 'inactive-computers' }
    '3'  = @{ Title = 'Паролі, що скоро спливають';      Fn = { Get-RptPasswordExpiry };    File = 'password-expiry' }
    '4'  = @{ Title = 'Заблоковані облікові записи';     Fn = { Get-RptLocked };            File = 'locked' }
    '5'  = @{ Title = 'Групи та їх учасники';            Fn = { Get-RptGroups };            File = 'groups' }
    '6'  = @{ Title = 'Користувачі та їх групи';         Fn = { Get-RptUserGroups };        File = 'user-groups' }
    '7'  = @{ Title = 'Адміністратори домену (привілейовані групи)'; Fn = { Get-RptPrivileged }; File = 'privileged' }
    '8'  = @{ Title = 'Пароль ніколи не спливає';        Fn = { Get-RptNeverExpires };      File = 'never-expires' }
    '9'  = @{ Title = 'Вимкнені облікові записи';        Fn = { Get-RptDisabled };          File = 'disabled' }
    '10' = @{ Title = 'Усі комп''ютери та версії Windows'; Fn = { Get-RptComputers };       File = 'computers' }
}

# =====================================================================
#  ФОРМУВАННЯ HTML
# =====================================================================
function Invoke-Reports([string[]]$Keys) {
    $sections = ''
    $counts = [ordered]@{}
    $toc = @()
    foreach ($k in $Keys) {
        $def = $Reports[$k]
        Write-Host ("  Формую: {0}..." -f $def.Title) -ForegroundColor Cyan
        try {
            $r = & $def.Fn
            $raw = @(); if ($r.Raw) { $raw = $r.Raw }
            $html = ''
            if ($r.Extra) { $html += $r.Extra }
            $html += ConvertTo-HtmlTable $r.Rows -RawColumns $raw -Name $def.File
            $sections += New-Section "$($def.Title) ($($r.Rows.Count))" $html $r.Desc "r$k"
            $counts[$def.Title] = $r.Rows.Count
            Write-Host ("    знайдено: {0}" -f $r.Rows.Count) -ForegroundColor DarkGray
        } catch {
            $sections += New-Section $def.Title "<p>$(New-Tag 'FAIL') $(Enc $_.Exception.Message)</p>" '' "r$k"
            $counts[$def.Title] = 'помилка'
            Write-Host "    помилка: $($_.Exception.Message)" -ForegroundColor Red
        }
        $toc += "<a href=""#r$k"">$(Enc $def.Title)</a>"
    }

    $pp = $null
    try { $pp = Get-ADDefaultDomainPasswordPolicy -ErrorAction Stop } catch { }
    $policy = [ordered]@{
        'Домен'                      = "$($Domain.DNSRoot) ($($Domain.NetBIOSName))"
        'Рівень домену'              = [string]$Domain.DomainMode
        'Основний контролер (PDC)'   = $Domain.PDCEmulator
    }
    if ($pp) {
        $policy['Макс. вік пароля']       = $(if ($pp.MaxPasswordAge.TotalDays -gt 0) { "$([int]$pp.MaxPasswordAge.TotalDays) дн." } else { 'не обмежено' })
        $policy['Мін. довжина пароля']    = "$($pp.MinPasswordLength) симв."
        $policy['Складність пароля']      = $(if ($pp.ComplexityEnabled) { 'Увімкнена' } else { 'Вимкнена' })
        $policy['Блокування після']       = $(if ($pp.LockoutThreshold -gt 0) { "$($pp.LockoutThreshold) невдалих спроб на $([int]$pp.LockoutDuration.TotalMinutes) хв" } else { 'Не налаштовано' })
    }

    $body  = New-Cards $counts
    if ($Keys.Count -gt 1) { $body += "<div class=""toc"">$($toc -join '')</div>" }
    $body += $sections
    $body += New-Section 'Домен і політика паролів' (ConvertTo-KvTable $policy)

    $title = "Звіт Active Directory: $($Domain.DNSRoot)"
    $sub = "Пороги: неактивні > $InactiveDays дн., паролі < $PasswordDays дн. · Сформував $env:USERDOMAIN\$env:USERNAME · $(Get-Date -Format 'dd.MM.yyyy HH:mm')"
    $html = New-HtmlPage $title $sub $body 'Get-ADReports'
    $name = 'AD'
    if ($Keys.Count -eq 1) { $name += "_$($Reports[$Keys[0]].File)" }
    $path = Save-Html $html "${name}_$($Domain.NetBIOSName)_$Stamp.html"
    Write-Host "  Звіт: $path" -ForegroundColor Green
    if (-not $NoOpen) { Start-Process $path }
}

# =====================================================================
#  РОЗБЛОКУВАННЯ
# =====================================================================
function Invoke-UnlockMenu {
    Write-Host ''
    $locked = @()
    try { $locked = @(Get-LockedUsers) } catch { Write-Host "Помилка: $($_.Exception.Message)" -ForegroundColor Red; return }
    if ($locked.Count -eq 0) { Write-Host '  Заблокованих облікових записів немає.' -ForegroundColor Green; return }
    Write-Host '  Заблоковані облікові записи:' -ForegroundColor White
    for ($i = 0; $i -lt $locked.Count; $i++) {
        $u = $locked[$i]
        $lt = ''
        if ($u.LockoutTime) { try { $lt = Fmt ([datetime]::FromFileTime([long]$u.LockoutTime)) } catch { } }
        Write-Host ('  {0,2}. {1,-20} {2,-30} {3}' -f ($i + 1), $u.SamAccountName, (Get-DisplayName $u), $lt)
    }
    $in = Read-Host '  Номери через кому, A - усі, Enter - назад'
    if ([string]::IsNullOrWhiteSpace($in)) { return }
    $targets = @()
    if ($in.Trim() -match '^[aAфФ]$') { $targets = $locked }
    else {
        foreach ($tok in ($in -split '[,;\s]+')) {
            $n = 0
            if ([int]::TryParse($tok, [ref]$n) -and $n -ge 1 -and $n -le $locked.Count) { $targets += $locked[$n - 1] }
        }
    }
    if ($targets.Count -eq 0) { return }
    $c = Read-Host ("  Розблокувати {0}: {1}? (Y/N)" -f $targets.Count, (($targets | ForEach-Object { $_.SamAccountName }) -join ', '))
    if ($c -notmatch '^[yYнН]') { return }
    foreach ($u in $targets) {
        try {
            Unlock-ADAccount -Identity $u.DistinguishedName -ErrorAction Stop
            Write-Host "  [OK] $($u.SamAccountName) розблоковано" -ForegroundColor Green
        } catch {
            Write-Host "  [!!] $($u.SamAccountName): $($_.Exception.Message)" -ForegroundColor Red
        }
    }
}

# =====================================================================
#  ЗАПУСК
# =====================================================================
function Resolve-Keys([string]$Text) {
    if ($Text -match '^\s*(all|[aAфФ])\s*$') { return @($Reports.Keys) }
    @($Text -split '[,;\s]+' | Where-Object { $Reports.Contains($_) } | Select-Object -Unique)
}

if ($Report) {
    $keys = Resolve-Keys ($Report -join ',')
    if ($keys.Count -eq 0) { Write-Host 'Невідомі номери звітів. Доступні: 1-10 або All.' -ForegroundColor Red; exit 1 }
    Invoke-Reports $keys
    exit 0
}

while ($true) {
    Clear-Host
    Write-Host '========================================================' -ForegroundColor Cyan
    Write-Host "  ЗВІТИ ACTIVE DIRECTORY - $($Domain.DNSRoot)" -ForegroundColor Cyan
    Write-Host '========================================================' -ForegroundColor Cyan
    foreach ($k in $Reports.Keys) { Write-Host ('  {0,2}. {1}' -f $k, $Reports[$k].Title) }
    Write-Host ''
    Write-Host '   A. Усі звіти одразу (один файл)' -ForegroundColor Yellow
    Write-Host '   U. Розблокувати обліковий запис' -ForegroundColor Yellow
    Write-Host "   S. Пороги (зараз: неактивні > $InactiveDays дн., паролі < $PasswordDays дн.)" -ForegroundColor Yellow
    Write-Host '   Q. Вихід' -ForegroundColor Yellow
    Write-Host ''
    $in = Read-Host '  Оберіть (можна кілька через кому, напр. 1,3,7)'
    if ([string]::IsNullOrWhiteSpace($in)) { continue }
    $t = $in.Trim()
    if ($t -match '^[qQйЙ]$') { break }
    if ($t -match '^[uUгГ]$') {
        Invoke-UnlockMenu
        Read-Host '  Enter - назад до меню' | Out-Null
        continue
    }
    if ($t -match '^[sSіІыЫ]$') {
        $v = Read-Host "  Днів без входу для «неактивних» [$InactiveDays]"
        $n = 0; if ([int]::TryParse($v, [ref]$n) -and $n -gt 0) { $InactiveDays = $n }
        $v = Read-Host "  За скільки днів попереджати про пароль [$PasswordDays]"
        $n = 0; if ([int]::TryParse($v, [ref]$n) -and $n -gt 0) { $PasswordDays = $n }
        continue
    }
    $keys = Resolve-Keys $t
    if ($keys.Count -eq 0) { continue }
    Write-Host ''
    Invoke-Reports $keys
    Write-Host ''
    Read-Host '  Enter - назад до меню' | Out-Null
}
