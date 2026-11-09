#Requires -Version 5.1
# Головне меню інструментів адміністратора. Шукає скрипти в підпапках поруч із собою.

$Root = Split-Path -Parent $PSCommandPath
try { $Host.UI.RawUI.WindowTitle = 'Інструменти адміністратора' } catch { }

$Tools = @(
    @{ Name = 'Інвентаризація ПК';            Desc = 'процесор, пам''ять, диски, Windows, ПЗ -> HTML';        Path = 'PC-Inventory\Get-PCInventory.ps1' }
    @{ Name = 'Аптечка ПК';                   Desc = 'чому гальмує/глючить і що робити (цей ПК або інший за ім''ям)'; Path = 'PC-Health\Test-PCHealth.ps1' }
    @{ Name = 'Налаштування нового ПК';       Desc = 'програми через winget + налаштування Windows (UAC)';    Path = 'PC-Setup\Setup-NewPC.ps1' }
    @{ Name = 'Діагностика мережі';           Desc = 'кабель, IP, шлюз, DNS, інтернет, трасування, швидкість'; Path = 'Network-Diag\Test-Network.ps1' }
    @{ Name = 'Діагностика мережі - швидка';  Desc = 'те саме без трасування й тесту швидкості (~20 с)';      Path = 'Network-Diag\Test-Network.ps1'; Args = @('-Quick') }
    @{ Name = 'Сканер мережі (безпека)';      Desc = 'пристрої, відкриті порти, що видно з інтернету, як закрити'; Path = 'Network-Scan\Scan-Network.ps1' }
    @{ Name = 'Звіти Active Directory';       Desc = 'неактивні, паролі, блокування, групи, адміни, ОС';      Path = 'AD-Reports\Get-ADReports.ps1' }
)

while ($true) {
    Clear-Host
    Write-Host '========================================================' -ForegroundColor Cyan
    Write-Host '   ІНСТРУМЕНТИ АДМІНІСТРАТОРА' -ForegroundColor Cyan
    Write-Host "   ПК: $env:COMPUTERNAME   Користувач: $env:USERDOMAIN\$env:USERNAME" -ForegroundColor DarkGray
    Write-Host '========================================================' -ForegroundColor Cyan
    Write-Host ''
    for ($i = 0; $i -lt $Tools.Count; $i++) {
        $t = $Tools[$i]
        $exists = Test-Path (Join-Path $Root $t.Path)
        $col = 'White'; $suffix = ''
        if (-not $exists) { $col = 'DarkGray'; $suffix = '  [немає файлу]' }
        Write-Host ('  {0}. {1}{2}' -f ($i + 1), $t.Name, $suffix) -ForegroundColor $col
        Write-Host ('     {0}' -f $t.Desc) -ForegroundColor DarkGray
    }
    Write-Host ''
    Write-Host '  Q. Вихід' -ForegroundColor Yellow
    Write-Host ''
    $in = Read-Host '  Оберіть номер'
    if ($in -match '^\s*[qQйЙ]\s*$') { break }
    $n = 0
    if (-not [int]::TryParse($in, [ref]$n) -or $n -lt 1 -or $n -gt $Tools.Count) { continue }
    $t = $Tools[$n - 1]
    $full = Join-Path $Root $t.Path
    if (-not (Test-Path $full)) {
        Write-Host "  Файл не знайдено: $full" -ForegroundColor Red
        Read-Host '  Enter - назад' | Out-Null
        continue
    }
    $argList = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $full)
    if ($t.Args) { $argList += $t.Args }
    Clear-Host
    & powershell.exe @argList
    $code = $LASTEXITCODE
    Write-Host ''
    if ($code -ne 0) {
        Write-Host "  Інструмент завершився з помилкою (код $code)." -ForegroundColor Red
        Read-Host '  Enter - назад до меню' | Out-Null
    } else {
        Read-Host '  Готово. Enter - назад до меню' | Out-Null
    }
}
