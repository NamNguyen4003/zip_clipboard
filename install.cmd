<# : phan cmd: chay PowerShell doc lai chinh file nay
@echo off
setlocal
set "SELF=%~f0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "& ([scriptblock]::Create((Get-Content -LiteralPath $env:SELF -Raw)))"
exit /b %errorlevel%
#>
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
try {
    [Net.ServicePointManager]::SecurityProtocol = 'Tls12'
    $shell = New-Object -ComObject Shell.Application
    $dir = $shell.NameSpace('shell:Downloads').Self.Path
    $bat = Join-Path $dir 'Nen-Clipboard.bat'
    # Tai ra file tam de mat mang giua chung khong lam hong ban dang dung.
    Invoke-WebRequest -UseBasicParsing 'https://raw.githubusercontent.com/NamNguyen4003/zip_clipboard/main/Nen-Clipboard.bat' -OutFile "$bat.tmp"
    Move-Item -LiteralPath "$bat.tmp" -Destination $bat -Force

    # Shell.Application thay cho WScript.Shell, vi WScript.Shell hong voi duong dan co dau tieng Viet.
    $desktop = [Environment]::GetFolderPath('Desktop')
    [IO.File]::WriteAllBytes((Join-Path $desktop 'Nen-Clipboard.lnk'), @())
    $link = $shell.NameSpace($desktop).ParseName('Nen-Clipboard.lnk').GetLink
    $link.Path = $bat
    $link.WorkingDirectory = $dir
    $link.Hotkey = 0x643    # Ctrl+Alt+C
    $link.ShowCommand = 7   # Minimized
    $link.Save()
    Write-Host "Da cai: $bat"
    Write-Host 'Copy trang OneNote, nhan Ctrl+Alt+C, cho thong bao, roi Ctrl+V vao Outlook.'
} catch {
    Write-Host "Cai that bai: $($_.Exception.Message)"
    exit 1
}
