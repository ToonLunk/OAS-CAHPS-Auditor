# register_context_menu.ps1
# Adds "Audit All Excel Files" context menu to folder background
# Must be run as Administrator

$installDir = "C:\OAS-CAHPS-Auditor"
$exePath = "$installDir\audit.exe"

# Check if audit.exe exists
if (-not (Test-Path $exePath)) {
    Write-Host "ERROR: audit.exe not found at $exePath" -ForegroundColor Red
    Write-Host "Please install the auditor first using deploy.bat" -ForegroundColor Yellow
    exit 1
}

Write-Host "Installing context menu for folders..." -ForegroundColor Cyan

# Register for folders (Right-click inside folder background)
$folderShellKey = "Registry::HKEY_CLASSES_ROOT\Directory\Background\shell\AuditAll"
try {
    New-Item -Path $folderShellKey -Force | Out-Null
    Set-ItemProperty -Path $folderShellKey -Name "(Default)" -Value "Audit All OAS Files"
    Set-ItemProperty -Path $folderShellKey -Name "Icon" -Value "$exePath,0"

    $folderCommandKey = "$folderShellKey\command"
    New-Item -Path $folderCommandKey -Force | Out-Null
    Set-ItemProperty -Path $folderCommandKey -Name "(Default)" -Value "cmd.exe /c cd /d `"%V`" && `"$exePath`" --all && pause"

    Write-Host "SUCCESS: Context menu installed!" -ForegroundColor Green
    Write-Host ""
    Write-Host "You can now right-click inside any folder and select" -ForegroundColor White
    Write-Host "'Audit All OAS Files' to audit all OAS .xlsx files in that folder." -ForegroundColor White
    exit 0
} catch {
    Write-Host "ERROR: Failed to register context menu" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    exit 1
}
