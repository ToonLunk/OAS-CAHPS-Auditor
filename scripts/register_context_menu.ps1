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

    Write-Host "✓ Registered for legacy context menu (Shift+Right-click)" -ForegroundColor Green

    # Add to Windows 11 new context menu
    # This uses the ExplorerCommandHandler which appears in the new Win11 menu
    Write-Host "✓ Attempting to register for Windows 11 new context menu..." -ForegroundColor Cyan
    
    # Create GUID for the command
    $guid = "{7C5A40EF-A0FB-4BFC-874A-C0F2E0B9FA8E}"
    
    # Register in Classes
    $win11Key = "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Classes\Directory\Background\shell\AuditAll"
    New-Item -Path $win11Key -Force | Out-Null
    Set-ItemProperty -Path $win11Key -Name "(Default)" -Value "Audit All OAS Files"
    Set-ItemProperty -Path $win11Key -Name "Icon" -Value "$exePath,0"
    
    $win11CommandKey = "$win11Key\command"
    New-Item -Path $win11CommandKey -Force | Out-Null
    $commandValue = "cmd.exe /c cd /d `"%V`" && `"$exePath`" --all && pause"
    Set-ItemProperty -Path $win11CommandKey -Name "(Default)" -Value $commandValue
    
    Write-Host "✓ Registered for Windows 11 new context menu" -ForegroundColor Green

    Write-Host ""
    Write-Host "SUCCESS: Context menu installed!" -ForegroundColor Green
    Write-Host ""
    Write-Host "Windows 10 users: Right-click inside any folder" -ForegroundColor White
    Write-Host "Windows 11 users: Right-click inside any folder (no Shift needed!)" -ForegroundColor White
    Write-Host "Then select 'Audit All OAS Files'" -ForegroundColor White
    exit 0
} catch {
    Write-Host "ERROR: Failed to register context menu" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    exit 1
}
