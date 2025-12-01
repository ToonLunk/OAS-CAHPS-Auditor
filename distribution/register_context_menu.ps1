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
    $command = 'cmd.exe /c cd /d "%V" && "' + $exePath + '" --all && pause'
    Set-ItemProperty -Path $folderCommandKey -Name "(Default)" -Value $command

    # Register for individual Excel files (.xlsx, .xls, .xlsm)
    Write-Host "Installing context menu for individual Excel files..." -ForegroundColor Cyan
    
    $extensions = @(".xlsx", ".xls", ".xlsm")
    foreach ($ext in $extensions) {
        $fileKey = "Registry::HKEY_CLASSES_ROOT\SystemFileAssociations\$ext\shell\AuditFile"
        New-Item -Path $fileKey -Force | Out-Null
        Set-ItemProperty -Path $fileKey -Name "(Default)" -Value "Audit This OAS File"
        Set-ItemProperty -Path $fileKey -Name "Icon" -Value "$exePath,0"
        
        $fileCommandKey = "$fileKey\command"
        New-Item -Path $fileCommandKey -Force | Out-Null
        Set-ItemProperty -Path $fileCommandKey -Name "(Default)" -Value "cmd.exe /c `"$exePath`" `"%1`" && pause"
    }
    Write-Host "Registered for Excel files (.xlsx, .xls, .xlsm)" -ForegroundColor Green

    Write-Host ""
    Write-Host "SUCCESS: Context menu installed!" -ForegroundColor Green
    Write-Host ""
    Write-Host "Folder auditing: Right-click inside any folder → 'Audit All OAS Files'" -ForegroundColor White
    Write-Host "File auditing: Right-click any Excel file → 'Audit This OAS File'" -ForegroundColor White
    exit 0
} catch {
    Write-Host "ERROR: Failed to register context menu" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    exit 1
}
