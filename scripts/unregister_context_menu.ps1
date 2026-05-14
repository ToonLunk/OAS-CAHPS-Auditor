# unregister_context_menu.ps1
# Removes "Audit All Excel Files" context menu from folder background

Write-Host "Removing context menu for folders..." -ForegroundColor Cyan

# Remove context menu entry
$folderShellKey = "Registry::HKEY_CURRENT_USER\Software\Classes\Directory\Background\shell\AuditAll"

try {
    $removed = $false
    
    if (Test-Path $folderShellKey) {
        Remove-Item -Path $folderShellKey -Recurse -Force
        Write-Host "Removed context menu" -ForegroundColor Green
        $removed = $true
    }
    
    # Remove file-level context menu entries
    $extensions = @(".xlsx", ".xls", ".xlsm")
    foreach ($ext in $extensions) {
        $fileKey = "Registry::HKEY_CURRENT_USER\Software\Classes\SystemFileAssociations\$ext\shell\AuditFile"
        if (Test-Path $fileKey) {
            Remove-Item -Path $fileKey -Recurse -Force
            $removed = $true
        }
    }
    
    if ($removed) {
        Write-Host "Removed Excel file context menu" -ForegroundColor Green
    }
    
    if ($removed) {
        Write-Host ""
        Write-Host "SUCCESS: Context menu removed!" -ForegroundColor Green
    } else {
        Write-Host "Context menu was not installed." -ForegroundColor Yellow
    }
    Write-Host ""
    pause
    exit 0
} catch {
    Write-Host "ERROR: Failed to remove context menu" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host ""
    pause
    exit 1
}
