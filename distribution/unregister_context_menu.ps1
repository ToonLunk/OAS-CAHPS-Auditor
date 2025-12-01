# unregister_context_menu.ps1
# Removes "Audit All Excel Files" context menu from folder background
# Must be run as Administrator

Write-Host "Removing context menu for folders..." -ForegroundColor Cyan

$folderShellKey = "Registry::HKEY_CLASSES_ROOT\Directory\Background\shell\AuditAll"
$win11Key = "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Classes\Directory\Background\shell\AuditAll"

try {
    $removed = $false
    
    if (Test-Path $folderShellKey) {
        Remove-Item -Path $folderShellKey -Recurse -Force
        Write-Host "Removed legacy context menu" -ForegroundColor Green
        $removed = $true
    }
    
    if (Test-Path $win11Key) {
        Remove-Item -Path $win11Key -Recurse -Force
        Write-Host "Removed Windows 11 context menu" -ForegroundColor Green
        $removed = $true
    }
    
    # Remove file-level context menu entries
    $extensions = @(".xlsx", ".xls", ".xlsm")
    foreach ($ext in $extensions) {
        $fileKey = "Registry::HKEY_CLASSES_ROOT\SystemFileAssociations\$ext\shell\AuditFile"
        if (Test-Path $fileKey) {
            Remove-Item -Path $fileKey -Recurse -Force
            $removed = $true
        }
    }
    
    if ($removed) {
        Write-Host "Removed Excel file context menu" -ForegroundColor Green
        Write-Host ""
        Write-Host "SUCCESS: Context menu removed!" -ForegroundColor Green
    } else {
        Write-Host "Context menu was not installed." -ForegroundColor Yellow
    }
    exit 0
} catch {
    Write-Host "ERROR: Failed to remove context menu" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    exit 1
}
