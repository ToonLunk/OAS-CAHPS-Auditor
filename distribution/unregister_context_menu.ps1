# unregister_context_menu.ps1
# Removes "Audit All Excel Files" context menu from folder background
# Must be run as Administrator

Write-Host "Removing context menu for folders..." -ForegroundColor Cyan

$folderShellKey = "Registry::HKEY_CLASSES_ROOT\Directory\Background\shell\AuditAll"

try {
    if (Test-Path $folderShellKey) {
        Remove-Item -Path $folderShellKey -Recurse -Force
        Write-Host "SUCCESS: Context menu removed!" -ForegroundColor Green
        exit 0
    } else {
        Write-Host "Context menu was not installed." -ForegroundColor Yellow
        exit 0
    }
} catch {
    Write-Host "ERROR: Failed to remove context menu" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    exit 1
}
