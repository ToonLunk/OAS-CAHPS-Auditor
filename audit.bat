@echo off
setlocal

REM If the first argument is --all, process every file in the current folder
if "%~1"=="--all" (
    for %%f in (*.*) do (
        echo Running audit on %%f
        python "C:\Users\Aaliy\OneDrive\Desktop\TYLER\Notes and Things\scripts\audit.py" "%%f"
    )
) else (
    REM Otherwise, just pass arguments through to audit.py
    python "C:\Users\Aaliy\OneDrive\Desktop\TYLER\Notes and Things\scripts\audit.py" %*
)

endlocal