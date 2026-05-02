@echo off
python scrape_rottentomatoes.py
set EXIT_CODE=%ERRORLEVEL%

echo Press any key to close . . .
pause >nul
exit /b %EXIT_CODE%
