@echo off
REM Stops the scraper completely: the main program, all 3 accounts and their
REM Chrome windows. Your own normal Chrome is NOT touched.
echo Stopping the Uber Eats scraper...
powershell -NoProfile -Command "Get-CimInstance Win32_Process | Where-Object { $_.ProcessId -ne $PID -and $_.CommandLine -and ($_.CommandLine -match 'run_all\.py|UberEatsScraper|\\profiles\\account|ue_chrome_profile|chromedriver') } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }"
echo Done. You can start it again any time.
timeout /t 3 >nul
