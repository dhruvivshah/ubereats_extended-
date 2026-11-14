@echo off
REM One-time setup (and again only if an account gets signed out).
cd /d "%~dp0"
UberEatsScraper.exe login
pause
