@echo off
REM Бърз старт на Chrome за скрипта (без копиране на профил).
REM Ползвай Start_Chrome.bat (пълния) само ако трябва да опресниш профила/станцията.

set "DST=%LOCALAPPDATA%\ClockAutomationProfile"
set "URL=https://sky-eu1.clock-software.com/spa/pms-composer/#/hotel/14539/home"

set "CHROME=C:\Program Files\Google\Chrome\Application\chrome.exe"
if not exist "%CHROME%" set "CHROME=C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
if not exist "%CHROME%" set "CHROME=%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"

start "" "%CHROME%" --remote-debugging-port=9222 --user-data-dir="%DST%" --profile-directory="Default" "%URL%"
