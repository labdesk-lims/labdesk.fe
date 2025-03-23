@echo off
for /f "delims=" %%a in ('dir /r /s /b \\labdesk-server\setup\*.ver') do set "name=%%~nxa"

mkdir %LOCALAPPDATA%\Labdesk

if exist %LOCALAPPDATA%\Labdesk\%name% (
    goto start
) else (
    echo "Update application . . ."
    del %LOCALAPPDATA%\Labdesk\.*
    robocopy "\\labdesk-server\setup" "%LOCALAPPDATA%\Labdesk" /mir >nul
    reg import %LOCALAPPDATA%\Labdesk\assets\security-patch.reg
)

:start
echo "Start application . . ."
start %LOCALAPPDATA%\Labdesk\bin\labdesk_fe.accdr