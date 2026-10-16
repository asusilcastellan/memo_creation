@echo off
REM Build the Windows app. Run from the project folder with the .venv activated:
REM     .venv\Scripts\activate
REM     build.bat
pyinstaller memo_app.spec --noconfirm --clean || goto :error

REM Editable files live next to the exe so the team can tune wording without rebuilding.
xcopy /E /I /Y resources dist\CastellanMemo\resources >nul
copy /Y .env.example dist\CastellanMemo\.env.example >nul
if exist .env copy /Y .env dist\CastellanMemo\.env >nul

echo.
echo Built: dist\CastellanMemo\CastellanMemo.exe
echo Put the .env file next to CastellanMemo.exe before sharing the folder.
goto :eof
:error
echo Build failed.
exit /b 1
