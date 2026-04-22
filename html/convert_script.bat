@echo off
setlocal enabledelayedexpansion

cd %~dp0..\scripts\

for %%x in (*.json) do del %%x

start /b /wait ExcelConverter.exe

set outDir=%~dp0chapters
for %%x in (*.json) do (
    set "filename=%%~nx"
    for /f "tokens=1* delims=_" %%a in ("!filename!") do (
        set "foldername=%%a"
        set "newname=%%b"
        if not exist "!outDir!\!foldername!" (mkdir "!outDir!\!foldername!")
        move "%%x" "!outDir!\!foldername!\!newname!.json"
    )
)

echo done.