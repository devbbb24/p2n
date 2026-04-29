@echo off
setlocal enabledelayedexpansion

pushd ..
set "p2n=%cd%"
set "targetDir=%p2n%\html\scripts"
set "scriptDir=%p2n%\excel"
popd

cd %scriptDir%
for %%x in (%scriptDir%\*.json) do del %%x
start /b /wait %scriptDir%\ExcelConverter.exe
cd %p2n%

for /r %scriptDir% %%F in (*.json) do (
    set "filePath=%%~dpF"
    set "relPath=!filePath:%scriptDir%\=!"

    set "dest=%targetDir%\!relPath!"
    
    if not exist "!dest!" (mkdir "!dest!")
    move "%%F" "!dest!%%~nxF"
)
endlocal

echo done.