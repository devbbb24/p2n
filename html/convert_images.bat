@echo off
setlocal

where py >nul 2>&1
if errorlevel 1 (
  echo [ERROR] Python launcher^(py^)를 찾을 수 없습니다.
  echo         https://www.python.org/ 에서 Python 3을 설치하세요.
  exit /b 1
)

py -3 -c "import PIL" 2>nul
if errorlevel 1 (
  echo Pillow가 설치되어 있지 않습니다. 설치를 진행합니다...
  py -3 -m pip install --user Pillow
  if errorlevel 1 (
    echo [ERROR] Pillow 설치에 실패했습니다.
    exit /b 1
  )
)

py -3 "%~dp0convert_images.py"
endlocal
