@echo off
chcp 65001 >nul
setlocal
echo === Tel.eng.br - Stick Corte v0.12 beta - Tickavel CAD ===
set PY_CMD=
py --version >nul 2>&1 && set PY_CMD=py
if not defined PY_CMD (python --version >nul 2>&1 && set PY_CMD=python)
if not defined PY_CMD (python3 --version >nul 2>&1 && set PY_CMD=python3)
if not defined PY_CMD (
  echo Python nao encontrado!
  pause
  exit /b
)
%PY_CMD% -m venv .venv
call .venv\Scripts\activate.bat
python -m pip install --upgrade pip
pip install -r requirements.txt
rmdir /s /q build 2>nul
rmdir /s /q dist 2>nul
del /q "Tel.eng.br - Stick Corte.spec" 2>nul
echo Gerando EXE v0.12 com icone LSF...
python -m PyInstaller --noconfirm --clean --onefile --windowed --name "Tel.eng.br - Stick Corte v0.12 beta" --icon=icon_profiles.ico --collect-all openpyxl --hidden-import openpyxl.cell._writer --add-data "assets;assets" main.py
if exist "dist\Tel.eng.br - Stick Corte v0.12 beta.exe" (
  echo SUCESSO! EXE v0.12 em dist\
) else (
  echo Tentando sem icone...
  python -m PyInstaller --noconfirm --clean --onefile --windowed --name "Tel.eng.br - Stick Corte v0.12 beta" --collect-all openpyxl main.py
)
pause
