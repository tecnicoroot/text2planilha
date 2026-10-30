@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ============================================
echo   Gerando o executavel "Conversor de Glosas"
echo ============================================
echo.
python -m pip install --upgrade pip
python -m pip install -r requirements.txt pyinstaller
echo.
python -m PyInstaller --noconfirm --clean --onefile --windowed ^
  --name "ConversorGlosas" ^
  --icon "icone.ico" ^
  --paths "." ^
  app_desktop.py
echo.
if exist "dist\ConversorGlosas.exe" (
  echo Pronto! O programa esta em: dist\ConversorGlosas.exe
) else (
  echo Algo deu errado. Veja as mensagens acima.
)
pause
