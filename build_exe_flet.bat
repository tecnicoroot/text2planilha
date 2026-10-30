@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ==================================================
echo   Gerando o executavel "Conversor de Glosas" (Flet)
echo ==================================================
echo.
python -m pip install --upgrade pip
python -m pip install -r requirements-flet.txt
echo.
flet pack app_flet.py ^
  --name "ConversorGlosas" ^
  --icon "icone.ico" ^
  --product-name "Conversor de Glosas" ^
  --product-version "2.0" ^
  --file-description "Converte o relatorio de glosas TXT em Excel" ^
  --hidden-import openpyxl ^
  -y
echo.
if exist "dist\ConversorGlosas.exe" (
  echo Pronto! O programa esta em: dist\ConversorGlosas.exe
) else (
  echo Algo deu errado. Veja as mensagens acima.
)
pause
