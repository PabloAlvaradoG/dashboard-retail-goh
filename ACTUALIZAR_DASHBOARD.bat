@echo off
chcp 65001 >nul
title Actualizar Dashboard Gerencial

cd /d "%~dp0"

python --version >nul 2>&1
if errorlevel 1 (
    echo Python no esta instalado. Descargalo en https://www.python.org
    pause
    exit /b 1
)

echo.
echo IMPORTANTE: cierra Historico_auditorias_retail.xlsx antes de continuar
echo (si esta abierto en Excel, la actualizacion va a fallar).
echo.
pause

echo Generando Dashboard_Gerencial.html con los datos actuales...
python generar_dashboard.py
if errorlevel 1 (
    echo.
    echo Hubo un error al generar el dashboard. Revisa el mensaje de arriba.
    pause
    exit /b 1
)

echo.
echo Publicando en GitHub Pages...
git add Dashboard_Gerencial.html generar_dashboard.py plantilla_dashboard.html
git commit -m "Actualizar dashboard" >nul 2>&1
git push
if errorlevel 1 (
    echo.
    echo No se pudo subir a GitHub. Revisa tu conexion o sesion de git e intenta de nuevo con: git push
    pause
    exit /b 1
)

echo.
echo Listo. Cambios en linea en unos segundos: https://pabloalvaradog.github.io/dashboard-retail-goh/Dashboard_Gerencial.html
echo Abriendo copia local...
start "" "Dashboard_Gerencial.html"
pause
