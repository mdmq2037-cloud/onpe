@echo off
chcp 65001 > nul
echo ============================================
echo  ONPE Consulta Electoral Masiva - Instalador
echo ============================================
echo.

echo [1/2] Instalando dependencias Python...
.venv\Scripts\pip.exe install undetected-chromedriver selenium openpyxl
if errorlevel 1 (
    echo ERROR: Fallo al instalar dependencias.
    pause
    exit /b 1
)

echo.
echo [2/2] Verificando instalacion...
.venv\Scripts\python.exe -c "import undetected_chromedriver, selenium, openpyxl; print('OK - Todas las dependencias instaladas')"

echo.
echo ============================================
echo  Instalacion completada con exito!
echo  Ejecuta: python onpe_consulta.py
echo ============================================
pause
