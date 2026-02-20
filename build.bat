@echo off
title PrinterDiag - Gerador de EXE
color 0A

echo.
echo  =====================================================
echo    PrinterDiag - Compilador de EXE
echo    Isso vai gerar o arquivo PrinterDiag.exe
echo  =====================================================
echo.

:: Verifica se Python esta instalado
python --version >nul 2>&1
if errorlevel 1 (
    echo  [ERRO] Python nao encontrado. Instale em python.org
    pause
    exit /b 1
)

echo  [1/4] Python encontrado. Instalando dependencias...
python -m pip install pyinstaller pywin32 --quiet
if errorlevel 1 (
    echo  [ERRO] Falha ao instalar dependencias.
    pause
    exit /b 1
)

echo  [2/4] Dependencias instaladas com sucesso!
echo  [3/4] Compilando PrinterDiag.exe (aguarde)...
echo.

python -m PyInstaller ^
    --onefile ^
    --windowed ^
    --name "PrinterDiag" ^
    printer_diagnostics.py

if errorlevel 1 (
    echo.
    echo  [ERRO] Falha na compilacao. Verifique os erros acima.
    pause
    exit /b 1
)

echo.
echo  [4/4] EXE gerado com sucesso!
echo.
echo  =====================================================
echo    Arquivo gerado em: dist\PrinterDiag.exe
echo    Execute como ADMINISTRADOR para funcionar!
echo  =====================================================
echo.

:: Abre a pasta dist automaticamente
explorer dist

pause
