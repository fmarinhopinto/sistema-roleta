@echo off
setlocal
cd /d "%~dp0"
title Atualizador Automatico - Sistema Roleta
color 0b

echo ======================================================================
echo    ATUALIZADOR AUTOMATICO - SISTEMA ROLETA
echo    Repositorio: https://github.com/fmarinhopinto/sistema-roleta
echo ======================================================================
echo.

set "GIT_EXE=git"

echo [*] Preparando arquivos...
"%GIT_EXE%" add .

echo [*] Registrando alteracoes locais...
"%GIT_EXE%" commit -m "Atualizacao automatica - %date% %time%" >nul 2>&1

echo.
echo [*] Sincronizando com o GitHub...
"%GIT_EXE%" push origin main

if %ERRORLEVEL% equ 0 (
    echo.
    color 0a
    echo ======================================================================
    echo   [SUCESSO] Sistema sincronizado com o GitHub com exito!
    echo   O Streamlit Cloud ja detectou e esta atualizando online.
    echo ======================================================================
) else (
    echo.
    color 0c
    echo ======================================================================
    echo   [ATENCAO] Nao foi possivel sincronizar automaticamente.
    echo   Verifique sua conexao ou login no GitHub.
    echo ======================================================================
)

pause
