@echo off
chcp 1252 >nul
title Criacao de Usuario - Active Directory + Microsoft 365

:MENU
cls
echo.
echo  ============================================
echo    CRIACAO DE USUARIO - AD + MICROSOFT 365
echo  ============================================
echo.

set "PRIMEIRO_NOME="
set "SOBRENOME="
set "TEMPLATE="
set "USERNAME="

set /p "PRIMEIRO_NOME=  [1/3] Digite o PRIMEIRO NOME do novo usuario: "
if "%PRIMEIRO_NOME%"=="" (
    echo.
    echo  [ERRO] O primeiro nome nao pode ser vazio!
    pause
    goto MENU
)

set /p "SOBRENOME=  [2/3] Digite o SOBRENOME do novo usuario: "
if "%SOBRENOME%"=="" (
    echo.
    echo  [ERRO] O sobrenome nao pode ser vazio!
    pause
    goto MENU
)

set /p "TEMPLATE=  [3/3] Digite o LOGIN do usuario TEMPLATE (ex: joao.silva): "
if "%TEMPLATE%"=="" (
    echo.
    echo  [ERRO] O template nao pode ser vazio!
    pause
    goto MENU
)

for /f "usebackq delims=" %%i in (`powershell -NoProfile -Command "('%PRIMEIRO_NOME%'.Trim() + '.' + '%SOBRENOME%'.Trim()).ToLower().Replace(' ','')"`) do set "USERNAME=%%i"

if "%USERNAME%"=="" (
    echo.
    echo  [ERRO] Falha ao gerar o username!
    pause
    goto MENU
)

echo.
echo  ============================================
echo   CONFIRME OS DADOS
echo  ============================================
echo   Nome completo : %PRIMEIRO_NOME% %SOBRENOME%
echo   Login         : %USERNAME%
echo   Template      : %TEMPLATE%
echo  ============================================
echo.
set /p "CONFIRMA=  Confirma a criacao? (S/N): "

if /i "%CONFIRMA%"=="S" goto CRIAR
if /i "%CONFIRMA%"=="N" goto CANCELAR
goto MENU

:CRIAR
echo.
echo  Iniciando criacao do usuario, aguarde...
echo.

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\criar_usuario_ad.ps1" "%PRIMEIRO_NOME%" "%SOBRENOME%" "%USERNAME%" "%TEMPLATE%"

echo.
echo  ============================================
echo   Pressione qualquer tecla para criar outro...
echo  ============================================
pause >nul
goto MENU

:CANCELAR
echo.
echo  Operacao cancelada pelo usuario.
echo.
pause
goto MENU
