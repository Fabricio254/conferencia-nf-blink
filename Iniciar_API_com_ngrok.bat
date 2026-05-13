@echo off
REM Script para iniciar API FastAPI + ngrok
REM Expõe a API com URL pública e domínio fixo

title Conferencia NF - API + ngrok
color 0A

setlocal enabledelayedexpansion

set APPDIR=%~dp0
set API_TOKEN=seu_token_super_secreto_aqui
set NGROK_DOMAIN=conferencia-nf-blink.ngrok-free.dev

echo ============================================
echo   Conferencia NF - API FastAPI + ngrok
echo ============================================
echo.

REM Verificar se ngrok está instalado
where ngrok >nul 2>&1
if %errorlevel% neq 0 (
    echo ERRO: ngrok nao encontrado no PATH
    echo Instale via: choco install ngrok (ou download de https://ngrok.com)
    pause
    exit /b 1
)

REM Verificar venv
if not exist "venv\Scripts\python.exe" (
    echo ERRO: nao encontrei o ambiente virtual
    pause
    exit /b 1
)

echo [1/3] Instalando dependencias...
venv\Scripts\pip.exe install -q -r "%APPDIR%requirements_api.txt" 2>nul

echo [2/3] Iniciando API FastAPI (porta 8000)...
start /B "API" cmd /k "cd /d %APPDIR% && set API_TOKEN=%API_TOKEN% && venv\Scripts\python.exe -m uvicorn conferencia_nf_api:app --host 0.0.0.0 --port 8000"

timeout /t 3 /nobreak

echo [3/3] Iniciando tunel ngrok...
echo.
echo ============================================
echo  API Local:   http://localhost:8000
echo  API Publico: https://%NGROK_DOMAIN%
echo  Swagger:     https://%NGROK_DOMAIN%/docs
echo  Token:       %API_TOKEN%
echo ============================================
echo.

start /B "ngrok" cmd /k "ngrok http --domain=%NGROK_DOMAIN% 8000"

timeout /t 5 /nobreak

echo.
echo PRONTO! Acesse a API em:
echo   https://%NGROK_DOMAIN%/api/health
echo.
echo Para parar, feche ambas as janelas (API e ngrok).
echo.

pause
