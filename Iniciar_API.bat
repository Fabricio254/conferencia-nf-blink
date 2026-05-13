@echo off
title Conferencia NF - API FastAPI
color 0A

setlocal enabledelayedexpansion

set APPDIR=%~dp0
set TOKEN=seu_token_super_secreto_aqui

echo ============================================
echo   Conferencia NF - API FastAPI
echo ============================================
echo.

if not exist "venv\Scripts\python.exe" (
    echo ERRO: nao encontrei o ambiente virtual em venv\Scripts\python.exe
    pause
    exit /b 1
)

echo [1/2] Verificando dependencias...
venv\Scripts\pip.exe install -q -r "%APPDIR%requirements_api.txt"
if errorlevel 1 (
    echo ERRO: Falha ao instalar dependencias
    pause
    exit /b 1
)

echo [2/2] Iniciando API FastAPI...
echo.
echo  URL LOCAL:  http://localhost:8000
echo  SWAGGER UI: http://localhost:8000/docs
echo  REDOC:      http://localhost:8000/redoc
echo.
echo  Token: %TOKEN%
echo.

set API_TOKEN=%TOKEN%
venv\Scripts\python.exe -m uvicorn conferencia_nf_api:app --host 0.0.0.0 --port 8000 --reload

if errorlevel 1 (
    echo.
    echo ERRO ao iniciar API
    pause
)
