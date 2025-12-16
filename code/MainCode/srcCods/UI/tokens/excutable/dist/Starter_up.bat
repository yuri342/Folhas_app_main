@echo off
setlocal

REM Caminho completo do EXE, baseado na pasta do .bat
set "EXE=%~dp0TokenRefresherSeniorX.exe"

REM Nome da chave no Registro
set "NAME=TokenRefresherSeniorX"

echo Registrando programa na inicialização do Windows...

REM Cria ou atualiza a chave de inicialização
reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Run" ^
    /v "%NAME%" ^
    /t REG_SZ ^
    /d "\"%EXE%\"" ^
    /f

if %errorlevel%==0 (
    echo Sucesso! O programa iniciará junto com o Windows.
) else (
    echo Falhou ao registrar o programa.
)

pause
