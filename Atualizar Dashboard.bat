@echo off
chcp 65001 >nul
setlocal
set PYTHONIOENCODING=utf-8
cd /d "%~dp0scripts"

:build
echo ==========================================================
echo  Atualizando Dashboard Horas - Medeiros
echo ==========================================================
echo.

python build_dashboard.py
if errorlevel 1 (
    echo.
    echo ==========================================================
    echo  ERRO ao gerar o dashboard.
    echo  Causa mais comum: o arquivo Excel esta aberto no Excel.
    echo ==========================================================
    echo.
    choice /c SN /m "Feche o Excel e digite S para tentar de novo, ou N para sair"
    if errorlevel 2 goto fim
    if errorlevel 1 goto build
) else (
    echo.
    echo ==========================================================
    echo  Dashboard atualizado com sucesso!
    echo  Arquivo: "%~dp0Dashboard Horas - Medeiros.html"
    echo ==========================================================

    echo.
    echo Atualizando Dashboard BI...
    python build_dashboard_bi.py
    if errorlevel 1 (
        echo.
        echo ==========================================================
        echo  AVISO: o Dashboard Horas foi atualizado, mas houve um erro
        echo  ao gerar o Dashboard BI. O Dashboard Horas abaixo esta OK;
        echo  rode "python build_dashboard_bi.py" na pasta scripts pra
        echo  ver o detalhe do erro do Dashboard BI.
        echo ==========================================================
    ) else (
        echo  Dashboard BI atualizado: "%~dp0Dashboard BI - Medeiros.html"
    )

    call :github

    start "" "%~dp0Dashboard Horas - Medeiros.html"
)

:fim
echo.
pause
goto :eof

:github
echo.
echo Atualizando STATUS.md e enviando ao GitHub...
python update_status.py
if errorlevel 1 (
    echo  AVISO: nao foi possivel gerar o STATUS.md.
    exit /b 0
)
git -C "%~dp0." add -A
git -C "%~dp0." diff --cached --quiet
if errorlevel 1 (
    git -C "%~dp0." commit -q -m "Atualiza dashboard e STATUS.md - %date% %time%"
    git -C "%~dp0." push -q
    if errorlevel 1 (
        echo  AVISO: commit feito, mas o envio ao GitHub falhou. Rode "git push" depois.
    ) else (
        echo  GitHub atualizado.
    )
) else (
    echo  Nada novo para enviar ao GitHub.
)
exit /b 0
