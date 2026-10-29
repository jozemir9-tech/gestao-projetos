@echo off
rem App Gestao de Projetos - sobe o servidor e abre o navegador quando o app
rem responder (se o servidor ja estiver rodando, apenas abre o navegador).
rem Requer Python no PATH com as dependencias instaladas (requirements.txt).
cd /d "%~dp0"
start "" /min python -m streamlit run app.py
powershell -NoProfile -Command "$fim=(Get-Date).AddSeconds(90); do { try { if ((Invoke-WebRequest -UseBasicParsing -TimeoutSec 1 -Uri 'http://localhost:8501/_stcore/health').Content -eq 'ok') { Start-Process 'http://localhost:8501'; exit 0 } } catch {}; Start-Sleep -Milliseconds 500 } while ((Get-Date) -lt $fim); Start-Process 'http://localhost:8501'"
