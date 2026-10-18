@echo off
cd /d "%~dp0"
start "" http://127.0.0.1:8765/
python -m uvicorn app.liste_sunucu:app --host 127.0.0.1 --port 8765
