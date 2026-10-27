@echo off
rem Resmi liste ekranini acar. Hata olursa pencere kapanmaz, mesaj okunabilir.
setlocal
cd /d "%~dp0"
title Resmi Listeler

rem Python: once py baslaticisi, sonra python. Microsoft Store kisayolu calismadigi icin denenerek secilir.
set "PY="
py -3 -c "import sys; assert sys.version_info >= (3, 10)" >nul 2>&1 && set "PY=py -3"
if not defined PY python -c "import sys; assert sys.version_info >= (3, 10)" >nul 2>&1 && set "PY=python"
if not defined PY (
  echo Python 3.10 veya ustu bulunamadi.
  echo python.org adresinden kurun ^(kurulumda "Add python.exe to PATH" isaretli olsun^) ya da BT'den isteyin.
  goto hata
)

%PY% -c "import fastapi, uvicorn, openpyxl" >nul 2>&1
if errorlevel 1 (
  echo Gerekli Python paketleri kuruluyor...
  %PY% -m pip install --user -r requirements.txt
  %PY% -c "import fastapi, uvicorn, openpyxl" >nul 2>&1
  if errorlevel 1 (
    echo.
    echo Paketler kurulamadi. Yukaridaki hatayi BT'ye iletin ya da su komutu deneyin:
    echo   %PY% -m pip install --user -r "%cd%\requirements.txt"
    goto hata
  )
)

rem Tarayici, sunucu ayaga kalktiktan sonra acilir.
start "" /min cmd /c "timeout /t 3 /nobreak >nul & start http://127.0.0.1:8765/"
echo Ekran: http://127.0.0.1:8765/   ^(kapatmak icin bu pencerede Ctrl+C^)
%PY% -m uvicorn app.liste_sunucu:app --host 127.0.0.1 --port 8765
if errorlevel 1 goto hata
exit /b 0

:hata
echo.
pause
exit /b 1
