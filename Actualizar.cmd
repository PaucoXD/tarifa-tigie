@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"
title Actualizar Tarifa TIGIE

set "PY="
where python >nul 2>nul && set "PY=python"
if not defined PY where py >nul 2>nul && set "PY=py"
if not defined PY (
  echo No encontré Python. Instálalo desde https://www.python.org y vuelve a intentar.
  pause & exit /b 1
)

echo Revisando librerías de Python...
%PY% -m pip install -q --disable-pip-version-check -r tools\requirements.txt
if errorlevel 1 ( echo No pude instalar las librerías. & pause & exit /b 1 )

echo.
echo Bajando archivos del SNICE y del DOF y regenerando la página...
echo.
%PY% tools\actualizar.py
if errorlevel 1 (
  echo.
  echo Hubo un problema. Revisa tools\ultimo_reporte.txt. No se subió nada.
  pause & exit /b 1
)

git diff --quiet -- index.html
if not errorlevel 1 (
  echo.
  echo No hay nada que subir.
  pause & exit /b 0
)

echo.
set /p OK=¿Subir la página actualizada a GitHub? (S/N):
if /i not "%OK%"=="S" (
  echo No se subió. Los cambios quedan en index.html; puedes volver a correr este archivo para subirlos.
  pause & exit /b 0
)
git add index.html tools\nom_mods.json tools\ultimo_reporte.txt
git commit -q -m "Actualización automática de la tarifa (%date%)"
git push -q
if errorlevel 1 ( echo No se pudo subir a GitHub. Revisa tu conexión o tu sesión de GitHub. & pause & exit /b 1 )
echo.
echo Listo. En 1 o 2 minutos se ve en https://paucoxd.github.io/tarifa-tigie/
pause
