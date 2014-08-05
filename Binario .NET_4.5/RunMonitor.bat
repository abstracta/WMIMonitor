@echo off
echo Para terminar: Ctrl + C

set rutalog=E:\WMILog

:dia
set dia=%date:~4,2%
set folderName=%rutalog%\%date:~10,4%-%date:~7,2%-%date:~4,2%
mkdir %folderName%

echo Directorio '%folderName%' creado

:top
Set fileName=%time:~0,2%.%time:~3,2%.%time:~6,2%-RMI.log
Abstracta.WMIMonitor_NET4.5.exe /File > %folderName%\%fileName%

echo Archivo log '%fileName%' creado en directorio '%folderName%'

: 5 minutos son 300 segundos

timeout /T 300 /nobreak
if NOT %dia%==%date:~4,2% goto dia
goto top