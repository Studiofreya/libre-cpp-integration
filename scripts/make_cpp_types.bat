rem Start path
set STARTDIR=%cd%

rem Libre paths. May have to be adjusted
set LIBRE=C:\Program Files (x86)\LibreOffice 5
set LIBREPROGRAM=%LIBRE%\program
set LIBRESDK=%LIBRE%\sdk\bin

set PATH=%LIBREPROGRAM%;%LIBRESDK%;%PATH%

set STAGE=%STARTDIR%\out
set TYPES=%LIBREPROGRAM%\types.rdb
set OOVBAAPI=%LIBREPROGRAM%\types\oovbaapi.rdb
set OFFAPI=%LIBREPROGRAM%\types\offapi.rdb

rem Alternate STAGE directory, uncomment and run as Administrator
rem to generate headers into the LibreOffice SDK incldue directory
rem set STAGE=C:\Program Files (x86)\LibreOffice 5\sdk\include

cppumaker -O "%STAGE%" "%TYPES%" "%OOVBAAPI%" "%OFFAPI%"


pause