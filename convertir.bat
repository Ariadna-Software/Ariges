@echo off
IF  EXIST "%PROGRAMFILES(X86)%" (
set CX="%ProgramFiles(x86)%\Softinterface, Inc\Convert XLS\ConvertXLS.EXE"
) ELSE (
set CX="C:\Program Files\Softinterface, Inc\Convert XLS\ConvertXLS.EXE"
)
set InputFolder=C:\PedidosAriadna
set OutputFolder=C:\PedidosAriadna
%CX% /S%InputFolder%\convertir.xlsx /T%OutputFolder%\convertir.CSV    /F51 /C6 /M1
set anio=%date:~6,4%

set mes=%date:~3,2%

set dia=%date:~0,2%

set hora=%time:~0,2%

set hora=%hora: =0%

set minuto=%time:~3,2%

set segundo=%time:~6,2%

move %InputFolder%\convertir.xlsx %InputFolder%\procesados\convertir-%anio%%mes%%dia%%hora%%minuto%%segundo%.xlsx

