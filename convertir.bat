IF  EXIST "%PROGRAMFILES(X86)%" (
set CX="%ProgramFiles(x86)%\Softinterface, Inc\Convert XLS\ConvertXLS.EXE"
) ELSE (
set CX="C:\Program Files\Softinterface, Inc\Convert XLS\ConvertXLS.EXE"
)
set CX="C:\Programas\TrasTaxi\Convertir.exe"
set InputFolder=C:\PedidosAriadna
cd C:\PedidosAriadna
%CX% %InputFolder%\convertir.xlsx