Attribute VB_Name = "LibImporTelMovil"
Option Explicit
Dim mdbTel As TelBaseDatos
Dim mdbFac As TelBaseDatos
Dim mdbCon As TelBaseDatos


Dim Contador As Long









'ANTIGUO
Public Function obtenNumeroFacturaOLD(serie As String, ano As Integer) As Integer
    '-- Obtiene el siguiente número de factura de la serie y año determinado
    '   actualizando en consonancia el contador
    Dim mdb As TelBaseDatos
    Dim mrs As ADODB.Recordset
    Dim mSQL As String
    Set mdb = New TelBaseDatos
    mdb.abrir "myGesSocial", "root", "aritel"
    mdb.tipo = "MYSQL"
    mSQL = "select Max(numero) from numfactelefonia " & _
            " where Serie = " & mdb.texto(serie) & _
            " and Ano = " & mdb.numero(ano)
    Set mrs = mdb.cursor(mSQL)
    If IsNull(mrs.Fields(0)) Then
        '-- No hay siquiera contador para la serie, lo creamos y devolvemos el
        '   valor adecuado
        obtenNumeroFacturaOLD = 1
        mSQL = "insert into numfactelefonia values ("
        mSQL = mSQL & mdb.texto(serie) & ","
        mSQL = mSQL & mdb.numero(ano) & ","
        mSQL = mSQL & mdb.numero(obtenNumeroFacturaOLD) & ")"
        mdb.ejecutar mSQL
    Else
        '-- Ya existe el contador lo incrementamos, devolvemos valor
        '   y guardamos el contador actualizado.
        obtenNumeroFacturaOLD = mrs.Fields(0) + 1
        mSQL = " update numfactelefonia set "
        mSQL = mSQL & " numero = " & mdb.numero(obtenNumeroFacturaOLD)
        mSQL = mSQL & " where serie = " & mdb.texto(serie)
        mSQL = mSQL & " and ano = " & mdb.numero(ano)
        mdb.ejecutar mSQL
    End If
    '-- Limpiamos la morralla
    mrs.Close
    Set mrs = Nothing
    Set mdb = Nothing
End Function
'AHORA
Public Function obtenNumeroFactura(serie As String, ano As Integer) As Long
    'Se cogeran los datos de contadores
    If Contador = 0 Then Contador = 1200000
    Contador = Contador + 1
    obtenNumeroFactura = Contador
End Function


