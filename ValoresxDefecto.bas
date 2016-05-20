Attribute VB_Name = "ValoresxDefecto"



'Dos metodos publicos
' CheckValueGuardar:
'                   para un form guardará el chcek de vistaprevia como esta
' CheckValueLeer:
'                   asiganr directamente a un check de vista previa su valor

'Guardar BYTE


Private Function DevNombreFichero(Nombre As String) As String

'Select Case NombreForm
'Case "frmDiag"
'
'Case ""
'
'Case Else
'    NombreFichero = ""
'End Select
DevNombreFichero = App.Path & "\" & Nombre & ".xdf"
End Function



Public Function CheckValueLeer(NombreForm As String) As Byte
Dim NombreFichero As String

On Error GoTo ECheckValueLeer
CheckValueLeer = 0  'UNCHECKED
'Se podria hacer un select para que no lie mucho los nombres en las carpetas
NombreFichero = DevNombreFichero(NombreForm)
If NombreFichero <> "" Then
    If Dir(NombreFichero) <> "" Then CheckValueLeer = 1
End If


Exit Function
ECheckValueLeer:
    Err.Clear
End Function



Public Sub CheckValueGuardar(NombreForm As String, ValorDelCheck As Byte)
Dim NombreFichero As String
'Dim ExisteFich As Boolean
On Error GoTo ECheckValueGuardar

'Se podria hacer un select para que no lie mucho los nombres en las carpetas
NombreFichero = DevNombreFichero(NombreForm)
If NombreFichero = "" Then Exit Sub
'ExisteFich = (Dir(NombreFichero) <> "")
If ValorDelCheck = 0 Then
    'Hay que eliminar si existe
    EliminaValoresPorDefecto NombreFichero
    Else
        CrearFichValoresPorDefecto NombreFichero
End If
    
Exit Sub
ECheckValueGuardar:
    Err.Clear
End Sub


Private Sub EliminaValoresPorDefecto(ByRef vPath As String)

On Error GoTo EeliminavaloresPorDefecto
If Dir(vPath, vbArchive) <> "" Then Kill vPath
EeliminavaloresPorDefecto:
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub CrearFichValoresPorDefecto(ByRef vPath As String)
Dim NF As Integer
On Error GoTo ECrearFichValoresPorDefecto
If Dir(vPath, vbArchive) = "" Then
    NF = FreeFile
    Open vPath For Output As #NF
    Print #NF, "Check = True"
    Close #NF
End If
Exit Sub
ECrearFichValoresPorDefecto:
    Err.Clear
End Sub



'Valores por defecto pero que no sean CHECK
'Son tipo BTE
Public Function ByteValueLeer(NombreForm As String) As Byte
Dim NombreFichero As String

On Error GoTo ECheckValueLeer
ByteValueLeer = 0
'Se podria hacer un select para que no lie mucho los nombres en las carpetas
NombreFichero = DevNombreFichero(NombreForm)
If NombreFichero <> "" Then
    If Dir(NombreFichero) <> "" Then
        FicheroByte True, NombreFichero, ByteValueLeer
    End If
End If


Exit Function
ECheckValueLeer:
    Err.Clear
End Function

Public Sub ByteValueGuardar(NombreForm As String, Valor As Byte)
 Dim NombreFichero  As String
    'Se podria hacer un select para que no lie mucho los nombres en las carpetas
    NombreFichero = DevNombreFichero(NombreForm)
    If NombreFichero = "" Then Exit Sub
    If Valor > 128 Then Valor = 128
    If Valor = 0 Then
        'Hay que eliminar si existe
        EliminaValoresPorDefecto NombreFichero
    Else
        FicheroByte False, NombreFichero, Valor
    End If
        
    

    
    
End Sub



Private Sub FicheroByte(Leer As Boolean, NomFich As String, ByRef Resultado As Byte)
Dim Cad As String
Dim NF As Integer
    
    On Error Resume Next
    NF = FreeFile
    If Leer Then
        Open NomFich For Input As #NF
        Line Input #NF, Cad
        Close #NF
        If Not IsNumeric(Cad) Then
            Cad = "0"
        Else
            If Val(Cad) > 128 Then Cad = "128"
        End If
        Resultado = CByte(Cad)
    Else
        Open NomFich For Output As #NF
        Print #NF, Resultado
        Close #NF
        
    End If
    
    Err.Clear
End Sub
