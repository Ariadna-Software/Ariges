Attribute VB_Name = "libWHose"
Option Explicit

Dim SQL As String
Dim RN As ADODB.Recordset



Public Function GenerarEstructuraPotencial(Cliente As Long) As Boolean

    On Error Resume Next
    
    GenerarEstructuraPotencial = False
    
    
    SQL = vParamAplic.PathDocsWHOSE & "\POTENC"
    If Dir(SQL, vbDirectory) = "" Then
        MkDir SQL
        If Err.Number <> 0 Then
                MuestraError Err.Number, Err.Description
                Exit Function
        End If
    End If
    
    SQL = SQL & "\P" & Format(Cliente, "000000")
    MkDir SQL
    If Err.Number <> 0 Then
        MuestraError Err.Number, Err.Description
    Else
        GenerarEstructuraPotencial = True
    End If
    
End Function

Public Function YaExisteExtructuraCliente(idCliente) As Boolean
    YaExisteExtructuraCliente = False
    
    'NO puede ser EOF
    SQL = vParamAplic.PathDocsWHOSE & "\" & Format(idCliente, "000000")
    If Dir(SQL, vbDirectory) <> "" Then YaExisteExtructuraCliente = True
End Function

'EsCrear
'   true:_  Crea la estructura y le pasa todos los archivos
'   false:  Borra la estructura para ele potencial
Public Function TratarExtructuraClienteConArchivos(EsCrear As Boolean, Potencial As Long, idCliente As Long) As Boolean
Dim V As String
Dim I As Byte

    TratarExtructuraClienteConArchivos = False
    V = vParamAplic.PathDocsWHOSE
    
    
    If Not EsCrear Then
        SQL = V & "\POTENC\P" & Format(Potencial, "000000")
        Kill SQL & "\*.*"
        If Err.Number <> 0 Then
            MsgBox Err.Description, vbExclamation
            Exit Function
        End If
        
        RmDir SQL
        If Err.Number <> 0 Then
            MsgBox Err.Description, vbExclamation
            Exit Function
        End If
        TratarExtructuraClienteConArchivos = True
        Exit Function
    End If
    
    
    'NO puede s
    SQL = V & "\" & Format(idCliente, "000000")
    On Error Resume Next
    MkDir SQL
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbExclamation
        Exit Function
    End If
    
    'RResto estructura
    CadenaDesdeOtroForm = "CONTRATO|OBRA|PI|EGD|ACTUA|"
    SQL = SQL & "\"
    For I = 1 To 5
        MkDir SQL & RecuperaValor(CadenaDesdeOtroForm, CInt(I))
        If Err.Number <> 0 Then
            MsgBox Err.Description & vbCrLf & SQL & RecuperaValor(CadenaDesdeOtroForm, CInt(I)), vbExclamation
            Exit Function
        End If
    Next
    
    'AHORA COPIAREMOS TODOS LOS ARCHIVOS QUE YA TIENE EL CLIENTE
    'DOCS\POTENC\P000002
    SQL = SQL & "CONTRATO\"
    V = V & "\POTENC\P" & Format(Potencial, "000000") & "\"
    CadenaDesdeOtroForm = Dir(V & "*.*")
    Do While CadenaDesdeOtroForm <> ""
        FileCopy V & CadenaDesdeOtroForm, SQL & CadenaDesdeOtroForm
        If Err.Number <> 0 Then
            MsgBox Err.Description & vbCrLf & V & vbCrLf & SQL, vbExclamation
            Exit Function
        End If
        CadenaDesdeOtroForm = Dir
    Loop
    
    TratarExtructuraClienteConArchivos = True

End Function




'pdf|xls|doc|docx
Public Function ExtensionSoportada(RutaArchivoCompleta As String, ParaPropContratos As Boolean) As String
Dim I As Integer

    I = InStrRev(RutaArchivoCompleta, ".")
    If I = 0 Then
        ExtensionSoportada = "Extension no soportada(I)"
    
    Else
        ExtensionSoportada = LCase(Mid(RutaArchivoCompleta, I + 1))
        If ParaPropContratos Then
            I = InStr(1, "pdf|xls|doc|docx|xlsx|", ExtensionSoportada & "|")
        End If
        If I = 0 Then ExtensionSoportada = "Extension no soportada(II)"
    End If
    
End Function





'Extension:
Public Function DevuelveNombreArhivo(Cliente As Long, PropuestaComercial As Boolean, Id As Long, extension As String) As String
Dim C As String
    C = vParamAplic.PathDocsWHOSE
    'NO puede ser EOF
    C = C & "\POTENC\P" & Format(Cliente, "000000") & "\"
    If PropuestaComercial Then
        C = C & "COM"
    Else
        C = C & "CON"
    End If
   ' C = C & Format(Id, "00000") & "." & ExtensionSoportada(ArchivoOrigen)
    
    C = C & Format(Id, "00000") & "." & extension
    DevuelveNombreArhivo = C
End Function


Public Function CopiaArchivoWHOSE(Cliente As Long, PropuestaComercial As Boolean, Id As Long, ArchivoOrigen As String) As Boolean
Dim Ext As String
    CopiaArchivoWHOSE = False
    On Error Resume Next
    
    If PropuestaComercial Then
        'Puede llevar el archivo a blancos
        If ArchivoOrigen = "" Then
            CopiaArchivoWHOSE = True
            Exit Function
        End If
    End If
    
    Ext = ExtensionSoportada(ArchivoOrigen, True)
    SQL = DevuelveNombreArhivo(Cliente, PropuestaComercial, Id, Ext)
    FileCopy ArchivoOrigen, SQL
    If Err.Number <> 0 Then
        MuestraError Err.Number, , Err.Description
    Else
        CopiaArchivoWHOSE = True
    End If
End Function


Public Function CopiaObraWHOSE(Destino As String, ArchivoOrigen As String) As Boolean
On Error Resume Next
    CopiaObraWHOSE = False
    
    
    SQL = vParamAplic.PathDocsWHOSE & Destino
    
    FileCopy ArchivoOrigen, SQL
    If Err.Number <> 0 Then
        MuestraError Err.Number, , Err.Description
    Else
        CopiaObraWHOSE = True
    End If
End Function


'Si PropuestaComercial=false--> CONTRATO
Public Sub CargaListviewWHOSE(ByRef ElListview As ListView, Potenciales As Boolean, PropuestaComercial As Boolean, idCliente As Long, PonerIcono As Boolean)
Dim It As ListItem

    ElListview.ListItems.Clear
    Set RN = New ADODB.Recordset
    
    
    If PropuestaComercial Then
        
        SQL = "f_preprop,f_rechazoprop,idPropComer ID,extension    FROM "
        If Potenciales Then
            SQL = SQL & "whoexpedientepotprocomer"
        Else
            SQL = SQL & "whoexpedientecliprocomer"
        End If
    Else
        
        If Potenciales Then
            SQL = "f_precont, f_rechazocon,idcontrato ID,extension FROM "
            SQL = SQL & "whoexpedientepotcontrato"
        Else
            SQL = "f_precont, f_rechazocon,idcontrato ID,extension,f_aceptado FROM "
            SQL = SQL & "whoexpedienteclicontrato"
        End If
    End If
    
    SQL = "SELECT " & SQL & " WHERE codclien =  " & idCliente & " ORDER BY   id desc"
    Set RN = New ADODB.Recordset
    RN.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RN.EOF
        Set It = ElListview.ListItems.Add(, "K" & Format(RN.Fields(2), "0000"))
        It.Text = Format(RN.Fields(0), "dd/mm/yyyy")
        If Not IsNull(RN.Fields(1)) Then
            SQL = Format(RN.Fields(1), "dd/mm/yyyy")
        Else
            SQL = " "
        End If
        It.SubItems(1) = SQL

        

        If Not PropuestaComercial And Not Potenciales Then
            SQL = " "
            If Not IsNull(RN!f_aceptado) Then SQL = Format(RN!f_aceptado, "dd/mm/yyyy")
            It.SubItems(2) = SQL
        End If


        If PropuestaComercial Then
            SQL = "COM"
        Else
            SQL = "CON"
        End If
        SQL = SQL & Format(RN!Id, "00000") & "." & RN!extension
        If PonerIcono Then It.SmallIcon = DevuelveIconoWHOSE(RN!extension)
        
        
        
        It.Tag = SQL
        RN.MoveNext
    Wend
    RN.Close
    Set RN = Nothing
End Sub



Public Sub CargaListviewExpPRI(ByRef ElListview As ListView, Ano As Integer, expediente As Long)
Dim It As ListItem

    ElListview.ListItems.Clear
   
        
    SQL = "SELECT * FROM whoobrasclipi WHERE expediente =  " & expediente & " AND anoexp = " & Ano & " ORDER BY   idPI desc"
   
    Set RN = New ADODB.Recordset
    RN.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RN.EOF
        'Fecha presntacion    contestacion    aceptada
        Set It = ElListview.ListItems.Add(, "K" & Format(RN.Fields(2), "0000"))
        It.Text = Format(RN.Fields(3), "dd/mm/yyyy")
        If Not IsNull(RN.Fields(4)) Then
            SQL = Format(RN.Fields(4), "dd/mm/yyyy")
            
        Else
            SQL = " "
        End If
        It.SubItems(1) = SQL
        If SQL <> " " Then
            If RN!aceptado = 0 Then
                SQL = "-"
            Else
                SQL = "SI"
            End If
        End If
        It.SubItems(2) = SQL
        'Monto la cadena NOMBRE SGDA
        If DBLet(RN!extension, "T") = "" Then
            SQL = ""
        Else
           
            SQL = Format(RN!expediente, "000000") & RN!anoexp & Format(RN!idpi, "000")
            
            SQL = SQL & "." & RN!extension
            It.SmallIcon = DevuelveIconoWHOSE(RN!extension)
        End If
        
        
        It.Tag = SQL
        RN.MoveNext
    Wend
    RN.Close
    Set RN = Nothing
End Sub


Public Sub CargaListviewExpSGD(ByRef ElListview As ListView, Ano As Integer, expediente As Long)
Dim It As ListItem

    ElListview.ListItems.Clear
        
    ' idempresa  NombreSGD
    SQL = "SELECT expediente,anoexp,NombreSGD,IdPres,f_preSGD,fcontesta,aceptado,extension,sgd FROM whoobrasclisgd left join whoegda  "
    SQL = SQL & " ON idempresa=sgd WHERE  expediente =  " & expediente & " AND anoexp = "
    SQL = SQL & Ano & " ORDER BY sgd,  IdPres desc"
    
    Set RN = New ADODB.Recordset
    RN.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RN.EOF
        'Fecha presntacion    contestacion    aceptada
        Set It = ElListview.ListItems.Add(, "K" & Format(RN!sgd, "00") & Format(RN!IdPres, "0000"))   'idpres
        
        It.Text = RN!sgd
        It.SubItems(1) = RN!NombreSGD
        
        It.SubItems(2) = Format(RN!f_preSGD, "dd/mm/yyyy")
        If Not IsNull(RN!fcontesta) Then
            SQL = Format(RN!fcontesta, "dd/mm/yyyy")
            
        Else
            SQL = " "
        End If
        It.SubItems(3) = SQL
        If SQL <> " " Then
            If RN!aceptado = 0 Then
                SQL = "-"
            Else
                SQL = "SI"
            End If
        End If
        It.SubItems(4) = SQL
        'Monto la cadena NOMBRE SGDA
        If DBLet(RN!extension, "T") = "" Then
            SQL = ""
        Else
            
            SQL = Format(RN!expediente, "000000") & RN!anoexp & Format(RN!IdPres, "000") & Format(RN!sgd, "00")
            SQL = SQL & "." & RN!extension
            It.SmallIcon = DevuelveIconoWHOSE(RN!extension)
        End If
        
        
        It.Tag = SQL
        RN.MoveNext
    Wend
    RN.Close
    Set RN = Nothing
End Sub




'Dado un item devuelve su nombre

Public Function DevuelveNombreArhivoITEM(ByRef It, PropuestaComercial As Boolean, Cliente As Long) As String
    
    DevuelveNombreArhivoITEM = vParamAplic.PathDocsWHOSE & "\POTENC\P" & Format(Cliente, "000000") & "\" & It.Tag
End Function

Public Function DevuelveNombreArhivoITEMClientes(ByRef It, Tipo As Byte, Cliente As Long) As String

    DevuelveNombreArhivoITEMClientes = ""
    If It.Tag = "" Then Exit Function


    SQL = vParamAplic.PathDocsWHOSE
    SQL = SQL & "\" & Format(Cliente, "000000") & "\"
    Select Case Tipo
    Case 0, 1
        'PROPUESTAS COMERCIALES
        SQL = SQL & "CONTRATO"
    Case 2
        'OBRAS
        SQL = SQL & "OBRA"
    Case 3
        'PROPIEDAD INTELECTUAL
        SQL = SQL & "PI"
    Case 4
        'Empresas de gestion de derechos de acutor
        SQL = SQL & "EGD"
    Case 5
        'ACtuaciones
        SQL = SQL & "ACTUA"
    End Select
    SQL = SQL & "\" & It.Tag
    DevuelveNombreArhivoITEMClientes = SQL
End Function




'Si se cambia el listimage hay que cambiar aqui tambien
Public Function DevuelveIconoWHOSE(extension As String) As Byte
    'A partir de la extension devolvera el icono
    ' 1 PDF,  2 XLS   3 DOC   4 VIDEO   5 Audio     6 OTROS
    extension = Trim(LCase(extension))
    If extension = "pdf" Then
        DevuelveIconoWHOSE = 1
    ElseIf Mid(extension, 1, 3) = "xls" Then
        DevuelveIconoWHOSE = 2
    ElseIf Mid(extension, 1, 3) = "doc" Then
        DevuelveIconoWHOSE = 3
    ElseIf extension = "avi" Or extension = "mpg" Or extension = "mpeg" Then
        DevuelveIconoWHOSE = 4
    ElseIf extension = "wav" Or extension = "mp3" Then
        DevuelveIconoWHOSE = 5
    ElseIf extension = "" Then
        DevuelveIconoWHOSE = 0
    Else
        DevuelveIconoWHOSE = 6
    End If
End Function

