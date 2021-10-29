Attribute VB_Name = "libEueler"
Option Explicit

'Si campo2="" entonces es una oferta
' Si campo2<>"" es un mantenimiento
Public Function ComprobarCarpetaPDFSMante2(campo1 As Long, campo2 As String) As String
Dim C As String
Dim Referencia As String
Dim i As Integer
    On Error GoTo eComprobarCarpetaOferta
    ComprobarCarpetaPDFSMante2 = ""
    C = EulerParam & "\"
    
    
    Referencia = CStr(campo2)
    For i = 1 To Len(Referencia)
        Referencia = Replace(Referencia, Mid("\/:*""?<>|", i, 1), " ")
    Next
    
        
    C = C & "Mante\" & Format(campo1, "000000") & Referencia
    
    If Dir(C, vbDirectory) = "" Then MkDir C
    
    ComprobarCarpetaPDFSMante2 = C & "\"
    
    
    
eComprobarCarpetaOferta:
    If Err.Number <> 0 Then
        MuestraError Err.Number, Err.Description
        C = ""
    End If
End Function



Public Function ComprobarExisteCarpetaPDFOferta(Anyo As Integer, codClien As Long, NumOfert As Long) As String
Dim C As String
Dim Aux As String

    On Error GoTo eComprobarCarpetaOferta
    ComprobarExisteCarpetaPDFOferta = ""
    Aux = EulerParam & "\Ofertas\" & Anyo & "\"
    C = Aux & Format(codClien, "000000") & "*"
    
    C = Dir(C, vbDirectory)
    If C = "" Then Exit Function
    Aux = Aux & C & "\"
    'Existe anoy y cliente.
    'Vamos a ver la oferta del cliente
    C = Aux & Format(NumOfert, "0000000") & "*"
    C = Dir(C, vbDirectory)
    
    If C = "" Then Exit Function
    Aux = Aux & C
    ComprobarExisteCarpetaPDFOferta = Aux & "\"
    
    
    
eComprobarCarpetaOferta:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
End Function





Public Function EliminarArhivoPDF(campo1 As Long, campo2 As String, Nombre As String) As Boolean
Dim C As String
Dim Referencia As String
Dim J As Integer
    On Error Resume Next
    
    EliminarArhivoPDF = False
    
    C = EulerParam & "\"
    
    
    
    Referencia = CStr(campo2)
    For J = 1 To Len(Referencia)
        Referencia = Replace(Referencia, Mid("\/:*""?<>|", J, 1), " ")
    Next
    
    
    
    C = C & "Mante\" & Format(campo1, "000000") & Referencia & "\" & Nombre

    
    
    If Dir(C, vbArchive) = "" Then
        MsgBox "No existe el archivo dentro de la oferta", vbExclamation
        EliminarArhivoPDF = True 'Para que borre la BD
    Else
        Kill C
        If Err.Number <> 0 Then
            MuestraError Err.Number, Err.Description
        Else
            EliminarArhivoPDF = True
        End If
    End If
End Function

Public Function EliminarArhivoPDFOferta(Destino As String) As Boolean
Dim C As String
    On Error Resume Next
    EliminarArhivoPDFOferta = False
    
    
    
    If Dir(Destino, vbArchive) = "" Then
        MsgBox "No existe el archivo dentro de la oferta", vbExclamation
        
    Else
        Kill Destino
        If Err.Number <> 0 Then
            MuestraError Err.Number, Err.Description
        Else
            EliminarArhivoPDFOferta = True
        End If
    End If
End Function



'Si campo2="" entonces es una oferta
' Si campo2<>"" es un mantenimiento
Public Function CopiaArhivoPDF2(campo1 As Long, campo2 As String, OrigenCompleto As String, Destino As String) As Boolean
Dim C As String
Dim extension As String  'MAYO 2019   Aceptamos todo tipo de ficvhero
Dim J As Integer
Dim Referencia As String
 
    On Error GoTo eCopiaArhivoOferta
    CopiaArhivoPDF2 = False
    
    C = EulerParam & "\"
    
    
    J = InStrRev(OrigenCompleto, ".")
    If J = 0 Then Err.Raise 513, , "Fichero si extension: " & OrigenCompleto
        
    extension = Mid(OrigenCompleto, J)
    If Len(extension) > 6 Then Err.Raise 513, , "Extension incorrecta: " & extension
        
        
    Referencia = CStr(campo2)
    For J = 1 To Len(Referencia)
        Referencia = Replace(Referencia, Mid("\/:*""?<>|", J, 1), " ")
    Next
        
        
        
        
    C = C & "Mante\" & Format(campo1, "000000") & Referencia & "\" & Destino & extension
    FileCopy OrigenCompleto, C
    
    CopiaArhivoPDF2 = True
    
    
    
eCopiaArhivoOferta:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
End Function


Public Function CopiaArhivoPDFOfertaEuler(CarpetaDestino As String, OrigenCompleto As String) As Boolean
Dim C As String
Dim J As Integer
Dim K As Integer

    On Error GoTo eCopiaArhivoOferta
    CopiaArhivoPDFOfertaEuler = False
    
    
    
    J = InStrRev(OrigenCompleto, "\")
    C = Mid(OrigenCompleto, J + 1)
    
    C = CarpetaDestino & C
    
    FileCopy OrigenCompleto, C
    
    CopiaArhivoPDFOfertaEuler = True
    
    
    
eCopiaArhivoOferta:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
End Function






Public Function EulerPathCompletoArchivoOfertas(campo1 As Long, NombreCorto As String) As String
Dim C As String

On Error GoTo eEulerPathCompletoArchivo
    
    EulerPathCompletoArchivoOfertas = ""
    C = davidCodtipom & NombreCorto
    If Dir(C, vbArchive) = "" Then Err.Raise 513, , "No existe fichero: " & C
    
    EulerPathCompletoArchivoOfertas = C
    
    Exit Function
eEulerPathCompletoArchivo:
    MuestraError Err.Number, Err.Description
End Function


Public Function EulerPathMante(campo1 As Long, campo2 As String, NombreCorto As String) As String
Dim C As String

On Error GoTo eEulerPathCompletoArchivo
    
    EulerPathMante = ""
    
    C = EulerParam & "\"
   
        
    C = C & "Mante\" & Format(campo1, "000000") & campo2 & "\" & NombreCorto
 
    
    
    
    
    If Dir(C, vbArchive) = "" Then Err.Raise 513, , "No existe fichero: " & C
    
    EulerPathMante = C
    
    Exit Function
eEulerPathCompletoArchivo:
    MuestraError Err.Number, Err.Description
End Function




Public Function NombreArchivoEULER(Descripcion As String) As String
Dim J As Integer
    
    J = InStrRev(Descripcion, "\")
    If J > 0 Then Descripcion = Mid(Descripcion, J + 1)
    
    For J = 1 To 9
        Descripcion = Replace(Descripcion, Mid("\/:*?""<>|", J, 1), "")
    Next
        
    NombreArchivoEULER = Descripcion
End Function






Public Function ImprimirLosCostesAlbaranEuler(ByRef ListView2 As ListView, hcoCodTipoM As String, numalbar As String) As Boolean

    
Dim C As String
Dim N As String
    
    ImprimirLosCostesAlbaranEuler = False
    
    C = "DELETE FROM tmpcommand WHERE codusu =" & vUsu.Codigo
    conn.Execute C


        
    'tmpcommand(codusu,cantidad,importel,fecrecep,nomprove,codfamia,nomfamia,nomartic,codartic)
    CadenaDesdeOtroForm = ""
    For NumRegElim = 1 To ListView2.ListItems.Count
        'Primera linea
        C = vUsu.Codigo & ","
        'Cantidad y precio
        C = C & DBSet(ListView2.ListItems(NumRegElim).SubItems(5), "N") & "," & DBSet(ListView2.ListItems(NumRegElim).SubItems(6), "N") & ","
        'Fecha
        N = Trim(Trim(ListView2.ListItems(NumRegElim).SubItems(3)))
        If N = "" Then N = Format(Now, "dd/mm/yyyy")
        C = C & DBSet(N, "F", "S") & ","
        
        'Resto campos  nomprove codfamia nomfamia,nomartic,codartic
        Select Case ListView2.ListItems(NumRegElim).Text
        Case "HOR"
            C = C & DBSet(ListView2.ListItems(NumRegElim).SubItems(1), "T") & ",1,'','',"
        
        Case "VEH"
            C = C & DBSet(ListView2.ListItems(NumRegElim).SubItems(1), "T") & ",0,"
            C = C & DBSet(ListView2.ListItems(NumRegElim).SubItems(2), "T") & ",'',"
        
        Case "ALV"
            C = C & DBSet("Venta. ", "T") & ",3,"
            C = C & DBSet(ListView2.ListItems(NumRegElim).SubItems(2), "T") & ","
            C = C & DBSet(ListView2.ListItems(NumRegElim).SubItems(4), "T") & ","
        Case "ALC"
            C = C & DBSet("Albaran. " & ListView2.ListItems(NumRegElim).SubItems(1), "T") & ",4,"
            C = C & DBSet(ListView2.ListItems(NumRegElim).SubItems(2), "T") & ","
            C = C & DBSet(ListView2.ListItems(NumRegElim).SubItems(4), "T") & ","
        Case "MAT"
            C = C & "'Material',2,'',"
            C = C & DBSet(ListView2.ListItems(NumRegElim).SubItems(4), "T") & ","
                
        Case "FAC"
            C = C & DBSet("Factura. " & ListView2.ListItems(NumRegElim).SubItems(1), "T") & ",5,"
            C = C & DBSet(ListView2.ListItems(NumRegElim).SubItems(2), "T") & ","
            C = C & DBSet(ListView2.ListItems(NumRegElim).SubItems(4), "T") & ","
     
     
        Case "PED"
            C = C & DBSet("Pedido. " & ListView2.ListItems(NumRegElim).SubItems(1), "T") & ",6,"
            C = C & DBSet(ListView2.ListItems(NumRegElim).SubItems(2), "T") & ","
            C = C & DBSet(ListView2.ListItems(NumRegElim).SubItems(4), "T") & ","
        Case Else
            MsgBox "No tratado. " & ListView2.ListItems(NumRegElim).Text, vbExclamation
            C = ""
        End Select
    
        If C <> "" Then
            C = C & DBSet(hcoCodTipoM, "T") & "," & DBSet(numalbar, "T") & ")"
            CadenaDesdeOtroForm = CadenaDesdeOtroForm & ", (" & C
        End If
    
    Next
    If CadenaDesdeOtroForm <> "" Then
        CadenaDesdeOtroForm = Mid(CadenaDesdeOtroForm, 2)
        C = "INSERT INTO tmpcommand(codusu,cantidad,importel,fecrecep,nomprove,codfamia,nomfamia,nomartic,codartic,codprove) VALUES "
        C = C & CadenaDesdeOtroForm
        conn.Execute C
        
        ImprimirLosCostesAlbaranEuler = True
    End If


End Function
