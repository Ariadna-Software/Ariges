Attribute VB_Name = "libEueler"
Option Explicit

'Si campo2="" entonces es una oferta
' Si campo2<>"" es un mantenimiento
Public Function ComprobarCarpetaPDFSMante2(campo1 As Long, campo2 As String) As String
Dim C As String
    
    On Error GoTo eComprobarCarpetaOferta
    ComprobarCarpetaPDFSMante2 = ""
    C = EulerParam & "\"
    
    
        
    C = C & "Mante\" & Format(campo1, "000000") & campo2
    
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
    On Error Resume Next
    EliminarArhivoPDF = False
    
    C = EulerParam & "\"
    
    C = C & "Mante\" & Format(campo1, "000000") & campo2 & "\" & Nombre

    
    
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
        Kill C
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
Dim Extension As String  'MAYO 2019   Aceptamos todo tipo de ficvhero
Dim J As Integer

 
    On Error GoTo eCopiaArhivoOferta
    CopiaArhivoPDF2 = False
    
    C = EulerParam & "\"
    
    
    J = InStrRev(OrigenCompleto, ".")
    If J = 0 Then Err.Raise 513, , "Fichero si extension: " & OrigenCompleto
        
    Extension = Mid(OrigenCompleto, J)
    If Len(Extension) > 6 Then Err.Raise 513, , "Extension incorrecta: " & Extension
        
    C = C & "Mante\" & Format(campo1, "000000") & campo2 & "\" & Destino & Extension
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
