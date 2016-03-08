Attribute VB_Name = "libEueler"
Option Explicit

'Si campo2="" entonces es una oferta
' Si campo2<>"" es un mantenimiento
Public Function ComprobarCarpetaPDFS(campo1 As Long, campo2 As String) As Boolean
Dim C As String
    
    On Error GoTo eComprobarCarpetaOferta
    ComprobarCarpetaPDFS = False
    C = EulerParam & "\"
    
    If campo2 = "" Then
        C = C & "Ofertas\" & Format(campo1, "00000")
    Else
        
         C = C & "Mante\" & Format(campo1, "000000") & campo2
    End If
    If Dir(C, vbDirectory) = "" Then MkDir C
    
    ComprobarCarpetaPDFS = True
    
    
    
eComprobarCarpetaOferta:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
End Function


Public Function EliminarArhivoPDF(campo1 As Long, campo2 As String, Nombre As String) As Boolean
Dim C As String
    On Error Resume Next
    EliminarArhivoPDF = False
    
    C = EulerParam & "\"
    If campo2 = "" Then
        C = C & "Ofertas\" & Format(campo1, "00000") & "\" & Nombre
    Else
        
        C = C & "Mante\" & Format(campo1, "000000") & campo2 & "\" & Nombre
    End If
    
    
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

'Si campo2="" entonces es una oferta
' Si campo2<>"" es un mantenimiento
Public Function CopiaArhivoPDF(campo1 As Long, campo2 As String, OrigenCompleto As String, Destino As String) As Boolean
Dim C As String
    
    On Error GoTo eCopiaArhivoOferta
    CopiaArhivoPDF = False
    
     C = EulerParam & "\"
    
    If campo2 = "" Then
        C = C & "Ofertas\" & Format(campo1, "00000") & "\" & Destino & ".pdf"
    Else
        
        C = C & "Mante\" & Format(campo1, "000000") & campo2 & "\" & Destino & ".pdf"
    End If
    FileCopy OrigenCompleto, C
    
    CopiaArhivoPDF = True
    
    
    
eCopiaArhivoOferta:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
End Function

'Si campo2="" entonces es una oferta
' Si campo2<>"" es un mantenimiento
Public Function EulerPathCompletoArchivo(campo1 As Long, campo2 As String, NombreCorto As String) As String
Dim C As String

On Error GoTo eEulerPathCompletoArchivo
    
    EulerPathCompletoArchivo = ""
    
    C = EulerParam & "\"
    If campo2 = "" Then
        C = C & "Ofertas\" & Format(campo1, "00000") & "\" & NombreCorto
    Else
        
        C = C & "Mante\" & Format(campo1, "000000") & campo2 & "\" & NombreCorto
    End If
    
    
    
    
    If Dir(C, vbArchive) = "" Then Err.Raise 513, , "No existe fichero: " & C
    
    EulerPathCompletoArchivo = C
    
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
