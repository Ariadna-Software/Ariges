Attribute VB_Name = "ImagenADO"
'------------------------------------------------------------------------------
' Código para grabar y leer imágenes en campos de bases             ( 9/Abr/98)
' Adaptado para usarlo con ADO                                      (11/Jul/01)
'
' Adaptado de un par de ejemplos de la ayuda de VB5
'
' ©Guillermo 'guille' Som, 1998-2001
' mensaje@elguille.info
'------------------------------------------------------------------------------
Option Explicit

Private nFile As Long
Private Chunk() As Byte
Private Const mBuffer As Long = 16384&

Public Function LeerBinary(ADOField As ADODB.Field, NombreArchivo As String) As Boolean
    ' Leer la imagen del campo de la base y asignarlo al Picture
    '--------------------------------------------
    ' Este procedimiento no es necesario usarlo
    ' si el Picture está ligado a un data control
    '--------------------------------------------
    Dim nChunks As Long
    Dim nSize As Long
    Dim Fragment As Long
    Dim I As Long
    
    
    On Error GoTo ELeerBinary
    LeerBinary = False
    
    ' Se usa un fichero temporal para guardar la imagen
    nFile = FreeFile
    Open NombreArchivo For Binary Access Write As nFile
    '
    ' Calcular los trozos completos y el resto
    nSize = ADOField.ActualSize
    nChunks = Int(nSize / mBuffer)
    Fragment = nSize Mod mBuffer
    Chunk() = ADOField.GetChunk(Fragment)
    Put nFile, , Chunk()
    For I = 1 To nChunks
        Chunk() = ADOField.GetChunk(mBuffer)
        Put nFile, , Chunk()
    Next
    Close nFile
    Erase Chunk
    LeerBinary = True
    Exit Function
ELeerBinary:
    MuestraError Err.Number, "Leer Binary"
    
    On Error Resume Next
    Close nFile
    Err.Clear
    
End Function

Public Sub GuardarBinary(ADOField As ADODB.Field, kImagen As String)   ' unImage As Image)
    ' Guardar el contenido del Picture en el campo de la base
    Dim I As Long
    Dim Fragment As Long
    Dim nSize As Long
    Dim nChunks As Long
    
    'Ahora Ya tengo el path del fichero
    '
    ' Guardar el contenido del picture en un fichero temporal
    'SavePicture unImage.Picture, "pictemp"
    
    ' Leer el fichero y guardarlo en el campo
    nFile = FreeFile
    Open kImagen For Binary Access Read As nFile
    nSize = LOF(nFile)    ' Longitud de los datos en el archivo
    If nSize = 0 Then
        Close nFile
        Exit Sub
    End If
    '
    ' Calcular el número de trozos y el resto
    nChunks = nSize \ mBuffer
    Fragment = nSize Mod mBuffer
    ReDim Chunk(Fragment)
    '
    Get nFile, , Chunk()
    ADOField.AppendChunk Chunk()
    ReDim Chunk(mBuffer)
    For I = 1 To nChunks
        Get nFile, , Chunk()
        ADOField.AppendChunk Chunk()
    Next I
    Close nFile
    
    
    ''
    '' Ya no necesitamos el fichero, así que borrarlo
    'On Local Error Resume Next
    'If Len(Dir$("pictemp")) Then
    '    Kill "pictemp"
    'End If
    Err = 0
End Sub




'PARA LOS BACKUPS
Public Sub LeerBinaryEnString(ADOField As ADODB.Field, CadenaFinal As String)
    ' Leer la imagen del campo de la base y asignarlo al Picture
    '--------------------------------------------
    ' Este procedimiento no es necesario usarlo
    ' si el Picture está ligado a un data control
    '--------------------------------------------
    Dim nChunks As Long
    Dim nSize As Long
    Dim Fragment As Long
    Dim I As Long
    '
    
    
    '
    ' Calcular los trozos completos y el resto
    nSize = ADOField.ActualSize
    nChunks = Int(nSize / mBuffer)
    Fragment = nSize Mod mBuffer
    Chunk() = ADOField.GetChunk(Fragment)
    
    'Put nFile, , Chunk()
    'For I = 1 To nChunks
    '    Chunk() = ADOField.GetChunk(mBuffer)
    '    Put nFile, , Chunk()
    'Next
    'Close nFile
    Erase Chunk
   

  
    ' Ya no necesitamos el fichero, así que borrarlo
    On Error Resume Next
    If Len(Dir$("pictemp")) Then
        Kill "pictemp"
    End If
    Err = 0
End Sub

