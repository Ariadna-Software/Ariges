Attribute VB_Name = "LibImagen"
Option Explicit


''  -- Modos de Trabajo
'Public Const vbNorm = 0  ' modo normal
'Public Const vbHistNue = 1  ' modo de recuperar historico
'Public Const vbHistAnt = 2  ' modo de recuperar historico de los antiguos
'
'
'Public Const vbMaxGrupos = 31
'
'Public ModoTrabajo As Byte  '---------------------
'
'Public FormatoFecha As String
'
'Public Conn As Connection
'Public vUsu As Cusuarios
'Public vConfig As CConfiguracion

Public miRsAux As ADODB.Recordset


Public listacod As Collection
Public listaimpresion As Collection  'Esta lista servira para cuando queramos imprimir

'Cuiado con esta varibale
Public DatosModificados As Boolean


'Saber si ha coipado el archivo al server
Public DatosCopiados As String
'
'Public SeHaEjecutadoFTP As Boolean


Public Type RegistroTipoMensaje   ' Crea un tipo definido por el usuario.
   Descripcion As String
   Color As Long
   Icono As Integer
End Type


Public ArrayTipoMen() As RegistroTipoMensaje
Public TotalTipos As Integer   'Menos 1. Es decir, si hay tres tipos la var vale 2



Public Sub PonerArrayTiposMensaje()
Dim L As Long
Dim fin As Integer
Dim I As Integer
Dim J As Integer
Dim Cortar11 As String
'Public Type RegistroTipoMensaje   ' Crea un tipo definido por el usuario.
'   Descripcion As String * 30
'   Color As Long
'End Type
'
'Public ArrayTipoMen() As RegistroTipoMensaje
    TotalTipos = 0
    Cortar11 = "Select count(*) from mailtipo"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open Cortar11, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    fin = 0
    If Not miRsAux.EOF Then fin = DBLet(miRsAux.Fields(0), "N")
    miRsAux.Close
    
    If fin = 0 Then Exit Sub
    
    
    ReDim ArrayTipoMen(fin)
    TotalTipos = fin
    
    Cortar11 = "Select * from mailtipo order by tipo "
    miRsAux.Open Cortar11, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    J = 0
    I = 0
    
    
    While Not miRsAux.EOF
        
        If miRsAux!Tipo - J > 1 Then
            J = J + 1
            For fin = J To miRsAux!Tipo - 1
                ArrayTipoMen(fin).Color = 0
                ArrayTipoMen(fin).Descripcion = ""
                ArrayTipoMen(fin).Icono = 0
            Next fin
            I = miRsAux!Tipo
        End If
        
        ArrayTipoMen(I).Color = DBLet(miRsAux!Color, "N")
        ArrayTipoMen(I).Descripcion = miRsAux!Descripcion
        ArrayTipoMen(I).Icono = miRsAux!numico
        J = miRsAux!Tipo
        
        miRsAux.MoveNext
        I = I + 1
    Wend
    miRsAux.Close
    Set miRsAux = Nothing

End Sub

'
'Public Sub CodificacionLinea(Leer As Boolean, ByRef Linea As String)
'Dim I As Integer
'Dim C As String
'Dim C2 As String
'    C = Linea
'    Linea = ""
'
'
'        'Escribir
'        For I = 1 To Len(C)
'            C2 = Mid(C, I, 1)
'            If Leer Then
'                C2 = Chr(Asc(C2) - 3)
'            Else
'                C2 = Chr(Asc(C2) + 3)
'            End If
'            Linea = Linea & C2
'        Next I
'End Sub
'
'
'Public Sub AsignarCampoMemo(ByRef Campo As String, ByRef nombrecampo As String, ByRef ADO As ADODB.Recordset)
'    On Error Resume Next
'    Campo = ADO.Fields(nombrecampo).Value
'    If Err.Number <> 0 Then
'        Err.Clear
'        Campo = ""
'    End If
'End Sub
'
