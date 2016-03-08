VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Contadores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'variables locales que contienen valores de propiedad

Private mvarTipoContador As String 'copia local
Private mvarNombreRegistro As String 'copia local
Private mvarContador As Long  'copia local

Private RS As ADODB.Recordset
Private SQL As String
Private Mens As String


Public Property Let Contador(ByVal vData As Long)
'se usa cuando se asigna un valor a una propiedad, en el lado izquierdo de la asignación.
'Syntax: X.Nivel = 5
    mvarContador = vData
End Property


Public Property Get Contador() As Long
'se usa cuando se asigna un valor a una propiedad, en el lado derecho de la asignación.
'Syntax: Debug.Print X.Nivel
    Contador = mvarContador
End Property


Public Property Let NombreRegistro(ByVal vData As String)
'se usa cuando se asigna un valor a una propiedad, en el lado izquierdo de la asignación.
'Syntax: X.Nombre= 5
    mvarNombreRegistro = vData
End Property


Public Property Get NombreRegistro() As String
'se usa cuando se asigna un valor a una propiedad, en el lado derecho de la asignación.
'Syntax: Debug.Print X.Nombre
    NombreRegistro = mvarNombreRegistro
End Property



Public Property Let TipoContador(ByVal vData As String)
'se usa cuando se asigna un valor a una propiedad, en el lado izquierdo de la asignación.
'Syntax: X.Nombre= 5
    mvarTipoContador = vData
End Property


Public Property Get TipoContador() As String
'se usa cuando se asigna un valor a una propiedad, en el lado derecho de la asignación.
'Syntax: Debug.Print X.Nombre
    TipoContador = mvarTipoContador
End Property



Public Function Eliminar() As Byte
    
    Exit Function
Err1:
    Mens = "Error: " & Err.Number & " : " & Err.Description
    MsgBox Mens, vbExclamation
    RS.Close
    Eliminar = 1
End Function


'///////////////////////////////////////////////////////////
'//
'//  Devuelve el contador que corresponda.
'//  Si es llamado desde dentro de un begin-trans
'//  entonces NO ejecuta las sentencias commit
Public Function ConseguirContador(LEtra As String, EjercicioActual As Boolean, DentroBeginTrans As Boolean) As Byte
Dim OK As Boolean
Dim C1 As Long
Dim C2 As Long
    On Error GoTo Err1
    'Abrimos bloqueando
    SQL = "Select * from contadores WHERE TipoRegi = '" & LEtra & "' FOR UPDATE"
    ConseguirContador = 1
    OK = False
    Set RS = New ADODB.Recordset
    RS.Open SQL, ConnConta, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not RS.EOF Then
        mvarTipoContador = LEtra
        mvarNombreRegistro = RS!nomregis
        C1 = RS!Contado1
        C2 = RS!Contado2
        OK = True
    End If
    RS.Close
    If Not OK Then
        If Not DentroBeginTrans Then ConnConta.Execute "commit"
        Exit Function
    End If
    
    'Actualizamos el contador
    SQL = "UPDATE contadores set "
    If EjercicioActual Then
        C1 = C1 + 1
        mvarContador = C1
        'c1 = c1 + 1 'ANTES
        SQL = SQL & " contado1=" & C1
    Else
        C2 = C2 + 1
        mvarContador = C2
        SQL = SQL & " contado2=" & C2
    End If
    SQL = SQL & " WHERE TipoRegi = '" & LEtra & "'"
    ConnConta.Execute SQL
    
    'Desbloqueamos
    If Not DentroBeginTrans Then ConnConta.Execute "commit"
    ConseguirContador = 0
    Exit Function
Err1:
    Mens = "Error: " & Err.Number & " : " & Err.Description
    MsgBox Mens, vbExclamation
End Function



'//////////////////////////////////////////////////////////////////7
'//
'//     Si es el ultimo numero de la serie devuelve el anterior
Public Function DevolverContador(LEtra As String, EjercicioActual As Boolean, Contador As Long) As Byte
Dim OK As Boolean
Dim C1 As Long
Dim C2 As Long
    On Error GoTo Err1
    'Abrimos bloqueando
    SQL = "Select * from contadores WHERE TipoRegi = '" & LEtra & "' FOR UPDATE"
    'Esto esba asi antes, sin comentar
    'connConta.Execute "Set autocommit = 0"
    DevolverContador = 1
    OK = False
    Set RS = New ADODB.Recordset
    RS.Open SQL, ConnConta, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not RS.EOF Then
        mvarTipoContador = LEtra
        mvarNombreRegistro = RS!nomregis
        C1 = RS!Contado1
        C2 = RS!Contado2
        OK = True
    End If
    RS.Close
    If Not OK Then
        ConnConta.Execute "commit"
        Exit Function
    End If

    OK = False
    'Contador = Contador + 1
    If EjercicioActual Then
        OK = (C1 = Contador)
    Else
        OK = (C2 = Contador)
    End If

    If OK Then
        'Actualizamos el contador
        SQL = "UPDATE contadores set "
        If EjercicioActual Then
            C1 = C1 - 1
            SQL = SQL & " contado1=" & C1
        Else
            C2 = C2 - 1
            SQL = SQL & " contado2=" & C2
        End If
        SQL = SQL & " WHERE TipoRegi = '" & LEtra & "'"
        ConnConta.Execute SQL
    End If

    'Desbloqueamos
    ConnConta.Execute "commit"
    DevolverContador = 0
    
Err1:
    If Err.Number <> 0 Then
        Mens = "Error: " & Err.Number & " : " & Err.Description
        MsgBox Mens, vbExclamation
    End If
    ConnConta.Execute "Set autocommit = 1"
End Function



Public Function Modificar() As Byte
 
    Exit Function
Err1:
    Mens = "Error: " & Err.Number & " : " & Err.Description
    MsgBox Mens, vbExclamation
    RS.Close
    Modificar = 1
End Function



