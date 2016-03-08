VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Cmailc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Variables locales que contienen valores de propiedad
Private mvarcodmail As Long 'copia local
Private mvarorigen As Integer 'copia local
Private mvardestino As Integer 'copia local
Private mvarleido As Byte 'copia local
Private mvaremail As Byte 'copia local
Private mvarTipo As Integer

'DE la otra tabla
Private mvarasunto As String 'copia local
Private mvarTexto As String 'copia local
Private mvarFecha As Date


'Solo para los enviados
Private mvarTextoPara As String 'copia local


Private EstaEnHco As Boolean
Private mvarRecibido As Boolean

Private SQL As String  'Para la sentencia
Private Mens As String    'Para los mensajes
Private Rs As ADODB.Recordset 'Para la base de datos
Private Rc As Byte



'------------------------------------------------
'Propiedades del modulo Cmailc
'------------------------------------------------


Public Property Let codmail(ByVal vData As Long)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.codmail = 5
     mvarcodmail = vData
End Property

Public Property Get codmail() As Long
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.codmail
     codmail = mvarcodmail
End Property


Public Property Let Origen(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.origen = 5
     mvarorigen = vData
End Property

Public Property Get Origen() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.origen
     Origen = mvarorigen
End Property


Public Property Let Destino(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.destino = 5
     mvardestino = vData
End Property

Public Property Get Destino() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.destino
     Destino = mvardestino
End Property


Public Property Let leido(ByVal vData As Byte)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.leido = 5
     mvarleido = vData
End Property

Public Property Get leido() As Byte
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.leido
     leido = mvarleido
End Property


Public Property Let email(ByVal vData As Byte)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.email = 5
     mvaremail = vData
End Property

Public Property Get email() As Byte
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.email
     email = mvaremail
End Property

Public Property Let Fecha(ByVal vData As Date)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.email = 5
     mvarFecha = vData
End Property

Public Property Get Fecha() As Date
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.email
     Fecha = mvarFecha
End Property



Public Property Let asunto(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.asunto = 5
     mvarasunto = vData
End Property

Public Property Get asunto() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.asunto
     asunto = mvarasunto
End Property


Public Property Let Texto(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.Texto = 5
     mvarTexto = vData
End Property

Public Property Get Texto() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.Texto
     Texto = mvarTexto
End Property


Public Property Let Textopara(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.Texto = 5
     mvarTextoPara = vData
End Property

Public Property Get Textopara() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.Texto
     Textopara = mvarTextoPara
End Property

Public Property Let Tipo(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.Texto = 5
     mvarTipo = vData
End Property

Public Property Get Tipo() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.Texto
     Tipo = mvarTipo
End Property


Public Property Get Recibido() As Boolean
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.Texto
     Recibido = mvarRecibido
End Property



Public Property Get EnHco() As Boolean
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.Texto
     EnHco = EstaEnHco
End Property


'------------------------------------------------
'FUNCIONES del modulo Cmailc
'------------------------------------------------



''****   AGREGAR  ***
Public Function GenerarMensaje(ByRef ListaDestinatarios As Collection, ByRef CadenaPara As String, ByRef ListaDestinatariosExternos As Collection) As Byte
Dim I As Integer
Dim CodigosDestinatarios As String
Dim TextoDestinatarios As String
    'En agregar , lo que haremos sera crear el mensaje con enviado a 0
    '
    On Error GoTo Err1
    GenerarMensaje = 1
    'OBTENEMOS CODIGO
    If Not Siguiente Then Exit Function
    EstaEnHco = False
    
    'Insertamos linea
    SQL = "INSERT INTO maill (codmail, asunto, Texto, Fecha) VALUES ("
    SQL = SQL & mvarcodmail & ",'" & DevNombreSQL(mvarasunto) & "',"
    If mvarTexto = "" Then
        SQL = SQL & "NULL"
    Else
        Mens = DevNombreSQL(mvarTexto)
        SQL = SQL & "'" & Mens & "'"
    End If
    SQL = SQL & ",'" & Format(mvarFecha, FormatoFecha) & "')"
    Conn.Execute SQL
    
    
    'ENVIADOS
    SQL = "INSERT INTO maile (codmail, origen, email, tipo, textoPara, Destinatarios) VALUES ("
    SQL = SQL & mvarcodmail & "," & mvarorigen & "," & mvaremail & "," & mvarTipo & ",'" & DevNombreSQL(CadenaPara) & "','')"
    Conn.Execute SQL
    
    
    'Ahora para cada destinatario
    SQL = "INSERT INTO mailc (codmail, origen,  leido, email,tipo, destino) VALUES ("
    SQL = SQL & mvarcodmail & "," & mvarorigen & ",0," & mvaremail & "," & mvarTipo & ","
    I = 1
    CodigosDestinatarios = ""

    While I <= ListaDestinatarios.Count
        CodigosDestinatarios = CodigosDestinatarios & ListaDestinatarios(I) & "|"
        Mens = SQL & ListaDestinatarios(I) & ")"
        Conn.Execute Mens
        I = I + 1
    Wend
    If Len(CodigosDestinatarios) > 255 Then
        MsgBox "La longituda para los destinatarios supera la permitida. GenerarMensaje", vbExclamation
        Exit Function
    End If
    SQL = "UPDATE maile SET Destinatarios = '" & CodigosDestinatarios & "' WHERE codmail = " & mvarcodmail
    Conn.Execute SQL
    
    
    
    If ListaDestinatariosExternos.Count > 0 Then InsertaDestinatario2 ListaDestinatariosExternos
        
    
    GenerarMensaje = 0
    Exit Function

Err1:
    Mens = "Se ha producido un error." & vbCrLf
    Mens = Mens & "Número: " & Err.Number & vbCrLf
    Mens = Mens & "Descripción: " & Err.Description
    MsgBox Mens, vbExclamation
    'Agregar = 1
End Function


Public Function InsertaDestinatario2(ByRef Lista As Collection)
Dim I As Integer
Dim C As String
    On Error Resume Next
    SQL = "INSERT INTO maildestext (codmail, nombre, mail) VALUES ("
    For I = 1 To Lista.Count
        C = RecuperaValor(Lista.item(I), 1)
        Mens = mvarcodmail & ",'" & C & "','"
        C = RecuperaValor(Lista.item(I), 2)
        Mens = Mens & C & "')"
        Mens = SQL & Mens
        Conn.Execute Mens
        If Err.Number <> 0 Then Err.Clear
    Next I
End Function


''****   ELIMINAR  ***
'Public Function Eliminar() As Byte
'    On Error GoTo Err2
'
'    Sql = "SELECT * FROM mailc"
'    Sql = Sql & " WHERE "
'    Sql = Sql & " codmail = " & mvarcodmail
'    Sql = Sql & "  and origen = " & mvarorigen
'    Sql = Sql & "  and destino = " & mvardestino
'    Rs.Open Sql, Conn, , , adCmdText
'    If Rs.EOF Then
'        Mens = "No existe un Registro en la Tabla mailc con esos valores."
'        MsgBox Mens, vbExclamation
'        Eliminar = 1
'        'ELSE
'        Else
'            Rs.Delete
'            Rs.MoveNext
'            Eliminar = 0
'    End If
'
'    Rs.Close
'    Exit Function
'
'Err2:
'    Mens = "Se ha producido un error." & vbCrLf
'    Mens = Mens & "Número: " & Err.Number & vbCrLf
'    Mens = Mens & "Descripción: " & Err.Description
'    MsgBox Mens, vbExclamation
'    Eliminar = 1
'End Function
'

'****   LEER  ***
Public Function Leer(vIdcodmail As Long, Recibido As Boolean, EnHco As Boolean) As Byte
Dim Tabla As String

    On Error GoTo Err3


    mvarRecibido = Recibido
    EstaEnHco = EnHco
    
    If Recibido Then
        Tabla = "mailc"
    Else
        Tabla = "maile"
    End If
    If EnHco Then Tabla = Tabla & "h"
    
    
    SQL = "SELECT maill.asunto,maill.texto,maill.fecha,T.*"
    SQL = SQL & " FROM maill," & Tabla & " as T"
    
    
    SQL = SQL & " WHERE  T.codmail = maill.codmail"
    SQL = SQL & " AND maill.codmail = " & vIdcodmail

    If Recibido Then SQL = SQL & "  and destino = " & vUsu.codigo

    
    
    Rs.Open SQL, Conn, , , adCmdText
    If Rs.EOF Then
        Leer = 1
        'ELSE
        Else
            mvarcodmail = Rs!codmail
            mvarorigen = Rs!Origen
            
            
            mvaremail = Rs!email
            mvarTipo = Rs!Tipo
            mvarasunto = Rs!asunto
            mvarTexto = DBLet(Rs!Texto, "T")
            mvarFecha = Rs!Fecha
            If Not Recibido Then
                'ENVIADOS
                mvarTextoPara = Rs!Textopara
                mvardestino = -1
                
            Else
                mvardestino = Rs!Destino
                mvarleido = Rs!leido
                mvarTextoPara = ""
            End If
            Leer = 0
    End If

    Rs.Close
    Exit Function

Err3:
    Mens = "Se ha producido un error." & vbCrLf
    Mens = Mens & "Número: " & Err.Number & vbCrLf
    Mens = Mens & "Descripción: " & Err.Description
    MsgBox Mens, vbExclamation
    Leer = 1
End Function


'****   MODIFICAR  ***
'Public Function Modificar() As Byte
'    On Error GoTo Err4
'
'    SQL = "SELECT * FROM mailc"
'    SQL = SQL & " WHERE "
'    SQL = SQL & " codmail = " & mvarcodmail
'    SQL = SQL & "  and origen = " & mvarorigen
'    SQL = SQL & "  and destino = " & mvardestino
'    RS.Open SQL, Conn, , , adCmdText
'    If RS.EOF Then
'        Mens = "No existe el Registro en la Tabla mailc con esos valores."
'        MsgBox Mens, vbExclamation
'        Modificar = 1
'        'ELSE
'        Else
'            'Rs.Edit
'            RS!leido = mvarleido
'            RS!email = mvaremail
'            RS.Update
'            Modificar = 0
'    End If
'
'    RS.Close
'    Exit Function
'
'Err4:
'    Mens = "Se ha producido un error." & vbCrLf
'    Mens = Mens & "Número: " & Err.Number & vbCrLf
'    Mens = Mens & "Descripción: " & Err.Description
'    MsgBox Mens, vbExclamation
'    Modificar = 1
'End Function



Public Function Eliminar()
    '----------------------------------------------------------
    '
    '
    'BORRAMOS EL MENSAJE
    '
    '------------------
    If Not mvarRecibido Then
        SQL = "DELETE FROM maile"
        If EstaEnHco Then SQL = SQL & "h"
        SQL = SQL & " WHERE codmail = " & mvarcodmail
    Else
        'RECIBIDOS
        SQL = "DELETE FROM mailc"
        If EstaEnHco Then SQL = SQL & "h"
        SQL = SQL & " WHERE codmail = " & mvarcodmail
        SQL = SQL & " AND destino = " & vUsu.codigo
    End If
    Conn.Execute SQL


    'Ahora comprobamos un par de cosas, para ver si hay que borrar de maill
    Rc = 0
    If Not CompruebaEnTabla("mailc") Then Rc = 1
    
    If Rc = 0 Then
        'NO QUEDAN EN DESTINO
        'Comprobamos en enviados
        If mvarRecibido Then
            If Not CompruebaEnTabla("maile") Then Rc = 1
        End If
    End If
    
    'EN HCO
    If Rc = 0 Then
        If Not CompruebaEnTabla("mailch") Then Rc = 1
    End If

    'HCO
    If Rc = 0 Then
        If Not CompruebaEnTabla("maileh") Then Rc = 1
    End If
        
        
    'Si rec=0 --> Esto significa que no hay registros en enviados, recibidos, era el ultimo....
    'Luego borramos en mail
    If Rc = 0 Then
        Conn.Execute "DELETE from maill where codmail =" & mvarcodmail
    End If
    
End Function




Public Function PasarAHistorico()
    If EstaEnHco Then Exit Function
    '----------------------------------------------------------
    '
    '
    'BORRAMOS EL MENSAJE
    '
    '------------------
    If Not mvarRecibido Then
    
        SQL = "INSERT INTO maileh Select * from maile"
        SQL = SQL & " WHERE codmail = " & mvarcodmail
        Conn.Execute SQL
        SQL = "DELETE FROM maile"
        SQL = SQL & " WHERE codmail = " & mvarcodmail
        Conn.Execute SQL
        
        SQL = "INSERT INTO maildestexth Select * from maildestext"
        SQL = SQL & " WHERE codmail = " & mvarcodmail
        Conn.Execute SQL
        SQL = "DELETE FROM maildestext"
        SQL = SQL & " WHERE codmail = " & mvarcodmail
        Conn.Execute SQL
    Else
        'RECIBIDOS
        SQL = "INSERT INTO mailch Select * from mailc"
        SQL = SQL & " WHERE codmail = " & mvarcodmail
        Conn.Execute SQL
        SQL = "DELETE FROM mailc"
        SQL = SQL & " WHERE codmail = " & mvarcodmail
        Conn.Execute SQL
    End If
    
End Function



Private Function CompruebaEnTabla(Tabla As String) As Boolean
    CompruebaEnTabla = False
    SQL = "Select * from " & Tabla & " WHERE codmail =" & mvarcodmail
    Rs.Open SQL, , , , adCmdText
    If Rs.EOF Then CompruebaEnTabla = True
    Rs.Close
End Function

Private Function Siguiente() As Boolean
    Siguiente = False
    mvarcodmail = 0
    Conn.Execute "commit"
    SQL = "SELECT * FROM contadorar for update"
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Rs.EOF Then
        MsgBox "ERROR GRAVE. Tabla vacia"
        End
    End If
    mvarcodmail = DBLet(Rs.Fields(0), "N")
    mvarcodmail = mvarcodmail + 1
    Rs.Close
    
    'UPDATEO CONTADOR
    SQL = "UPDATE contadorar SET codmail =" & mvarcodmail
    Conn.Execute SQL
    Conn.Execute "commit"
    
    Siguiente = True
    
End Function



Public Sub MarcarComoLeido()
    SQL = "UPDATE mailc SET leido =1 WHERE codmail = " & mvarcodmail
    SQL = SQL & " AND destino = " & vUsu.codigo
    Conn.Execute SQL
End Sub


Private Sub Class_Initialize()
    Set Rs = New ADODB.Recordset
    'Ponemos el recordset en modo:
    Rs.CursorType = adOpenKeyset
    Rs.LockType = adLockOptimistic
End Sub


Private Sub Class_Terminate()
    Set Rs = Nothing
End Sub

