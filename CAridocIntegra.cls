VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAridocIntegra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const vbPermisoTotal = 2147483647   'DE aridoc


'-----------------------------------------------------------------------
'
'
'   Aridoc Integra.
'                Leeremos los valores para integrar de cada tipo de integracion
'                   que realizemos. Facturas....
'
'                Lee de las tabla paramaridoc_lin


'Variables locales que contienen valores de propiedad
Private mvarcodigo As Long 'Codigo u opcion de integracion
Private mvarcodext As Integer 'copia local
Private mvarcodcarpeta As Long 'copia local
Private mvarUsuario As String


Private mvarCampoBD As String
Private mvarCampoSel As String
Private mvarfechabd As String
Private mvarfechaSel As String


Private mvarCodUsuario As Integer
Private mvarGrupo As Long
Private mvarRutaAlmacen As String

'Internos


Private SQL As String
Private Rs As ADODB.Recordset


Public Property Let Codigo(ByVal vData As Long)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.codigo = 5
     mvarcodigo = vData
End Property

Public Property Get Codigo() As Long
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.codigo
     Codigo = mvarcodigo
End Property

'--------------------------------
'SOLO GET

Public Property Get Usuario() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.codigo
     Usuario = mvarUsuario
End Property


Public Property Get Grupo() As Long
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.codigo
     Grupo = mvarGrupo
End Property


Public Property Get RutaAlmacen() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.codigo
     RutaAlmacen = mvarRutaAlmacen
End Property


















Public Function Leer(vIdcodigo As Byte) As Byte
    On Error GoTo Err3


    'Leeremos los parametros generales de aridoc
    Leer = 1
    Set Rs = New ADODB.Recordset
    SQL = "SELECT * FROM paramaridoc"
    
    
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not Rs.EOF Then
            mvarcodext = Rs!extension
            mvarUsuario = Rs!Usuario
            Leer = 0
    End If
    Rs.Close
    
    
    If Leer = 1 Then Exit Function
    
    SQL = "SELECT * FROM paramaridoc_lin "
    SQL = SQL & " WHERE codigo = " & vIdcodigo
    
    
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Rs.EOF Then
        Leer = 1
        'ELSE
    Else
            mvarcodigo = Rs!Codigo
            

            mvarCampoBD = Rs!CampoBD
            mvarCampoSel = Rs!camposel
            mvarfechabd = Rs!fechabd
            mvarfechaSel = Rs!fechaSel
            mvarcodcarpeta = Rs!carpeta
            
    End If
    Rs.Close
Err3:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbExclamation
    End If
    Set Rs = Nothing
End Function

Public Function EstablecerValoresARidoc(ByRef Cnn As Connection) As Boolean
    
    On Error GoTo EEst
    EstablecerValoresARidoc = False
    
    Set Rs = New ADODB.Recordset
    
    'Dado un usuario: ej root, comprobar que existe
    SQL = "select codusu from usuarios where login='" & mvarUsuario & "'"
    Rs.Open SQL, Cnn, adOpenForwardOnly, adLockPessimistic, adCmdText
    SQL = ""
    If Not Rs.EOF Then SQL = Rs!CodUsu
    Rs.Close
    If SQL = "" Then
        MsgBox "NO existe usuario: " & mvarUsuario, vbExclamation
        Exit Function
    Else
        mvarCodUsuario = Val(SQL)
    End If
    
    
    'Obtenner el grupo
    SQL = "select codgrupo from usuariosgrupos where codusu=" & mvarCodUsuario & " order by orden desc"
    Rs.Open SQL, Cnn, adOpenForwardOnly, adLockPessimistic, adCmdText
    SQL = ""
    If Not Rs.EOF Then SQL = Rs!codgrupo
    Rs.Close
    If SQL = "" Then
        MsgBox "NO existe grupo para el usuario: " & mvarUsuario, vbExclamation
        Exit Function
    Else
         mvarGrupo = Val(SQL)
    End If
            
    'Obtener almacen carpeta
    SQL = "select carpetas.codcarpeta,almacen.pathreal,version from carpetas , almacen where carpetas.almacen = almacen.codalma AND codcarpeta=" & mvarcodcarpeta
    Rs.Open SQL, Cnn, adOpenForwardOnly, adLockPessimistic, adCmdText
    SQL = ""
    If Not Rs.EOF Then
        If Rs!Version = 0 Then
            MsgBox "Capeta accesible via FTP.", vbExclamation
        Else
            SQL = Rs!pathreal
        End If
    Else
        MsgBox "La carpeta / almacen: " & mvarcodcarpeta, vbExclamation
    End If
    Rs.Close
    If SQL = "" Then Exit Function
            

    mvarRutaAlmacen = SQL

    
    'Hago dir por si acaso
    If Dir(mvarRutaAlmacen, vbDirectory) = "" Then
        MsgBox "Host incacesible: " & RutaAlmacen, vbExclamation
        Exit Function
    End If
    
    
    'Si llga aqui: todo bien
    EstablecerValoresARidoc = True
    Exit Function
EEst:
    MsgBox Err.Description, vbExclamation
End Function




Public Function DevuelveSQL() As String
Dim I As Integer
Dim J As Integer
    DevuelveSQL = ""
    For I = 1 To 4
        SQL = RecuperaValor(mvarCampoSel, I)
        J = Val(SQL)
        SQL = RecuperaValor(mvarCampoBD, J + 1)
        DevuelveSQL = DevuelveSQL & ", " & SQL
    Next I
    
    
    'Añadimos la fecha1
    SQL = RecuperaValor(mvarfechaSel, 1)
    J = Val(SQL)
    SQL = RecuperaValor(mvarfechabd, J + 1)
    DevuelveSQL = DevuelveSQL & ", " & SQL
    
    'la fecha2 (si procede)
    SQL = RecuperaValor(mvarfechaSel, 2)
    If SQL <> "" Then
        J = Val(SQL)
        SQL = SQL = RecuperaValor(mvarfechabd, J)
        DevuelveSQL = DevuelveSQL & ", " & SQL
    End If
    
    
    DevuelveSQL = Mid(DevuelveSQL, 2)
End Function



'Base: codigo a partir del cual empezara a insertar dcumentos
'tamayo: del fichero
Public Function GeneraSQLTimagen(ByRef CA As CAridoc, miCodigo As Long, Tamanyo As Currency) As String


    SQL = "INSERT INTO timagen (codigo, codext, codcarpeta, campo1, campo2, campo3, campo4,"
    SQL = SQL & "fecha1, fecha2, fecha3, importe1, tamnyo,"
    SQL = SQL & "  userprop, groupprop, lecturau, lecturag, escriturau, escriturag) VALUES ("
    'Valores
    SQL = SQL & miCodigo & "," & mvarcodext & "," & mvarcodcarpeta
    SQL = SQL & ",'" & (CA.campo1) & "',"
    SQL = SQL & DBSet(CA.campo2, "T") & ","       '
    SQL = SQL & DBSet(CA.campo3, "T") & ","       '
    SQL = SQL & DBSet(CA.campo4, "T") & ",'"       '
    SQL = SQL & Format(CA.Fecha1, FormatoFecha) & "',"
    If CA.Fecha2 = "0:00:00" Then
        SQL = SQL & "NULL,"
    Else
        SQL = SQL & DBSet(CA.Fecha2, "F") & ","
    End If

    SQL = SQL & DBSet(Now, "F") & ","
    
    SQL = SQL & DBSet(CA.importe1, "N") & ","
    
    SQL = SQL & TransformaComasPuntos(CStr(Tamanyo))
    
    SQL = SQL & "," & mvarCodUsuario
    SQL = SQL & "," & mvarGrupo
    SQL = SQL & "," & vbPermisoTotal
    SQL = SQL & "," & vbPermisoTotal
    SQL = SQL & "," & vbPermisoTotal
    SQL = SQL & "," & vbPermisoTotal
    SQL = SQL & ")"
    GeneraSQLTimagen = SQL
End Function


Public Function DevMaxTimagen(ByRef Cnn As ADODB.Connection) As Long
    SQL = "Select max(codigo) from  timagen"
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Cnn, adOpenForwardOnly, adLockPessimistic, adCmdText
    DevMaxTimagen = 0
    If Not Rs.EOF Then DevMaxTimagen = DBLet(Rs.Fields(0), "N")
    Rs.Close
    Set Rs = Nothing
    
End Function
