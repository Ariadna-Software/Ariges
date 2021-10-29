VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Usuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Soporta entidades financieras usadas en los ingresos"
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'variables locales que contienen valores de propiedad
Private mvarCodigo As Long 'copia local
Private mvarNombre As String 'copia local
Private mvarPasswd As String 'copia local
Private mvarPasswdPROPIO As String 'copia local
Private mvarlogin As String 'copia local
Private mvarDirFich As String 'copia local
Private mvarNivel As Long 'copia local
Private mvarCadenaConexion As String
Private mvarPC As String
Private mvarSkin As Integer

Private mvarAlmacenPorDefecto As String

Private mvarCodigoAgente As Integer

Private mvarNivel2 As Byte   'Aplicara un segundo nivel para tema de permisos
Private mvarId As Long    'Aplicara un segundo nivel para tema de permisos

Private mvarTabPorDefecto As Integer



Private RS As ADODB.Recordset
Private Sql As String
Private Mens As String

Public Property Let Nivel(ByVal vData As Long)
    mvarNivel = vData
End Property


Public Property Get Nivel() As Long
    Nivel = mvarNivel
End Property

Public Property Let Passwd(ByVal vData As String)
Attribute Passwd.VB_Description = "Password de usuario"
    mvarPasswd = vData
End Property


Public Property Get Passwd() As String
    Passwd = mvarPasswd
End Property



Public Property Let CadenaConexion(ByVal vData As String)
    mvarCadenaConexion = vData
End Property


Public Property Get CadenaConexion() As String
    CadenaConexion = mvarCadenaConexion
End Property



Public Property Let PasswdPROPIO(ByVal vData As String)
'se usa cuando se asigna un valor a una propiedad, en el lado izquierdo de la asignación.
'Syntax: X.Passwd = 5
    mvarPasswdPROPIO = vData
End Property


Public Property Get PasswdPROPIO() As String
'se usa cuando se asigna un valor a una propiedad, en el lado derecho de la asignación.
'Syntax: Debug.Print X.Passwd
    PasswdPROPIO = mvarPasswdPROPIO
End Property


Public Property Let Nombre(ByVal vData As String)
Attribute Nombre.VB_Description = "Nombre del grupo de formas de cobro."
'se usa cuando se asigna un valor a una propiedad, en el lado izquierdo de la asignación.
'Syntax: X.Nombre= 5
    mvarNombre = vData
End Property


Public Property Get Nombre() As String
'se usa cuando se asigna un valor a una propiedad, en el lado derecho de la asignación.
'Syntax: Debug.Print X.Nombre
    Nombre = mvarNombre
End Property

'---------------------------
'Dirfich
Public Property Let Dirfich(ByVal vData As String)
'se usa cuando se asigna un valor a una propiedad, en el lado izquierdo de la asignación.
'Syntax: X.Nombre= 5
    mvarDirFich = vData
End Property


Public Property Get Dirfich() As String
'se usa cuando se asigna un valor a una propiedad, en el lado derecho de la asignación.
'Syntax: Debug.Print X.Nombre
    Dirfich = mvarDirFich
End Property

'Login
Public Property Let Login(ByVal vData As String)
'se usa cuando se asigna un valor a una propiedad, en el lado izquierdo de la asignación.
'Syntax: X.Nombre= 5
    mvarlogin = vData
End Property


'Solo tiene GET, se fija desde una funcion
Public Property Get AlmacenPorDefecto2() As String
'se usa cuando se asigna un valor a una propiedad, en el lado derecho de la asignación.
'Syntax: Debug.Print X.Nombre
    AlmacenPorDefecto2 = mvarAlmacenPorDefecto
End Property

Public Property Get CodigoAgente() As Integer
'se usa cuando se asigna un valor a una propiedad, en el lado derecho de la asignación.
'Syntax: Debug.Print X.Nombre
    CodigoAgente = mvarCodigoAgente
End Property


Public Property Get Login() As String
'se usa cuando se asigna un valor a una propiedad, en el lado derecho de la asignación.
'Syntax: Debug.Print X.Nombre
    Login = mvarlogin
End Property



Public Property Let Codigo(ByVal vData As Long)
Attribute Codigo.VB_Description = "Código del grupo de formas de cobro"
'se usa cuando se asigna un valor a una propiedad, en el lado izquierdo de la asignación.
'Syntax: X.Codigo= 5
    mvarCodigo = vData
End Property


Public Property Get Codigo() As Long
'se usa cuando se asigna un valor a una propiedad, en el lado derecho de la asignación.
'Syntax: Debug.Print X.Codigo
    Codigo = mvarCodigo
End Property

Public Property Get PC() As String
'se usa cuando se asigna un valor a una propiedad, en el lado derecho de la asignación.
'Syntax: Debug.Print X.Codigo
    PC = mvarPC
End Property

'Aplicara un segundo nivel para tema de permisos
' Taxco.  Quien tenga ese nivel a 2 NO puede modificar precio
Public Property Get Nivel2() As Byte
    Nivel2 = mvarNivel2
End Property


Public Property Let Id(ByVal vData As Long)
    mvarId = vData
End Property


Public Property Get Id() As Long
    Id = mvarId
End Property



Public Property Get Skin() As Integer
    Skin = mvarSkin
End Property
Public Property Let Skin(ByVal vData As Integer)
    mvarSkin = vData
End Property

Public Property Let TabPorDefecto(ByVal vData As Integer)
    mvarTabPorDefecto = vData
End Property


Public Property Get TabPorDefecto() As Integer
    TabPorDefecto = mvarTabPorDefecto
End Property







Public Function Leer(ByVal vlogin As String) As Byte
Attribute Leer.VB_Description = "Lee un grupo de formas de cobro en función del código que se le ha pasado. Si devuelve 0 se ha conseguido leer el dato."
Dim i As Long
    On Error GoTo Err1
    
    Sql = "SELECT * FROM usuarios.usuarios WHERE login = '" & vlogin & "'"
    Sql = Sql & " AND NivelAriges >=0"
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    If RS.EOF Then
        Leer = 1
    Else
        mvarId = DBLet(RS!CodUsu)
        i = DevuelveAumentoPC
        mvarCodigo = Val(DBLet(RS!CodUsu)) + i
        mvarNombre = DBLet(RS!nomusu)
        mvarNivel = Val(DBLet(RS!nivelariges))
        mvarPasswdPROPIO = RS!passwordpropio
        mvarlogin = vlogin  'por si acasp esta escroto en mayus, minusclu....
        mvarlogin = RS!Login
        mvarCadenaConexion = ""  'Se le asigna en tiempo de diseño
        
        
        mvarSkin = -1
        
        
        mvarTabPorDefecto = 2 'De momento pongo esto
        
        'If Not IsNull(RS!skinariges) Then mvarSkin = RS!skinariges
        FijarSkinAriges
        
        Leer = 0
    End If
    RS.Close
    Set RS = Nothing
    Exit Function
Err1:
    Mens = "Error: " & Err.Number & " : " & Err.Description
    MsgBox Mens, vbExclamation
    RS.Close
    Leer = 1
End Function

Private Function FijarSkinAriges()

    On Error Resume Next
    Sql = RS!skinariges
    If Err.Number <> 0 Then
        Err.Clear
        conn.Errors.Clear
    Else
         mvarSkin = RS!skinariges
    End If
End Function



Private Function DevuelveAumentoPC() As Long

DevuelveAumentoPC = 32000
'mvarPC = ComputerName
mvarPC = ComputerName
Sql = DevuelveDesdeBD(conAri, "codpc", "usuarios.pcs", "nompc", mvarPC, "T")
If Sql <> "" Then DevuelveAumentoPC = Val(Sql) * 1000

End Function


'Almacen por defecto y CODAGENT,   Nivel2:Aplicara un segundo nivel para tema de permisos
Public Sub FijarOtrosValoresUsuario()
Dim R As ADODB.Recordset

    
    On Error GoTo eFijarOtrosValoresUsuario
    
    mvarAlmacenPorDefecto = "1"
    mvarCodigoAgente = 0
    mvarNivel2 = 0
    Set R = New ADODB.Recordset
    Sql = "Select codalmac,codagent,nrosegur FROM straba WHERE login = " & DBSet(Login, "T")
    R.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not R.EOF Then
        If Not IsNull(R!codAlmac) Then mvarAlmacenPorDefecto = R!codAlmac
        If Not IsNull(R!CodAgent) Then mvarCodigoAgente = R!CodAgent
        
        If vParamAplic.NumeroInstalacion = vbTaxco Then
            If UCase(DBLet(R!nrosegur, "T")) = "TALLER" Then mvarNivel2 = 2
        End If
    End If
    R.Close
eFijarOtrosValoresUsuario:
    If Err.Number <> 0 Then
        MsgBox "Leyendo datos trabajador " & vbCrLf & Err.Description, vbExclamation
        conn.Errors.Clear
        Err.Clear
    End If
    Set R = Nothing
End Sub



Public Sub ActualizarSkin()
    Sql = "UPDATE usuarios.usuarios SET skin=" & mvarSkin & " WHERE codusu = " & vUsu.Id
   
    ejecutar Sql, False
End Sub
Public Sub LeerFiltros(aplicacion As String, idPrg As Integer)

End Sub

Public Sub InicializaFiltrosEmpresa()
        
        'Filtros por defecto
       ' mvarFiltroCompensacionHco = 0
       ' mvarFiltroCompensacionHco2 = 0
    
End Sub



Public Sub LeerTabPorDefecto()
    'Las tabs estan en mysql,menus, desde el 1 hasta el 14
    If mvarTabPorDefecto = 0 Then mvarTabPorDefecto = 2
    
    Sql = "aplicacion = 'ariges' and filtro =1  and codigo in (select codigo from menus where padre=0 and ocultarPorInstalacion=0)  and codusu  "

    Sql = DevuelveDesdeBD(conAri, "codigo", "menus_usuarios", Sql, CStr(mvarId))
    If Sql <> "" Then
        If Val(Sql) > 0 Then mvarTabPorDefecto = CInt(Sql)
    End If
    
End Sub

Public Sub GuardarTabPorDefecto()
    On Error GoTo eGuardarTabPorDefecto
    'En todas las contabilidades
    Sql = "SELECT codempre FROM usuarios.empresasariges "
    Sql = Sql & " where mid(ariges,1,6) = 'ariges' "
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If RS.EOF Then
        'SOLO en esta. Deberia dar error
        Sql = "UPDATE menus_usuarios SET filtro=if(codigo=" & vUsu.TabPorDefecto & ",1,0) WHERE codusu = " & vUsu.Id
        Sql = Sql & " AND aplicacion = 'ariges' and codigo in (select codigo from menus where padre=0 and ocultarPorInstalacion=0) "
        conn.Execute Sql
    Else
        While Not RS.EOF
            Sql = "UPDATE ariges" & IIf(RS!codempre = 0, "", RS!codempre) & ".menus_usuarios "
            Sql = Sql & " SET filtro=if(codigo=" & vUsu.TabPorDefecto & ",1,0) WHERE codusu = " & vUsu.Id
            Sql = Sql & " AND aplicacion = 'ariges' and codigo in (select codigo from menus where padre=0 and ocultarPorInstalacion=0) "
            conn.Execute Sql
            RS.MoveNext
        Wend
    End If
    RS.Close
    
eGuardarTabPorDefecto:
    If Err.Number <> 0 Then Err.Clear

    Set RS = Nothing
    
End Sub

