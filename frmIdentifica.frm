VERSION 5.00
Begin VB.Form frmIdentifica 
   BackColor       =   &H00800000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5655
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8790
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5655
   ScaleWidth      =   8790
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      ItemData        =   "frmIdentifica.frx":0000
      Left            =   4440
      List            =   "frmIdentifica.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   2880
      Width           =   2925
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5760
      Top             =   480
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   4440
      PasswordChar    =   "*"
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   4560
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   4320
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "vers"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   4200
      TabIndex        =   8
      Top             =   2400
      Width           =   2775
   End
   Begin VB.Image imgBl 
      Height          =   240
      Left            =   4080
      Picture         =   "frmIdentifica.frx":0004
      Top             =   4920
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label lblMay 
      BackStyle       =   0  'Transparent
      Caption         =   "Tecla Bloq. Mayús esta activada"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   375
      Left            =   4440
      TabIndex        =   7
      Top             =   5280
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Label lblInd 
      BackStyle       =   0  'Transparent
      Caption         =   "Cargando ..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   285
      Left            =   360
      TabIndex        =   6
      Top             =   2880
      Width           =   3495
   End
   Begin VB.Label lblTiempo 
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   7335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   375
      Index           =   1
      Left            =   4320
      TabIndex        =   4
      Top             =   4560
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   375
      Index           =   0
      Left            =   4320
      TabIndex        =   3
      Top             =   3600
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   5655
      Left            =   0
      Top             =   0
      Width           =   8775
   End
End
Attribute VB_Name = "frmIdentifica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim PrimeraVez As Boolean
Dim T1 As Single
Dim Segundos As Integer




Dim UltUsu_ As String
Dim UltEmpre_ As String






Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean
    
    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub Combo1_LostFocus()
    Text1(0).Text = Combo1.Text
End Sub







Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        
        Espera 0.5
        Me.Refresh
        
        'Vemos datos de configAriges.ini
        Set vConfig = New Configuracion
        If vConfig.Leer = 1 Then
             vConfig.SERVER = InputBox("Servidor: ")
             vConfig.User = InputBox("Usuario: ")
             vConfig.password = InputBox("Password: ")
'             vConfig.Integraciones = InputBox("Path integraciones: ")
             vConfig.Grabar
             MsgBox "Reinicie AriGes", vbCritical
             End
             Exit Sub
        End If
        
         'Abrimos conexion para comprobar el usuario
         'Luego, en funcion del nivel de usuario que tenga cerraremos la conexion
         'y la abriremos con usuario-codigo ajustado a su nivel
         If AbrirConexionUsuarios() = False Then
             MsgBox "La aplicación no puede continuar sin acceso a los datos. ", vbCritical
             End
        End If
         
         
         'Para que borre de la tabla temporal
        PrepararCarpetasEnvioMail
        DoEvents
         
         'Gestionar el nombre del PC para la asignacion de PC en el entorno de red
         GestionaPC
        
         'Leemos el ultimo usuario conectado
         UltimoUsuarioLogado True, UltUsu_
         Text1(0).Text = UltUsu_
         
         
         CargaCombo
         PosicionarCombo2 Combo1, Text1(0)
         
         
         T1 = T1 + 2.5 - Timer
         If T1 > 0 Then Espera T1
         
         CadenaDesdeOtroForm = ""
         Segundos = 60
         Timer1.Enabled = True
        
         
         PonerVisible True
         If Text1(0).Text <> "" Then
            PonerFoco Text1(1)
         Else
            PonerFoco Text1(0)
         End If
        
        
        LeeMayusculas_
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    LeeMayusculas_
End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    PonerVisible False
    T1 = Timer
    Text1(0).Text = ""
    Text1(1).Text = ""
    lblTiempo.Caption = ""
    lblVersion.Caption = "Ver. " & App.Major & "." & App.Minor & "." & App.Revision
    lblInd.Caption = "Cargando .."
    
    PrimeraVez = True
    CargaImagen
End Sub


Private Sub CargaImagen()
    On Error Resume Next
    Me.Image1 = LoadPicture(App.Path & "\arifon.dat")
    Me.Height = Me.Image1.Height
    Me.Width = Me.Image1.Width
    FijarText
        
    If Err.Number <> 0 Then
        MsgBox Err.Description & vbCrLf & vbCrLf & "Error cargando", vbCritical
        Set conn = Nothing
        End
    End If
End Sub


Private Sub FijarText()
Dim L As Long
    On Error GoTo EF
    L = Me.Width - Text1(1).Width - 360
    Text1(0).Left = L
    Text1(1).Left = L
    Me.Label1(0).Left = L
    Me.Label1(1).Left = L

    lblMay.Left = L
    Combo1.Left = L
    Me.imgBl.Left = L - Me.imgBl.Width - 30
    
    'L = Label1(2).Height + 220
    L = Me.Height - 720
    L = IIf(L <= 500, 500, L)
    
    Text1(1).Top = L
    imgBl.Top = L
    L = L - 360   '375 + algo
    Label1(1).Top = L + 60
    L = L - 300
    Text1(0).Top = L
    Combo1.Top = L - 90
    L = L - 360
    Label1(0).Top = L
    
    
    
    
    
    lblMay.Top = Me.Height - 300
    
    
    
EF:
    If Err.Number <> 0 Then MuestraError Err.Number
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If Text1(0).Text <> UltUsu_ Then UltimoUsuarioLogado False, Text1(0).Text
End Sub













Private Sub Text1_GotFocus(Index As Integer)
    With Text1(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub


Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        
        If Index = 1 Then
            If Me.Text1(Index).Text = "" Then PonerFocoOBj Me.Combo1: Exit Sub
        End If
        KeyAscii = 0
        SendKeys "{tab}"
    Else
        If KeyAscii = 27 Then
            CadenaDesdeOtroForm = ""
            Unload Me
        End If
    End If
End Sub

Private Sub Text1_LostFocus(Index As Integer)
    Text1(Index).Text = Trim(Text1(Index).Text)
    
    'Comprobamos si los dos estan con datos
    If Text1(0).Text <> "" And Text1(1).Text <> "" Then
        'Probar conexion usuario
        Validar
    
    Else
        If Index = 1 Then
            If Text1(1).Text = "" Then PonerFocoOBj Me.Combo1
        End If
    
    
    End If
        
    
End Sub



Private Sub Validar()
Dim NuevoUsu As Usuario
Dim OK As Byte

    'Validaremos el usuario y despues el password
    Set vUsu = New Usuario
    
    If vUsu.Leer(Text1(0).Text) = 0 Then
        'Con exito
        If vUsu.PasswdPROPIO = Text1(1).Text Then
            OK = 0
        Else
            OK = 1
        End If

    Else
        OK = 2
    End If
    
    If OK <> 0 Then
        MsgBox "Usuario-Clave Incorrecto", vbExclamation
            LeeMayusculas_
            Text1(1).Text = ""
            PonerFoco Text1(1)
            
    Else
        'OK
        Timer1.Enabled = False
       
        
                                'Con codejcok
        If vUsu.Skin >= 0 Then UsuarioCorrecto
        
        CadenaDesdeOtroForm = "OK"
        Unload Me
    End If

End Sub


Private Sub PonerVisible(visible As Boolean)
    'Label1(2).visible = Not visible  'Cargando
    lblInd.visible = Not visible
    Text1(0).visible = visible
    Combo1.visible = visible
    Text1(1).visible = visible
    Label1(0).visible = visible
    Label1(1).visible = visible
    
    
End Sub




Private Sub Timer1_Timer()
    Segundos = Segundos - 1
    If Segundos > 55 Then
        lblTiempo.Caption = ""
    Else
    
        lblTiempo.Caption = "Si no hace login, la pantalla se cerrará automáticamente en " & " " & Segundos & " segundos."
        lblTiempo.Refresh
        If Segundos < 1 Then
            Timer1.Enabled = False
            Unload Me
        End If
    End If
End Sub

Private Sub LeeMayusculas_()
Dim Tmp
Dim keys(0 To 255) As Byte
Dim VK_CAPITAL 'As Byte
    On Error GoTo el
    imgBl.visible = False
    lblMay.visible = False
    
    
       ' Tmp = GetKeyState(vbKeyCapital)
       ' If Tmp = 1 Then
       '     Image2.visible = True
       '     Me.Label1(4).visible = True
       ' End If
        
        GetKeyboardState keys(0)
        VK_CAPITAL = &H14
       ' Debug.Print Timer & " " & keys(VK_CAPITAL)
        If keys(VK_CAPITAL) = 1 Or keys(VK_CAPITAL) = 129 Then
            imgBl.visible = True
            lblMay.visible = True
        End If
   
el:
    Err.Clear
End Sub

Private Sub CargaCombo()
'Dim miRsAux As ADODB.Recordset

    Combo1.Clear
    'Conceptos
    Set miRsAux = New ADODB.Recordset
    
    miRsAux.Open "Select * from usuarios.usuarios where nivelariges <> -1 order by login", conn, adOpenForwardOnly, adLockOptimistic, adCmdText

    
    While Not miRsAux.EOF
        Combo1.AddItem miRsAux!Login
        Combo1.ItemData(Combo1.NewIndex) = miRsAux!CodUsu
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    'Aprovecho aqui para leer unas para el calendario
    TextosLabelEspanol = "select texto from usuarios.calendaretiquetas order by id"
    miRsAux.Open TextosLabelEspanol, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    TextosLabelEspanol = ""
    While Not miRsAux.EOF
        TextosLabelEspanol = TextosLabelEspanol & miRsAux!texto & "|"
        miRsAux.MoveNext
    Wend
    miRsAux.Close
        
    If TextosLabelEspanol = "" Then
        TextosLabelEspanol = "Ninguna|Importante|Negocios|Personal|Vacaciones|Atender|Viaje|"
        TextosLabelEspanol = TextosLabelEspanol & "Preparar|Cumpleaños|Aniversario|Llamada|"
    End If

    
    
    Set miRsAux = Nothing
    
End Sub


Public Sub pLabel(texto As String)

    Me.Label1(2).Caption = texto
    Label1(2).Refresh
    Espera 0.1
End Sub



Private Sub UsuarioCorrecto()
Dim Sql As String
Dim PrimeraBD As String
Dim EmpreProhibid As String

        Screen.MousePointer = vbHourglass
        CadenaDesdeOtroForm = "OK"
        Label1(2).Caption = "Leyendo ."  'Si tarda pondremos texto aquin
        
        PonerVisible False
        Me.Refresh
        Espera 0.1
        Me.Refresh

        Screen.MousePointer = vbHourglass
        
        
        pLabel "Conectando BD"
        
       Screen.MousePointer = vbHourglass
       
       EmpreProhibid = DevuelveProhibidasSys
        
       UltimoEmpresaLogada True, UltEmpre_
       
                                    'Va por modo antiguo. Vamos a buscar su empresa vinculada (arigesXX y guardarlo sin mas)
       If InStr(1, UltEmpre_, "|") > 0 Then
            DevuelveArigesXXX
            If UltEmpre_ <> "" Then UltimoEmpresaLogada False, UltEmpre_
        Else
            
        End If
        ' antes de cerrar la conexion cojo de usuarios.empresasariconta la primera que encuentre
        ' que no este bloqueada
        Sql = "select min(codempre) from usuarios.empresasariges  "
        Sql = Sql & " WHERE codempre>0 and not codempre in (select codempre from usuarios.usuarioempresasariges where codusu =" & vUsu.ID & ")"
        
        PrimeraBD = DevuelveValor(Sql)
    
        If UltEmpre_ = "" And PrimeraBD <> "" Then
            UltEmpre_ = PrimeraBD
            UltimoEmpresaLogada False, UltEmpre_
        End If
        CadenaDesdeOtroForm = UltEmpre_
        
        'Veo si la empresa prohibida es esta
        If EmpreProhibid <> "" Then
            Sql = Trim(Replace(CadenaDesdeOtroForm, "ariges", ""))
            Sql = "|" & Sql & "|"
            If InStr(1, EmpreProhibid, Sql) > 0 Then
                'Empresa entre las prohibidas. BUscamois otra
                If PrimeraBD = 0 Then
                    MsgBox "NO teiene acceso a empresas del sistema", vbCritical
                    Set conn = Nothing
                    End
                Else
                    CadenaDesdeOtroForm = "ariges" & PrimeraBD
                End If
            End If
                 
        End If
        pLabel "Abriendo " & CadenaDesdeOtroForm
        vUsu.CadenaConexion = CadenaDesdeOtroForm
        If AbrirConexion() = False Then
            CadenaDesdeOtroForm = "ariges" & PrimeraBD
            vUsu.CadenaConexion = CadenaDesdeOtroForm
            If AbrirConexion() = False Then
                End
            End If
        End If
        
        Screen.MousePointer = vbHourglass
        pLabel "Leyendo parametros"
        LeerDatosEmpresa
        LeerParametros
        
        If AbrirConexionConta(False) = False Then
            MsgBox "La aplicación no puede continuar sin acceso a los datos contables. ", vbCritical
            End
        End If
        
        vUsu.LeerTabPorDefecto
        
        'Otras acciones
        OtrasAcciones

        'La madre de todas las batallas
        pLabel "Cargando principal"

        Load frmPpal
        Load frmPpalN
End Sub


Private Function DevuelveProhibidasSys() As String
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open "Select * from usuarios.usuarioempresasariges WHERE codusu =" & vUsu.ID, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    DevuelveProhibidasSys = ""
    While Not miRsAux.EOF
        DevuelveProhibidasSys = DevuelveProhibidasSys & miRsAux.Fields(1) & "|"
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    If DevuelveProhibidasSys <> "" Then DevuelveProhibidasSys = "|" & DevuelveProhibidasSys
End Function

Private Sub DevuelveArigesXXX()
Dim B As Boolean
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open "Select nomempre,ariges from usuarios.empresasariges WHERE ariges<>''", conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    B = False
    UltEmpre_ = Replace(UltEmpre_, "|", "")
    While Not miRsAux.EOF
        If Not B Then
            If miRsAux!nomempre = UltEmpre_ Then
                UltEmpre_ = miRsAux!AriGes
                B = True
            End If
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    If Not B Then UltEmpre_ = ""
End Sub

