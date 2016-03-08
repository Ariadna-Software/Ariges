VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmConfParamGral 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Parámetros Generales"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   6720
   Icon            =   "frmConfParamGral.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   6720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   4755
      TabIndex        =   28
      Top             =   5160
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Left            =   720
      TabIndex        =   26
      Top             =   5040
      Width           =   2355
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   240
         Width           =   1920
      End
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   10
      Left            =   1665
      MaxLength       =   40
      TabIndex        =   10
      Tag             =   "eMail|T|S|||sparam|maiempre|||"
      Text            =   "Text1"
      Top             =   4320
      Width           =   4125
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   9
      Left            =   1665
      MaxLength       =   40
      TabIndex        =   9
      Tag             =   "Web|T|S|||sparam|wwwempre|||"
      Text            =   "Text1"
      Top             =   3870
      Width           =   4125
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   8
      Left            =   4065
      MaxLength       =   20
      TabIndex        =   8
      Tag             =   "Fax|T|S|||sparam|faxempre|||"
      Text            =   "Text1"
      Top             =   3420
      Width           =   1725
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   7
      Left            =   1665
      MaxLength       =   20
      TabIndex        =   7
      Tag             =   "Teléfono|T|S|||sparam|telempre|||"
      Text            =   "Text1"
      Top             =   3420
      Width           =   1725
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   6
      Left            =   1665
      MaxLength       =   9
      TabIndex        =   6
      Tag             =   "C.I.F.|T|N|||sparam|cifempre|||"
      Text            =   "Text1"
      Top             =   3000
      Width           =   1605
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   5
      Left            =   1665
      MaxLength       =   30
      TabIndex        =   5
      Tag             =   "Provincia|T|N|||sparam|proempre|||"
      Text            =   "Text1"
      Top             =   2535
      Width           =   1605
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   4
      Left            =   3465
      MaxLength       =   30
      TabIndex        =   4
      Tag             =   "Población|T|N|||sparam|pobempre|||"
      Text            =   "Text1"
      Top             =   2085
      Width           =   2325
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   3
      Left            =   1665
      MaxLength       =   6
      TabIndex        =   3
      Tag             =   "CPostal|T|N|||sparam|codpobla|||"
      Text            =   "Text1"
      Top             =   2085
      Width           =   765
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   2
      Left            =   1665
      MaxLength       =   40
      TabIndex        =   2
      Tag             =   "Domicilio de la Empresa|T|N|||sparam|domempre|||"
      Text            =   "Text1"
      Top             =   1650
      Width           =   4125
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   3600
      TabIndex        =   11
      Top             =   5160
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   4755
      TabIndex        =   12
      Top             =   5160
      Width           =   1035
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   1
      Left            =   1680
      MaxLength       =   40
      TabIndex        =   1
      Tag             =   "Nombre de la Empresa|T|N|||sparam|nomempre|||"
      Text            =   "Text1"
      Top             =   1200
      Width           =   4125
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   3360
      MaxLength       =   15
      TabIndex        =   0
      Tag             =   "Código Parámetros Generales|N|N|||sparam|codigo||S|"
      Text            =   "Text1"
      Top             =   1200
      Width           =   645
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   6720
      _ExtentX        =   11853
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "1"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   4440
      Top             =   600
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Image ImgMail 
      Height          =   240
      Index           =   0
      Left            =   1395
      Tag             =   "-1"
      ToolTipText     =   "Enviar e-mail"
      Top             =   4320
      Width           =   240
   End
   Begin VB.Image imgWeb 
      Height          =   255
      Left            =   1395
      Picture         =   "frmConfParamGral.frx":000C
      Stretch         =   -1  'True
      Tag             =   "-1"
      ToolTipText     =   "Abrir web"
      Top             =   3880
      Width           =   255
   End
   Begin VB.Image imgBuscar 
      Enabled         =   0   'False
      Height          =   240
      Left            =   1395
      Picture         =   "frmConfParamGral.frx":0596
      Tag             =   "-1"
      ToolTipText     =   "Buscar población"
      Top             =   2100
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Datos de la Empresa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   11
      Left            =   720
      TabIndex        =   25
      Top             =   600
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "E-mail"
      Height          =   255
      Index           =   10
      Left            =   780
      TabIndex        =   24
      Top             =   4320
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Web"
      Height          =   255
      Index           =   9
      Left            =   780
      TabIndex        =   23
      Top             =   3870
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Fax"
      Height          =   255
      Index           =   8
      Left            =   3700
      TabIndex        =   22
      Top             =   3420
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "Teléfono"
      Height          =   255
      Index           =   7
      Left            =   780
      TabIndex        =   21
      Top             =   3420
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "C.I.F."
      Height          =   255
      Index           =   6
      Left            =   780
      TabIndex        =   20
      Top             =   2985
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Provincia"
      Height          =   255
      Index           =   5
      Left            =   780
      TabIndex        =   19
      Top             =   2535
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Población"
      Height          =   255
      Index           =   4
      Left            =   2660
      TabIndex        =   18
      Top             =   2085
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "CPostal"
      Height          =   255
      Index           =   3
      Left            =   780
      TabIndex        =   17
      Top             =   2085
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Domicilio"
      Height          =   255
      Index           =   2
      Left            =   780
      TabIndex        =   16
      Top             =   1650
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Código"
      Height          =   255
      Index           =   0
      Left            =   2460
      TabIndex        =   14
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre"
      Height          =   255
      Index           =   1
      Left            =   780
      TabIndex        =   13
      Top             =   1200
      Width           =   615
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu mnModificar 
         Caption         =   "&Modificar"
         HelpContextID   =   1
         Shortcut        =   ^M
      End
      Begin VB.Menu mnBarra1 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmConfParamGral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private NombreTabla As String  'Nombre de la tabla o de la
Private Ordenacion As String
Private CadenaConsulta As String

Private VieneDeBuscar As Boolean
'Para cuando devuelve dos poblaciones con el mismo codigo Postal. Si viene de pulsar prismatico
'de busqueda poner el valor de poblacion seleccionado y no volver a recuperar de la Base de Datos

Private WithEvents frmCP As frmCPostal 'Codigos Postales
Attribute frmCP.VB_VarHelpID = -1

Dim Modo As Byte
'Solo hay Modo=0 Visualizacion y Modo=4 para Modificar datos


Private Sub cmdAceptar_Click()
    If DatosOk Then
        'Modifica datos en la Tabla: sparam
'        I = ModificaDesdeFormulario(Me)
        
        'Actualizar campos de la clase
        vEmpresa.nomempre = Text1(1).Text
        vEmpresa.ModificarDatos
        
        vParam.NombreEmpresa = Text1(1).Text
        vParam.DomicilioEmpresa = Text1(2).Text
        vParam.CPostal = Text1(3).Text
        vParam.Poblacion = Text1(4).Text
        vParam.Provincia = Text1(5).Text
        vParam.CifEmpresa = Text1(6).Text
        vParam.Telefono = Text1(7).Text
        vParam.Fax = Text1(8).Text
        vParam.WebEmpresa = Text1(9).Text
        vParam.MailEmpresa = Text1(10).Text
        vParam.Modificar
        TerminaBloquear
        
        Me.imgBuscar.Enabled = False
        PonerModo 0
        PonerFocoBtn Me.cmdSalir
    End If
End Sub


Private Sub cmdCancelar_Click()
    TerminaBloquear
    PonerCampos
    PonerModo 0
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    If Modo <> 4 Then PonerCadenaBusqueda 'Modo 4: MOdificar
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub Form_Load()
    'Icono del formulario
    Me.Icon = frmPpal.Icon
    
    'Icono de e-mail
    Me.ImgMail(0).Picture = frmPpal.imgListComun.ListImages(20).Picture


    ' ICONITOS DE LA BARRA
    With Me.Toolbar1
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 4   'Modificar
        .Buttons(3).Image = 15  'Salir
    End With
    
    VieneDeBuscar = False
    
    '## A mano
    NombreTabla = "sparam"
    Ordenacion = " ORDER BY codigo"
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = Conn
    CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    PonerModo 0
'    PonerCadenaBusqueda
End Sub


Private Sub PonerCadenaBusqueda()
On Error GoTo EEPonerBusq
    Screen.MousePointer = vbHourglass
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        MsgBox "No hay ningún registro en la tabla " & NombreTabla, vbInformation
        Screen.MousePointer = vbDefault
        Exit Sub
    Else
        Data1.Recordset.MoveFirst
        PonerCampos
        'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
        'Si estamos en Insertar además limpia los campos Text1
        BloquearText1 Me, Modo
    End If

    Screen.MousePointer = vbDefault
    Exit Sub
EEPonerBusq:
    MuestraError Err.Number, "PonerCadenaBusqueda"
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerCampos()
    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma Me, Data1
    If Trim(Text1(3).Text) = "0" Then Text1(3).Text = ""
    If Trim(Text1(6).Text) = "0" Then Text1(6).Text = ""
End Sub


Private Sub frmCP_DatoSeleccionado(CadenaSeleccion As String)
'Formulario Mantenimiento C. Postales
Dim Indice As Byte
Dim devuelve As String

    Indice = 3
    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'CPostal
    Text1(Indice + 1).Text = ObtenerPoblacion(Text1(Indice).Text, devuelve)  'Poblacion
    'provincia
    Text1(Indice + 2).Text = devuelve
End Sub

Private Sub imgBuscar_Click()
    If Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass

    'Codigo Postal
    Set frmCP = New frmCPostal
    frmCP.DatosADevolverBusqueda = "0"
    frmCP.Show vbModal
    Set frmCP = Nothing

    VieneDeBuscar = True
    PonerFoco Text1(3)
    Screen.MousePointer = vbDefault
End Sub


Private Sub ImgMail_Click(Index As Integer)
'Abrir Outlook para enviar e-mail
Dim dirMail As String

'    If Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass
    
    Select Case Index
        Case 0: dirMail = Text1(10).Text
    End Select

    If LanzaMailGnral(dirMail) Then espera 2
    Screen.MousePointer = vbDefault
End Sub

Private Sub imgWeb_Click()
'Abrimos el explorador de windows con la pagina Web del cliente

'    If Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass

    If LanzaHomeGnral(Text1(9).Text) Then espera 2
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnModificar_Click()
    If BLOQUEADesdeFormulario(Me) Then BotonModificar
End Sub

Private Sub mnSalir_Click()
    Unload Me
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    ConseguirFoco Text1(Index), 3
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
End Sub


Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress (KeyAscii)
End Sub


Private Sub Text1_LostFocus(Index As Integer)
Dim devuelve As String

    Select Case Index
        Case 3 'CPostal
            If Text1(Index).Text <> "" And Not VieneDeBuscar Then
                Text1(4).Text = ObtenerPoblacion(Text1(Index).Text, devuelve)
                Text1(5).Text = devuelve
            End If
            VieneDeBuscar = False
    End Select
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1  'Modificar
                mnModificar_Click
        Case 3 'Salir
            mnSalir_Click
    End Select
End Sub


Private Sub BotonModificar()
    'Añadiremos el boton de aceptar y demas objetos para insertar
'    Me.lblIndicador.Caption = "MODIFICAR"
    PonerModo 4
    Me.imgBuscar.Enabled = True
    PonerFoco Text1(1)
End Sub


Private Function DatosOk() As Boolean
Dim b As Boolean
    DatosOk = False
    b = CompForm(Me, 1)
    DatosOk = b
End Function


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub PonerBotonCabecera(b As Boolean)
    Me.cmdAceptar.visible = Not b
    Me.cmdCancelar.visible = Not b
    Me.cmdSalir.visible = b
    If b Then Me.lblIndicador.Caption = ""
End Sub


Private Sub PonerModo(vModo As Byte)
Dim b As Boolean

    Modo = vModo
    b = (Modo = 0)
    PonerIndicador Me.lblIndicador, Modo
'    If b Then Me.lblIndicador.Caption = ""
    
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1
    BloquearText1 Me, Modo
    
    'Poner Botones Aceptar/Cancelar si estamos Modificando datos
    PonerBotonCabecera b
    
    'Solo si es root o administrador puede modificar el registro
    cmdAceptar.Enabled = (vUsu.Nivel <= 1)
    
    'Modificar
    Toolbar1.Buttons(1).Enabled = b
    Me.mnModificar.Enabled = b

    PonerOpcionesMenu   'Activar opciones de menu según nivel
                        'de permisos del usuario
End Sub


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub

