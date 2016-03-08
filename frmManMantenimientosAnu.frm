VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmManMantenimientosAnu 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimientos A N U L A D O S"
   ClientHeight    =   7425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11490
   Icon            =   "frmManMantenimientosAnu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7425
   ScaleWidth      =   11490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   910
      Left            =   120
      TabIndex        =   79
      Top             =   425
      Width           =   11175
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   2
         Left            =   6960
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "Nº Mantenimiento|T|N|||scamana|nummante||S|"
         Text            =   "Text1"
         Top             =   160
         Width           =   1365
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   3
         Left            =   6960
         MaxLength       =   15
         TabIndex        =   3
         Tag             =   "Fecha Inicio|F|N|||scamana|fechaini|dd/mm/yyyy|N|"
         Text            =   "Text1"
         Top             =   510
         Width           =   1365
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   1
         Left            =   2240
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   81
         Text            =   "Text2"
         Top             =   510
         Width           =   3165
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   0
         Left            =   2240
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   80
         Text            =   "Text2"
         Top             =   160
         Width           =   3165
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   1
         Left            =   1400
         MaxLength       =   3
         TabIndex        =   1
         Tag             =   "Cód. Dirección|N|S|0|999|scamana|coddirec|000|N|"
         Text            =   "Text1"
         Top             =   510
         Width           =   765
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   0
         Left            =   1400
         MaxLength       =   6
         TabIndex        =   0
         Tag             =   "Código Cliente|N|N|0|999999|scamana|codclien|000000|S|"
         Text            =   "Text"
         Top             =   160
         Width           =   750
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         Index           =   1
         X1              =   8760
         X2              =   11040
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         Index           =   0
         X1              =   8760
         X2              =   11040
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label Label2 
         Caption         =   "ANULADOS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   495
         Index           =   0
         Left            =   8880
         TabIndex        =   92
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Mantenim."
         Height          =   255
         Index           =   13
         Left            =   5640
         TabIndex        =   85
         Top             =   165
         Width           =   1335
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   6600
         Picture         =   "frmManMantenimientosAnu.frx":000C
         ToolTipText     =   "Buscar fecha"
         Top             =   510
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Inicio"
         Height          =   255
         Index           =   14
         Left            =   5640
         TabIndex        =   84
         Top             =   510
         Width           =   975
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1120
         ToolTipText     =   "Buscar direc./dpto"
         Top             =   520
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1120
         Picture         =   "frmManMantenimientosAnu.frx":0097
         ToolTipText     =   "Buscar cliente"
         Top             =   170
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Cod. Direc."
         Height          =   255
         Index           =   1
         Left            =   160
         TabIndex        =   83
         Top             =   510
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Cód. Cliente"
         Height          =   255
         Index           =   0
         Left            =   160
         TabIndex        =   82
         Top             =   160
         Width           =   855
      End
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   21
      Left            =   2040
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   56
      Text            =   "Text2"
      Top             =   6960
      Visible         =   0   'False
      Width           =   4365
   End
   Begin VB.Frame Frame1 
      Height          =   475
      Index           =   0
      Left            =   120
      TabIndex        =   43
      Top             =   6855
      Width           =   1815
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   240
         Left            =   120
         TabIndex        =   44
         Top             =   180
         Width           =   1515
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   10290
      TabIndex        =   39
      Top             =   6960
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   9000
      TabIndex        =   38
      Top             =   6960
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   5640
      Top             =   6960
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   45
      Top             =   0
      Width           =   11490
      _ExtentX        =   20267
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   20
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ver todos"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Nuevo"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Revisiones"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Histórico"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Componentes"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Último"
            ImageIndex      =   9
         EndProperty
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   8400
         TabIndex        =   46
         Top             =   120
         Width           =   1215
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5520
      Left            =   120
      TabIndex        =   47
      Top             =   1335
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   9737
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Datos básicos"
      TabPicture(0)   =   "frmManMantenimientosAnu.frx":0199
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Text1(37)"
      Tab(0).Control(1)=   "Text1(36)"
      Tab(0).Control(2)=   "Text1(35)"
      Tab(0).Control(3)=   "Text1(34)"
      Tab(0).Control(4)=   "Text1(7)"
      Tab(0).Control(5)=   "Text1(6)"
      Tab(0).Control(6)=   "chkBaterias"
      Tab(0).Control(7)=   "cboTipoPago"
      Tab(0).Control(8)=   "Text2(5)"
      Tab(0).Control(9)=   "Text2(4)"
      Tab(0).Control(10)=   "Text1(5)"
      Tab(0).Control(11)=   "Text1(4)"
      Tab(0).Control(12)=   "Frame2"
      Tab(0).Control(13)=   "Label1(9)"
      Tab(0).Control(14)=   "Label1(6)"
      Tab(0).Control(15)=   "Label1(4)"
      Tab(0).Control(16)=   "Label1(2)"
      Tab(0).Control(17)=   "Label1(54)"
      Tab(0).Control(18)=   "imgBuscar(3)"
      Tab(0).Control(19)=   "imgBuscar(2)"
      Tab(0).Control(20)=   "Label1(7)"
      Tab(0).Control(21)=   "Label1(36)"
      Tab(0).Control(22)=   "Label1(15)"
      Tab(0).Control(23)=   "Label1(34)"
      Tab(0).ControlCount=   24
      TabCaption(1)   =   "Observaciones"
      TabPicture(1)   =   "frmManMantenimientosAnu.frx":01B5
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1(3)"
      Tab(1).Control(1)=   "Label1(5)"
      Tab(1).Control(2)=   "Text1(32)"
      Tab(1).Control(3)=   "Text1(33)"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Anulación / histórico"
      TabPicture(2)   =   "frmManMantenimientosAnu.frx":01D1
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "imgFlecha(1)"
      Tab(2).Control(1)=   "imgFlecha(0)"
      Tab(2).Control(2)=   "Label1(53)"
      Tab(2).Control(3)=   "Label1(52)"
      Tab(2).Control(4)=   "Label1(51)"
      Tab(2).Control(5)=   "Label1(50)"
      Tab(2).Control(6)=   "Label1(49)"
      Tab(2).Control(7)=   "Label1(48)"
      Tab(2).Control(8)=   "Label1(47)"
      Tab(2).Control(9)=   "Label1(46)"
      Tab(2).Control(10)=   "Label1(45)"
      Tab(2).Control(11)=   "Label1(44)"
      Tab(2).Control(12)=   "Label1(43)"
      Tab(2).Control(13)=   "Label1(42)"
      Tab(2).Control(14)=   "Label1(41)"
      Tab(2).Control(15)=   "Label1(40)"
      Tab(2).Control(16)=   "Label1(39)"
      Tab(2).Control(17)=   "Label1(38)"
      Tab(2).Control(18)=   "Label1(37)"
      Tab(2).Control(19)=   "Label1(10)"
      Tab(2).Control(20)=   "Label1(11)"
      Tab(2).Control(21)=   "Label1(12)"
      Tab(2).Control(22)=   "Label1(18)"
      Tab(2).Control(23)=   "Label1(55)"
      Tab(2).Control(24)=   "Text2(28)"
      Tab(2).Control(25)=   "Text2(29)"
      Tab(2).Control(26)=   "Text2(30)"
      Tab(2).Control(27)=   "Text2(31)"
      Tab(2).Control(28)=   "Text2(32)"
      Tab(2).Control(29)=   "Text2(33)"
      Tab(2).Control(30)=   "Text2(41)"
      Tab(2).Control(31)=   "Text2(42)"
      Tab(2).Control(32)=   "Text2(43)"
      Tab(2).Control(33)=   "Text2(44)"
      Tab(2).Control(34)=   "Text2(45)"
      Tab(2).Control(35)=   "Text2(46)"
      Tab(2).Control(36)=   "Text2(47)"
      Tab(2).Control(37)=   "Text2(34)"
      Tab(2).Control(38)=   "Text2(22)"
      Tab(2).Control(39)=   "Text2(23)"
      Tab(2).Control(40)=   "Text2(24)"
      Tab(2).Control(41)=   "Text2(25)"
      Tab(2).Control(42)=   "Text2(26)"
      Tab(2).Control(43)=   "Text2(27)"
      Tab(2).Control(44)=   "Text2(35)"
      Tab(2).Control(45)=   "Text2(36)"
      Tab(2).Control(46)=   "Text2(37)"
      Tab(2).Control(47)=   "Text2(38)"
      Tab(2).Control(48)=   "Text2(39)"
      Tab(2).Control(49)=   "Text2(40)"
      Tab(2).Control(50)=   "Text2(6)"
      Tab(2).Control(51)=   "Text1(38)"
      Tab(2).Control(52)=   "Text1(39)"
      Tab(2).Control(53)=   "Text1(40)"
      Tab(2).ControlCount=   54
      TabCaption(3)   =   "Seguimiento"
      TabPicture(3)   =   "frmManMantenimientosAnu.frx":01ED
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Text1(41)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      Begin VB.TextBox Text1 
         Height          =   4515
         Index           =   41
         Left            =   360
         MaxLength       =   500
         MultiLine       =   -1  'True
         TabIndex        =   146
         Tag             =   "Observ. seguimiento|T|S|||scamana|obserseg||N|"
         Text            =   "frmManMantenimientosAnu.frx":0209
         Top             =   720
         Width           =   10245
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   40
         Left            =   -74040
         MaxLength       =   15
         TabIndex        =   140
         Tag             =   "Anticipado Sig.|F|S|0||scamana|fechabaj|dd/mm/yyyy|S|"
         Text            =   "Text1"
         Top             =   960
         Width           =   1365
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   39
         Left            =   -71880
         MaxLength       =   30
         TabIndex        =   139
         Tag             =   "P|T|S|||scamana|usuario||N|"
         Text            =   "WWWWWWWWWWWWWWW"
         Top             =   960
         Width           =   3045
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   38
         Left            =   -67560
         MaxLength       =   30
         TabIndex        =   138
         Tag             =   "Tipo Contrato|T|N|||scamana|codincid||N|"
         Text            =   "Text1"
         Top             =   960
         Width           =   540
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   6
         Left            =   -66990
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   137
         Text            =   "Text2"
         Top             =   960
         Width           =   2805
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   40
         Left            =   -71160
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   119
         Text            =   "Text2"
         Top             =   4335
         Width           =   1485
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   39
         Left            =   -71160
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   118
         Text            =   "Text2"
         Top             =   3975
         Width           =   1485
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   38
         Left            =   -71160
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   117
         Text            =   "Text2"
         Top             =   3615
         Width           =   1485
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   37
         Left            =   -71160
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   116
         Text            =   "Text2"
         Top             =   3255
         Width           =   1485
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   36
         Left            =   -71160
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   115
         Text            =   "Text2"
         Top             =   2895
         Width           =   1485
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   35
         Left            =   -71160
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   114
         Text            =   "Text2"
         Top             =   2535
         Width           =   1485
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   27
         Left            =   -72795
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   113
         Text            =   "Text2"
         Top             =   4335
         Width           =   1485
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   26
         Left            =   -72795
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   112
         Text            =   "Text2"
         Top             =   3975
         Width           =   1485
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   25
         Left            =   -72795
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   111
         Text            =   "Text2"
         Top             =   3615
         Width           =   1485
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   24
         Left            =   -72795
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   110
         Text            =   "Text2"
         Top             =   3255
         Width           =   1485
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   23
         Left            =   -72795
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   109
         Text            =   "Text2"
         Top             =   2895
         Width           =   1485
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   22
         Left            =   -72795
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   108
         Text            =   "Text2"
         Top             =   2535
         Width           =   1485
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   34
         Left            =   -67800
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   107
         Text            =   "Text2"
         Top             =   4815
         Width           =   1485
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   47
         Left            =   -66120
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   106
         Text            =   "Text2"
         Top             =   4815
         Width           =   1485
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   46
         Left            =   -66120
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   105
         Text            =   "Text2"
         Top             =   4335
         Width           =   1485
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   45
         Left            =   -66120
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   104
         Text            =   "Text2"
         Top             =   3975
         Width           =   1485
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   44
         Left            =   -66120
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   103
         Text            =   "Text2"
         Top             =   3615
         Width           =   1485
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   43
         Left            =   -66120
         MaxLength       =   15
         TabIndex        =   102
         Text            =   "Text2"
         Top             =   3255
         Width           =   1485
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   42
         Left            =   -66120
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   101
         Text            =   "Text2"
         Top             =   2895
         Width           =   1485
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   41
         Left            =   -66120
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   100
         Text            =   "Text2"
         Top             =   2535
         Width           =   1485
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   33
         Left            =   -67800
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   99
         Text            =   "Text2"
         Top             =   4335
         Width           =   1485
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   32
         Left            =   -67800
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   98
         Text            =   "Text2"
         Top             =   3975
         Width           =   1485
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   31
         Left            =   -67800
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   97
         Text            =   "Text2"
         Top             =   3615
         Width           =   1485
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   30
         Left            =   -67800
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   96
         Text            =   "Text2"
         Top             =   3255
         Width           =   1485
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   29
         Left            =   -67800
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   95
         Text            =   "Text2"
         Top             =   2895
         Width           =   1485
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   28
         Left            =   -67800
         MaxLength       =   15
         TabIndex        =   94
         Text            =   "Text2"
         Top             =   2535
         Width           =   1485
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   37
         Left            =   -71520
         MaxLength       =   30
         TabIndex        =   8
         Tag             =   "P|T|S|||scamana|attetiqu||N|"
         Text            =   "WWWWWWWWWWWWWWW"
         Top             =   840
         Width           =   2805
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   36
         Left            =   -73320
         MaxLength       =   60
         TabIndex        =   11
         Tag             =   "P|T|S|||scamana|concefac||N|"
         Text            =   "WWWWWWWWW0WWWWWWWWW0WWWWWWWWW0WWWWWWWWW0WWWWWWWWW0WWWWWWWW60"
         Top             =   1320
         Width           =   8085
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   35
         Left            =   -67920
         MaxLength       =   15
         TabIndex        =   13
         Tag             =   "P|T|S|||scamana|producto||N|"
         Text            =   "WWWWWWWWWWWWWWW"
         Top             =   1800
         Width           =   2805
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   34
         Left            =   -73320
         MaxLength       =   35
         TabIndex        =   12
         Tag             =   "P|T|S|||scamana|persconta||N|"
         Text            =   "WWWWWWWWW0WWWWWWWWW0WWWWWWWWW0"
         Top             =   1800
         Width           =   4245
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   7
         Left            =   -73800
         MaxLength       =   15
         TabIndex        =   5
         Tag             =   "Anticipado Sig.|N|S|0||scamana|anticip2|##,###,##0.00|N|"
         Text            =   "Text1"
         Top             =   840
         Width           =   1365
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   6
         Left            =   -73800
         MaxLength       =   15
         TabIndex        =   4
         Tag             =   "Anticipado Act.|N|S|0||scamana|anticip1|##,###,##0.00|N|"
         Text            =   "Text1"
         Top             =   480
         Width           =   1365
      End
      Begin VB.TextBox Text1 
         Height          =   1755
         Index           =   33
         Left            =   -74400
         MaxLength       =   500
         MultiLine       =   -1  'True
         TabIndex        =   41
         Tag             =   "Observ. Técnico|T|S|||scamana|obsertec||N|"
         Text            =   "frmManMantenimientosAnu.frx":020F
         Top             =   3240
         Width           =   9885
      End
      Begin VB.TextBox Text1 
         Height          =   1755
         Index           =   32
         Left            =   -74400
         MaxLength       =   500
         MultiLine       =   -1  'True
         TabIndex        =   40
         Tag             =   "Observ. Comercial|T|S|||scamana|observac||N|"
         Text            =   "frmManMantenimientosAnu.frx":0217
         Top             =   840
         Width           =   9885
      End
      Begin VB.CheckBox chkBaterias 
         Caption         =   "Baterias"
         Height          =   255
         Left            =   -69840
         TabIndex        =   7
         Tag             =   "Baterías|N|N|||scamana|baterias||N|"
         Top             =   480
         Width           =   975
      End
      Begin VB.ComboBox cboTipoPago 
         Height          =   315
         Left            =   -71520
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Tag             =   "Tipo de Pago|N|N|||scamana|tipopago||N|"
         Top             =   495
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   5
         Left            =   -66840
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   52
         Text            =   "Text2"
         Top             =   860
         Width           =   2805
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   4
         Left            =   -66840
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   51
         Text            =   "Text2"
         Top             =   500
         Width           =   2805
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   5
         Left            =   -67395
         MaxLength       =   30
         TabIndex        =   10
         Tag             =   "Forma de Pago|N|N|0|999|scamana|codforpa|000|N|"
         Text            =   "Text1"
         Top             =   860
         Width           =   525
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   4
         Left            =   -67410
         MaxLength       =   30
         TabIndex        =   9
         Tag             =   "Tipo Contrato|T|N|||scamana|codtipco||N|"
         Text            =   "Text1"
         Top             =   500
         Width           =   540
      End
      Begin VB.Frame Frame2 
         Height          =   3195
         Left            =   -74880
         TabIndex        =   59
         Top             =   2160
         Width           =   10740
         Begin VB.ComboBox cmbMes 
            Height          =   315
            Left            =   2640
            Style           =   2  'Dropdown List
            TabIndex        =   89
            Tag             =   "Ultimo mes facturado|N|N|1||scamana|ulmesfac||N|"
            Top             =   2760
            Width           =   2175
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   20
            Left            =   6720
            MaxLength       =   15
            TabIndex        =   20
            Tag             =   "Julio Actual|N|S|0||scamana|mes07act|##,###,##0.00|N|"
            Text            =   "Text1"
            Top             =   480
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   22
            Left            =   6720
            MaxLength       =   15
            TabIndex        =   21
            Tag             =   "Agosto Actual|N|S|0||scamana|mes08act|##,###,##0.00|N|"
            Text            =   "Text1"
            Top             =   840
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   24
            Left            =   6720
            MaxLength       =   15
            TabIndex        =   22
            Tag             =   "Septiembre Actual|N|S|0||scamana|mes09act|##,###,##0.00|N|"
            Text            =   "Text1"
            Top             =   1200
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   26
            Left            =   6720
            MaxLength       =   15
            TabIndex        =   23
            Tag             =   "Octubre Actual|N|S|0||scamana|mes10act|##,###,##0.00|N|"
            Text            =   "Text1"
            Top             =   1560
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   28
            Left            =   6720
            MaxLength       =   15
            TabIndex        =   24
            Tag             =   "Noviembre Actual|N|S|0||scamana|mes11act|##,###,##0.00|N|"
            Text            =   "Text1"
            Top             =   1920
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   30
            Left            =   6720
            MaxLength       =   15
            TabIndex        =   25
            Tag             =   "Diciembre Actual|N|S|0||scamana|mes12act|##,###,##0.00|N|"
            Text            =   "Text1"
            Top             =   2280
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   21
            Left            =   8400
            MaxLength       =   15
            ScrollBars      =   2  'Vertical
            TabIndex        =   32
            Tag             =   "Julio Siguiente|N|S|0||scamana|mes07sig|##,###,##0.00|N|"
            Text            =   "Text1"
            Top             =   480
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   23
            Left            =   8400
            MaxLength       =   15
            TabIndex        =   33
            Tag             =   "Agosto Siguiente|N|S|0||scamana|mes08sig|##,###,##0.00|N|"
            Text            =   "Text1"
            Top             =   840
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   25
            Left            =   8400
            MaxLength       =   15
            TabIndex        =   34
            Tag             =   "Septiembre Siguiente|N|S|0||scamana|mes09sig|##,###,##0.00|N|"
            Text            =   "Text1"
            Top             =   1200
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   27
            Left            =   8400
            MaxLength       =   15
            TabIndex        =   35
            Tag             =   "Octubre Siguiente|N|S|0||scamana|mes10sig|##,###,##0.00|N|"
            Text            =   "Text1"
            Top             =   1560
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   29
            Left            =   8400
            MaxLength       =   15
            TabIndex        =   36
            Tag             =   "Noviembre Siguiente|N|S|0||scamana|mes11sig|##,###,##0.00|N|"
            Text            =   "Text1"
            Top             =   1920
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   31
            Left            =   8400
            MaxLength       =   15
            TabIndex        =   37
            Tag             =   "Diciembre Siguiente|N|S|0||scamana|mes12sig|##,###,##0.00|N|"
            Text            =   "Text1"
            Top             =   2280
            Width           =   1485
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   3
            Left            =   8400
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   69
            Text            =   "Text2"
            Top             =   2760
            Width           =   1485
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   2
            Left            =   6720
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   68
            Text            =   "Text2"
            Top             =   2760
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   8
            Left            =   1605
            MaxLength       =   15
            TabIndex        =   14
            Tag             =   "Enero Actual|N|S|0||scamana|mes01act|##,###,##0.00|N|"
            Text            =   "Text1"
            Top             =   480
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   10
            Left            =   1605
            MaxLength       =   15
            TabIndex        =   15
            Tag             =   "Febrero Actual|N|S|0||scamana|mes02act|##,###,##0.00|N|"
            Text            =   "Text1"
            Top             =   840
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   12
            Left            =   1605
            MaxLength       =   15
            TabIndex        =   16
            Tag             =   "Marzo Actual|N|S|0||scamana|mes03act|##,###,##0.00|N|"
            Text            =   "Text1"
            Top             =   1200
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   14
            Left            =   1605
            MaxLength       =   15
            TabIndex        =   17
            Tag             =   "Abril Actual|N|S|0||scamana|mes04act|##,###,##0.00|N|"
            Text            =   "Text1"
            Top             =   1560
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   16
            Left            =   1605
            MaxLength       =   15
            TabIndex        =   18
            Tag             =   "Mayo Actual|N|S|0||scamana|mes05act|##,###,##0.00|N|"
            Text            =   "Text1"
            Top             =   1920
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   18
            Left            =   1605
            MaxLength       =   15
            TabIndex        =   19
            Tag             =   "Junio Actual|N|S|0||scamana|mes06act|##,###,##0.00|N|"
            Text            =   "Text1"
            Top             =   2280
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   9
            Left            =   3240
            MaxLength       =   15
            TabIndex        =   26
            Tag             =   "Enero Siguiente|N|S|0||scamana|mes01sig|##,###,##0.00|N|"
            Text            =   "Text1"
            Top             =   480
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   11
            Left            =   3240
            MaxLength       =   15
            TabIndex        =   27
            Tag             =   "Febrero Siguiente|N|S|0||scamana|mes02sig|##,###,##0.00|N|"
            Text            =   "Text1"
            Top             =   840
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   13
            Left            =   3240
            MaxLength       =   15
            TabIndex        =   28
            Tag             =   "Marzo Siguiente|N|S|0||scamana|mes03sig|##,###,##0.00|N|"
            Text            =   "Text1"
            Top             =   1200
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   15
            Left            =   3240
            MaxLength       =   15
            TabIndex        =   29
            Tag             =   "Abril Siguiente|N|S|0||scamana|mes04sig|##,###,##0.00|N|"
            Text            =   "Text1"
            Top             =   1560
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   17
            Left            =   3240
            MaxLength       =   15
            TabIndex        =   30
            Tag             =   "Mayo Siguiente|N|S|0||scamana|mes05sig|##,###,##0.00|N|"
            Text            =   "Text1"
            Top             =   1920
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   19
            Left            =   3240
            MaxLength       =   15
            TabIndex        =   31
            Tag             =   "Junio Siguiente|N|S|0||scamana|mes06sig|##,###,##0.00|N|"
            Text            =   "Text1"
            Top             =   2280
            Width           =   1485
         End
         Begin VB.Label Label1 
            Caption         =   "Ultimo mes facturado"
            Height          =   195
            Index           =   8
            Left            =   960
            TabIndex        =   90
            Top             =   2840
            Width           =   1485
         End
         Begin VB.Label Label1 
            Caption         =   "Julio"
            Height          =   255
            Index           =   24
            Left            =   5880
            TabIndex        =   78
            Top             =   585
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Agosto"
            Height          =   255
            Index           =   25
            Left            =   5880
            TabIndex        =   77
            Top             =   900
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Septiembre"
            Height          =   255
            Index           =   26
            Left            =   5880
            TabIndex        =   76
            Top             =   1260
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Octubre"
            Height          =   255
            Index           =   27
            Left            =   5880
            TabIndex        =   75
            Top             =   1620
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Noviembre"
            Height          =   255
            Index           =   28
            Left            =   5880
            TabIndex        =   74
            Top             =   1965
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Diciembre"
            Height          =   255
            Index           =   29
            Left            =   5880
            TabIndex        =   73
            Top             =   2325
            Width           =   855
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "ACTUAL"
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
            Index           =   32
            Left            =   6840
            TabIndex        =   72
            Top             =   240
            Width           =   1485
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "SIGUIENTE"
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
            Index           =   33
            Left            =   8400
            TabIndex        =   71
            Top             =   240
            Width           =   1485
         End
         Begin VB.Label Label1 
            Caption         =   "TOTAL"
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
            Index           =   21
            Left            =   5880
            TabIndex        =   70
            Top             =   2760
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Enero"
            Height          =   255
            Index           =   16
            Left            =   960
            TabIndex        =   67
            Top             =   480
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Febrero"
            Height          =   255
            Index           =   17
            Left            =   960
            TabIndex        =   66
            Top             =   840
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Marzo"
            Height          =   255
            Index           =   19
            Left            =   960
            TabIndex        =   65
            Top             =   1200
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Abril"
            Height          =   255
            Index           =   20
            Left            =   960
            TabIndex        =   64
            Top             =   1560
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Mayo"
            Height          =   255
            Index           =   22
            Left            =   960
            TabIndex        =   63
            Top             =   1920
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Junio"
            Height          =   255
            Index           =   23
            Left            =   960
            TabIndex        =   62
            Top             =   2280
            Width           =   615
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "ACTUAL"
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
            Index           =   30
            Left            =   1680
            TabIndex        =   61
            Top             =   240
            Width           =   1365
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "SIGUIENTE"
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
            Index           =   31
            Left            =   3240
            TabIndex        =   60
            Top             =   240
            Width           =   1485
         End
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Motivo"
         Height          =   255
         Index           =   55
         Left            =   -68400
         TabIndex        =   145
         Top             =   990
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Histórico"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   495
         Index           =   18
         Left            =   -74760
         TabIndex        =   144
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Anulación"
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
         Height          =   495
         Index           =   12
         Left            =   -74760
         TabIndex        =   143
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha"
         Height          =   255
         Index           =   11
         Left            =   -74640
         TabIndex        =   142
         Top             =   990
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Usuario"
         Height          =   255
         Index           =   10
         Left            =   -72480
         TabIndex        =   141
         Top             =   990
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "SIGUIENTE"
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
         Index           =   37
         Left            =   -71160
         TabIndex        =   136
         Top             =   2175
         Width           =   1125
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "ACTUAL"
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
         Index           =   38
         Left            =   -72240
         TabIndex        =   135
         Top             =   2175
         Width           =   885
      End
      Begin VB.Label Label1 
         Caption         =   "Junio"
         Height          =   255
         Index           =   39
         Left            =   -73680
         TabIndex        =   134
         Top             =   4335
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Mayo"
         Height          =   255
         Index           =   40
         Left            =   -73680
         TabIndex        =   133
         Top             =   3975
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Abril"
         Height          =   255
         Index           =   41
         Left            =   -73680
         TabIndex        =   132
         Top             =   3615
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Marzo"
         Height          =   255
         Index           =   42
         Left            =   -73680
         TabIndex        =   131
         Top             =   3255
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Febrero"
         Height          =   255
         Index           =   43
         Left            =   -73680
         TabIndex        =   130
         Top             =   2895
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Enero"
         Height          =   255
         Index           =   44
         Left            =   -73680
         TabIndex        =   129
         Top             =   2535
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "TOTAL"
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
         Index           =   45
         Left            =   -68640
         TabIndex        =   128
         Top             =   4815
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "SIGUIENTE"
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
         Index           =   46
         Left            =   -66120
         TabIndex        =   127
         Top             =   2175
         Width           =   1485
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "ACTUAL"
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
         Index           =   47
         Left            =   -67800
         TabIndex        =   126
         Top             =   2175
         Width           =   1485
      End
      Begin VB.Label Label1 
         Caption         =   "Diciembre"
         Height          =   255
         Index           =   48
         Left            =   -68640
         TabIndex        =   125
         Top             =   4335
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Noviembre"
         Height          =   255
         Index           =   49
         Left            =   -68640
         TabIndex        =   124
         Top             =   3975
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Octubre"
         Height          =   255
         Index           =   50
         Left            =   -68640
         TabIndex        =   123
         Top             =   3615
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Septiembre"
         Height          =   255
         Index           =   51
         Left            =   -68640
         TabIndex        =   122
         Top             =   3255
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Agosto"
         Height          =   255
         Index           =   52
         Left            =   -68640
         TabIndex        =   121
         Top             =   2895
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Julio"
         Height          =   255
         Index           =   53
         Left            =   -68640
         TabIndex        =   120
         Top             =   2535
         Width           =   855
      End
      Begin VB.Image imgFlecha 
         Height          =   480
         Index           =   0
         Left            =   -72840
         Picture         =   "frmManMantenimientosAnu.frx":021D
         Top             =   2055
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image imgFlecha 
         Height          =   480
         Index           =   1
         Left            =   -70080
         Picture         =   "frmManMantenimientosAnu.frx":065F
         Top             =   2055
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label Label1 
         Caption         =   "Att ETIQ."
         Height          =   255
         Index           =   9
         Left            =   -72360
         TabIndex        =   91
         Top             =   885
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Concepto factura"
         Height          =   255
         Index           =   6
         Left            =   -74760
         TabIndex        =   88
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Producto"
         Height          =   255
         Index           =   4
         Left            =   -68760
         TabIndex        =   87
         Top             =   1840
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Persona contacto"
         Height          =   255
         Index           =   2
         Left            =   -74760
         TabIndex        =   86
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Anticipado 2"
         Height          =   255
         Index           =   54
         Left            =   -74760
         TabIndex        =   58
         Top             =   840
         Width           =   975
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   -67680
         ToolTipText     =   "Buscar forma de pago"
         Top             =   885
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   -67680
         ToolTipText     =   "Buscar tipo contrato"
         Top             =   510
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Anticipado 1"
         Height          =   255
         Index           =   7
         Left            =   -74760
         TabIndex        =   55
         Top             =   495
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Observación Técnico"
         Height          =   255
         Index           =   5
         Left            =   -74400
         TabIndex        =   54
         Top             =   3000
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Observación Comercial"
         Height          =   255
         Index           =   3
         Left            =   -74400
         TabIndex        =   53
         Top             =   540
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Pago"
         Height          =   255
         Index           =   36
         Left            =   -72360
         TabIndex        =   50
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Forma Pago"
         Height          =   255
         Index           =   15
         Left            =   -68640
         TabIndex        =   49
         Top             =   855
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Contrato"
         Height          =   255
         Index           =   34
         Left            =   -68640
         TabIndex        =   48
         Top             =   495
         Width           =   975
      End
   End
   Begin MSAdodcLib.Adodc Data3 
      Height          =   330
      Left            =   7560
      Top             =   7080
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   10290
      TabIndex        =   42
      Top             =   6960
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label Label2 
      Caption         =   "A N U L A D O S"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   615
      Index           =   1
      Left            =   5280
      TabIndex        =   93
      Top             =   6960
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "Técnico"
      Height          =   255
      Index           =   35
      Left            =   3165
      TabIndex        =   57
      Top             =   6960
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu mnBuscar 
         Caption         =   "&Buscar"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnVerTodos 
         Caption         =   "&Ver Todos"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnBarra1 
         Caption         =   "-"
      End
      Begin VB.Menu mnNuevo 
         Caption         =   "&Nuevo"
         Enabled         =   0   'False
         HelpContextID   =   2
         Shortcut        =   ^N
         Visible         =   0   'False
      End
      Begin VB.Menu mnModificar 
         Caption         =   "&Modificar"
         Enabled         =   0   'False
         HelpContextID   =   2
         Shortcut        =   ^M
         Visible         =   0   'False
      End
      Begin VB.Menu mnEliminar 
         Caption         =   "&Eliminar"
         HelpContextID   =   2
         Shortcut        =   ^E
      End
      Begin VB.Menu mnBarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu mnMtoLineas 
      Caption         =   "&Mantenimiento Lineas"
      Begin VB.Menu mnRevisiones 
         Caption         =   "&Revisiones"
         HelpContextID   =   2
         Shortcut        =   ^R
         Visible         =   0   'False
      End
      Begin VB.Menu mnHistorico 
         Caption         =   "&Histórico"
         Shortcut        =   ^H
      End
   End
End
Attribute VB_Name = "frmManMantenimientosAnu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

Private WithEvents frmB As frmBuscaGrid 'Form para busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmF As frmCal 'Form Calendario Fecha
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmC As frmFacClientes 'Form Mto Clientes
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmTC As frmManTiposContrato 'Form Mto Tipos Contrato
Attribute frmTC.VB_VarHelpID = -1
Private WithEvents frmFP As frmFacFormasPago 'Form Mto Formas de Pago
Attribute frmFP.VB_VarHelpID = -1
'Private WithEvents frmT As frmAdmTrabajadores

'-----------------------------
'Se distinguen varios modos
'   0.-  Formulario limpio sin nungun campo rellenado
'   1.-  Preparando para hacer la busqueda
'   2.-  Ya tenemos registros y los vamos a recorrer
'        y podemos editarlos Edicion del campo
'   3.-  Insercion de nuevo registro
'   4.-  Modificar
'   5.-  Mantenimiento Lineas
'-------------------------------------------------------------------------
'-------------------------------------------------------------------------
Private Modo As Byte

Dim ModificaLineas As Byte
'1.- Añadir,  2.- Modificar,  3.- Borrar,  0.-Pasar control a Lineas

Dim NumTabMto As Byte
'Indica que numero de Tab que esta en modo Mantenimiento

Dim TituloLinea As String 'Descripcion de la linea que estamos en Mantenimiento

Dim PrimeraVez As Boolean

Dim EsCabecera As Boolean
'Para saber en MandaBusquedaPrevia si busca en la tabla scamanao en la tabla sdirec

'SQL de la tabla principal del formulario
Private CadenaConsulta As String

Private Ordenacion As String
Private NombreTabla As String  'Nombre de la tabla o de la
Private NomTablaLineas As String 'Nombre de la Tabla de lineas del Mantenimiento en que estemos
Private kCampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean

Dim btnAnyadir As Byte
'Variable que indica el número del Boton  Anyadir en la Toolbar1
Dim btnPrimero As Byte
'Variable que indica el número del Boton  PrimerRegistro en la Toolbar1


Private Sub cboTipoPago_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkBaterias_KeyDown(KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub chkBaterias_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Function ModificaMantenimiento() As Boolean
Dim sql As String

    On Error GoTo ErrMod
    
    If Me.Data1.Recordset.EOF Then Exit Function
    
    sql = "UPDATE scamana SET obserseg=" & DBSet(Text1(41).Text, "T", "S")
    sql = sql & " WHERE codclien=" & Data1.Recordset!codClien & " and nummante=" & Data1.Recordset!nummante
    sql = sql & " AND fechabaj=" & DBSet(Data1.Recordset!fechabaj, "F")
    conn.Execute sql
    ModificaMantenimiento = True
    Exit Function
    
ErrMod:
    MuestraError Err.Number, "Modificar mantenimiento.", Err.Description
End Function


Private Sub cmdAceptar_Click()
    Screen.MousePointer = vbHourglass
    On Error GoTo Error1

    Select Case Modo
        Case 1  'BUSQUEDA
            HacerBusqueda
            
        Case 3 'INSERTAR
            If DatosOk Then
              If InsertarDesdeForm(Me) Then
                    Colorines
                    PosicionarData
              End If
            End If
            
        Case 4  'MODIFICAR
            If DatosOk Then
'                If ModificaDesdeFormulario(Me, 1) Then
                If ModificaMantenimiento Then
                    Colorines
                    TerminaBloquear
                    PosicionarData
                End If
            End If
            
         Case 5 'InsertarModificar linea
            'Actualizar el registro en la tabla de lineas 'slima1' (Revisiones)
'            If ModificaLineas = 1 Then 'INSERTAR lineas
'                If InsertarLinea Then 'Revisiones
'                   ' If Me.SSTab1.Tab = 2 Then CargaGrid DataGrid1, Data2, True
'                    'BotonAnyadirLinea
'                End If
'            ElseIf ModificaLineas = 2 Then 'MODIFICAR lineas
'                If ModificarLinea Then
'                    TerminaBloquear
'                    PonerBotonCabecera True
'                    ModificaLineas = 0
'                    If Me.SSTab1.Tab = 2 Then 'Habilidades
'                        LLamaLineas 10
'                        CargaGrid2 DataGrid1, Data2
'                    End If
'                    PonerFocoBtn Me.cmdRegresar
'                End If
'            End If
    End Select
    Screen.MousePointer = vbDefault
 
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


'Private Sub cmdAux_Click()
'    'Abre Formulario de Mantenimiento de Trabajadores
'    Set frmT = New frmAdmTrabajadores
'    frmT.DatosADevolverBusqueda = "0|1|"
'    frmT.Show vbModal
'    Set frmT = Nothing
'    'PonerFoco Me.TxtAux1(1)
'End Sub


Private Sub cmdCancelar_Click()
    Select Case Modo
        Case 1, 3 'Busqueda, Insertar
            LimpiarCampos
            'Poner los grid sin apuntar a nada
            LimpiarDataGrids
            PonerModo 0
            PonerFoco Text1(0)
            
        Case 4  'Modificar
            TerminaBloquear
            PonerModo 2
            PonerCampos
            PonerFoco Text1(0)
            
        Case 5 'Lineas Detalle
'            TerminaBloquear
'            If Me.SSTab1.Tab = 2 Then 'Revisiones
'                If ModificaLineas = 1 Then 'INSERTAR
'                    ModificaLineas = 0
'                    DataGrid1.AllowAddNew = False
'                    If Not Data2.Recordset.EOF Then Data2.Recordset.MoveFirst
'                End If
'                ModificaLineas = 0
'                LLamaLineas 10
'                DataGrid1.Enabled = True
'            End If
'            PonerBotonCabecera True
    End Select
End Sub


Private Sub BotonAnyadir()
'Añadir registro en tabla de trabajadores: straba (Cabecera)

    LimpiarCampos 'Vacía los TextBox
    LimpiarCamposHistorico
    'Poner los grid sin apuntar a nada
    LimpiarDataGrids
    Colorines
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 3
    Text1(3).Text = Format(Now, "dd/mm/yyyy")
    PonerFoco Text1(0)
    cmbMes.ListIndex = 0
End Sub


Private Sub BotonAnyadirLinea()
Dim anc As Single
Dim I As Byte

    'Si no estaba modificando lineas salimos
    ' Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 2 Then Exit Sub
        
    If NumTabMto <> Me.SSTab1.Tab Then
        MsgBox "No puede Añadir. Esta en Modo Mantenimiento de otra linea.", vbExclamation
        Exit Sub
    End If
    
    ModificaLineas = 1 'Ponemos Modo Añadir Linea
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerBotonCabecera False
    lblIndicador.Caption = "INSERTAR"
    
'    If Me.SSTab1.Tab = 2 Then 'Revisiones
'        AnyadirLinea DataGrid1, Data2
'        For I = 0 To Me.TxtAux1.Count - 1
'            Me.TxtAux1(I).Text = ""
'        Next I
'        anc = ObtenerAlto(Me.DataGrid1) + 10
'        LLamaLineas anc
'        BloquearTxt TxtAux1(0), False
'        PonerFoco TxtAux1(0)
'    End If
End Sub


Private Sub BotonBuscar()
    'Buscar
    If Modo <> 1 Then
        LimpiarCampos
        'Poner los grid sin apuntar a nada
        LimpiarDataGrids
        PonerModo 1
        'Si pasamos el control aqui lo ponemos en amarillo
        PonerFoco Text1(0)
        Text1(0).BackColor = vbYellow
    Else
        HacerBusqueda
        If Data1.Recordset.EOF Then
            Text1(kCampo).Text = ""
            Text1(kCampo).BackColor = vbYellow
            PonerFoco Text1(kCampo)
        End If
    End If
End Sub


Private Sub BotonVerTodos()
'Ver todos
    LimpiarCampos
    LimpiarCamposHistorico
    LimpiarDataGrids
    
    If chkVistaPrevia.Value = 1 Then
        EsCabecera = True
        MandaBusquedaPrevia ""
    Else
        CadenaConsulta = "Select * from " & NombreTabla & " " & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Sub Desplazamiento(Index As Integer)
'Botones de Desplazamiento de la Toolbar
'Para desplazarse por los registros de control Data
    If Data1.Recordset.EOF Then Exit Sub
    Text2(21).Text = ""
    DesplazamientoData Data1, Index
    PonerCampos
End Sub


Private Sub BotonModificar()
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 4
'    PonerFoco Text1(3)
    
    BloquearTxt Text1(41), False
    Me.SSTab1.Tab = 3
    PonerFoco Text1(41)
End Sub


Private Sub BotonModificarLinea()
'Modificar una linea
'Dim vWhere As String
'Dim anc As Single
'
'    On Error GoTo EModificarLinea
'
'    'Si no estaba modificando lineas salimos
'    'Es decir, si estaba insertando linea no podemos hacer otra cosa
'    If ModificaLineas = 1 Then Exit Sub '1= Insertar
'
'    If NumTabMto <> Me.SSTab1.Tab Then
'        MsgBox "No puede Modificar. Esta en Modo Mantenimiento de otra linea.", vbExclamation
'        Exit Sub
'    End If
'
'    ModificaLineas = 2 'Modificar
'
'    If Me.SSTab1.Tab = 2 Then 'Revisiones
'         If Data2.Recordset.EOF Then Exit Sub
'          vWhere = ObtenerWhereCP(False) & " and fecharev='" & Format(Data2.Recordset!FechaRev, FormatoFecha) & "'"
'         If Not BloqueaRegistro(NomTablaLineas, vWhere) Then Exit Sub
'         anc = ObtenerAlto(Me.DataGrid1) + 10
'         'Llamamos al form
'         Me.TxtAux1(0).Text = DataGrid1.Columns(2).Text
'         Me.TxtAux1(1).Text = DataGrid1.Columns(3).Text
'         Me.TxtAux1(2).Text = DataGrid1.Columns(4).Text
'         LLamaLineas anc
'         DataGrid1.Enabled = False
'         BloquearTxt TxtAux1(0), True
'         PonerFoco TxtAux1(1)
'    End If
'
'
'    'Añadiremos el boton de aceptar y demas objetos para insertar
'    Me.lblIndicador.Caption = "MODIFICAR"
'    PonerBotonCabecera False
'
'EModificarLinea:
'    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub BotonEliminar()
'Eliminar Registro de la Cabecera: Tabla de Mantenimientos (scaman)
' y los registros correspondientes de las tablas de lineas (sliman y slima1)
Dim cad As String
    
    On Error GoTo EEliminar

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
    
    MsgBox "De momento no podemos eliminar de anulados", vbExclamation
    
'    cad = "Cabecera de Mantenimientos." & vbCrLf
'    cad = cad & "-----------------------------------" & vbCrLf & vbCrLf
'    cad = cad & "Va a eliminar el Mantenimiento:            "
'    cad = cad & vbCrLf & "Cliente:  " & Format(Text1(0).Text, "000000") & " - " & Text2(0).Text
''    cad = cad & vbCrLf & "Direc.:  " & Format(Text1(1).Text, "000") & " - " & Text2(1).Text
'    cad = cad & vbCrLf & "Nº Mante.:  " & Text1(2).Text
'    cad = cad & vbCrLf & vbCrLf & " ¿Desea Eliminarlo? "
'
'    'Borramos
'    If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then
'        'Hay que eliminar
'        On Error GoTo EEliminar
'        Screen.MousePointer = vbHourglass
'        NumRegElim = Data1.Recordset.AbsolutePosition
'        If Not Eliminar Then
'            Exit Sub
'        ElseIf SituarDataTrasEliminar(Data1, NumRegElim) Then
'            PonerCampos
'        Else
'            LimpiarCampos
'            'Poner los grid sin apuntar a nada
'            LimpiarDataGrids
'            PonerModo 0
'        End If
'    End If
'    Screen.MousePointer = vbDefault
    
EEliminar:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Mantenimiento", Err.Description
End Sub


Private Sub BotonEliminarLinea()
'Eliminar una linea De Mantenimiento. Tabla: slima1
'Dim SQL As String
'Dim FechaRev As Date
'
'    On Error GoTo EEliminarLinea
'
'    'Si no estaba modificando lineas salimos
'    'Es decir, si estaba insertando linea no podemos hacer otra cosa
'    If ModificaLineas = 1 Or ModificaLineas = 2 Then Exit Sub '1= Insertar, 2=Modificar
'
'     If NumTabMto <> Me.SSTab1.Tab Then
'        MsgBox "No puede eliminar. Esta en Modo Mantenimiento de otra linea.", vbExclamation
'        Exit Sub
'    End If
'
''    If Me.SSTab1.Tab = 2 Then 'Revisiones
''        If Data2.Recordset.EOF Then Exit Sub
''        FechaRev = Data2.Recordset!FechaRev
''    End If
'
'    ModificaLineas = 3 'Eliminar
'    SQL = "¿Seguro que desea eliminar la línea de " & TituloLinea & "?      " & vbCrLf
'    SQL = SQL & vbCrLf & "Fec. Rev.: " & FechaRev
'    SQL = SQL & vbCrLf & " Técnico: " & Format(Data2.Recordset!CodTraba, "0000") & " - " & Text2(21).Text
'
'    If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then
'        'Hay que eliminar
'        SQL = "Delete from " & NomTablaLineas & ObtenerWhereCP(True)
'        SQL = SQL & " and fecharev='" & Format(FechaRev, FormatoFecha) & "'"
'        Conn.Execute SQL
'        ModificaLineas = 0
''        If Me.SSTab1.Tab = 2 Then CargaGrid2 DataGrid1, Data2 'Revisiones
''        CancelaADODC
'    End If
'    PonerFocoBtn Me.cmdRegresar
'
'EEliminarLinea:
'    Screen.MousePointer = vbDefault
'    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Lineas Mantenimientos", Err.Description
End Sub


Private Sub cmdRegresar_Click()
'Este es el boton Cabecera
Dim cad As String

    'Quitar lineas y volver a la cabecera
    If Modo = 5 Then  'modo 5: Mantenimientos Lineas
        PonerModo 2
        Me.lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    Else 'Se llama desde algún Prismatico de otro Form al Mantenimiento de Trabajadores
        If Data1.Recordset.EOF Then
            MsgBox "Ningún registro devuelto.", vbExclamation
            Exit Sub
        End If
        cad = Data1.Recordset.Fields(0) & "|"
        cad = cad & Data1.Recordset.Fields(1) & "|"
        RaiseEvent DatoSeleccionado(cad)
        Unload Me
    End If
End Sub


Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error GoTo Error1
'
'    If Not Data2.Recordset.EOF And ModificaLineas <> 1 Then 'No en Insertar
'        'Poner descripcion del Trabajador
'        Text2(21).Text = DevuelveDesdeBDNew(conAri, "straba", "nomtraba", "codtraba", Data2.Recordset!CodTraba.Value, "N")
'    Else
'        Text2(21).Text = ""
'    End If
'
Error1:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub Form_Load()
    'Icono del formulario
    Me.Icon = frmPpal.Icon
    
    'Icono de busqueda
    For kCampo = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(kCampo).Picture = frmPpal.imgListComun.ListImages(19).Picture
    Next kCampo

    ' ICONITOS DE LA BARRA
    btnAnyadir = 5
    btnPrimero = 17
    With Me.Toolbar1
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 1   'Botón Buscar
        .Buttons(2).Image = 2   'Botón Todos
        .Buttons(5).Image = 3   'Insertar Nuevo
        .Buttons(6).Image = 4   'Modificar
        .Buttons(7).Image = 5   'Borrar
        .Buttons(10).Image = 10 'Mto Lineas Revisiones
        .Buttons(11).Image = 38 'Mto Lineas Histórico
        .Buttons(12).Image = 34 'Componentes
        .Buttons(14).Image = 15  'Salir
        .Buttons(btnPrimero).Image = 6  'Primero
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
        .Buttons(btnPrimero + 3).Image = 9 'Último
    End With
    Me.SSTab1.Tab = 0
          
    CargaComboMes
    LimpiarCampos   'Limpia los campos TextBox
    CargarComboTipoPago
   
    'Comprobar si es Departamento o Direccion
    Me.Label1(1).Caption = "Cod. " & DevuelveTextoDepto(True)
    
        
    '## A mano
    NombreTabla = "scamana"
    Ordenacion = " ORDER BY codclien, nummante "
         
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = conn
    Data1.RecordSource = "Select * from " & NombreTabla & " where codclien=-1"
    Data1.Refresh
    If DatosADevolverBusqueda = "" Then
        PonerModo 0
    Else
        PonerModo 1
        Text1(0).BackColor = vbYellow
    End If
    
    'Cargar el DataGrid de lineas de Revisiones inicialmente a nada DATA2
'    PrimeraVez = True
'    NomTablaLineas = "slima1" 'Tabla lineas de Revisiones de MAntenimientos
'    Data2.ConnectionString = Conn
'    Data2.RecordSource = "Select * from " & NomTablaLineas & " where codclien=-1"
'    Data2.Refresh
''    CargaGrid DataGrid1, Data2, False
    
    'Cargamos inicialmente el DATA3 a nada
    Data3.ConnectionString = conn
    Data3.RecordSource = "select * from slimana where codclien=-1"
    Data3.Refresh
End Sub


Private Sub LimpiarCampos()
On Error Resume Next

    limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador.Caption = ""
    'Aqui va el especifico de cada form es
    '### a mano
    Me.cboTipoPago.ListIndex = -1
    Me.chkBaterias.Value = False
    Me.cmbMes.ListIndex = -1
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
End Sub


Private Sub frmB_Selecionado(CadenaDevuelta As String)
Dim cadB As String
Dim Aux As String
      
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        If EsCabecera Then 'Llama desde VerTodos del Form
            cadB = ""
            Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
'            Aux = Aux & " and " & ValorDevueltoFormGrid(Text1(1), CadenaDevuelta, 2)
            Aux = Aux & " and " & ValorDevueltoFormGrid(Text1(2), CadenaDevuelta, 3)
            cadB = Aux
            CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
            PonerCadenaBusqueda
        Else 'Llama desde Prismatico Direcciones/Departamentos
            Text1(1).Text = Format(RecuperaValor(CadenaDevuelta, 1), "000")
            Text2(1).Text = RecuperaValor(CadenaDevuelta, 2)
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub frmC_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Clientes
    Text1(0).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000") 'Cod Clien
    Text2(0).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Clien
End Sub


Private Sub frmF_Selec(vFecha As Date) 'Calendario Fechas
    Text1(3).Text = Format(vFecha, "dd/mm/yyyy")
End Sub


Private Sub frmFP_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Formas de Pago
    Text1(5).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000") 'Cod Forma Pago
    Text2(5).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Forma Pago
End Sub


'Private Sub frmT_DatoSeleccionado(CadenaSeleccion As String)
''Form Mantenimiento de Trabajadores
'    Me.TxtAux1(1).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod trabajador
'    FormateaCampo Me.TxtAux1(1)
'    Text2(21).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom trabajador
'End Sub

Private Sub frmTC_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Tipos de Contrato
    Text1(4).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod Tipo Contrato
    Text2(4).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Tipo Contrato
End Sub


Private Sub imgBuscar_Click(Index As Integer)
Dim indice As Byte

    If Modo = 2 Or Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass

    Select Case Index
        Case 0 'Cod. Cliente
            Set frmC = New frmFacClientes
            frmC.DatosADevolverBusqueda = "0"
            frmC.Show vbModal
            Set frmC = Nothing
            indice = Index
            
        Case 1 'Cod. Direc.
             'Mostrar las Direc. o Dptos del cliente seleccionado
             If Trim(Text1(0).Text) = "" Then
                MsgBox "Debe seleccionar un cliente.", vbInformation
                Screen.MousePointer = vbDefault
                Exit Sub
             Else
                EsCabecera = False
                MandaBusquedaPrevia " codclien= " & Val(Text1(0).Text)
                indice = 1
             End If
             
        Case 2 'Cod. Tipo Contrato
            Set frmTC = New frmManTiposContrato
            frmTC.DatosADevolverBusqueda = "0"
            frmTC.Show vbModal
            Set frmTC = Nothing
            indice = 4
            
        Case 3 'Cod. Formas de Pago
            Set frmFP = New frmFacFormasPago
            frmFP.DatosADevolverBusqueda = "0"
            frmFP.Show vbModal
            Set frmFP = Nothing
            indice = 5
    End Select
    PonerFoco Text1(indice)
    Screen.MousePointer = vbDefault
End Sub


Private Sub imgFecha_Click(Index As Integer) 'Abre calendario Fechas
Dim indice As Byte

   If Modo = 2 Or Modo = 0 Then Exit Sub
   Screen.MousePointer = vbHourglass
   Set frmF = New frmCal
   frmF.Fecha = Now
   indice = Index + 3
   
   PonerFormatoFecha Text1(indice)
   If Text1(indice).Text <> "" Then frmF.Fecha = CDate(Text1(indice).Text)

   Screen.MousePointer = vbDefault
   frmF.Show vbModal
   Set frmF = Nothing
   PonerFoco Text1(indice)
End Sub


Private Sub imgFlecha_Click(Index As Integer)
'Desplazarse por los dos registros siguientes del histórico
On Error GoTo EDesplazamiento
    
    Select Case Index
        Case 0 'Anterior
            If Data3.Recordset.BOF Then Exit Sub
            If (Data3.Recordset.AbsolutePosition Mod 2) = 0 Then 'numero par
                Data3.Recordset.Move -3
            Else 'numero impar
                Data3.Recordset.Move -2
            End If
            If Data3.Recordset.BOF Then Exit Sub
        Case 1 'Siguiente
            If Data3.Recordset.EOF Then Exit Sub
            If Not Data3.Recordset.BOF Then
                Data3.Recordset.MoveNext
            Else
                Data3.Recordset.Move 3
            End If
    End Select
    PonerCamposHistorico
    
EDesplazamiento:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub mnBuscar_Click()
    Me.SSTab1.Tab = 0
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    If Modo = 5 Then 'Eliminar lineas de Mantenimiento
         BotonEliminarLinea
    Else   'Eliminar Mantenimiento
         BotonEliminar
    End If
End Sub


Private Sub mnHistorico_Click()
Dim b As Boolean
    
    'Cargar el data3 con los datos de la tabla "sliman"
    NomTablaLineas = "slimana"
    Me.SSTab1.Tab = 2
    'ASignamos un SQL al DATA3
'    Data3.ConnectionString = Conn
    Data3.RecordSource = "Select anomante,mes01man,mes02man,mes03man,mes04man,mes05man,mes06man,mes07man,mes08man,mes09man,mes10man,mes11man,mes12man from " & NomTablaLineas & ObtenerWhereCP(True)
    Data3.CursorType = adOpenStatic
    Data3.Refresh
    If Data3.Recordset.EOF Then
        MsgBox "No existen datos en el Histórico para ese cliente y Direc./Dpto.", vbInformation
        Exit Sub
    Else
        b = Data3.Recordset.RecordCount > 2
        Me.imgFlecha(0).visible = b
        Me.imgFlecha(1).visible = b
    End If
    PonerCamposHistorico
End Sub


Private Sub mnModificar_Click()
    If Modo = 5 Then 'Modificar lineas
         BotonModificarLinea
    Else   'Modificar Mantenimiento
         If SSTab1.Tab <> 0 And Me.SSTab1.Tab <> 1 Then Me.SSTab1.Tab = 0
         If BLOQUEADesdeFormulario(Me) Then BotonModificar
    End If
End Sub


Private Sub mnNuevo_Click()
    If Modo = 5 Then 'Añadir lineas
         BotonAnyadirLinea
    Else 'Añadir Mantenimiento
         Me.SSTab1.Tab = 0
         BotonAnyadir
    End If
End Sub

Private Sub mnRevisiones_Click()
'Abre Mantenimiento de lineas  de Revisiones
    BotonMtoLineas 2, "Revisiones"
    NomTablaLineas = "slima1"
End Sub

Private Sub mnSalir_Click()
    Screen.MousePointer = vbDefault
    If (Modo = 5) Then 'Modo 5: Mto Lineas
        '1:Insertar linea, 2: Modificar
        If ModificaLineas = 1 Or ModificaLineas = 2 Then cmdCancelar_Click
        cmdRegresar_Click
        Exit Sub
    End If
    Unload Me
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub


Private Sub SSTab1_Click(PreviousTab As Integer)
Dim b As Boolean
    
    b = (Me.SSTab1.Tab = 2)
    'Poner Visible el Nombre del Técnico si estamos en Mantenimiento Lineas
    'Me.Text2(21).visible = b
    'Me.Label1(35).visible = b
End Sub


'### A mano
'Los metodos del text tendran que estar
'Los descomentamos cuando esten puestos ya los controles
Private Sub Text1_GotFocus(Index As Integer)
    kCampo = Index
    If Not Text1(Index).MultiLine Then
        ConseguirFoco Text1(Index), Modo
    Else
        SituarMultiTextFinal Text1(Index)
    End If
End Sub


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    If Not Text1(Index).MultiLine Then KEYdown KeyCode
End Sub


Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If Not Text1(Index).MultiLine Then KEYpress KeyAscii
End Sub

'----------------------------------------------------------------
'----------------------------------------------------------------
' Cunado el campo de texto pierde el enfoque
' Es especifico de cada formulario y en el podremos controlar
' lo que queramos, desde formatear un campo si asi lo deseamos
' hasta pedir que nos devuelva los datos de la empresa
'----------------------------------------------------------------
'----------------------------------------------------------------
Private Sub Text1_LostFocus(Index As Integer)
Dim devuelve As String
Dim cadMen As String
Dim cadDpto As String
        
    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
    
    'Si queremos hacer algo ..
    Select Case Index
        Case 0 'Cod. Clien
            devuelve = PonerNombreDeCod(Text1(Index), conAri, "sclien", "nomclien")
            Text2(Index).Text = devuelve

        Case 1 'Cod. Direc
            If Text1(Index).Text = "" Then
                Text2(Index).Text = ""
            Else
                'Comprobar que el cliente seleccionada tiene esa direccion
                If Text1(0).Text <> "" Then 'HAY cliente
                    devuelve = DevuelveDesdeBDNew(conAri, "sdirec", "coddirec", "codclien", Text1(0).Text, "N", , "coddirec", Text1(1).Text, "N")
                    If devuelve = "" Then 'NO EXISTE el dpto para el cliente
                        Text2(Index).Text = ""
                        cadDpto = DevuelveTextoDepto(False)
                        cadMen = "No existe" & cadDpto & Text1(1).Text & " para el cliente: " & vbCrLf
                        cadMen = cadMen & vbCrLf & Text1(0).Text & " - " & Text2(0).Text
                        MsgBox cadMen, vbInformation
                        
                        Text1(Index).Text = ""
                        PonerFoco Text1(Index)
                    Else 'EXISTE el dpto
                        '---- Laura: modificado 27/09/2006
                        'devuelve = PonerNombreDeCod(Text1(Index), conAri, "sdirec", "nomdirec")
                        devuelve = DevuelveDesdeBDNew(conAri, "sdirec", "nomdirec", "codclien", Text1(0).Text, "N", , "coddirec", Text1(1).Text, "N")
                        Text2(Index).Text = devuelve
                        FormateaCampo Text1(Index)
                    End If
                Else
                    Text1(Index).Text = ""
                    cadDpto = DevuelveTextoDepto(False)
                    MsgBox "Debe introducir un Cliente, antes de seleccionar su " & cadDpto, vbInformation
                End If
            End If
            
        Case 2 'Nº Mantenimiento
            'Comprobar si ya existe un registro con esa clave Primaria si Insertando
            If Modo = 3 And Text1(0).Text <> "" And Text1(2).Text <> "" Then
                devuelve = "select count(*) from scamana" & ObtenerWhereCP(True)
                If RegistrosAListar(devuelve) > 0 Then
                    MsgBox "Ya existe el registro en la Tabla de Mantenimientos.", vbExclamation
                    PonerFoco Text1(2)
                End If
            End If

        Case 3 'Fecha Inicio
            'Si no es modo de Busqueda poner el formato
            If Text1(Index).Text <> "" Then PonerFormatoFecha Text1(Index)
        Case 4 'Tipo Contrato
            Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "stipco", "nomtipco")
        Case 5 'Forma de Pago
            Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "sforpa", "nomforpa")
        Case 6, 7 'Anticipado
            'Formato 3: Decimal(10,2)
            If Text1(Index).Text <> "" Then PonerFormatoDecimal Text1(Index), 3
        Case 38
            Text2(6).Text = PonerNombreDeCod(Text1(Index), conAri, "sincid", "nomincid")
        Case 8 To 31 'Meses
            If Text1(Index).Text <> "" Then
                'Formato 3: Decimal(10,2)
                PonerFormatoDecimal Text1(Index), 3
                If EsNumerico(Text1(Index).Text) Then
                    SumarTotal
                Else
                    PonerFoco Text1(Index)
                End If
            End If
        Case 33
            PonerFocoBtn Me.cmdAceptar
    End Select
End Sub


Private Sub HacerBusqueda()
Dim cadB As String

    cadB = ObtenerBusqueda(Me, False)
    If chkVistaPrevia = 1 Then
        EsCabecera = True
        MandaBusquedaPrevia cadB
    ElseIf cadB <> "" Then
        'Se muestran en el mismo form
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Sub MandaBusquedaPrevia(cadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim cad As String
Dim Tabla As String
Dim Titulo As String
Dim Desc As String, devuelve As String
    'Llamamos a al form
    '##A mano
    cad = ""
    If EsCabecera Then
        cad = cad & ParaGrid(Text1(0), 10, "Cliente")
        cad = cad & "Desc. Cliente|sclien|nomclien|T||36·"
        Desc = DevuelveTextoDepto(True)
        cad = cad & ParaGrid(Text1(1), 7, Desc)
        cad = cad & "Desc. " & Desc & "|sdirec|nomdirec|T||35·"
        cad = cad & ParaGrid(Text1(2), 13, "Nº Mant.")
        
        Tabla = "(" & NombreTabla & " LEFT JOIN sclien ON " & NombreTabla & ".codclien=sclien.codclien" & ")"
        Tabla = Tabla & " LEFT JOIN sdirec ON " & NombreTabla & ".codclien=sdirec.codclien and "
        Tabla = Tabla & NombreTabla & ".coddirec=sdirec.coddirec"
        Titulo = "Mantenimientos"
        devuelve = "0|2|4|"
    Else
        If vParamAplic.HayDeparNuevo = 1 Then
            Titulo = "Departamentos Cliente: "
            Desc = "Dpto."
        ElseIf vParamAplic.HayDeparNuevo = 0 Then
            Titulo = "Direcciones Cliente: "
            Desc = "Direc."
        Else
            Titulo = "Obras Cliente: "
            Desc = "Obra"
        End If
        Titulo = Titulo & Text1(0).Text & " - " & Text2(0).Text
        cad = cad & "Cod. " & Desc & "|sdirec|coddirec|N|000|15·"
        cad = cad & "Desc. " & Desc & "|sdirec|nomdirec|T||60·"
        Tabla = "sdirec"
        devuelve = "0|1|"
    End If
           
    If cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = cad
        frmB.vTabla = Tabla
        frmB.vSQL = cadB
        HaDevueltoDatos = False
        '###A mano
'        frmB.vDevuelve = "0|1|"
        frmB.vDevuelve = devuelve
        frmB.vTitulo = Titulo
        frmB.vselElem = 1
        frmB.vConexionGrid = conAri 'Conexión a BD: Ariges
        If Not EsCabecera Then frmB.Label1.FontSize = 11
'        frmB.vBuscaPrevia = chkVistaPrevia
        '#
        frmB.Show vbModal
        Set frmB = Nothing
        'Si ha puesto valores y tenemos que es formulario de busqueda entonces
        'tendremos que cerrar el form lanzando el evento
'        If HaDevueltoDatos Then
'''            If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
'''                cmdRegresar_Click
'        Else   'de ha devuelto datos, es decir NO ha devuelto datos
'            PonerFoco Text1(kCampo)
        'End If
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerCadenaBusqueda()
    Screen.MousePointer = vbHourglass
    On Error GoTo EEPonerBusq

    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        MsgBox "No hay ningún registro en la tabla " & NombreTabla, vbInformation
        Screen.MousePointer = vbDefault
        PonerFoco Text1(0)
        Exit Sub
    Else
        Data1.Recordset.MoveFirst
        PonerModo 2
        PonerCampos
    End If

    Screen.MousePointer = vbDefault
    Exit Sub
    
EEPonerBusq:
    MuestraError Err.Number, "PonerCadenaBusqueda"
    PonerModo 0
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerCamposLineas()
'Carga las Pestañas con las tablas de lineas del Trabajador seleccionado para mostrar
    On Error GoTo EPonerLineas

    Screen.MousePointer = vbHourglass
   
    'Revisiones - Datos de la tabla slima1
    'CargaGrid DataGrid1, Data2, True
    
    PrimeraVez = False
    Screen.MousePointer = vbDefault
    Exit Sub
    
EPonerLineas:
    MuestraError Err.Number, "PonerCamposLineas"
    PonerModo 0
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerCampos()
Dim cCli As CCliente
Dim nomDirDpto As String
    On Error Resume Next

    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma Me, Data1
    
    'Colorines de los meses facturados
    Colorines
    
    'Desc. Tipo Contrato
    Text2(4).Text = PonerNombreDeCod(Text1(4), conAri, "stipco", "nomtipco")
    'Desc. Forma PAgo
    Text2(5).Text = PonerNombreDeCod(Text1(5), conAri, "sforpa", "nomforpa")
    'Cod. Cliente
    Text2(0).Text = PonerNombreDeCod(Text1(0), conAri, "sclien", "nomclien")
    'Motivo anulacion
    Modo = 3
    Text1_LostFocus 38
    Modo = 2
    
    
    '---- Laura: modificado 28/09/2006
    'Cod Dir./Dpto
    'Text2(1).Text = DevuelveDesdeBDNew(Text1(1), conAri, "sdirec", "nomdirec")
    Text2(1).Text = ""
    If Text1(1).Text <> "" Then
        Set cCli = New CCliente
        cCli.Codigo = Text1(0).Text
        If cCli.DptoCliente(Text1(1).Text, nomDirDpto) Then Text2(1).Text = nomDirDpto
        Set cCli = Nothing
    End If
    '----

       
    PonerCamposLineas 'Pone los datos de las tablas de lineas asociadas al trabajador
    LimpiarCamposHistorico
    SumarTotal
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
    If Err.Number <> 0 Then Err.Clear
End Sub

'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
Private Sub PonerModo(Kmodo As Byte)
Dim I As Byte, NumReg As Byte
Dim b As Boolean
    
    On Error GoTo EPonerModo

    'Actualiza Iconos Insertar,Modificar,Eliminar
    ActualizarToolbarGnral Me.Toolbar1, Modo, Kmodo, btnAnyadir
    
    Modo = Kmodo
    PonerIndicador lblIndicador, Modo
    
    'Modo 2. Hay datos y estamos visualizandolos
    '=========================================
    b = (Modo = 2)

    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.visible = b
    Else
        cmdRegresar.visible = False
    End If
        
    'Poner Flechas de desplazamiento visibles
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, NumReg
        
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1
    If Modo = 4 Then
        BloquearTxt Text1(41), False
    Else
        BloquearText1 Me, Modo
    End If
  
              
    '---------------------------------------------
    b = Modo <> 0 And Modo <> 2 And Modo <> 5
    cmdCancelar.visible = b
    cmdAceptar.visible = b
    
    
    b = Modo <> 0 And Modo <> 2 And Modo <> 5 And Modo <> 4
    Me.cmbMes.Enabled = b
    
    For I = 0 To Me.imgFecha.Count - 1
        Me.imgFecha(I).Enabled = b
    Next I
    
'    For i = 0 To Me.imgBuscar.Count - 1
'        Me.imgBuscar(i).Enabled = b
'    Next i
   For I = 0 To Me.imgBuscar.Count - 1
        BloquearImg Me.imgBuscar(I), Not b
    Next I
    
    
    If Modo = 4 Then 'Modificar. Bloquear clave Primaria
        Me.imgBuscar(0).Enabled = False
'        Me.imgBuscar(1).Enabled = False
    End If
    
    Me.chkVistaPrevia.visible = (Modo <> 5)
    Me.cboTipoPago.Enabled = (Modo = 3) 'Or (Modo = 4)
    Me.chkBaterias.Enabled = (Modo = 3) 'Or (Modo = 4)
       
    PonerModoOpcionesMenu (Modo) 'Activar opciones de menu según modo
    PonerOpcionesMenu   'Activar opciones de menu según nivel
                        'de permisos del usuario
                        
    Exit Sub
    
EPonerModo:
    MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Function DatosOk() As Boolean
Dim b As Boolean
    
    On Error GoTo EDatosOK

    DatosOk = False
    b = True
    b = CompForm(Me, 1) 'Comprobar formato datos ok
    If Not b Then Exit Function
    
    DatosOk = b
    
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


'Private Function DatosOkLinea() As Boolean
'Dim b As Boolean
'
'    On Error GoTo EDatosOkLinea
'
'    DatosOkLinea = False
'    b = True
'
'    If Me.SSTab1.Tab = 2 Then 'Fecha Revision
'        If Trim(TxtAux1(0).Text) = "" Then
'            MsgBox "El campo Fecha Revisión no puede ser nulo", vbExclamation
'            b = False
'        End If
'
'        If Trim(TxtAux1(1).Text) = "" Then 'Tecnico
'            MsgBox "El campo Cod. Técnico no puede ser nulo", vbExclamation
'            b = False
'        End If
'    End If
'
'    DatosOkLinea = b
'
'EDatosOkLinea:
'    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
'End Function


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1  'Buscar
            mnBuscar_Click
        Case 2  'Todos
            BotonVerTodos
        Case 5  'Nuevo
            mnNuevo_Click
        Case 6  'Modificar
            mnModificar_Click
        Case 7  'Borrar
            mnEliminar_Click
        Case 10  'Lineas Revisiones
            mnRevisiones_Click
        Case 11 'Líneas Histórico
            mnHistorico_Click
        Case 12 'Componentes
            BotonComponentes
        Case 14    'Salir
            mnSalir_Click
        Case btnPrimero To btnPrimero + 3 'Flechas Desplazamiento
            Desplazamiento (Button.Index - btnPrimero)
    End Select
End Sub


Private Sub PonerOpcionesMenu()
Dim J As Byte

    On Error Resume Next
    
    PonerOpcionesMenuGeneral Me
    
    J = Val(Me.mnRevisiones.HelpContextID)
    If J < vUsu.Nivel Then Me.mnRevisiones.Enabled = False
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub

    
'Private Function InsertarLinea() As Boolean
''Inserta un registro en la tabla de Revisiones: slima1
'Dim SQL As String
'
'    On Error GoTo EInsertarLinea
'
'    InsertarLinea = False
'    SQL = ""
'    If DatosOkLinea And Me.SSTab1.Tab = 2 Then 'Revisiones
'        SQL = "INSERT INTO slima1 "
'        SQL = SQL & "(codclien, nummante, fecharev, codtraba, observac) "
'        SQL = SQL & "VALUES (" & Val(Text1(0).Text) & ", "
'        SQL = SQL & DBSet(Text1(2).Text, "T") & ", " & DBSet(TxtAux1(0).Text, "F") & ", " & TxtAux1(1).Text & ", "
'        SQL = SQL & QuitarCaracterEnter(DBSet(TxtAux1(2).Text, "T")) & ")"
'     End If
'
'    If SQL <> "" Then
'        Conn.Execute SQL
'        InsertarLinea = True
'    End If
'    Exit Function
'
'EInsertarLinea:
'    MuestraError Err.Number, "Insertar Lineas Mantenimiento" & vbCrLf & Err.Description
'End Function


'Private Function ModificarLinea() As Boolean
''Modifica un registro en la tabla de Revisiones: slima1
'Dim SQL As String
'
'    On Error GoTo EModificarLinea
'
'    ModificarLinea = False
'    SQL = ""
'    If DatosOkLinea And Me.SSTab1.Tab = 2 Then 'Habilidades
'        SQL = "UPDATE slima1 Set codtraba = " & TxtAux1(1).Text & ", observac='" & QuitarCaracterEnter(TxtAux1(2).Text) & "'"
'        SQL = SQL & ObtenerWhereCP(True) & " AND fecharev='" & Format(Data2.Recordset!FechaRev, FormatoFecha) & "'"
'    End If
'
'    If SQL <> "" Then
'        Conn.Execute SQL
'        ModificarLinea = True
'    End If
'    Exit Function
'
'EModificarLinea:
'    MuestraError Err.Number, "Modificar Lineas Trabajador" & vbCrLf & Err.Description
'End Function


Private Sub PonerBotonCabecera(b As Boolean)
'Pone el boton de Regresar a la Cabecera si pasamos a MAntenimiento de Lineas
'o Pone los botones de Aceptar y cancelar en Insert,update o delete lineas
On Error Resume Next
    Me.cmdAceptar.visible = Not b
    Me.cmdCancelar.visible = Not b
    Me.cmdRegresar.visible = b
    Me.cmdRegresar.Caption = "Cabecera"
    If b Then
        Me.lblIndicador.Caption = "Líneas " & TituloLinea
        PonerFocoBtn Me.cmdRegresar
    End If
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub CargaGrid(ByRef vDataGrid As DataGrid, ByRef vData As Adodc, enlaza As Boolean)
Dim b As Boolean
Dim sql As String

    On Error GoTo ECargaGrid

    b = vDataGrid.Enabled
    
    sql = MontaSQLCarga(enlaza)
    CargaGridGnral vDataGrid, vData, sql, PrimeraVez
    vDataGrid.RowHeight = 470
    CargaGrid2 vDataGrid, vData
   
        
    vDataGrid.ScrollBars = dbgAutomatic
    vDataGrid.Enabled = b
    
    Exit Sub
    
ECargaGrid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub


Private Sub CargaGrid2(ByRef vDataGrid As DataGrid, ByRef vData As Adodc)
Dim tots As String

    On Error GoTo ECargaGrid
  
    vData.Refresh
    
    'codclien,nummante
    tots = "N||||0|;N||||0|;"
    'fecharev,codtraba,observa
    tots = tots & "S|TxtAux1(0)|T|Fecha Rev.|1100|;S|TxtAux1(1)|T|Técnico|900|;S|cmdAux|B||0|;S|TxtAux1(2)|T|Observaciones|8100|;"
    arregla tots, vDataGrid, Me
    
     vDataGrid.Columns(3).NumberFormat = "0000"
     vDataGrid.Columns(4).WrapText = True
     
     vDataGrid.RowHeight = 470

     vDataGrid.Enabled = (Modo = 0) Or (Modo = 2) Or (Modo = 5 And ModificaLineas = 0)
     Exit Sub
     
ECargaGrid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub


'Private Sub LLamaLineas(alto As Single)
'Dim jj As Byte
'Dim b As Boolean
'
'    DeseleccionaGrid Me.DataGrid1
'
'    b = (Modo = 5) And (ModificaLineas = 1 Or ModificaLineas = 2) 'Insertar o Modificar Lineas
'
'    For jj = 0 To Me.TxtAux1.Count - 1
'        Me.TxtAux1(jj).Height = DataGrid1.RowHeight
'        Me.TxtAux1(jj).Top = alto
'        Me.TxtAux1(jj).visible = b
'    Next jj
'
'    Me.cmdAux.Height = DataGrid1.RowHeight
'    Me.cmdAux.Top = alto
'    Me.cmdAux.visible = b
'End Sub





Private Sub BotonMtoLineas(numTab As Integer, cad As String)
        Me.SSTab1.Tab = numTab
        NumTabMto = numTab
        TituloLinea = cad
        PonerModo 5
        PonerBotonCabecera True
End Sub


Private Function Eliminar() As Boolean
Dim sql As String

    On Error GoTo FinEliminar

    conn.BeginTrans
    sql = " WHERE  codclien=" & Data1.Recordset!codClien
'        SQL = SQL & " AND coddirec=" & Data1.Recordset!CodDirec
    sql = sql & " AND nummante='" & Data1.Recordset!nummante & "'"

    'Lineas Mantenimiento (Histórico)
    conn.Execute "Delete from sliman " & sql
    'Lineas Revisiones
    conn.Execute "Delete from slima1 " & sql
    
    'Cabecera
    conn.Execute "Delete from scamana" & sql

FinEliminar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar"
        conn.RollbackTrans
        Eliminar = False
    Else
        conn.CommitTrans
        Eliminar = True
    End If
End Function


Private Sub LimpiarDataGrids()
'Pone los Grids sin datos, apuntando a ningún registro
On Error Resume Next
    'CargaGrid DataGrid1, Data2, False
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub CargarComboTipoPago()
'### Combo Tipo Pago
'Cargaremos el combo, o bien desde una tabla o con valores fijos o como
'se quiera, la cuestion es cargarlo
' El estilo del combo debe de ser 2 - Dropdown List
' Si queremos que este ordenado, o lo ordenamos por la sentencia sql
' o marcamos la opcion sorted del combo
'0-Mensual, 1-Trimestral, 2-Semestral, 3-Anual

    cboTipoPago.Clear
    cboTipoPago.AddItem "Mensual"
    cboTipoPago.ItemData(cboTipoPago.NewIndex) = 0
    
    cboTipoPago.AddItem "Trimestral"
    cboTipoPago.ItemData(cboTipoPago.NewIndex) = 1
    
    cboTipoPago.AddItem "Semestral"
    cboTipoPago.ItemData(cboTipoPago.NewIndex) = 2
    
    cboTipoPago.AddItem "Anual"
    cboTipoPago.ItemData(cboTipoPago.NewIndex) = 3
    
End Sub


Private Sub PosicionarData()
Dim Indicador As String
Dim vWhere As String

    vWhere = "(" & ObtenerWhereCP(False) & ")"
    If SituarDataMULTI(Data1, vWhere, Indicador) Then
        PonerModo 2
        lblIndicador.Caption = Indicador
   Else
'        LimpiarCampos
        'Poner los grid sin apuntar a nada
'        LimpiarDataGrids
        PonerModo 0
    End If
End Sub


Private Function ObtenerWhereCP(conW As Boolean) As String
Dim sql As String

    sql = ""
    If conW Then sql = " WHERE "
    sql = sql & " codclien=" & Text1(0).Text
    sql = sql & " AND nummante=" & DBSet(Text1(2).Text, "T")
    ObtenerWhereCP = sql
End Function


Private Function MontaSQLCarga(enlaza As Boolean) As String
'--------------------------------------------------------------------
' MontaSQlCarga:
'   Basándose en la información proporcionada por el vector de campos
'   crea un SQl para ejecutar una consulta sobre la base de datos que los
'   devuelva.
' Si ENLAZA -> Enlaza con el data1
'           -> Si no lo cargamos sin enlazar a ningun campo
'--------------------------------------------------------------------
Dim sql As String
    
    sql = "SELECT * FROM slima1"
    If enlaza Then
        sql = sql & " " & ObtenerWhereCP(True)
    Else
        sql = sql & " WHERE codclien = -1"
    End If
    sql = sql & Ordenacion
    MontaSQLCarga = sql
End Function


Private Sub SumarTotal()
Dim total As Currency
Dim I As Byte

    'Calcular el total actual
    total = 0
    For I = 8 To 30 Step 2
        If Text1(I).Text <> "" Then total = total + CCur(Text1(I).Text)
    Next I
    If total <> 0 Then
        Text2(2).Text = CStr(total)
        PonerFormatoDecimal Text2(2), 3
    Else
        Text2(2).Text = ""
    End If
    
    'Calcular el total Siguiente
    total = 0
    For I = 9 To 31 Step 2
        If Text1(I).Text <> "" Then total = total + CCur(Text1(I).Text)
    Next I
    If total <> 0 Then
        Text2(3).Text = CStr(total)
        PonerFormatoDecimal Text2(3), 3
    Else
        Text2(3).Text = ""
    End If
End Sub


Private Sub SumarTotalHistorico()
Dim total As Currency
Dim I As Byte

    'Calcular el total actual
    total = 0
    For I = 22 To 33
        If Text2(I).Text <> "" Then total = total + CCur(Text2(I).Text)
    Next I
    If total <> 0 Then
        Text2(I).Text = CStr(total)
        PonerFormatoDecimal Text2(I), 3
    Else
        Text2(I).Text = ""
    End If
    
    'Calcular el total Siguiente
    total = 0
    For I = 35 To 46
        If Text2(I).Text <> "" Then total = total + CCur(Text2(I).Text)
    Next I
    If total <> 0 Then
        Text2(I).Text = CStr(total)
        PonerFormatoDecimal Text2(I), 3
    Else
        Text2(I).Text = ""
    End If
End Sub



Private Sub PonerModoOpcionesMenu(Modo As Byte)
'Activas unas Opciones de Menu y Toolbar según el modo en que estemos
Dim b As Boolean

    b = (Modo = 2) Or (Modo = 5)
    'Insertar
    Toolbar1.Buttons(5).Enabled = (b Or Modo = 0)
    Me.mnNuevo.Enabled = (b Or Modo = 0)
    'Modificar
    Toolbar1.Buttons(6).Enabled = b
    Me.mnModificar.Enabled = b
    'eliminar
    Toolbar1.Buttons(7).Enabled = b
    Me.mnEliminar.Enabled = b
    
    
    b = (Modo = 2)
    'Mantenimiento lineas Revisiones
    Toolbar1.Buttons(10).Enabled = b
    Me.mnRevisiones.Enabled = b
    'Lineas Histórico
    Toolbar1.Buttons(11).Enabled = b
    Me.mnHistorico.Enabled = b
    Me.mnOpciones.Enabled = b Or (Modo = 0)
    Me.mnMtoLineas.Enabled = b Or (Modo = 0)
    'Componentes
    Me.Toolbar1.Buttons(12).Enabled = b
    
    
    b = (Modo >= 3) Or Modo = 1
    'Buscar
    Toolbar1.Buttons(1).Enabled = Not b
    Me.mnBuscar.Enabled = Not b
    'Ver Todos
    Toolbar1.Buttons(2).Enabled = Not b
    Me.mnVerTodos.Enabled = Not b
End Sub


Private Sub PonerCamposHistorico()
Dim I As Integer
    
    On Error Resume Next
    
    If Data3.Recordset.EOF Then
        Data3.Recordset.MoveLast
        Exit Sub
    End If
    'Pone 2 años (2 registros) cada vez
    'Primer Año
    '----------------------------------------------------------------------------
    Me.Label1(38).Caption = Data3.Recordset.Fields(0).Value
    Me.Label1(47).Caption = Me.Label1(38).Caption
    
    For I = 1 To 12
        
        'Text2(22).Text = Format(Data3.Recordset.Fields(4).Value, FormatoCantidad)
        Text2(21 + I).Text = Format(Data3.Recordset.Fields(I).Value, FormatoCantidad)
    Next I
    
       
    'Segundo Año
    '----------------------------------------------------------------------------
    Data3.Recordset.MoveNext
    If Not Data3.Recordset.EOF Then
        'Poner el año siguiente
        Me.Label1(37).Caption = Data3.Recordset.Fields(0).Value
        Me.Label1(46).Caption = Me.Label1(37).Caption
        
        
        For I = 1 To 12
            
            'Text2(22).Text = Format(Data3.Recordset.Fields(4).Value, FormatoCantidad)
            Text2(34 + I).Text = Format(Data3.Recordset.Fields(I).Value, FormatoCantidad)
        Next I
        
        
        
    Else
        Data3.Recordset.MoveLast
        Me.Label1(37).Caption = ""
        Me.Label1(46).Caption = ""
        For I = 4 To 15
            Text2(I + 31).Text = ""
        Next I
    End If
    SumarTotalHistorico
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub LimpiarCamposHistorico()
Dim I As Byte

    Me.imgFlecha(0).visible = False
    Me.imgFlecha(1).visible = False
    Me.Label1(38).Caption = "Año"
    Me.Label1(37).Caption = "Año"
    Me.Label1(46).Caption = "Año"
    Me.Label1(47).Caption = "Año"
    For I = 4 To 15
        Text2(I + 18).Text = ""
        Text2(I + 31).Text = ""
    Next I
    'Limpiar el total del Histórico
    Text2(34).Text = ""
    Text2(47).Text = ""
End Sub


Private Sub BotonComponentes()
'Muestra un form de Mensaje para seleccionar el tipo de resumen que queremos mostrar:
'Por Mantenimiento, Por Departamento, Por Cliente
Dim vWhere As String

    'Codigo Cliente
    If Text1(0).Text = "" Then Exit Sub
    vWhere = " WHERE codclien = " & Text1(0).Text
    
    frmMensajes.cadWHERE = vWhere
    
    'cadClien= codclien - nomclien
    vWhere = Text1(0).Text & " - " & Text2(0).Text
    'vCampos= Mantenimiento|coddirec|Desc. coddirec| cadCliente
    frmMensajes.vCampos = Text1(2).Text & "|" & Text1(1).Text & "|" & Text2(1).Text & "|" & vWhere & "|"
    frmMensajes.OpcionMensaje = 5 'Componentes
    frmMensajes.Show vbModal
End Sub



Private Sub CargaComboMes()
Dim F As Date
    Me.cmbMes.Clear
    'Sin especificar
        cmbMes.AddItem "    "
        cmbMes.ItemData(0) = 0
    For kCampo = 1 To 12
        NombreTabla = Format(CDate("01/" & kCampo & "/2000"), "mmmm")
        NombreTabla = UCase(Mid(NombreTabla, 1, 1)) & Mid(NombreTabla, 2)
        cmbMes.AddItem NombreTabla
        cmbMes.ItemData(kCampo) = kCampo
    Next kCampo
    
End Sub


Private Sub Colorines()
Dim I As Integer
Dim J As Integer
    
    I = cmbMes.ListIndex
    For J = 1 To I
        Text1(((J - 1) * 2) + 8).ForeColor = vbRed
    Next J
    For J = I + 1 To 12
        Text1(((J - 1) * 2) + 8).ForeColor = vbBlack
    Next J
End Sub
