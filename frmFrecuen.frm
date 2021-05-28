VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFrecuencias 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Frecuencias"
   ClientHeight    =   6780
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   11055
   ClipControls    =   0   'False
   Icon            =   "frmFrecuen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   11055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   33
      Left            =   9120
      MaxLength       =   10
      TabIndex        =   19
      Tag             =   "Tipo cable|T|S|||scafre|cablerep|||"
      Text            =   "Text1"
      Top             =   3540
      Width           =   1695
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   1
      Left            =   6360
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Tag             =   "Prop. ubicacion|N|S|||scafre|propubic|||"
      Top             =   3540
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   0
      Left            =   3000
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Tag             =   "Prop. equip|N|S|||scafre|propirep|||"
      Top             =   3540
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   31
      Left            =   4080
      MaxLength       =   6
      TabIndex        =   35
      Tag             =   "Altura|N|S|||scafre|alturrep|||"
      Text            =   "Text1"
      Top             =   5820
      Width           =   615
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   6720
      TabIndex        =   53
      Text            =   "Text2"
      Top             =   600
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   1920
      TabIndex        =   52
      Text            =   "Text2"
      Top             =   600
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   1395
      Index           =   32
      Left            =   6360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   36
      Tag             =   "O|T|S|||scafre|obs01rep|||"
      Text            =   "frmFrecuen.frx":000C
      Top             =   4800
      Width           =   4455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   30
      Left            =   1560
      MaxLength       =   6
      TabIndex        =   34
      Tag             =   "Metros|N|S|||scafre|mcablrep|||"
      Text            =   "Text1"
      Top             =   5820
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   29
      Left            =   4080
      MaxLength       =   35
      TabIndex        =   33
      Tag             =   "Pote.|N|S|||scafre|potenrep|||"
      Text            =   "Text1"
      Top             =   5400
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   28
      Left            =   1560
      MaxLength       =   6
      TabIndex        =   32
      Tag             =   "Cota|N|S|||scafre|mcotarep|||"
      Text            =   "Text1"
      Top             =   5380
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   27
      Left            =   4920
      MaxLength       =   30
      TabIndex        =   31
      Tag             =   "Coor|T|S|||scafre|coo24rep|||"
      Text            =   "Text1"
      Top             =   4920
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   26
      Left            =   4560
      MaxLength       =   30
      TabIndex        =   30
      Tag             =   "Coor|N|S|||scafre|coo23rep|00||"
      Text            =   "Text1"
      Top             =   4920
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   25
      Left            =   4200
      MaxLength       =   30
      TabIndex        =   29
      Tag             =   "Coor|N|S|||scafre|coo22rep|00||"
      Text            =   "Text1"
      Top             =   4920
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   24
      Left            =   3720
      MaxLength       =   30
      TabIndex        =   28
      Tag             =   "Coor|N|S|||scafre|coo21rep|000||"
      Text            =   "Text1"
      Top             =   4920
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   23
      Left            =   2760
      MaxLength       =   30
      TabIndex        =   27
      Tag             =   "Coor|T|S|||scafre|coo14rep|||"
      Text            =   "Text1"
      Top             =   4920
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   22
      Left            =   2400
      MaxLength       =   30
      TabIndex        =   26
      Tag             =   "Coor|N|S|||scafre|coo13rep|00||"
      Text            =   "Text1"
      Top             =   4920
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   21
      Left            =   2040
      MaxLength       =   30
      TabIndex        =   25
      Tag             =   "Coor|N|S|||scafre|coo12rep|00||"
      Text            =   "Text1"
      Top             =   4920
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   20
      Left            =   1560
      MaxLength       =   30
      TabIndex        =   24
      Tag             =   "Coor|N|S|||scafre|coo11rep|000||"
      Text            =   "Text1"
      Top             =   4920
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   19
      Left            =   6360
      MaxLength       =   50
      TabIndex        =   23
      Tag             =   "A|T|S|||scafre|antenrep|||"
      Text            =   "DAVIDGANDULCASTELLSDAVIDGANDUL"
      Top             =   4380
      Width           =   4455
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   18
      Left            =   1560
      MaxLength       =   50
      TabIndex        =   22
      Tag             =   "U|T|S|||scafre|ubicarep|||"
      Text            =   "DAVIDGANDULCASTELLSDAVIDGANDUL"
      Top             =   4440
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   17
      Left            =   3720
      MaxLength       =   35
      TabIndex        =   21
      Tag             =   "NSerie|T|S|||scafre|nomserie|||"
      Text            =   "Text1"
      Top             =   3960
      Width           =   4815
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   16
      Left            =   1560
      MaxLength       =   15
      TabIndex        =   20
      Tag             =   "NSerie|T|S|||scafre|numserie|||"
      Text            =   "Text1"
      Top             =   3960
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   15
      Left            =   9000
      TabIndex        =   16
      Tag             =   "Certif|F|S|||scafre|feccambi|dd/mm/yyyy||"
      Text            =   "Text1"
      Top             =   2700
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   14
      Left            =   5160
      TabIndex        =   15
      Tag             =   "Certif|F|S|||scafre|feccerti|||"
      Text            =   "Text1"
      Top             =   2700
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   13
      Left            =   1440
      TabIndex        =   14
      Tag             =   "Proyecto|F|S|||scafre|fecproye|dd/mm/yyyy||"
      Text            =   "99/99/9999"
      Top             =   2700
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   12
      Left            =   9840
      MaxLength       =   30
      TabIndex        =   13
      Tag             =   "Freq|N|S|||scafre|cantxrpt|0.00000||"
      Text            =   "999,9999"
      Top             =   2100
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   11
      Left            =   6720
      MaxLength       =   30
      TabIndex        =   12
      Tag             =   "Freq|N|S|||scafre|subtxrpt|0.00000||"
      Text            =   "999,9999"
      Top             =   2100
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   10
      Left            =   3720
      MaxLength       =   30
      TabIndex        =   11
      Tag             =   "Freq|N|S|||scafre|fretxrpt|0.00000||"
      Text            =   "Text1"
      Top             =   2100
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   9
      Left            =   9840
      MaxLength       =   30
      TabIndex        =   10
      Tag             =   "Freq|N|S|||scafre|canrxrpt|0.00000||"
      Text            =   "999,9999"
      Top             =   1680
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   8
      Left            =   6720
      MaxLength       =   30
      TabIndex        =   9
      Tag             =   "Freq|N|S|||scafre|subrxrpt|0.00000||"
      Text            =   "999,9999"
      Top             =   1680
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   7
      Left            =   3720
      MaxLength       =   30
      TabIndex        =   8
      Tag             =   "Freq|N|S|||scafre|frerxrpt|0.00000||"
      Text            =   "999,9999"
      Top             =   1680
      Width           =   975
   End
   Begin VB.CheckBox Check1 
      Height          =   315
      Left            =   9840
      TabIndex        =   2
      Tag             =   "D|N|S|||scafre|legalsno||S|"
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   6
      Left            =   10320
      MaxLength       =   6
      TabIndex        =   7
      Tag             =   "Año|N|S|||scafre|anorenov|||"
      Text            =   "Text1"
      Top             =   1080
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   5
      Left            =   6480
      MaxLength       =   30
      TabIndex        =   6
      Tag             =   "D|T|S|||scafre|nomcanal||N|"
      Text            =   "Text1"
      Top             =   1080
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   4
      Left            =   5640
      MaxLength       =   6
      TabIndex        =   5
      Tag             =   "NºCanal|N|N|||scafre|numcanal||S|"
      Text            =   "Text1"
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   3
      Left            =   3720
      TabIndex        =   4
      Tag             =   "Fecha inicio|F|N|||scafre|fechaini|dd/mm/yyyy|S|"
      Text            =   "Text1"
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   2
      Left            =   1080
      MaxLength       =   15
      TabIndex        =   3
      Tag             =   "Numexp|T|N|||scafre|numexped||S|"
      Text            =   "Text1"
      Top             =   1080
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   1
      Left            =   6000
      MaxLength       =   40
      TabIndex        =   1
      Tag             =   "Dpto|N|N|||scafre|coddirec||S|"
      Text            =   "Text1"
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   8040
      TabIndex        =   37
      Top             =   6360
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   9360
      TabIndex        =   38
      Top             =   6360
      Width           =   1035
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   9315
      TabIndex        =   39
      Top             =   6360
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Index           =   0
      Left            =   120
      TabIndex        =   44
      Top             =   6240
      Width           =   2535
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   240
         Left            =   240
         TabIndex        =   45
         Top             =   180
         Width           =   2115
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   0
      Left            =   1080
      MaxLength       =   6
      TabIndex        =   0
      Tag             =   "Cod. clien|N|N|||scafre|codclien|00000|S|"
      Text            =   "Text1"
      Top             =   585
      Width           =   735
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   40
      Top             =   0
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   18
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ver Todos"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
            Object.ToolTipText     =   "Modificar expediente"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cambiar cliente /departamento"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Último"
         EndProperty
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   315
         Left            =   5520
         TabIndex        =   43
         Top             =   0
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   480
      Top             =   6360
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
   Begin VB.Label Label6 
      Caption         =   "Tipo cable"
      Height          =   195
      Index           =   29
      Left            =   8280
      TabIndex        =   82
      Top             =   3540
      Width           =   750
   End
   Begin VB.Line Line1 
      X1              =   5400
      X2              =   5400
      Y1              =   4320
      Y2              =   6240
   End
   Begin VB.Label Label3 
      Caption         =   "Legal"
      Height          =   255
      Left            =   10200
      TabIndex        =   81
      Top             =   660
      Width           =   615
   End
   Begin VB.Image imgWeb 
      Height          =   255
      Left            =   1320
      Picture         =   "frmFrecuen.frx":0012
      Stretch         =   -1  'True
      Tag             =   "-1"
      ToolTipText     =   "Abrir web"
      Top             =   4920
      Width           =   255
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   1
      Left            =   5640
      Picture         =   "frmFrecuen.frx":059C
      Tag             =   "-1"
      ToolTipText     =   "Buscar Dir./Dpto."
      Top             =   600
      Width           =   240
   End
   Begin VB.Image imgFecha 
      Height          =   240
      Index           =   15
      Left            =   8640
      Picture         =   "frmFrecuen.frx":069E
      ToolTipText     =   "Buscar fecha"
      Top             =   2730
      Width           =   240
   End
   Begin VB.Image imgFecha 
      Height          =   240
      Index           =   14
      Left            =   4800
      Picture         =   "frmFrecuen.frx":0729
      ToolTipText     =   "Buscar fecha"
      Top             =   2730
      Width           =   240
   End
   Begin VB.Image imgFecha 
      Height          =   240
      Index           =   13
      Left            =   1080
      Picture         =   "frmFrecuen.frx":07B4
      ToolTipText     =   "Buscar fecha"
      Top             =   2730
      Width           =   240
   End
   Begin VB.Image imgFecha 
      Height          =   240
      Index           =   3
      Left            =   3360
      Picture         =   "frmFrecuen.frx":083F
      ToolTipText     =   "Buscar fecha"
      Top             =   1117
      Width           =   240
   End
   Begin VB.Shape Shape1 
      Height          =   975
      Left            =   30
      Top             =   1560
      Width           =   10935
   End
   Begin VB.Label Label6 
      Caption         =   "Antena"
      Height          =   195
      Index           =   28
      Left            =   5640
      TabIndex        =   80
      Top             =   4380
      Width           =   585
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Obser:"
      Height          =   195
      Index           =   27
      Left            =   5520
      TabIndex        =   79
      Top             =   4860
      Width           =   600
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Altura"
      Height          =   195
      Index           =   26
      Left            =   3360
      TabIndex        =   78
      Top             =   5820
      Width           =   570
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Potencia"
      Height          =   195
      Index           =   25
      Left            =   3120
      TabIndex        =   77
      Top             =   5400
      Width           =   810
   End
   Begin VB.Label Label6 
      Caption         =   "metros."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   24
      Left            =   4680
      TabIndex        =   76
      Top             =   5880
      Width           =   705
   End
   Begin VB.Label Label6 
      Caption         =   "Metros:"
      Height          =   195
      Index           =   23
      Left            =   240
      TabIndex        =   75
      Top             =   5820
      Width           =   570
   End
   Begin VB.Label Label6 
      Caption         =   "Watios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   22
      Left            =   4680
      TabIndex        =   74
      Top             =   5400
      Width           =   870
   End
   Begin VB.Label Label6 
      Caption         =   "metros."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   21
      Left            =   2160
      TabIndex        =   73
      Top             =   5880
      Width           =   690
   End
   Begin VB.Label Label6 
      Caption         =   "metros."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   20
      Left            =   2160
      TabIndex        =   72
      Top             =   5400
      Width           =   705
   End
   Begin VB.Label Label6 
      Caption         =   "Coordenadas"
      Height          =   195
      Index           =   19
      Left            =   240
      TabIndex        =   71
      Top             =   4920
      Width           =   945
   End
   Begin VB.Label Label6 
      Caption         =   "Cota"
      Height          =   195
      Index           =   18
      Left            =   240
      TabIndex        =   70
      Top             =   5400
      Width           =   570
   End
   Begin VB.Label Label6 
      Caption         =   "Nº Serie"
      Height          =   195
      Index           =   17
      Left            =   240
      TabIndex        =   69
      Top             =   3960
      Width           =   585
   End
   Begin VB.Label Label6 
      Caption         =   "Ubicación"
      Height          =   195
      Index           =   16
      Left            =   240
      TabIndex        =   68
      Top             =   4440
      Width           =   720
   End
   Begin VB.Label Label6 
      Caption         =   "Propiedad equipo"
      Height          =   195
      Index           =   15
      Left            =   1560
      TabIndex        =   67
      Top             =   3540
      Width           =   1410
   End
   Begin VB.Label Label6 
      Caption         =   "Propiedad ubicación"
      Height          =   195
      Index           =   13
      Left            =   4800
      TabIndex        =   66
      Top             =   3540
      Width           =   1590
   End
   Begin VB.Label Label6 
      Caption         =   "Frecuencia (Tx de Rpt)"
      Height          =   195
      Index           =   11
      Left            =   0
      TabIndex        =   65
      Top             =   0
      Width           =   1650
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Cambio de canal"
      Height          =   195
      Index           =   14
      Left            =   7200
      TabIndex        =   64
      Top             =   2760
      Width           =   1410
   End
   Begin VB.Label Label6 
      Caption         =   "Proyecto"
      Height          =   195
      Index           =   12
      Left            =   240
      TabIndex        =   63
      Top             =   2760
      Width           =   810
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Certificación"
      Height          =   195
      Index           =   10
      Left            =   3720
      TabIndex        =   62
      Top             =   2760
      Width           =   1050
   End
   Begin VB.Label Label6 
      Caption         =   "Canalizacion  (ca de tx)"
      Height          =   195
      Index           =   9
      Left            =   8040
      TabIndex        =   61
      Top             =   2160
      Width           =   1650
   End
   Begin VB.Label Label6 
      Caption         =   "Canalizacion  (ca de Rx)"
      Height          =   195
      Index           =   8
      Left            =   8040
      TabIndex        =   60
      Top             =   1740
      Width           =   1770
   End
   Begin VB.Label Label6 
      Caption         =   "Subtono (Sb de Tx)"
      Height          =   195
      Index           =   7
      Left            =   5160
      TabIndex        =   59
      Top             =   2160
      Width           =   1650
   End
   Begin VB.Label Label6 
      Caption         =   "Subtono (Sb de Rx)"
      Height          =   195
      Index           =   6
      Left            =   5160
      TabIndex        =   58
      Top             =   1740
      Width           =   1530
   End
   Begin VB.Label Label6 
      Caption         =   "Frecuencia (Tx de Rpt)"
      Height          =   195
      Index           =   2
      Left            =   2040
      TabIndex        =   57
      Top             =   2160
      Width           =   1650
   End
   Begin VB.Label Label6 
      Caption         =   "REPETIDOR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   300
      Index           =   5
      Left            =   120
      TabIndex        =   56
      Top             =   3240
      Width           =   1440
   End
   Begin VB.Label Label6 
      Caption         =   "TRANSMISION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   300
      Index           =   4
      Left            =   120
      TabIndex        =   55
      Top             =   2100
      Width           =   1560
   End
   Begin VB.Label Label6 
      Caption         =   "RECEPCION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   300
      Index           =   3
      Left            =   120
      TabIndex        =   54
      Top             =   1680
      Width           =   1560
   End
   Begin VB.Label Label6 
      Caption         =   "Frecuencia (Rx de Rpt)"
      Height          =   195
      Index           =   1
      Left            =   2040
      TabIndex        =   51
      Top             =   1740
      Width           =   1890
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   0
      Left            =   720
      Picture         =   "frmFrecuen.frx":08CA
      Tag             =   "-1"
      ToolTipText     =   "Buscar cliente"
      Top             =   615
      Width           =   240
   End
   Begin VB.Label Label8 
      Caption         =   "Renov."
      Height          =   195
      Left            =   9720
      TabIndex        =   50
      Top             =   1140
      Width           =   525
   End
   Begin VB.Label Label6 
      Caption         =   "Nº Canal"
      Height          =   255
      Index           =   0
      Left            =   4920
      TabIndex        =   49
      Top             =   1110
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "F.Incio"
      Height          =   255
      Left            =   2760
      TabIndex        =   48
      Top             =   1110
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Expediente"
      Height          =   195
      Left            =   120
      TabIndex        =   47
      Top             =   1140
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "Dpto."
      Height          =   195
      Left            =   5160
      TabIndex        =   46
      Top             =   645
      Width           =   645
   End
   Begin VB.Label Label5 
      Caption         =   "Cliente"
      Height          =   255
      Left            =   120
      TabIndex        =   42
      Top             =   615
      Width           =   615
   End
   Begin VB.Label Label10 
      Caption         =   "Cargando datos ........."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   41
      Top             =   8220
      Visible         =   0   'False
      Width           =   3495
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
         HelpContextID   =   2
         Shortcut        =   ^N
      End
      Begin VB.Menu mnModificar 
         Caption         =   "&Modificar"
         HelpContextID   =   2
         Shortcut        =   ^M
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
End
Attribute VB_Name = "frmFrecuencias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

Private WithEvents frmB As frmBuscaGrid 'Form para busquedas (frmBuscaGrid)
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmF As frmCal
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmMtoCliente As frmBasico2 'frmFacClientesGr
Attribute frmMtoCliente.VB_VarHelpID = -1

Dim NombreTabla As String
Dim Ordenacion As String
Private Modo As Byte
Private ModoAnterior As Byte
Dim kCampo As Integer

Dim btnPrimero As Byte
'Variable que indica el número del Boton  PrimerRegistro en la Toolbar1

'Si lanzamos el google earth o el google maps
Dim GoogleMaps As Boolean

Dim CadenaConsulta As String

Private HaDevueltoDatos As Boolean

Private VieneDeBuscar As Boolean
'Para cuando devuelve dos poblaciones con el mismo codigo Postal. Si viene de pulsar prismatico
'de busqueda poner el valor de poblacion seleccionado y no volver a recuperar de la Base de Datos


'Private Sub cboTipoDirec_KeyPress(KeyAscii As Integer)
'    KEYpress KeyAscii
'End Sub


Private Sub Check1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()
On Error GoTo EAceptar
    
    Screen.MousePointer = vbHourglass
    Select Case Modo
        Case 1 'BUSCAR
            HacerBusqueda
            
        Case 3 'INSERTAR
            If DatosOk Then
                If InsertarDesdeForm(Me) Then PosicionarData
            End If
            
        Case 4 'MODIFICAR
            If DatosOk Then
                 If ModificaDesdeFormulario(Me, 1) Then
                     TerminaBloquear
                     PosicionarData
                 End If
            End If
    End Select
EAceptar:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub cmdCancelar_Click()
On Error GoTo ECancelar

    Select Case Modo
        Case 1 'Buscar
            LimpiarCampos
            PonerModo 0
        Case 3 'Insertar
            If ModoAnterior = 0 Then
                LimpiarCampos
                PonerModo 0
            Else
                PonerModo 2
                PonerCampos
            End If
        Case 4  'Modificar
            TerminaBloquear
            PonerModo 2
            PonerCampos
            PonerFoco Text1(5)
    End Select
ECancelar:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub cmdRegresar_Click()
Dim cad As String

    If Data1.Recordset.EOF Then
        MsgBox "Ningún registro devuelto.", vbExclamation
        Exit Sub
    End If

    cad = Data1.Recordset.Fields(0) & "|"
    cad = cad & Data1.Recordset.Fields(2) & "|"
    RaiseEvent DatoSeleccionado(cad)
    Unload Me
End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    If Modo = 1 Then PonerFoco Text1(0)
End Sub


Private Sub Form_Load()
    'Icono del formulario
    Me.Icon = frmPpal.Icon
    
    'ICONOS de La toolbar
    btnPrimero = 15 'Posicion del Boton Primero en la toolbar (+ 3 siguientes)
    With Toolbar1
        .ImageList = frmPpal.imgListComun
        'ASignamos botones
        .Buttons(1).Image = 1   'Buscar
        .Buttons(2).Image = 2 'Ver Todos
        .Buttons(5).Image = 3 'Añadir
        .Buttons(6).Image = 4 'Modificar
        .Buttons(7).Image = 5 'Eliminar
        .Buttons(9).Image = 43 'Mod cabecera
        
        .Buttons(10).Image = 17 'Mod cabecera
        
        .Buttons(12).Image = 16 'Imprimir
        .Buttons(13).Image = 15 'Salir
        .Buttons(btnPrimero).Image = 6 'Primero
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
        .Buttons(btnPrimero + 3).Image = 9 'Ultimo
    End With
    
    LimpiarCampos   'Limpia los campos TextBox
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    CargarComboTipoDirec
    
    GoogleMaps = True
    ComprobarGoogleEarth
    Me.imgWeb.Tag = NombreTabla
    If imgWeb.Tag = "" Then imgWeb.Enabled = False
    
    NombreTabla = "scafre" 'Frecuencias
    Ordenacion = " ORDER BY codclien,numexped"
    CadenaConsulta = "Select * from " & NombreTabla & " WHERE codclien = -1" 'No recupera datos
    Data1.ConnectionString = conn
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    
    If DatosADevolverBusqueda = "" Then
        PonerModo 0
    Else
        PonerModo 1 'Modo Busqueda
    End If
    
    
    Label1.Caption = "Dpto."
    
    
    Screen.MousePointer = vbDefault
End Sub


Private Sub frmB_Selecionado(CadenaDevuelta As String)
'Formulario para Busqueda
Dim cadB As String
Dim Aux As String
      
      
    
    If CadenaDevuelta <> "" Then
            HaDevueltoDatos = True
            Screen.MousePointer = vbHourglass
            
            
            If frmB.vTabla = "sdirec" Then
                Text1(1).Text = RecuperaValor(CadenaDevuelta, 1)
                Text2(1).Text = RecuperaValor(CadenaDevuelta, 2)
            Else
                'Estamos en Cabecera
                'Recupera todo el registro de Tarifas de Precios
                'Sabemos que campos son los que nos devuelve
                'Creamos una cadena consulta y ponemos los datos
                cadB = ""
                Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
                cadB = Aux
                Aux = ValorDevueltoFormGrid(Text1(1), CadenaDevuelta, 2)
                cadB = cadB & " and " & Aux
                Aux = ValorDevueltoFormGrid(Text1(2), CadenaDevuelta, 3)
                cadB = cadB & " and " & Aux
                Aux = ValorDevueltoFormGrid(Text1(3), CadenaDevuelta, 4)
                cadB = cadB & " and " & Aux
                CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
                PonerCadenaBusqueda
            End If
    End If
    Screen.MousePointer = vbDefault
End Sub


'Private Sub frmCP_DatoSeleccionado(CadenaSeleccion As String)
''Formulario Mantenimiento C. Postales
'Dim Indice As Byte
'Dim devuelve As String
'
'    Indice = 3
'    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'CPostal
'    Text1(Indice + 1).Text = ObtenerPoblacion(Text1(Indice).Text, devuelve) 'poblacion
'    'provincia
'    Text1(Indice + 2).Text = devuelve
'End Sub



Private Sub frmF_Selec(vFecha As Date)


    Text1(CInt(Me.imgFecha(3).Tag)).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmMtoCliente_DatoSeleccionado(CadenaSeleccion As String)
    Text1(0).Text = RecuperaValor(CadenaSeleccion, 1)
    Text2(0).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub imgBuscar_Click(Index As Integer)
Dim cad As String

    If Modo <> 3 And Modo <> 1 Then Exit Sub 'SOLO INSERTAR,  o buscar y o
 
    Screen.MousePointer = vbHourglass
    
    VieneDeBuscar = True
        
    If Index = 0 Then
        'NOMBRE CLIENTE
'        Set frmMtoCliente = New frmFacClientesGr
'        frmMtoCliente.DatosADevolverBusqueda = "0|1|"
        If Not IsNumeric(Text1(0).Text) Then Text1(0).Text = ""
'        frmMtoCliente.Show vbModal
        
        Set frmMtoCliente = New frmBasico2
        AyudaClientes frmMtoCliente, Text1(0).Text
        Set frmMtoCliente = Nothing
        
    Else
        'DEPARTAMENTO
        
        If Text1(0).Text = "" Then
            MsgBox "Seleccione el cliente", vbExclamation
            Exit Sub
        End If
                
        Set frmB = New frmBuscaGrid
        
'        If vParamAplic.Departamento Then
            cad = "Dptos."
'        Else
'            Cad = "Direc."
'        End If
        
        cad = cad & " Cliente: " & Text1(0).Text & " - " & Text2(0).Text
        frmB.vTitulo = cad
        cad = "Codigo|sdirec|coddirec|N|000|15·"
        cad = cad & "Descripcion|sdirec|nomdirec|T||55·"
        frmB.vCampos = cad
        frmB.vTabla = "sdirec"
        frmB.vSQL = " codclien =" & Text1(0).Text
        frmB.vCargaFrame = False
        frmB.vConexionGrid = conAri 'Conexion a BD Ariges
        frmB.vDevuelve = "0|1|"
        frmB.vselElem = 0
        frmB.Show vbModal
        Set frmB = Nothing
        
        
    End If
    'PonerFoco Text1(3)
    Screen.MousePointer = vbDefault
End Sub


Private Sub imgFecha_Click(Index As Integer) 'Abre calendario Fechas


   If Modo = 2 Or Modo = 0 Then Exit Sub
   If Modo = 4 And Index = 3 Then Exit Sub 'La fec
   Screen.MousePointer = vbHourglass
   Set frmF = New frmCal
   frmF.Fecha = Now

   Me.imgFecha(3).Tag = Index
   
   PonerFormatoFecha Text1(Index)
   If Text1(Index).Text <> "" Then frmF.Fecha = CDate(Text1(Index).Text)

   Screen.MousePointer = vbDefault
   frmF.Show vbModal
   Set frmF = Nothing
   PonerFoco Text1(Index)
End Sub

Private Sub imgWeb_Click()
Dim L As Double
Dim La As Double


    If Modo <> 2 Then Exit Sub
    If Data1.Recordset.EOF Then Exit Sub
    'Voy a lanzar el google earth
    On Error GoTo EGoo
    
    
    
            'Convertimos G,m,s a grados decimal
        If Text1(23).Text = "S" Or Text1(23).Text = "N" Then
            'LONGITUD
            La = Val(Text1(22).Text) / 3600
            La = La + Val(Text1(21).Text) / 60
            La = Val(Text1(20).Text) + La
            L = Val(Text1(24).Text) + Val(Text1(25).Text) / 60 + Val(Text1(26).Text) / 3600
            If Text1(23).Text = "S" Then La = -1 * La
            If Text1(27).Text <> "E" Then L = -1 * L
        Else
            La = Val(Text1(20).Text) + Val(Text1(21).Text) / 60 + Val(Text1(22).Text) / 3600
            L = Val(Text1(24).Text) + Val(Text1(25).Text) / 60 + Val(Text1(26).Text) / 3600
            If Text1(27).Text = "S" Then La = -1 * La
            If Text1(23).Text <> "E" Then L = -1 * L
        End If
        La = Round(La, 5)
        L = Round(L, 5)
    
    
    
    
    
    
    
    If Not GoogleMaps Then
        'GOOGLE EARTH
        CadenaDesdeOtroForm = App.Path & "\Antena.kml"
        If Dir(CadenaDesdeOtroForm, vbArchive) <> "" Then Kill CadenaDesdeOtroForm
        
        
        NumRegElim = FreeFile
        Open CadenaDesdeOtroForm For Output As NumRegElim
        Print #NumRegElim, "<?xml version=""1.0"" encoding=""UTF-8""?>"
        Print #NumRegElim, " <kml xmlns=""http://earth.google.com/kml/2.0"">"
        Print #NumRegElim, "  <Placemark>"
        Print #NumRegElim, "    <name>" & Text2(0).Text & "</name>"
        Print #NumRegElim, "    <visibility>0</visibility>"
        Print #NumRegElim, "    <LookAt id=""khLookAt786"">"

        Print #NumRegElim, "        <longitude>" & TransformaComasPuntos(Format(L, "0.0000000000")) & "</longitude>"
        'Convertimos G,m,s a grados decimal
        
        Print #NumRegElim, "        <latitude>" & TransformaComasPuntos(Format(La, "0.0000000000")) & "</latitude>"
        
        Print #NumRegElim, "        <range>392.9086289641584</range>"
        Print #NumRegElim, "        <tilt>3.915988552288592e-011</tilt>"
        Print #NumRegElim, "        <heading>21.24266674690592</heading>"
        Print #NumRegElim, "    </LookAt>"
        Print #NumRegElim, "    <styleUrl>root://styleMaps#default+nicon=0x307+hicon=0x317</styleUrl>"
        Print #NumRegElim, "    <Point id=""khPoint787"">"
        Print #NumRegElim, "    <coordinates>" & TransformaComasPuntos(Format(L, "0.0000000000")) & "," & TransformaComasPuntos(Format(La, "0.0000000000")) & ",0</coordinates>"
        Print #NumRegElim, "    </Point>"
        Print #NumRegElim, "   </Placemark>"
        Print #NumRegElim, "</kml>"
        Close NumRegElim
    
    
        
        
        
        
        Else
            'GOOGLE MAPs
            CadenaDesdeOtroForm = "lat=" & Trim(TransformaComasPuntos(CStr(La))) & "&lng=" & Trim(TransformaComasPuntos(CStr(L))) & "&zoom=18"
            CadenaDesdeOtroForm = "www.goolzoom.com/mapa.html?" & CadenaDesdeOtroForm
        End If
    CadenaDesdeOtroForm = Me.imgWeb.Tag & " " & CadenaDesdeOtroForm
    Shell CadenaDesdeOtroForm, vbNormalFocus
    Espera 0.5
    DoEvents
    Exit Sub
EGoo:
    MuestraError Err.Number, "Mostrando google earth"
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnModificar_Click()
     If BLOQUEADesdeFormulario(Me) Then BotonModificar
End Sub

Private Sub mnNuevo_Click()
    BotonAnyadir
End Sub

Private Sub mnSalir_Click()
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    kCampo = Index
    If Index <> 32 Then ConseguirFoco Text1(Index), Modo
End Sub


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
        If Index <> 32 Then KEYdown KeyCode
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index <> 32 Then KEYpress KeyAscii
End Sub


Private Sub Text1_LostFocus(Index As Integer)
Dim devuelve As String
On Error Resume Next

    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
    Text1(Index).Text = Trim(Text1(Index).Text)
    If Text1(Index).Text = "" Then
        If Index = 0 Then
                Text1(0).Text = ""
                Text1(1).Text = ""
                Text2(1).Text = ""
        Else
            If Index = 1 Then Text2(1).Text = ""
        End If
        Exit Sub
    End If
    With Text1(Index)
        Select Case Index
            Case 0 'Codigo Direccion
                devuelve = ""
                If PonerFormatoEntero(Text1(Index)) Then
                    'Comprobar si ya existe el cod de direccion en la tabla
                    If Modo = 3 Then 'Insertar
                        devuelve = DevuelveDesdeBD(conAri, "nomclien", "sclien", "codclien", Text1(0).Text, "N")
                        If devuelve = "" Then
                            Text1(0).Text = ""
                            Text1(1).Text = ""
                            Text2(1).Text = ""
                            PonerFoco Text1(0)
                        End If
                    End If
                End If
                Text2(0).Text = devuelve
            Case 1
                If Modo = 3 Then
                    devuelve = ""
                    If Text1(0).Text = "" Then
                        MsgBox "Ponga  el cliente", vbExclamation
                    Else
                        If Text1(1).Text = "0" Then
                            'Si es el CERO no pasa nada. NO es ningun departamento
                            
                        Else
                            If PonerFormatoEntero(Text1(Index)) Then
                                devuelve = DevuelveDesdeBDNew(conAri, "sdirec", "nomdirec", "codclien", Text1(0).Text, "N", "", "coddirec", Text1(1).Text, "N")
                                If devuelve = "" Then
                                    Text1(1).Text = "0"
                                    PonerFoco Text1(1)
                                End If
                            End If
                        End If
                    End If
                    Text2(1).Text = devuelve
                End If
            Case 20, 21, 22
                If Not PonerFormatoEntero(Text1(Index)) Then PonerFoco Text1(Index)
            Case 24, 25, 26
                If Not PonerFormatoEntero(Text1(Index)) Then PonerFoco Text1(Index)
            Case 23, 27
                'Letra de coordenadas
                devuelve = "NSEWO"
                Text1(Index).Text = UCase(Text1(Index).Text)
                If InStr(1, devuelve, Text1(Index).Text) = 0 Then
                    MsgBox "Letra coordenadas incorrectas", vbExclamation
                    PonerFoco Text1(Index)
                End If
            
            Case 3, 13, 14, 15
                devuelve = Text1(Index).Text
                If Not EsFechaOK(devuelve) Then devuelve = ""
                Text1(Index).Text = devuelve
                
            Case 7 To 12, 20, 21, 22, 24, 25, 26
            
                '8.- Siginica que el formato lo coje del tag
                If Not PonerFormatoDecimal(Text1(Index), 8) Then
                    Text1(Index).Text = ""
                    PonerFoco Text1(Index)
                End If
                
            
        End Select
    End With
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Index
        Case 1 'Busqueda
            mnBuscar_Click
        Case 2 'Ver Todos
            mnVerTodos_Click
        Case 5 'Nuevo
            mnNuevo_Click
        Case 6  'Modificar
            mnModificar_Click
        Case 7 'Eliminar
            mnEliminar_Click
        Case 9
            ModificarExpediente
            Screen.MousePointer = vbDefault
            
        Case 10
            'Cambiar cliente dpto
            CadenaDesdeOtroForm = ""
            frmListado2.Opcion = 45
            frmListado2.Show vbModal
            If CadenaDesdeOtroForm <> "" Then
                If Modo = 2 Then
                    LimpiarCampos
                    PonerCadenaBusqueda
                End If
                
            End If
            
        Case 12 'Imprimir
            AbrirListado 96
        Case 13  'Salir
            mnSalir_Click
        Case btnPrimero To btnPrimero + 3 'Flechas de Desplazamiento
            Desplazamiento (Button.Index - btnPrimero)
    End Select
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub PonerModo(Kmodo As Byte)
Dim b As Boolean
Dim NumReg As Byte 'Solo para saber que hay + de 1 Registro

    Modo = Kmodo
    PonerIndicador lblIndicador, Modo
    
    '--------------------------------------------
    'Modo 2. Hay datos y estamos visualizandolos
    b = (Kmodo = 2)
    
    'Visualizar flechas de desplazamiento en la toolbar si modo=2
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, NumReg
    
    
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.visible = b
        If Modo = 1 Then Me.lblIndicador.Caption = "BUSQUEDA"
    Else
        cmdRegresar.visible = False
    End If
    
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1 y bloquea clave primaria
    BloquearText1 Me, Modo
    
    Check1.Enabled = Modo = 3 Or Modo = 1
    
    'Bloquear Registro sino es Insert o Update
    b = (Modo = 0) Or (Modo = 2)
    
    
           
    '------------------------------------
    b = Modo <> 0 And Modo <> 2
    cmdCancelar.visible = b
    cmdAceptar.visible = b
    
   ' Me.Check1.Enabled = b
    Me.imgBuscar(0).Enabled = b
    Me.imgBuscar(1).Enabled = b
    Me.Combo1(0).Enabled = b
    Me.Combo1(1).Enabled = b
    
    chkVistaPrevia.Enabled = (Modo <= 2)
    
    'Poner el tamaño de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos
    
    PonerModoOpcionesMenu 'Activa las Opciones de menu según Modo
    PonerOpcionesMenu   'Activar opciones de menu según nivel
                        'de permisos del usuario
End Sub


Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
    PonerLongCamposGnral Me, Modo, 1
End Sub


Private Sub PonerModoOpcionesMenu()
'Activas unas Opciones de Menu y Toolbar según el modo en que estemos
Dim b As Boolean

    b = (Modo = 2 Or Modo = 0 Or Modo = 1)
    'Insertar
    Toolbar1.Buttons(5).Enabled = b
    Me.mnNuevo.Enabled = b
    
    b = (Modo = 2)
    'Modificar
    Toolbar1.Buttons(6).Enabled = b
    Me.mnModificar.Enabled = b
    'eliminar
    Toolbar1.Buttons(7).Enabled = b
    Me.mnEliminar.Enabled = b

    '-------------------------------------
    b = (Modo >= 3)
    'Buscar
    Toolbar1.Buttons(1).Enabled = Not b
    Me.mnBuscar.Enabled = Not b
    'Ver Todos
    Toolbar1.Buttons(2).Enabled = Not b
    Me.mnVerTodos.Enabled = Not b
End Sub


Private Sub LimpiarCampos()
    limpiar Me   'Metodo general: Limpia los controles TextBox
    'Aqui va el especifico de cada form es
    '### a mano
    Me.Combo1(0).ListIndex = -1
    Me.Combo1(1).ListIndex = -1
    Me.Check1.Value = 0
End Sub


Private Sub Desplazamiento(Index As Integer)
'Botones de Desplazamiento de la Toolbar
'Para desplazarse por los registros de control Data
    DesplazamientoData Data1, Index
    PonerCampos
End Sub


Private Sub BotonBuscar()
    'Buscar
    If Modo <> 1 Then
        LimpiarCampos
        PonerModo 1
    
        'Si pasamos el control
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
    
    If chkVistaPrevia.Value = 1 Then
        MandaBusquedaPrevia ""
    Else
        CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Sub BotonAnyadir()

    LimpiarCampos 'Vacía los TextBox
    ModoAnterior = Modo 'Para el botón Cancelar en Modo Insertar
    PonerModo 3
    


    PonerFoco Text1(0)
End Sub


Private Sub BotonModificar()
    If Data1.Recordset.EOF Then Exit Sub

    'Escondemos el navegador y ponemos Modo Modificar
    PonerModo 4
    PonerFoco Text1(5)
End Sub


Private Sub BotonEliminar()
Dim SQL As String

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
    SQL = SQL & "¿Seguro que desea eliminar la frecuencia?"
    SQL = SQL & vbCrLf & "Cliente: " & Format(Text1(0).Text, "000") & " " & Text2(0).Text
    SQL = SQL & vbCrLf & "Dpto : " & Text1(1).Text
    SQL = SQL & vbCrLf & "Exped : " & Text1(2).Text
    
    
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then
        'Hay que eliminar
        On Error GoTo Error2
        NumRegElim = Data1.Recordset.AbsolutePosition
        If Not Eliminar Then Exit Sub
        
        If SituarDataTrasEliminar(Data1, NumRegElim) Then
            PonerCampos
        Else
            LimpiarCampos
            PonerModo 0
        End If
    End If
    
Error2:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar Dirección", Err.Description
        Data1.Recordset.CancelUpdate
    End If
End Sub


Private Function Eliminar() As Boolean
Dim SQL As String
On Error GoTo FinEliminar
        
        SQL = " WHERE coddirec=" & Data1.Recordset!CodDirec
        SQL = SQL & " AND codclien =" & Data1.Recordset!codClien
        SQL = SQL & " AND numexped =" & DBSet(Data1.Recordset!numexped, "T")
        SQL = SQL & " AND numcanal =" & Data1.Recordset!numcanal
        SQL = SQL & " AND legalsno =" & Data1.Recordset!legalsno
        'Cabeceras
        conn.Execute "Delete  from " & NombreTabla & SQL
                      
FinEliminar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar"
        Eliminar = False
    Else
        Eliminar = True
    End If
End Function


Private Function DatosOk() As Boolean
Dim b As Boolean

    DatosOk = False
    b = CompForm(Me, 1)
    If Not b Then Exit Function
        
    If Text1(1).Text = "" Then
        MsgBox "El departamento tiene que tener valor(0 Si no tiene asignados)", vbExclamation
        Exit Function
    End If
    DatosOk = True
End Function


Private Sub MandaBusquedaPrevia(cadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim cad As String
Dim tabla As String
Dim Titulo As String

    'Llamamos a al form
    cad = ""
    'Estamos en Modo de Cabeceras
    'Registro de la tabla de cabeceras: slista
    cad = cad & ParaGrid(Text1(0), 25)
    cad = cad & ParaGrid(Text1(1), 25)
    cad = cad & ParaGrid(Text1(2), 25)
    cad = cad & ParaGrid(Text1(3), 25)
    tabla = "scafre"
    Titulo = "Frecuencias"
               
    If cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = cad
        frmB.vTabla = tabla
        frmB.vSQL = cadB
        HaDevueltoDatos = False
        '###A mano
        frmB.vDevuelve = "0|1|2|3|"
        frmB.vTitulo = Titulo
        frmB.vselElem = 0
        frmB.vCargaFrame = False
        
        frmB.vConexionGrid = conAri 'Conexion a BD Ariges
'        frmB.vBuscaPrevia = chkVistaPrevia
        '#
        frmB.Show vbModal
        Set frmB = Nothing
        'Si ha puesto valores y tenemos que es formulario de busqueda entonces
        'tendremos que cerrar el form lanzando el evento
        If HaDevueltoDatos Then
''            If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
''                cmdRegresar_Click
'        Else   'de ha devuelto datos, es decir NO ha devuelto datos
'            If Modo = 5 Then
'                PonerFoco txtAux(0)
'            Else
                PonerFoco Text1(kCampo)
'            End If
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub HacerBusqueda()
Dim cadB As String

    cadB = ObtenerBusqueda(Me, False)
    If chkVistaPrevia = 1 Then
        MandaBusquedaPrevia cadB
    ElseIf cadB <> "" Then 'Se muestran en el mismo form
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Sub PonerCadenaBusqueda()
    Screen.MousePointer = vbHourglass
    On Error GoTo EEPonerBusq

    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        If Modo = 1 Then
            MsgBox "No hay ningún registro en la tabla " & NombreTabla & " para ese criterio de Búsqueda.", vbInformation
        Else
            MsgBox "No hay ningún registro en la tabla " & NombreTabla & ".", vbInformation
        End If
        Screen.MousePointer = vbDefault
        PonerModo Modo
        PonerFoco Text1(0)
        Exit Sub
    Else
        PonerModo 2
        Data1.Recordset.MoveFirst
        PonerCampos
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
    
EEPonerBusq:
    MuestraError Err.Number, "PonerCadenaBusqueda"
    PonerModo 0
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerCampos()
On Error GoTo EPonerCampos

    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma Me, Data1
    
    'FALTA####
    'Este trozo deberia hacero en PonerCamposForma
    'Si el combo tiene valor NULL entonces deberia ponerlo a listindex=-1
    If IsNull(Data1.Recordset!propirep) Then Combo1(0).ListIndex = -1
    If IsNull(Data1.Recordset!propubic) Then Combo1(1).ListIndex = -1
    
    Modo = 3  'Para que el lostfcous ponga los no bres del cliente y/o departmento
    Text1_LostFocus 0
    If Text1(1).Text <> "" Then
        Text1_LostFocus 1
    Else
        Text1(1).Text = "0"
        Text2(1).Text = ""
    End If
     Modo = 2
    
    
    
    
    
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    PonerFoco Text1(5)
EPonerCampos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poniendo Campos", Err.Description
End Sub


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub


Private Sub CargarComboTipoDirec()
'### Combo Tipo Direccion
'Cargaremos el combo, o bien desde una tabla o con valores fijos o como
'se quiera, la cuestion es cargarlo

    For kCampo = 0 To 1
        Me.Combo1(kCampo).Clear
        Combo1(kCampo).AddItem "CLIENTE"
        Combo1(kCampo).ItemData(Combo1(kCampo).NewIndex) = 0

        Combo1(kCampo).AddItem "Propia"
        Combo1(kCampo).ItemData(Combo1(kCampo).NewIndex) = 1
    Next kCampo
End Sub


Private Sub PosicionarData()
Dim vWhere As String, Indicador As String

    vWhere = "codclien = " & Text1(0).Text & " and coddirec = " & Val(Text1(1).Text) & " and  numexped = '" & Text1(2).Text & "' and numcanal = " & Text1(4).Text & " and legalsno = " & Abs(Val(Check1.Value))
    If SituarDataMULTI(Data1, vWhere, Indicador) Then
    'If SituarData(Data1, vWhere, Indicador) Then
        PonerModo 2
        lblIndicador.Caption = Indicador
    Else
'        LimpiarCampos
        PonerModo 0
    End If
End Sub

Private Sub ComprobarGoogleEarth()
    On Error GoTo EComprobarGoogleEarth
    
    
    If GoogleMaps Then
        
        CadenaConsulta = "C:\Archivos de programa\Internet Explorer\iexplore.exe"
        NombreTabla = CadenaConsulta
        If Dir(CadenaConsulta, vbArchive) = "" Then
            CadenaConsulta = "C:\Program files\Internet Explorer\iexplore.exe"
            NombreTabla = CadenaConsulta
            If Dir(CadenaConsulta, vbArchive) = "" Then NombreTabla = ""
        End If
    
    Else
        'google earth
        CadenaConsulta = "C:\Archivos de programa\Google\Google Earth\GoogleEarth.exe"
        NombreTabla = CadenaConsulta
        If Dir(CadenaConsulta, vbArchive) = "" Then
            CadenaConsulta = "C:\Program files\Google\Google Earth\GoogleEarth.exe"
            NombreTabla = CadenaConsulta
            If Dir(CadenaConsulta, vbArchive) = "" Then NombreTabla = ""
        End If
    End If
    
    Exit Sub
EComprobarGoogleEarth:
    
        MuestraError Err.Number, "Comprobando carpeta(1)"
        NombreTabla = ""
End Sub



Private Sub ModificarExpediente()
Dim SQL As String

    If Modo <> 2 Then Exit Sub
    
    If Me.Data1.Recordset.EOF Then Exit Sub
    
    If vUsu.Nivel > 1 Then
        MsgBox "No tiene suficientes privilegios. Consulte al administrador del sistema. ", vbExclamation
        Exit Sub
    End If
    CadenaDesdeOtroForm = Text1(2).Text & "|" & Abs(Me.Check1.Value) & "|"
    frmListado2.Opcion = 24
    frmListado2.Show vbModal
    
    If CadenaDesdeOtroForm <> "" Then
        'OK. Ha actualizado
        Screen.MousePointer = vbHourglass
        SQL = RecuperaValor(CadenaDesdeOtroForm, 1)
        CadenaDesdeOtroForm = RecuperaValor(CadenaDesdeOtroForm, 2)
        CadenaConsulta = "numexped = " & DBSet(SQL, "T") & " AND legalsno = " & CadenaDesdeOtroForm
        SQL = "UPDATE scafre SET numexped = " & DBSet(SQL, "T") & ", legalsno = " & CadenaDesdeOtroForm
        SQL = SQL & " WHERE codclien = " & Data1.Recordset!codClien & " AND coddirec = " & Data1.Recordset!CodDirec
        SQL = SQL & " AND numexped = " & DBSet(Data1.Recordset!numexped, "T")
        SQL = SQL & " AND numcanal = " & Data1.Recordset!numcanal
        SQL = SQL & " AND legalsno = " & Data1.Recordset!legalsno
        
        
        If ejecutar(SQL, False) Then
            Espera 0.5
            
            SQL = " AND codclien = " & Data1.Recordset!codClien & " AND coddirec = " & Data1.Recordset!CodDirec
            'SQL = SQL & " AND numexped = " & DBSet(Data1.Recordset!numexped, "T") esta en el update
            SQL = SQL & " AND numcanal = " & Data1.Recordset!numcanal
            'SQL = SQL & " AND legalsno = " & Data1.Recordset!legalsno esta tb en el update
            SQL = CadenaConsulta & SQL
            
            CadenaConsulta = "select * from " & NombreTabla & " WHERE " & SQL & " " & Ordenacion
            PonerCadenaBusqueda
        End If
        CadenaConsulta = ""
    End If
    
    
End Sub
