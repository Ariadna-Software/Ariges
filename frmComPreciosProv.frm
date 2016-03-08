VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmComPreciosProv2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Precios Proveedor"
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   8160
   ClipControls    =   0   'False
   Icon            =   "frmComPreciosProv.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   8160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Exposición"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00972E0B&
      Height          =   735
      Left            =   5160
      TabIndex        =   44
      Top             =   5400
      Width           =   2895
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   12
         Left            =   1320
         MaxLength       =   13
         TabIndex        =   16
         Tag             =   "Precio|N|S|0|999999.0000|slispr|precioexp|###,##0.0000|N|"
         Text            =   "Text1"
         Top             =   255
         Width           =   1350
      End
      Begin VB.Label Label4 
         Caption         =   "Precio"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   45
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1575
      Left            =   120
      TabIndex        =   37
      Top             =   440
      Width           =   7935
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   14
         Left            =   3000
         MaxLength       =   40
         TabIndex        =   3
         Tag             =   "R|T|S|||slispr|descripprov|||"
         Text            =   "Text1"
         Top             =   1200
         Width           =   4215
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   13
         Left            =   1320
         MaxLength       =   16
         TabIndex        =   2
         Tag             =   "R|T|S|||slispr|referprov|||"
         Text            =   "Text1"
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   0
         Left            =   1320
         MaxLength       =   6
         TabIndex        =   1
         Tag             =   "Cod. Proveedor|N|N|0|999999|slispr|codprove|000000|S|"
         Text            =   "Text1"
         Top             =   675
         Width           =   1005
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   1
         Left            =   1320
         MaxLength       =   16
         TabIndex        =   0
         Tag             =   "Cod. Artículo|T|N|||slispr|codartic||S|"
         Text            =   "Text1"
         Top             =   180
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   1
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   39
         Text            =   "Text2"
         Top             =   180
         Width           =   4005
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   0
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   38
         Text            =   "Text2"
         Top             =   675
         Width           =   4785
      End
      Begin VB.Label Label3 
         Caption         =   "Ref. Proveedor"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   46
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   960
         Picture         =   "frmComPreciosProv.frx":000C
         ToolTipText     =   "Buscar proveedor"
         Top             =   675
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Proveedor"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   41
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Artículo"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   180
         Width           =   735
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   960
         Picture         =   "frmComPreciosProv.frx":010E
         ToolTipText     =   "Buscar artículo"
         Top             =   180
         Width           =   240
      End
   End
   Begin VB.Frame FrameOtros 
      Caption         =   "Otros"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00972E0B&
      Height          =   1160
      Left            =   5160
      TabIndex        =   34
      Top             =   2160
      Width           =   2895
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   6
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   11
         Tag             =   "Cantidad Minima|N|S|0|999999.00|slispr|cantmini|###,##0.00|N|"
         Text            =   "Text1"
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   5
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   10
         Tag             =   "Cantidad Fija|N|S|0|999999.00|slispr|cantfija|###,##0.00|N|"
         Text            =   "123456.25"
         Top             =   250
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Cant. Minima"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   36
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Cant. Fija"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   35
         Top             =   250
         Width           =   975
      End
   End
   Begin VB.Frame FramePromo 
      Caption         =   "Promoción"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00972E0B&
      Height          =   1815
      Left            =   5160
      TabIndex        =   28
      Top             =   3480
      Width           =   2895
      Begin VB.CheckBox chkPermiteDto 
         Caption         =   "Permite Dto."
         Enabled         =   0   'False
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   15
         Tag             =   "Permite Descuento|N|N|||slispr|dtoperm1||N|"
         Top             =   1485
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   8
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   13
         Tag             =   "Fecha Fin|F|S|||slispr|fechafin|dd/mm/yyyy|N|"
         Text            =   "Text1"
         Top             =   692
         Width           =   1350
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   7
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   12
         Tag             =   "Fecha Inicio|F|S|||slispr|fechaini|dd/mm/yyyy|N|"
         Text            =   "Text1"
         Top             =   320
         Width           =   1350
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   9
         Left            =   1320
         MaxLength       =   13
         TabIndex        =   14
         Tag             =   "Precio|N|S|0|999999.0000|slispr|preciopr|###,##0.0000|N|"
         Text            =   "Text1"
         Top             =   1080
         Width           =   1350
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   2
         Left            =   1050
         Picture         =   "frmComPreciosProv.frx":0210
         ToolTipText     =   "Buscar fecha"
         Top             =   692
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha Fin"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   33
         Top             =   692
         Width           =   855
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   1
         Left            =   1050
         Picture         =   "frmComPreciosProv.frx":029B
         ToolTipText     =   "Buscar fecha"
         Top             =   320
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha Inicio"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   32
         Top             =   320
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Precio"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   29
         Top             =   1065
         Width           =   615
      End
   End
   Begin VB.Frame FrameActuales 
      Caption         =   "Valores Actuales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00972E0B&
      Height          =   1560
      Left            =   120
      TabIndex        =   26
      Top             =   2160
      Width           =   4935
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   11
         Left            =   2160
         MaxLength       =   5
         TabIndex        =   7
         Tag             =   "Descuento 2|N|S|0|99.00|slispr|dtoline2|#0.00|N|"
         Text            =   "Text1"
         Top             =   670
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   10
         Left            =   740
         MaxLength       =   5
         TabIndex        =   6
         Tag             =   "Descuento 1|N|S|0|99.00|slispr|dtoline1|#0.00|N|"
         Text            =   "Text1"
         Top             =   670
         Width           =   735
      End
      Begin VB.CheckBox chkPermiteDto 
         Caption         =   "Permite Dto."
         Enabled         =   0   'False
         Height          =   195
         Index           =   0
         Left            =   2880
         TabIndex        =   5
         Tag             =   "Permite Descuento|N|N|||slispr|dtopermi||N|"
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   3
         Left            =   3720
         MaxLength       =   10
         TabIndex        =   9
         Tag             =   "Fecha Cambio|F|S|||slispr|fechanue|dd/mm/yyyy|N|"
         Text            =   "25/12/2004"
         Top             =   1090
         Width           =   1050
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   4
         Left            =   840
         MaxLength       =   13
         TabIndex        =   8
         Tag             =   "Precio Nuevo|N|S|0|999999.0000|slispr|precionu|###,##0.0000|N|"
         Text            =   "Text1"
         Top             =   1090
         Width           =   1200
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   2
         Left            =   740
         MaxLength       =   13
         TabIndex        =   4
         Tag             =   "Precio Actual|N|N|0|999999.0000|slispr|precioac|###,##0.0000|N|"
         Text            =   "Text1"
         Top             =   250
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "Dto 2"
         Height          =   255
         Left            =   1680
         TabIndex        =   43
         Top             =   670
         Width           =   495
      End
      Begin VB.Label Label7 
         Caption         =   "Dto 1"
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   670
         Width           =   495
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   3435
         Picture         =   "frmComPreciosProv.frx":0326
         ToolTipText     =   "Buscar fecha"
         Top             =   1095
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha Cambio"
         Height          =   255
         Index           =   0
         Left            =   2280
         TabIndex        =   31
         Top             =   1155
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Precio Nuevo"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   30
         Top             =   1035
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Precio"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   255
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   5760
      TabIndex        =   17
      Top             =   6480
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   6915
      TabIndex        =   18
      Top             =   6480
      Width           =   1035
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   6915
      TabIndex        =   19
      Top             =   6480
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Index           =   0
      Left            =   360
      TabIndex        =   24
      Top             =   6315
      Width           =   2655
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   240
         Left            =   240
         TabIndex        =   25
         Top             =   180
         Width           =   2115
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   8160
      _ExtentX        =   14393
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   22
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
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cambio de precio masivo"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Último"
         EndProperty
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   315
         Left            =   5880
         TabIndex        =   23
         Top             =   0
         Width           =   1215
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmComPreciosProv.frx":03B1
      Height          =   2295
      Left            =   120
      TabIndex        =   20
      Top             =   3960
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   4048
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   3240
      Top             =   6480
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
   Begin MSAdodcLib.Adodc Data2 
      Height          =   330
      Left            =   3240
      Top             =   5280
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
      TabIndex        =   22
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
Attribute VB_Name = "frmComPreciosProv2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public NuevoDato As String 'Por si esta insertando algun articulo y viene aqui

Private WithEvents frmB As frmBuscaGrid 'Form para busquedas (frmBuscaGrid)
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmF As frmCal 'Calendario de Fechas
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmA As frmAlmArticu2  'Form Mantenimiento Articulos
Attribute frmA.VB_VarHelpID = -1
Private WithEvents frmP As frmComProveedores 'Form MantenimientoProveedores
Attribute frmP.VB_VarHelpID = -1

Dim NombreTabla As String 'Nombre tabla Cabecera
Dim NombreTablaLin As String 'Nombre tabla Lineas
Dim PrimeraVez As Boolean
Dim Ordenacion As String
Private Modo As Byte
Private ModoAnterior As Byte
Dim kCampo As Integer

Dim btnPrimero As Byte
'Variable que indica el número del Boton  PrimerRegistro en la Toolbar1

Dim CadenaConsulta As String

Private HaDevueltoDatos As Boolean


'===========================================================================

Private Sub chkPermiteDto_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub chkPermiteDto_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub chkVistaPrevia_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub cmdAceptar_Click()
Dim b As Boolean
    On Error GoTo Error1
    Screen.MousePointer = vbHourglass
    
    Select Case Modo
        Case 1 'BUSQUEDA
            HacerBusqueda
        Case 3 'INSERTAR
            If DatosOk Then
                If InsertarDesdeForm(Me) Then PosicionarData
            End If
        Case 4 'MODIFICAR
            If DatosOk Then
                    
                conn.BeginTrans
                b = ModificarRegistro
                
                If b Then
                     conn.CommitTrans
                     TerminaBloquear
                     PosicionarData
                Else
                    conn.RollbackTrans
                End If
            End If
    End Select
    Screen.MousePointer = vbDefault
Error1:
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
            PonerFoco Text1(0)
    End Select
    
ECancelar:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub Form_Activate()
    If PrimeraVez Then
       PrimeraVez = False
       If Me.NuevoDato <> "" Then BotonAnyadir
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
    PrimeraVez = True
    'Icono del formulario
    Me.Icon = frmPpal.Icon
    
    'ICONOS de La toolbar
    btnPrimero = 19 'Posicion del Boton Primero en la toolbar (+ 3 siguientes)
    With Toolbar1
        .ImageList = frmPpal.imgListComun
        'ASignamos botones
        .Buttons(1).Image = 1   'Buscar
        .Buttons(2).Image = 2 'Ver Todos
        .Buttons(5).Image = 3 'Añadir
        .Buttons(6).Image = 4 'Modificar
        .Buttons(7).Image = 5 'Eliminar
        .Buttons(10).Image = 21 'Para cambiar precios
        .Buttons(15).Image = 16 'Imprimir
        .Buttons(16).Image = 15 'Salir
        .Buttons(btnPrimero).Image = 6 'Primero
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
        .Buttons(btnPrimero + 3).Image = 9 'Ultimo
    End With
    
    LimpiarCampos   'Limpia los campos TextBox
    DataGrid1.ClearFields
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    NombreTabla = "slispr" 'Tabla Cabecera Precios Proveedor
    NombreTablaLin = "slisp1" 'Tabla Lineas Precios Proveedor
    Ordenacion = " ORDER BY codartic, codprove "
    CadenaConsulta = "Select * from " & NombreTabla & " WHERE codartic = -1" 'No recupera datos
    
    Data1.ConnectionString = conn
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    
    PonerModo 0
    CargaGrid (Modo = 2)
End Sub


Private Sub CargaGrid(enlaza As Boolean)
Dim b As Boolean
Dim I As Byte
Dim SQL As String

    On Error GoTo ECarga

    b = DataGrid1.Enabled
    
    SQL = MontaSQLCarga(enlaza)
    CargaGridGnral DataGrid1, Me.Data2, SQL, False
    
    DataGrid1.Columns(0).visible = False 'Cod. Articulo
    DataGrid1.Columns(1).visible = False 'Cod. Proveedor
    DataGrid1.Columns(2).visible = False 'Numero linea
    I = 2
       
    'Fecha Cambio
    DataGrid1.Columns(I + 1).Caption = "Fecha Cambio"
    DataGrid1.Columns(I + 1).Width = 1600
    
    'Precio Unidad
    DataGrid1.Columns(I + 2).Caption = "Precio"
    DataGrid1.Columns(I + 2).Width = 1800
    DataGrid1.Columns(I + 2).Alignment = dbgRight
    DataGrid1.Columns(I + 2).NumberFormat = FormatoPrecio
       
    
    For I = 0 To DataGrid1.Columns.Count - 1
        DataGrid1.Columns(I).AllowSizing = False
    Next I
    DataGrid1.Enabled = b
    DataGrid1.ScrollBars = dbgAutomatic
    
ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGrid1.Tag, Err.Description
End Sub


Private Sub frmA_DatoSeleccionado(CadenaSeleccion As String)
'Formulario Mantenimiento Articulos
    Text1(1).Text = RecuperaValor(CadenaSeleccion, 1)
    Text2(1).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub frmB_Selecionado(CadenaDevuelta As String)
'Formulario para Busqueda
Dim cadB As String
Dim Aux As String
      
    If CadenaDevuelta <> "" Then
            HaDevueltoDatos = True
            Screen.MousePointer = vbHourglass
            
            'Estamos en Cabecera
            'Recupera todo el registro de Tarifas de Precios
            'Sabemos que campos son los que nos devuelve
            'Creamos una cadena consulta y ponemos los datos
            cadB = ""
            Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
            cadB = Aux
            Aux = ValorDevueltoFormGrid(Text1(1), CadenaDevuelta, 2)
            cadB = cadB & " and " & Aux
            CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
            PonerCadenaBusqueda
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub frmF_Selec(vFecha As Date)
'Calendario de Fecha
Dim indice As Byte
    Select Case Me.imgFecha(0).Tag
        Case 0: indice = 3
        Case 1: indice = 7
        Case 2: indice = 8
    End Select
    Text1(indice).Text = Format(vFecha, "dd/mm/yyyy")
End Sub


Private Sub frmP_DatoSeleccionado(CadenaSeleccion As String)
'Formulario Mantenimiento Proveedores
    Text1(0).Text = RecuperaValor(CadenaSeleccion, 1)
    FormateaCampo Text1(0)
    Text2(0).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub imgBuscar_Click(Index As Integer)
    If Modo = 2 Or Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass
   
    Select Case Index
        Case 0  'Cod. Proveedor
            Set frmP = New frmComProveedores
            frmP.DatosADevolverBusqueda = "0"
            frmP.Show vbModal
            Set frmP = Nothing
        Case 1 'Codigo Articulo
            Set frmA = New frmAlmArticu2
            'frmA.DatosADevolverBusqueda3 = "@1@" 'Abre en Modo Busqueda
            frmA.DesdeTPV = False
            frmA.Show vbModal
            Set frmA = Nothing
    End Select
    
    PonerFoco Text1(Index)
    Screen.MousePointer = vbDefault
End Sub


Private Sub imgFecha_Click(Index As Integer)
Dim indice As Byte

   If Modo = 2 Or Modo = 0 Then Exit Sub
   
   Screen.MousePointer = vbHourglass
   Set frmF = New frmCal
   frmF.Fecha = Now
   
   Me.imgFecha(0).Tag = Index
   Select Case Index
    Case 0: indice = 3
    Case 1: indice = 7
    Case 2: indice = 8
   End Select
   
   PonerFormatoFecha Text1(indice)
   If Text1(indice).Text <> "" Then frmF.Fecha = CDate(Text1(indice).Text)
   
   Screen.MousePointer = vbDefault
   frmF.Show vbModal
   Set frmF = Nothing
   
   PonerFoco Text1(indice)
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
    ConseguirFoco Text1(Index), Modo
End Sub


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
     KEYpress KeyAscii
End Sub


Private Sub Text1_LostFocus(Index As Integer)

    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub

    Select Case Index
        Case 0 'Codigo Proveedor
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "sprove", "nomprove")
            Else
                Text2(Index).Text = ""
            End If

        Case 1 'Codigo Articulo
            Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "sartic", "nomartic")
            If Modo = 3 Then
                If vParamAplic.NumeroInstalacion = 4 And Text1(14).Text = "" Then Text1(14).Text = Text2(Index).Text
                
            End If
        Case 2, 4, 9, 12 'Precios Actuales y Nuevos y exposicion
            'Formato tipo 2: Decimal(10,4)
            If Text1(Index).Text <> "" Then PonerFormatoDecimal Text1(Index), 2
        
        Case 5, 6 'cantidades Decimal(8,2)
            If Text1(Index).Text <> "" Then PonerFormatoDecimal Text1(Index), 6
            BloquearTxt Text1(5), (Text1(6).Text <> "")
            BloquearTxt Text1(6), (Text1(5).Text <> "")
            
        Case 3, 7, 8 'Fecha Cambio
            If Text1(Index).Text <> "" Then PonerFormatoFecha Text1(Index)
            
        Case 10, 11 'descuentos
            'Formato tipo 4: Decimal(4,2)
            If Text1(Index).Text <> "" Then PonerFormatoDecimal Text1(Index), 4
    End Select
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
        Case 10
            If Modo = 2 Or Modo = 0 Then
            
                frmListado2.Opcion = 44
                frmListado2.Show vbModal
            End If
            
        Case 15 'Imprimir
            AbrirListado (309) '309: Informe Precios Compras
        Case 16  'Salir
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
Dim I As Byte
Dim b As Boolean
Dim NumReg As Byte

    Modo = Kmodo
    PonerIndicador Me.lblIndicador, Modo
    
    '===========================================
    'Modo 2. Hay datos y estamos visualizandolos
    b = (Kmodo = 2)
    'Visualizar flechas de desplazamiento en la toolbar si modo=2
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, NumReg
          
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1 y bloquea clave primaria
    BloquearText1 Me, Modo
    
    If Modo = 4 Then 'Modificar
        BloquearTxt Text1(5), (Text1(6).Text <> "")
        BloquearTxt Text1(6), (Text1(5).Text <> "")
        
        
        'Permitiremos cambiar el proveedor, de momento para euler
        If vParamAplic.NumeroInstalacion = 4 Then BloquearTxt Text1(0), False
        
    End If
    
    'Modo Insertar
    If Kmodo = 3 Then Me.chkPermiteDto(0).Value = 1
    Me.chkPermiteDto(0).Enabled = (Modo = 3) Or (Modo = 4) 'Insert o Modificar
    Me.chkPermiteDto(1).Enabled = (Modo = 3) Or (Modo = 4) 'Insert o Modificar
    
    '==============================
    b = Modo <> 0 And Modo <> 2
    cmdCancelar.visible = b
    cmdAceptar.visible = b
    
    
    Me.imgBuscar(1).Enabled = Modo = 1 Or Modo = 3 'Si modificar no activado pq son claves ajenas
    Me.imgBuscar(0).Enabled = Modo = 1 Or Modo >= 3 'Si modificar no activado pq son claves ajenas
    
    
    For I = 0 To Me.imgFecha.Count - 1
        Me.imgFecha(I).Enabled = b
    Next I
    
    'Poner el tamaño de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCamposGnral Me, Modo, 1
    
    PonerModoOpcionesMenu 'Activar opciones de menu según modo
    PonerOpcionesMenu   'Activar opciones de menu según nivel
                        'de permisos del usuario
End Sub

'Private Sub PonerLongCampos()
''Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
''para los campos que permitan introducir criterios más largos del tamaño del campo
'    PonerLongCamposGnral Me, Modo, 1
'End Sub


Private Sub PonerModoOpcionesMenu()
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
    
    
    '===============================
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
    Me.chkPermiteDto(0).Value = 0
    Me.chkPermiteDto(1).Value = 0
End Sub


Private Sub Desplazamiento(Index As Integer)
'Botones de Desplazamiento de la Toolbar
'Para desplazarse por los registros de control Data
    DesplazamientoData Data1, Index
    PonerCampos
End Sub


Private Function MontaSQLCarga(enlaza As Boolean) As String
'--------------------------------------------------------------------
' MontaSQlCarga:
'   Basándose en la información proporcionada por el vector de campos
'   crea un SQl para ejecutar una consulta sobre la base de datos que los
'   devuelva.
' Si ENLAZA -> Enlaza con el data1
'           -> Si no lo cargamos sin enlazar a ningun campo
'--------------------------------------------------------------------
Dim SQL As String
Dim Tabla As String
    
    Tabla = "slisp1" 'Tabla de lineas
    SQL = "SELECT * FROM " & Tabla
    
    If enlaza Then
        SQL = SQL & " WHERE codartic=" & DBSet(Data1.Recordset!codartic, "T") & " AND codprove=" & Data1.Recordset!Codprove
    Else
        SQL = SQL & " WHERE codprove = -1"
    End If
    
    SQL = SQL & " ORDER BY " & Tabla & ".numlinea desc"
    MontaSQLCarga = SQL
End Function


Private Sub BotonBuscar()
    If Modo <> 1 Then
        LimpiarCampos
        PonerModo 1
        'Ponemos el grid lineasfacturas enlazando a ningun sitio
        CargaGrid False

        'Si pasamos el control aqui lo ponemos en amarillo
        PonerFoco Text1(1)
        Text1(1).BackColor = vbYellow
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
    'Ponemos el grid lineasfacturas enlazando a ningun sitio
    CargaGrid False
    
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
           
    'Ponemos el grid de lineas enlazando a ningun sitio
    CargaGrid False
    If Me.NuevoDato = "" Then
        PonerFoco Text1(1)
    Else
        Text1(1).Text = RecuperaValor(NuevoDato, 1)
        Text2(1).Text = RecuperaValor(NuevoDato, 2)
        Text1(0).Text = RecuperaValor(NuevoDato, 3)
        Text2(0).Text = RecuperaValor(NuevoDato, 4)
        
        If vParamAplic.NumeroInstalacion = 4 Then
            Text1(14).Text = Text2(1)
            PonerFoco Text1(13)
        Else
            PonerFoco Text1(2)
        End If
    End If
End Sub


Private Sub BotonModificar()
    If Data1.Recordset.EOF Then Exit Sub
    
    'Escondemos el navegador y ponemos Modo Modificar
    PonerModo 4
    PonerFoco Text1(13)
End Sub


Private Sub BotonEliminar()
Dim SQL As String

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
    SQL = "Precios Proveedor." & vbCrLf
    SQL = SQL & "--------------------------" & vbCrLf & vbCrLf
    
    SQL = SQL & "Va a Eliminar El Precio de Proveedor:"
    SQL = SQL & vbCrLf & "Proveedor : " & Text1(0).Text & " - " & Text2(0).Text
    SQL = SQL & vbCrLf & "Articulo : " & Text1(1).Text & " - " & Text2(1).Text
    
    SQL = SQL & vbCrLf & vbCrLf & "¿Desea continuar ? "
    
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then
        'Hay que eliminar
        On Error GoTo Error2
        
        NumRegElim = Data1.Recordset.AbsolutePosition
        If Not Eliminar Then Exit Sub
        
        If SituarDataTrasEliminar(Data1, NumRegElim) Then
            PonerCampos
        Else
            LimpiarCampos
            CargaGrid False
            PonerModo 0
        End If
    End If
    
Error2:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar Precio Proveedor", Err.Description
        Data1.Recordset.CancelUpdate
    End If
End Sub


Private Function Eliminar() As Boolean
Dim SQL As String
    
    On Error GoTo FinEliminar
        
    conn.BeginTrans
    SQL = " WHERE codartic=" & DBSet(Data1.Recordset!codartic, "T")
    SQL = SQL & " AND codprove=" & Val(Data1.Recordset!Codprove)
    
    'Lineas
    conn.Execute "Delete  from " & NombreTablaLin & SQL
    
    'Cabeceras
    conn.Execute "Delete  from " & NombreTabla & SQL
                      
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


Private Function DatosOk() As Boolean
Dim b As Boolean
Dim Aux As String

    On Error Resume Next

    DatosOk = False
    b = CompForm(Me, 1)
    If Not b Then Exit Function
    
    'Comprobar que si hay valores nuevos, la fecha de cambio no es nulo
    If (Not EsVacio(Text1(4))) Then b = (Not EsVacio(Text1(3)))
    
    If Not b Then
        MsgBox "La Fecha de Cambio debe tener valor.", vbInformation
        Exit Function
    End If
    
    'Comprobar que si no hay valores nuevos no haya fecha de Cambio
    If EsVacio(Text1(4)) Then b = (EsVacio(Text1(3)))
    
    If Not b Then
        MsgBox "No hay precio nuevo para la fecha de cambio", vbInformation
        Exit Function
    End If
    
    'Febrero 2014
    '--------------
    ' Si da de alta el articulo-proveedor para el articulo, el proveedor es el mismo que en la ficha del articulo,
    'entoces verificamos la REFERENCIA
    If Modo = 3 Then
        CadenaConsulta = "referprov"
        Aux = DevuelveDesdeBD(conAri, "codprove", "sartic", "codartic", Text1(1).Text, "T", CadenaConsulta)
        If Aux <> "" And CadenaConsulta <> "" Then
            If Val(Aux) = Val(Text1(0).Text) Then
                'OK. Mismo articulo, mismo proveedor
                If Text1(13).Text <> CadenaConsulta Then
                    Aux = "Referencia proveedor: " & vbCrLf & vbCrLf & "Ficha articulo: " & CadenaConsulta
                    Aux = Aux & vbCrLf & "Precios prov: " & Text1(13).Text
                    Aux = Aux & vbCrLf & vbCrLf & vbCrLf & "¿Desea que se guarde la de la ficha del articulo como referencia?"
                    Select Case MsgBox(Aux, vbQuestion + vbYesNoCancel)
                    Case vbYes
                        Text1(13).Text = CadenaConsulta
                    Case vbCancel
                        b = False
                    Case Else
                        'NADA
                    End Select
                End If
            End If
        End If
        CadenaConsulta = Data1.RecordSource
    End If
    
    
    If b And Modo = 4 And vParamAplic.NumeroInstalacion = 4 Then
        'EULER. Puede cambiar el codprove
        If Val(Data1.Recordset!Codprove) <> Val(Text1(0).Text) Then
            'Veamos que no existe un articulo-proveedor , ya que lo han cambiado
            Aux = "codprove = " & Text1(0).Text & " AND codartic"
            Aux = DevuelveDesdeBD(conAri, "codprove", "slispr", Aux, Text1(1).Text, "T")
            If Aux <> "" Then
                MsgBox "Ya existe un precio para el ese articulo con el proveedor " & Text1(0).Text, vbExclamation
                b = False
            End If
        End If
    End If
    DatosOk = b
End Function



Private Sub MandaBusquedaPrevia(cadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim Cad As String
Dim Tabla As String
Dim Titulo As String

    'Llamamos a al form
    Cad = ""
    'Estamos en Modo de Cabeceras
    'Registro de la tabla de cabeceras: slista
    Cad = Cad & ParaGrid(Text1(0), 9, "Prov.")
    Cad = Cad & "Nombre Prov.|sprove|nomprove|T||33·"
    Cad = Cad & ParaGrid(Text1(1), 20, "Articulo")
    Cad = Cad & "Desc. Artic|sartic|nomartic|T||38·"
    
    Tabla = "(" & NombreTabla & " LEFT JOIN sprove ON " & NombreTabla & ".codprove=sprove.codprove" & ")"
    Tabla = Tabla & " LEFT JOIN sartic ON " & NombreTabla & ".codartic=sartic.codartic"
    
    Titulo = "Precios Proveedor"
           
    If Cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = Cad
        frmB.vTabla = Tabla
        frmB.vSQL = cadB
        HaDevueltoDatos = False
        '###A mano
        frmB.vDevuelve = "0|2|"
        frmB.vTitulo = Titulo
        frmB.vselElem = 0
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
    ElseIf cadB <> "" Then
        'Se muestran en el mismo form
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Sub PonerCadenaBusqueda()
On Error GoTo EEPonerBusq
    
    Screen.MousePointer = vbHourglass

    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    
    If Data1.Recordset.RecordCount <= 0 Then
        MsgBox "No hay ningún registro en la tabla " & NombreTabla & " para ese criterio de Búsqueda.", vbInformation
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
    'Poner el nombre del cod. cliente
    Text2(0).Text = PonerNombreDeCod(Text1(0), 1, "sprove", "nomprove")
    'Poner el nombre del cod. Articulo
    Text2(1).Text = PonerNombreDeCod(Text1(1), 1, "sartic", "nomartic")
    
    'Si los campos de precios nuevos son cero mostrar cadena vacia
    If Text1(2).Text <> "" Then
        If Text1(2).Text = 0 Then Text1(2).Text = ""
    End If
    If Text1(4).Text <> "" Then
        If Text1(4).Text = 0 Then Text1(4).Text = ""
    End If
    
    CargaGrid True
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
EPonerCampos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poniendo Campos", Err.Description
End Sub


Private Sub BotonActualizar()
'Actualizar Precios Especiales
Dim SQL As String

    If Data1.Recordset.EOF Then
        MsgBox "Ningún Precio Especial para actualizar.", vbExclamation
        Exit Sub
    End If
    
    If Data2 Is Nothing Then Exit Sub
   
    SQL = "Actualización Precios Especiales de Artículos." & vbCrLf
    SQL = SQL & "---------------------------------------------" & vbCrLf & vbCrLf
    
    SQL = SQL & "Va a Actualizar el Precio Especial para:"
    SQL = SQL & vbCrLf & " Cod. Clien. :  " & CStr(Format(Data1.Recordset.Fields(0), "000000"))
    SQL = SQL & vbCrLf & " Cod. Artic. :  " & Data1.Recordset.Fields(1)
    SQL = SQL & vbCrLf & vbCrLf & " ¿Desea continuar ? "
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    
    NumRegElim = Data1.Recordset.AbsolutePosition
    If ActualizarPreEspecial Then
        SituarDataTrasEliminar Data1, NumRegElim
    End If
End Sub


Private Function ActualizarPreEspecial() As Boolean
'Actualiza los Precios Especiales insertando los precios actuales con la fecha de cambio en el hostórico
' y modificando el la tabla de precios especiales pasando los valores nuevos a ser los actuales.
Dim Donde As String
Dim bol As Boolean
On Error GoTo EActualizarPreEspecial
    
   
    'Aqui empieza transaccion
    conn.BeginTrans
    bol = ActualizarElPrecio(Donde)

EActualizarPreEspecial:
        If Err.Number <> 0 Then
            Donde = "Actualizar Precio Especial." & vbCrLf & "----------------------------" & vbCrLf & Donde
            MuestraError Err.Number, Donde, Err.Description
            bol = False
        End If
        If bol Then
            conn.CommitTrans
            ActualizarPreEspecial = True
        Else
            conn.RollbackTrans
            ActualizarPreEspecial = False
        End If
End Function


Private Function ActualizarElPrecio(ByRef ADonde As String) As Boolean

    ActualizarElPrecio = False
    
     
    'Insertamos en lineas Historico
    ADonde = "Insertando datos en Historico lineas Precios Especiales"
    If Not InsertarLineasHistorico Then Exit Function
'    IncrementarProgres 2
    
    
    'Modificamos en cabeceras de Tarifas
    ADonde = "Modificando datos en cabecera de Precios Especiales"
    If Not ModificarCabecera Then Exit Function
'    IncrementarProgres 2
    ActualizarElPrecio = True
End Function


Private Function ModificarCabecera() As Boolean
'Modifica la tabla de cabeceras de Tarifas
Dim SQL As String

    On Error GoTo ErrModCab

    SQL = "UPDATE " & NombreTabla & " SET precioac=precionu, precioa1=precion1, dtoespec=dtoespe1, fechanue=null, precionu=0, precion1=0"
    SQL = SQL & " WHERE codclien=" & Data1.Recordset!codclien & " AND codartic=" & DBSet(Data1.Recordset!codartic, "T")
   
    conn.Execute SQL
    ModificarCabecera = True
    Exit Function
    
ErrModCab:
'    If Err.Number <> 0 Then
         'Hay error , almacenamos y salimos
        ModificarCabecera = False
'    Else
'        ModificarCabecera = True
'    End If
End Function


Private Function InsertarLineasHistorico() As Boolean
Dim SQL As String
Dim NumF As String

    On Error GoTo ErrInsLin

    'Obtenemos la siguiente numero de linea de tarifa
    SQL = "codclien=" & Data1.Recordset!codclien & " AND codartic=" & DBSet(Data1.Recordset!codartic, "T")
    NumF = SugerirCodigoSiguienteStr("spree1", "numlinea", SQL)

    SQL = "INSERT INTO spree1 (codclien, codartic, numlinea, fechanue, precioac, precioa1, dtoespec)"
    SQL = SQL & " VALUES (" & Data1.Recordset.Fields(0).Value & ", " & DBSet(Data1.Recordset.Fields(1).Value, "T") & ", "
    SQL = SQL & NumF & ", " & DBSet(Text1(4).Text, "F") & ", "
    SQL = SQL & DBSet(Data1.Recordset!precioac, "N") & ", " & DBSet(Data1.Recordset!precioa1, "N") & ", "
    SQL = SQL & DBSet(Data1.Recordset!dtoespec, "N") & ") "
    conn.Execute SQL
    
    InsertarLineasHistorico = True
    Exit Function
    
ErrInsLin:
'    If Err.Number <> 0 Then
'        'Hay error , almacenamos y salimos
        InsertarLineasHistorico = False
'    Else
'        InsertarLineasHistorico = True
'    End If
End Function


Private Sub BotonImprimir()
        frmListado.NumCod = Text1(0).Text
        AbrirListado (8) '8: Informe Movimientos Almacen
End Sub


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub


Private Sub PosicionarData()
Dim Indicador As String
Dim vWhere As String

    vWhere = "(codartic=" & DBSet(Text1(1).Text, "T") & " AND codprove=" & Text1(0).Text & ")"
    If SituarDataMULTI(Data1, vWhere, Indicador) Then
        PonerModo 2
        lblIndicador.Caption = Indicador
    Else
        PonerModo 0
    End If
End Sub





Private Function ModificarRegistro() As Boolean
Dim b As Boolean

    On Error GoTo eModificarRegistro

    ModificarRegistro = False

    b = ModificaDesdeFormulario(Me, 1)  'Esto actualizara los campos (sin codprove)
    
    
    If b And vParamAplic.NumeroInstalacion = 4 Then
        'EULER. Puede cambiar el codprove
        If Val(Data1.Recordset!Codprove) <> Val(Text1(0).Text) Then
            Me.Tag = "UPDATE @@@ SET codprove=" & Text1(0).Text & " WHERE codartic=" & DBSet(Data1.Recordset!codartic, "T") & " AND codprove=" & Data1.Recordset!Codprove & ";"
            
            conn.Execute "SET FOREIGN_KEY_CHECKS=0;"
            'Cabecera
            conn.Execute Replace(Me.Tag, "@@@", NombreTabla)
            conn.Execute Replace(Me.Tag, "@@@", NombreTablaLin)
            conn.Execute "SET FOREIGN_KEY_CHECKS=1;"
    
            b = True
    
    
            Me.Tag = ""
        End If
    End If
    ModificarRegistro = b
    Exit Function
eModificarRegistro:
    MuestraError Err.Number
    ejecutar "SET FOREIGN_KEY_CHECKS=1;", True  'Por si acaso lo ha dejado fuera
End Function
