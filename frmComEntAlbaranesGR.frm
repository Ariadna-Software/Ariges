VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmComEntAlbaranesGR 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   10920
   ClientLeft      =   45
   ClientTop       =   4035
   ClientWidth     =   15390
   Icon            =   "frmComEntAlbaranesGR.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10920
   ScaleWidth      =   15390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameDesplazamiento 
      Height          =   705
      Left            =   6600
      TabIndex        =   122
      Top             =   90
      Width           =   2415
      Begin MSComctlLib.Toolbar ToolbarDes 
         Height          =   330
         Left            =   240
         TabIndex        =   123
         Top             =   180
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Primero"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Anterior"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Siguiente"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "�ltimo"
               ImageIndex      =   9
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FrameBotonGnral2 
      Height          =   705
      Left            =   3780
      TabIndex        =   120
      Top             =   90
      Width           =   2715
      Begin MSComctlLib.Toolbar Toolbar5 
         Height          =   330
         Left            =   210
         TabIndex        =   121
         Top             =   180
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   5
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Cambiar Proveedor"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Mover a Hist�rico"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Recepci�n Factura"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "N�mero Series"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Imprimir Etiquetas Estanter�a"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   90
      TabIndex        =   118
      Top             =   90
      Width           =   3585
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   180
         TabIndex        =   119
         Top             =   180
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   10
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Nuevo"
               Object.Tag             =   "2"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Modificar"
               Object.Tag             =   "2"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Eliminar"
               Object.Tag             =   "2"
               Object.Width           =   1e-4
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Buscar"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Ver Todos"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Imprimir"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Object.ToolTipText     =   "Salir"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Style           =   3
            EndProperty
         EndProperty
      End
   End
   Begin VB.CheckBox chkVistaPrevia 
      Caption         =   "Vista previa"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   13770
      TabIndex        =   104
      Top             =   315
      Width           =   1530
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   285
      Index           =   0
      Left            =   9810
      MaxLength       =   15
      TabIndex        =   103
      Text            =   "BASE IMPONIBLE"
      Top             =   285
      Width           =   1920
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   50
      Left            =   11805
      MaxLength       =   15
      TabIndex        =   102
      Text            =   "Text1 7"
      Top             =   270
      Width           =   1845
   End
   Begin VB.TextBox txtAux2 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   9
      Left            =   3690
      Locked          =   -1  'True
      MaxLength       =   60
      TabIndex        =   97
      Text            =   "nom ccoste"
      Top             =   10440
      Visible         =   0   'False
      Width           =   6330
   End
   Begin VB.Frame Frame2 
      Height          =   1020
      Left            =   75
      TabIndex        =   80
      Top             =   870
      Width           =   15200
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   30
         Left            =   3210
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "Fecha entrada mercancia|F|N|||scaalp|fentrada|dd/mm/yyyy|N|"
         Text            =   "99/99/9999"
         Top             =   450
         Width           =   1290
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   9360
         MaxLength       =   30
         TabIndex        =   6
         Tag             =   "Cod. Proveedor|N|N|0|999999|scaalp|codprove|000000|S|"
         Text            =   "Text1"
         Top             =   540
         Width           =   825
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   5
         Left            =   10185
         MaxLength       =   40
         TabIndex        =   7
         Tag             =   "Nombre Proveedor|T|N|||scaalp|nomprove||N|"
         Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwwwwwwwwww aq"
         Top             =   540
         Width           =   4710
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   1635
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Fecha Albaran|F|N|||scaalp|fechaalb|dd/mm/yyyy|S|"
         Text            =   "99/99/9999"
         Top             =   450
         Width           =   1290
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   210
         MaxLength       =   10
         TabIndex        =   0
         Tag             =   "N� Albaran|T|N|0||scaalp|numalbar||S|"
         Text            =   "0000000000"
         Top             =   450
         Width           =   1305
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   9360
         MaxLength       =   30
         TabIndex        =   5
         Tag             =   "Realizada Por|N|N|0|9999|scaalp|codtraba|0000|N|"
         Text            =   "Text1"
         Top             =   180
         Width           =   825
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   10185
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   81
         Text            =   "Text2"
         Top             =   180
         Width           =   4710
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   4815
         MaxLength       =   7
         TabIndex        =   3
         Tag             =   "N� Pedido|N|S|0||scaalp|numpedpr|0000000|N|"
         Text            =   "Text1 7"
         Top             =   450
         Width           =   1155
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   20
         Left            =   6195
         MaxLength       =   10
         TabIndex        =   4
         Tag             =   "Fecha Pedido|F|S|||scaalp|fecpedpr|dd/mm/yyyy|N|"
         Top             =   450
         Width           =   1290
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   9090
         Picture         =   "frmComEntAlbaranesGR.frx":000C
         Tag             =   "-1"
         ToolTipText     =   "Buscar proveedor"
         Top             =   585
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   3
         Left            =   4230
         Picture         =   "frmComEntAlbaranesGR.frx":0A0E
         ToolTipText     =   "Buscar fecha"
         Top             =   165
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "F.Entrada"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   3210
         TabIndex        =   101
         ToolTipText     =   "Fecha entrada mercancia"
         Top             =   165
         Width           =   1035
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   9075
         ToolTipText     =   "Buscar trabajador"
         Top             =   195
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Proveedor"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   7695
         TabIndex        =   87
         Top             =   540
         Width           =   1185
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   2715
         Picture         =   "frmComEntAlbaranesGR.frx":0A99
         ToolTipText     =   "Buscar fecha"
         Top             =   180
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "N� Albaran"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   50
         Left            =   210
         TabIndex        =   85
         Top             =   165
         Width           =   1125
      End
      Begin VB.Label Label1 
         Caption         =   "Realizada Por"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   21
         Left            =   7695
         TabIndex        =   84
         Top             =   180
         Width           =   1320
      End
      Begin VB.Label Label1 
         Caption         =   "N� Pedido"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   4845
         TabIndex        =   83
         Top             =   165
         Width           =   1005
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Pedido"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   6180
         TabIndex        =   82
         Top             =   165
         Width           =   1365
      End
      Begin VB.Label Label1 
         Caption         =   "F.Albaran"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   14
         Left            =   1635
         TabIndex        =   86
         ToolTipText     =   "Fecha albaran"
         Top             =   165
         Width           =   990
      End
   End
   Begin VB.Frame Frame1 
      Height          =   475
      Index           =   0
      Left            =   75
      TabIndex        =   52
      Top             =   10350
      Width           =   2175
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   225
         TabIndex        =   53
         Top             =   135
         Width           =   1755
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   14220
      TabIndex        =   50
      Top             =   10425
      Width           =   1065
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13050
      TabIndex        =   49
      Top             =   10425
      Width           =   1065
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   495
      Top             =   5175
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
      Left            =   2835
      Top             =   4455
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   8280
      Left            =   90
      TabIndex        =   54
      Tag             =   "Fecha Oferta|F|N|||scapre|fecentre|dd/mm/yyyy|N|"
      Top             =   1980
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   14605
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Datos b�sicos"
      TabPicture(0)   =   "frmComEntAlbaranesGR.frx":0B24
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(2)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(35)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "imgBuscar(9)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "DataGrid1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtAux(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtAux(1)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtAux(3)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtAux(4)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtAux(5)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtAux(6)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtAux(7)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtAux(2)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmdAux(0)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cmdAux(1)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "FrameCliente"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Text2(17)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Text2(16)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtAux(8)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtAux(9)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "cmdAux(2)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "FrameToolAux0"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).ControlCount=   21
      TabCaption(1)   =   "Observaciones / Totales"
      TabPicture(1)   =   "frmComEntAlbaranesGR.frx":0B40
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Text1(29)"
      Tab(1).Control(1)=   "Text1(28)"
      Tab(1).Control(2)=   "FrameFactura"
      Tab(1).Control(3)=   "FrameHco"
      Tab(1).Control(4)=   "Text1(19)"
      Tab(1).Control(5)=   "Text1(18)"
      Tab(1).Control(6)=   "Text1(17)"
      Tab(1).Control(7)=   "Text1(16)"
      Tab(1).Control(8)=   "Text1(15)"
      Tab(1).Control(9)=   "Label1(48)"
      Tab(1).Control(10)=   "Label1(47)"
      Tab(1).Control(11)=   "Label1(45)"
      Tab(1).ControlCount=   12
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   29
         Left            =   -65730
         MaxLength       =   80
         TabIndex        =   23
         Tag             =   "T|T|S|||scaalp|SReferencia||N|"
         Top             =   945
         Width           =   5505
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   28
         Left            =   -73005
         MaxLength       =   80
         TabIndex        =   22
         Tag             =   "O|T|S|||scaalp|NReferencia||N|"
         Top             =   930
         Width           =   5505
      End
      Begin VB.Frame FrameToolAux0 
         Height          =   645
         Left            =   225
         TabIndex        =   130
         Top             =   2970
         Width           =   1500
         Begin MSComctlLib.Toolbar ToolAux 
            Height          =   330
            Index           =   0
            Left            =   150
            TabIndex        =   131
            Top             =   180
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   582
            ButtonWidth     =   609
            ButtonHeight    =   582
            AllowCustomize  =   0   'False
            Style           =   1
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   3
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Nuevo"
                  Object.Tag             =   "2"
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Modificar"
                  Object.Tag             =   "2"
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Eliminar"
                  Object.Tag             =   "2"
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame FrameFactura 
         Height          =   2400
         Left            =   -74640
         TabIndex        =   105
         Top             =   3795
         Width           =   14625
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   33
            Left            =   150
            MaxLength       =   15
            TabIndex        =   29
            Text            =   "Text1 7"
            Top             =   555
            Width           =   1665
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   34
            Left            =   2205
            MaxLength       =   15
            TabIndex        =   30
            Text            =   "Text1 7"
            Top             =   555
            Width           =   1590
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   35
            Left            =   4185
            MaxLength       =   15
            TabIndex        =   31
            Text            =   "Text1 7"
            Top             =   555
            Width           =   1635
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   36
            Left            =   6165
            MaxLength       =   15
            TabIndex        =   32
            Text            =   "Text1 7"
            Top             =   555
            Width           =   1665
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   43
            Left            =   10305
            MaxLength       =   15
            TabIndex        =   35
            Text            =   "Text1 7"
            Top             =   540
            Width           =   1800
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   37
            Left            =   8325
            MaxLength       =   4
            TabIndex        =   33
            Text            =   "Text1 7"
            Top             =   540
            Width           =   750
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   40
            Left            =   9360
            MaxLength       =   5
            TabIndex        =   34
            Text            =   "Text1 7"
            Top             =   540
            Width           =   750
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   46
            Left            =   12510
            MaxLength       =   15
            TabIndex        =   36
            Text            =   "Text1 7"
            Top             =   540
            Width           =   1935
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   44
            Left            =   10305
            MaxLength       =   15
            TabIndex        =   39
            Text            =   "Text1 7"
            Top             =   900
            Width           =   1800
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   38
            Left            =   8325
            MaxLength       =   4
            TabIndex        =   37
            Text            =   "Text1 7"
            Top             =   900
            Width           =   750
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   41
            Left            =   9360
            MaxLength       =   5
            TabIndex        =   38
            Text            =   "Text1 7"
            Top             =   900
            Width           =   750
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   47
            Left            =   12510
            MaxLength       =   15
            TabIndex        =   40
            Text            =   "Text1 7"
            Top             =   900
            Width           =   1935
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   45
            Left            =   10305
            MaxLength       =   15
            TabIndex        =   43
            Text            =   "Text1 7"
            Top             =   1275
            Width           =   1800
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   39
            Left            =   8325
            MaxLength       =   4
            TabIndex        =   41
            Text            =   "Text1 7"
            Top             =   1275
            Width           =   750
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   42
            Left            =   9360
            MaxLength       =   5
            TabIndex        =   42
            Text            =   "Text1 7"
            Top             =   1275
            Width           =   750
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   48
            Left            =   12510
            MaxLength       =   15
            TabIndex        =   44
            Text            =   "Text1 7"
            Top             =   1275
            Width           =   1935
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0FF&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   49
            Left            =   12285
            MaxLength       =   15
            TabIndex        =   45
            Text            =   "Text1 7"
            Top             =   1830
            Width           =   2160
         End
         Begin VB.Label Label1 
            Caption         =   "Base Imponible"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   27
            Left            =   10305
            TabIndex        =   117
            Top             =   255
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Imp. Bruto"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   24
            Left            =   150
            TabIndex        =   116
            Top             =   270
            Width           =   1260
         End
         Begin VB.Label Label1 
            Caption         =   "Imp.Dto. PP"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   23
            Left            =   2205
            TabIndex        =   115
            Top             =   270
            Width           =   1260
         End
         Begin VB.Label Label1 
            Caption         =   "Imp.Dto.Gnral"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   22
            Left            =   4185
            TabIndex        =   114
            Top             =   270
            Width           =   1500
         End
         Begin VB.Label Label1 
            Caption         =   "Base Imponible"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   18
            Left            =   6165
            TabIndex        =   113
            Top             =   270
            Width           =   1620
         End
         Begin VB.Label Label1 
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   30
            Left            =   1965
            TabIndex        =   112
            Top             =   480
            Width           =   135
         End
         Begin VB.Label Label1 
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   31
            Left            =   3945
            TabIndex        =   111
            Top             =   480
            Width           =   135
         End
         Begin VB.Label Label1 
            Caption         =   "="
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   32
            Left            =   5925
            TabIndex        =   110
            Top             =   480
            Width           =   135
         End
         Begin VB.Label Label1 
            Caption         =   "Importe IVA"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   33
            Left            =   12510
            TabIndex        =   109
            Top             =   285
            Width           =   1515
         End
         Begin VB.Label Label1 
            Caption         =   "TOTAL ALBARAN"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   285
            Index           =   39
            Left            =   10320
            TabIndex        =   108
            Top             =   1845
            Width           =   1785
         End
         Begin VB.Label Label1 
            Caption         =   "% IVA"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   41
            Left            =   9360
            TabIndex        =   107
            Top             =   240
            Width           =   765
         End
         Begin VB.Label Label1 
            Caption         =   "Cod.IVA"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   42
            Left            =   8325
            TabIndex        =   106
            Top             =   255
            Width           =   960
         End
      End
      Begin VB.CommandButton cmdAux 
         Appearance      =   0  'Flat
         Caption         =   "+"
         Height          =   315
         Index           =   2
         Left            =   10920
         TabIndex        =   99
         ToolTipText     =   "Buscar centro coste"
         Top             =   5835
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   9
         Left            =   10440
         MaxLength       =   4
         TabIndex        =   64
         Tag             =   "centro coste"
         Text            =   "cc"
         Top             =   5835
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   8
         Left            =   10320
         MaxLength       =   3
         TabIndex        =   96
         Tag             =   "IVA"
         Text            =   "IVA"
         Top             =   5355
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   16
         Left            =   1755
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   65
         Text            =   "frmComEntAlbaranesGR.frx":0B5C
         Top             =   7755
         Width           =   8700
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   17
         Left            =   11490
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   66
         Text            =   "ABCDKFJADKSFJAK"
         Top             =   7755
         Visible         =   0   'False
         Width           =   3390
      End
      Begin VB.Frame FrameHco 
         Caption         =   "Datos  Eliminaci�n"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   1170
         Left            =   -74640
         TabIndex        =   88
         Top             =   6390
         Width           =   14640
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   22
            Left            =   135
            MaxLength       =   10
            TabIndex        =   46
            Top             =   585
            Width           =   1290
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   23
            Left            =   1680
            MaxLength       =   30
            TabIndex        =   47
            Text            =   "Text1"
            Top             =   585
            Width           =   840
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   23
            Left            =   2565
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   90
            Text            =   "Text2"
            Top             =   585
            Width           =   4740
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   24
            Left            =   7485
            MaxLength       =   30
            TabIndex        =   48
            Text            =   "Text1"
            Top             =   600
            Width           =   885
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   24
            Left            =   8415
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   89
            Text            =   "Text2"
            Top             =   600
            Width           =   6000
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   37
            Left            =   135
            TabIndex        =   93
            Top             =   300
            Width           =   1440
         End
         Begin VB.Label Label1 
            Caption         =   "Trabajador"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   38
            Left            =   1695
            TabIndex        =   92
            Top             =   315
            Width           =   1140
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   6
            Left            =   2925
            ToolTipText     =   "Buscar trabajador"
            Top             =   315
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Incidencia"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   40
            Left            =   7500
            TabIndex        =   91
            Top             =   330
            Width           =   1185
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   7
            Left            =   8730
            ToolTipText     =   "Buscar incidencia"
            Top             =   330
            Width           =   285
         End
      End
      Begin VB.Frame FrameCliente 
         Height          =   2475
         Left            =   220
         TabIndex        =   70
         Top             =   465
         Width           =   14775
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   21
            Left            =   9090
            MaxLength       =   30
            TabIndex        =   19
            Tag             =   "Trab. Pedido|N|S|0|9999|scaalp|codtrab1|0000|N|"
            Text            =   "Text1"
            Top             =   1035
            Width           =   840
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   21
            Left            =   9930
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   125
            Text            =   "Text2"
            Top             =   1035
            Width           =   4710
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   25
            Left            =   5790
            MaxLength       =   10
            TabIndex        =   15
            Tag             =   "Fecha recepcion|F|S|||scaalp|fecenvio|dd/mm/yyyy||"
            Top             =   1890
            Width           =   1425
         End
         Begin VB.CheckBox chkDocArchi 
            Caption         =   "Documento archivado"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   7425
            TabIndex        =   21
            Tag             =   "Ar|N|S|||scaalp|docarchiv|||"
            Top             =   1890
            Width           =   2625
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   26
            Left            =   9930
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   124
            Text            =   "Text2"
            Top             =   1425
            Width           =   4710
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   26
            Left            =   9090
            MaxLength       =   30
            TabIndex        =   20
            Tag             =   "Envio|N|S|0|9999|scaalp|codenvio|0000|N|"
            Text            =   "Text1"
            Top             =   1425
            Width           =   840
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   27
            Left            =   1575
            MaxLength       =   10
            TabIndex        =   14
            Tag             =   "Fecha entraga|F|S|||scaalp|fecentrega|dd/mm/yyyy|N|"
            Top             =   1890
            Width           =   1425
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   11
            Left            =   1395
            MaxLength       =   30
            TabIndex        =   13
            Tag             =   "Provincia|T|N|||scaalp|proprove||N|"
            Text            =   "Text1 Text1 Text1 Text1 Text22"
            Top             =   1410
            Width           =   5820
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   9
            Left            =   1380
            MaxLength       =   6
            TabIndex        =   11
            Tag             =   "CPostal|T|N|||scaalp|codpobla||N|"
            Text            =   "Text15"
            Top             =   1005
            Width           =   945
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   10
            Left            =   2370
            MaxLength       =   30
            TabIndex        =   12
            Tag             =   "Poblaci�n|T|N|||scaalp|pobprove||N|"
            Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwaq"
            Top             =   1005
            Width           =   4845
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   7
            Left            =   4515
            MaxLength       =   20
            TabIndex        =   9
            Tag             =   "tel�fono Proveedor|T|S|||scaalp|telprove||N|"
            Text            =   "12345678911234567899"
            Top             =   190
            Width           =   2700
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   6
            Left            =   1380
            MaxLength       =   15
            TabIndex        =   8
            Tag             =   "NIF Proveedor|T|N|||scaalp|nifprove||N|"
            Text            =   "123456789"
            Top             =   190
            Width           =   2115
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   12
            Left            =   9090
            MaxLength       =   30
            TabIndex        =   16
            Tag             =   "Forma de Pago|N|N|0|999|scaalp|codforpa|000|N|"
            Text            =   "Text1"
            Top             =   190
            Width           =   855
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   12
            Left            =   9945
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   72
            Text            =   "Text2"
            Top             =   190
            Width           =   4695
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   13
            Left            =   9090
            MaxLength       =   7
            TabIndex        =   17
            Tag             =   "Descuento P.Pago|N|N|0|99.90|scaalp|dtoppago|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   600
            Width           =   840
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   14
            Left            =   11190
            MaxLength       =   7
            TabIndex        =   18
            Tag             =   "Descuento General|N|N|0|99.90|scaalp|dtognral|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   600
            Width           =   930
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   8
            Left            =   1380
            MaxLength       =   35
            TabIndex        =   10
            Tag             =   "Domicilio|T|N|||scaalp|domprove||N|"
            Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwwwww aq"
            Top             =   600
            Width           =   5835
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   8
            Left            =   8820
            ToolTipText     =   "Buscar trabajador"
            Top             =   1440
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "F.Recogida"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   44
            Left            =   135
            TabIndex        =   129
            Top             =   1935
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Trab. Pedido"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   7425
            TabIndex        =   128
            Top             =   1035
            Width           =   1365
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   4
            Left            =   8820
            ToolTipText     =   "Buscar trabajador"
            Top             =   1050
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "F.Archivo"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   4455
            TabIndex        =   127
            Top             =   1935
            Width           =   990
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   1
            Left            =   5490
            Picture         =   "frmComEntAlbaranesGR.frx":0B99
            ToolTipText     =   "Buscar fecha"
            Top             =   1935
            Width           =   240
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Forma Envio"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   4
            Left            =   7425
            TabIndex        =   126
            Top             =   1425
            Width           =   1200
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   2
            Left            =   1275
            Picture         =   "frmComEntAlbaranesGR.frx":0C24
            ToolTipText     =   "Buscar fecha"
            Top             =   1935
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Enabled         =   0   'False
            Height          =   240
            Index           =   5
            Left            =   1095
            ToolTipText     =   "Buscar proveedor vario"
            Top             =   240
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Enabled         =   0   'False
            Height          =   240
            Index           =   2
            Left            =   1110
            ToolTipText     =   "Buscar poblaci�n"
            Top             =   1005
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Provincia"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   17
            Left            =   120
            TabIndex        =   79
            Top             =   1410
            Width           =   960
         End
         Begin VB.Label Label1 
            Caption         =   "Poblaci�n"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   16
            Left            =   120
            TabIndex        =   78
            Top             =   1005
            Width           =   1005
         End
         Begin VB.Label Label1 
            Caption         =   "Tel�fono"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   19
            Left            =   3555
            TabIndex        =   77
            Top             =   195
            Width           =   960
         End
         Begin VB.Label Label1 
            Caption         =   "NIF"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   20
            Left            =   120
            TabIndex        =   76
            Top             =   190
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Forma Pago"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   15
            Left            =   7455
            TabIndex        =   75
            Top             =   195
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Dto. P. Pago"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   25
            Left            =   7455
            TabIndex        =   74
            Top             =   600
            Width           =   1290
         End
         Begin VB.Label Label1 
            Caption         =   "Dto. Gral"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   26
            Left            =   10200
            TabIndex        =   73
            Top             =   600
            Width           =   930
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   3
            Left            =   8805
            ToolTipText     =   "Buscar forma de pago"
            Top             =   195
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Domicilio"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   71
            Top             =   600
            Width           =   1005
         End
      End
      Begin VB.CommandButton cmdAux 
         Appearance      =   0  'Flat
         Caption         =   "+"
         Height          =   315
         Index           =   1
         Left            =   2655
         TabIndex        =   69
         ToolTipText     =   "Buscar art�culo"
         Top             =   5895
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.CommandButton cmdAux 
         Appearance      =   0  'Flat
         Caption         =   "+"
         Height          =   315
         Index           =   0
         Left            =   960
         TabIndex        =   68
         ToolTipText     =   "Buscar almacen"
         Top             =   5895
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   2880
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   58
         Tag             =   "Nombre Art�culo"
         Text            =   "nomArtic"
         Top             =   5835
         Visible         =   0   'False
         Width           =   3165
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   7
         Left            =   9360
         MaxLength       =   16
         TabIndex        =   63
         Tag             =   "Importe"
         Text            =   "Importe"
         Top             =   5835
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   6
         Left            =   8880
         MaxLength       =   30
         TabIndex        =   62
         Tag             =   "Descuento 2"
         Text            =   "Dto2"
         Top             =   5835
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   8280
         MaxLength       =   5
         TabIndex        =   61
         Tag             =   "Descuento 1"
         Text            =   "Dto1"
         Top             =   5835
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   7320
         MaxLength       =   12
         TabIndex        =   60
         Tag             =   "Precio"
         Text            =   "123,456.7879"
         Top             =   5835
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   6120
         MaxLength       =   16
         TabIndex        =   59
         Tag             =   "Cantidad"
         Text            =   "1,234,567,891.25"
         Top             =   5835
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   1200
         MaxLength       =   18
         TabIndex        =   57
         Tag             =   "C�digo Art�culo"
         Text            =   "Artic Artic Artic5"
         Top             =   5895
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   360
         MaxLength       =   15
         TabIndex        =   56
         Tag             =   "C�digo Almacen"
         Text            =   "codalmac"
         Top             =   5895
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   19
         Left            =   -73005
         MaxLength       =   80
         TabIndex        =   28
         Tag             =   "Observaci�n 5|T|S|||scaalp|observa5||N|"
         Top             =   3210
         Width           =   12810
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   18
         Left            =   -73005
         MaxLength       =   80
         TabIndex        =   27
         Tag             =   "Observaci�n 4|T|S|||scaalp|observa4||N|"
         Top             =   2820
         Width           =   12810
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   17
         Left            =   -73005
         MaxLength       =   80
         TabIndex        =   26
         Tag             =   "Observaci�n 3|T|S|||scaalp|observa3||N|"
         Top             =   2430
         Width           =   12810
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   16
         Left            =   -73005
         MaxLength       =   80
         TabIndex        =   25
         Tag             =   "Observaci�n 2|T|S|||scaalp|observa2||N|"
         Top             =   2040
         Width           =   12810
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   15
         Left            =   -73005
         MaxLength       =   80
         TabIndex        =   24
         Tag             =   "Observaci�n 1|T|S|||scaalp|observa1||N|"
         Top             =   1650
         Width           =   12810
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmComEntAlbaranesGR.frx":0CAF
         Height          =   3840
         Left            =   225
         TabIndex        =   67
         Top             =   3720
         Width           =   14595
         _ExtentX        =   25744
         _ExtentY        =   6773
         _Version        =   393216
         AllowUpdate     =   -1  'True
         AllowArrows     =   -1  'True
         ColumnHeaders   =   -1  'True
         HeadLines       =   1
         RowHeight       =   19
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
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
      Begin VB.Label Label1 
         Caption         =   "S/Referencia"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   48
         Left            =   -67350
         TabIndex        =   133
         Top             =   945
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "N/Referencia"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   47
         Left            =   -74670
         TabIndex        =   132
         Top             =   945
         Width           =   2175
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   9
         Left            =   1470
         ToolTipText     =   "Buscar forma de pago"
         Top             =   7755
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Ampliaci�n L�nea"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   35
         Left            =   240
         TabIndex        =   95
         Top             =   7755
         Width           =   1065
      End
      Begin VB.Label Label1 
         Caption         =   "N� Lote"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   10635
         TabIndex        =   94
         Top             =   7755
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label Label1 
         Caption         =   "Observaciones"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   45
         Left            =   -74685
         TabIndex        =   55
         Top             =   1665
         Width           =   1590
      End
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   14220
      TabIndex        =   51
      Top             =   10440
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Label lblF 
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10305
      TabIndex        =   100
      Top             =   10470
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Centro coste"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   46
      Left            =   2355
      TabIndex        =   98
      Top             =   10470
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Enabled         =   0   'False
      Visible         =   0   'False
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
      Begin VB.Menu mnLineas 
         Caption         =   "&Lineas"
         HelpContextID   =   2
         Shortcut        =   ^L
      End
      Begin VB.Menu mnBarra3 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmComEntAlbaranesGR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el n� de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

'Si se llama de la busqueda en el frmAlmMovimArticulos se accede
'a las tablas del Albaran o de Facturas de movimiento seleccionado (solo consulta)
Public hcoCodMovim As String 'cod. movim
Public hcoFechaMovim As Date 'Fecha del Movim
Public hcoCodProve As Long 'Codigo de Proveedor

Public EsHistorico As Boolean 'Si es true abrir el formulario con la tabla de
                              'de historico schalb, y solo en modo de consulta
                              
'cadena que selecciona los albaranes de un proveedor para mostrar
'antes de facturarlos
Public cadSelAlbaranes As String

Private WithEvents frmB As frmBasico2 'Form para busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmCCos As frmBasico2 'ayuda centros de coste
Attribute frmCCos.VB_VarHelpID = -1
Private WithEvents frmF As frmCal 'Form Calendario Fecha
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmCP As frmCPostal 'Codigos Postales
Attribute frmCP.VB_VarHelpID = -1

Private WithEvents frmProve As frmBasico2  'Form Mto Proveedores
Attribute frmProve.VB_VarHelpID = -1
Private WithEvents frmPV As frmComProveV   'Form Mto Proveedores Varios
Attribute frmPV.VB_VarHelpID = -1

Private WithEvents frmFP As frmBasico2 'frmFacFormasPago 'Form Mto Formas de Pago
Attribute frmFP.VB_VarHelpID = -1
Private WithEvents frmT As frmBasico2 'frmAdmTrabajadores  'Form Mto Trabajadores
Attribute frmT.VB_VarHelpID = -1
Private WithEvents frmAlm As frmAlmAlPropios   'Form Almacenes Propios
Attribute frmAlm.VB_VarHelpID = -1
Private WithEvents FrmArt As frmBasico2   'Form Articulos
Attribute FrmArt.VB_VarHelpID = -1

Private WithEvents frmNSerie As frmRepCargarNSerie  'Form Cargar n� Series
Attribute frmNSerie.VB_VarHelpID = -1
Private WithEvents frmMen As frmMensajes  'Form Mensajes
Attribute frmMen.VB_VarHelpID = -1
Private WithEvents frmList As frmListadoOfer
Attribute frmList.VB_VarHelpID = -1
Private WithEvents frmFE As frmFacFormasEnvio
Attribute frmFE.VB_VarHelpID = -1


'-------------------------------------------------------------------------
Private Modo As Byte
'-----------------------------
'Se distinguen varios Modos
'   0.-  Formulario limpio sin nungun campo rellenado
'   1.-  Preparando para hacer la busqueda
'   2.-  Ya tenemos registros y los vamos a recorrer
'        y podemos editarlos Edicion del campo
'   3.-  Insercion de nuevo registro
'   4.-  Modificar
'   5.-  Mantenimiento Lineas
'-------------------------------------------------------------------------

Dim ModificaLineas As Byte
'1.- A�adir,  2.- Modificar,  3.- Borrar,  0.-Pasar control a Lineas
'4.- Mantenimiento de N� de Serie

Dim TituloLinea As String 'Descripcion de la linea que estamos en Mantenimiento

Dim PrimeraVez As Boolean
Dim PrimeraVezForm As Boolean

Dim CodTipoMov As String
'Codigo tipo de movimiento en funci�n del valor en la tabla de par�metros: stipom

Dim EsDeVarios As Boolean
'Si el Proveedor mostrado es de Varios o No

'SQL de la tabla principal del formulario
Private CadenaConsulta As String

Private Ordenacion As String
Private NombreTabla As String  'Nombre de la tabla o de la
Private NomTablaLineas As String 'Nombre de la Tabla de lineas
Private kCampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean

Dim btnAnyadir As Byte
'Variable que indica el n�mero del Boton  Anyadir en la Toolbar1
Dim btnPrimero As Byte
'Variable que indica el n�mero del Boton  PrimerRegistro en la Toolbar1


Private VieneDeBuscar As Boolean
'Para cuando devuelve dos poblaciones con el mismo codigo Postal. Si viene de pulsar prismatico
'de busqueda poner el valor de poblacion seleccionado y no volver a recuperar de la Base de Datos
Private HaCambiadoCP As Boolean
'Para saber si tras haber vuelto de prismaticos ha cambiado el valor del CPostal

Dim cadList As String 'cadena para pasar al historico

'Febrero 2010
'Despues de meter la primera linea, continua ofertando el almacen de la linea anterior
Dim AlmacenLineas As Integer  '

Dim PulsadoMas2 As Boolean


'Diciembre 2010.
'Para ver si al volver a cabecera muestra el listado o no
Dim HaModifEnLineas As Boolean


Dim Indice As Byte




Private Sub chkDocArchi_Click()
   '  If Modo = 1 Then CheckCadenaBusqueda chkDocArchi, BuscaChekc
End Sub
Private Sub chkDocArchi_KeyDown(KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub
Private Sub chkDocArchi_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub










Private Sub cmdAceptar_Click()
Dim PrimeraLin As Boolean 'Si se inserta la primera linea no esta creado el datagrid1 entonces llamar
                          ' a DataGrid, sino llamar solo a DataGrid2
Dim numlinea As String
Dim precioUC As Currency
Dim FechaUltCompra As Date
Dim EnPromocionOPrecioEspecial As String
Dim ParaElMargen As Currency

    Screen.MousePointer = vbHourglass
    On Error GoTo Error1

    Select Case Modo
        Case 1  'BUSQUEDA
            HacerBusqueda
            
        Case 3 'INSERTAR CABECERA
            If DatosOk Then InsertarCabecera
            
        Case 4  'MODIFICAR CABECERA
            If DatosOk Then
                If ModificarCabAlbaran Then
                    If cadSelAlbaranes = "" Then TerminaBloquear
                    PosicionarData
                End If
            End If
            
         Case 5 'InsertarModificar LINEAS
            'Antes de insertar la linea guardamos el sartic.preciouc actual
            'para aplicar margen despues, pq en Insertar linea se actualiza ya el preciouc
            If txtAux(1).Text = "" Then
                Screen.MousePointer = vbDefault
                MsgBox "Campos obligatorios", vbExclamation
                PonerFoco txtAux(1)
                Exit Sub
            End If
            numlinea = "ultfecco"
            precioUC = ComprobarCero(DevuelveDesdeBDNew(conAri, "sartic", "preciouc", "codartic", txtAux(1).Text, "T", numlinea))
            'Fecha ultima compra
            FechaUltCompra = "01/01/1900"
            If numlinea <> "" Then
                FechaUltCompra = CDate(numlinea)
                numlinea = ""
            End If
                
         
            'Actualizar el registro en la tabla de lineas 'slialb'
            If ModificaLineas = 1 Then 'INSERTAR lineas Albaran
                PrimeraLin = False
                If Data2.Recordset.EOF = True Then PrimeraLin = True
                          
                If InsertarLinea(numlinea) Then
                    'Comprobar si Hay N� SERIE en compras y Mostrar
                    'ventana para pedir los N� Serie de la cantidad introducida
                    If vParamAplic.NumSeries Then
                        ComprobarNumSeries (numlinea)
                    End If
                    
                    'Comprobar si se ha modificado el precio desde la ultima compra
                    'y preguntar quiere modificar el PVP del articulo aplicandole su margen
                    'y el precio de las TArifas aplicandole el margen
                    '-- Laura 19/12/2006: el precio de compra es el precio con los descuentos (importe/cantidad)
                    'If precioUC <> CCur(txtAux(4).Text) Then
                    If CCur(txtAux(3).Text) <> 0 Then
                            If precioUC <> Round2(CCur(txtAux(7).Text) / CCur(txtAux(3).Text), 4) Then
                                If CDate(Text1(1).Text) >= FechaUltCompra Then
                                
                                    'Marzo. Para saber si esta en pormociones/ofertas
                                    precioUC = 0
                                    
                                    EnPromocionOPrecioEspecial = "fechaini<=" & DBSet(Text1(1).Text, "F")
                                    EnPromocionOPrecioEspecial = EnPromocionOPrecioEspecial & " AND fechafin>=" & DBSet(Text1(1).Text, "F") & " and codartic"
                                    EnPromocionOPrecioEspecial = DevuelveDesdeBD(conAri, "codartic", "spromo", EnPromocionOPrecioEspecial, txtAux(1).Text, "T")
                                    If EnPromocionOPrecioEspecial <> "" Then precioUC = 1
                                    EnPromocionOPrecioEspecial = DevuelveDesdeBD(conAri, "codartic", "sprees", "codartic", txtAux(1).Text, "T")
                                    If EnPromocionOPrecioEspecial <> "" Then precioUC = precioUC + 2
                                    If precioUC = 0 Then
                                        EnPromocionOPrecioEspecial = ""
                                    Else
                                        EnPromocionOPrecioEspecial = "ATENCION. Art�culo en:"
                                        If precioUC = 1 Or precioUC = 3 Then EnPromocionOPrecioEspecial = EnPromocionOPrecioEspecial & vbCrLf & " - PROMOCIONES"
                                        If precioUC = 2 Or precioUC = 3 Then EnPromocionOPrecioEspecial = EnPromocionOPrecioEspecial & vbCrLf & " - PRECIOS ESPECIALES"
                                        EnPromocionOPrecioEspecial = vbCrLf & String(20, "*") & vbCrLf & vbCrLf & EnPromocionOPrecioEspecial & vbCrLf & String(20, "*")
                                        EnPromocionOPrecioEspecial = vbCrLf & vbCrLf & vbCrLf & EnPromocionOPrecioEspecial
                                    End If
    
                                    EnPromocionOPrecioEspecial = "Se ha modificado el precio �ltima compra." & vbCrLf & "�Desea actualizar los precios de venta?" & EnPromocionOPrecioEspecial
                                    If MsgBox(EnPromocionOPrecioEspecial, vbQuestion + vbYesNo) = vbYes Then
                                        'Comprobar que el art�culo tiene margen comercial
                                        If ArticuloTieneMargen(txtAux(1).Text) Then
          
                                                frmComActPrecios.parCodArtic = txtAux(1).Text
                                                frmComActPrecios.parNomArtic = txtAux(2).Text
                                                frmComActPrecios.Show vbModal
            
                                        End If
                                    Else
                                    
                                        If vParamAplic.RecalculoMargen Then ActualizacionAutomaticaMargen txtAux(1).Text
                                              
                                    End If
                                    
                                    'Veremos si es un articulo "escandallo"
                                    EnPromocionOPrecioEspecial = DevuelveDesdeBD(conAri, "count(*)", "sarti1", "codarti1", txtAux(1).Text, "T")
                                    If EnPromocionOPrecioEspecial = "" Then EnPromocionOPrecioEspecial = "0"
                                    If Val(EnPromocionOPrecioEspecial) > 0 Then
                                        frmListado4.Opcion = 1
                                        frmListado4.vCadena = txtAux(1).Text & "|"
                                        frmListado4.Show vbModal
                                    End If
                                    EnPromocionOPrecioEspecial = ""
                                    
                                End If   'Fecha ultima compra
                            End If  'Precio ultima compra
                    End If  'Por si han puesto cantidad=0
                    
                    
                    'Abril 2013
                    'INsertar linea de componentes
                    InsertarSubComponentes
                     
                     
                    If PrimeraLin Then
                        CargaGrid DataGrid1, Data2, True
                    Else
                        CargaGrid2 DataGrid1, Data2
                    End If
                    BotonAnyadirLinea
                End If
                
            ElseIf ModificaLineas = 2 Then 'MODIFICAR lineas
                If DatosOkLinea() Then
                    If ModificarLinea Then
                        
                        
                        
                        If precioUC <> Round2(CCur(txtAux(7).Text) / CCur(txtAux(3).Text), 4) Then
                        
                               'Marzo. Para saber si esta en pormociones/ofertas
                                precioUC = 0
                                EnPromocionOPrecioEspecial = DevuelveDesdeBD(conAri, "codartic", "spromo", "codartic", txtAux(1).Text, "T")
                                If EnPromocionOPrecioEspecial <> "" Then precioUC = 1
                                EnPromocionOPrecioEspecial = DevuelveDesdeBD(conAri, "codartic", "spromo", "codartic", txtAux(1).Text, "T")
                                If EnPromocionOPrecioEspecial <> "" Then precioUC = precioUC + 2
                                If precioUC = 0 Then
                                    EnPromocionOPrecioEspecial = ""
                                Else
                                    EnPromocionOPrecioEspecial = "ATENCION. Art�culo en:"
                                    If precioUC = 1 Or precioUC = 3 Then EnPromocionOPrecioEspecial = EnPromocionOPrecioEspecial & vbCrLf & " - PROMOCIONES"
                                    If precioUC = 3 Then EnPromocionOPrecioEspecial = EnPromocionOPrecioEspecial & vbCrLf & " - PRECIOS ESPECIALES"
                                    EnPromocionOPrecioEspecial = vbCrLf & String(20, "*") & vbCrLf & vbCrLf & EnPromocionOPrecioEspecial & vbCrLf & String(20, "*")
                                    EnPromocionOPrecioEspecial = vbCrLf & vbCrLf & vbCrLf & EnPromocionOPrecioEspecial
                                End If

                                EnPromocionOPrecioEspecial = "Se ha modificado el precio �ltima compra." & vbCrLf & "�Desea actualizar los precios de venta?" & EnPromocionOPrecioEspecial
                    
                        
                        
                        
                            If MsgBox(EnPromocionOPrecioEspecial, vbQuestion + vbYesNo) = vbYes Then
                                'Comprobar que el art�culo tiene margen comercial
                                If ArticuloTieneMargen(txtAux(1).Text) Then
                                    'Aplicar margen comercial a los precios
                                    'Modificar precios de venta en articulo y tarifas
                                    frmComActPrecios.parCodArtic = txtAux(1).Text
                                    frmComActPrecios.parNomArtic = txtAux(2).Text
        '                            frmcomactprecios.parPrecioUC =
                                    frmComActPrecios.Show vbModal
                                End If
                            Else
                                If vParamAplic.RecalculoMargen Then ActualizacionAutomaticaMargen txtAux(1).Text
                            End If
                            
                            'Veremos si es un articulo "escandallo"
                            EnPromocionOPrecioEspecial = DevuelveDesdeBD(conAri, "count(*)", "sarti1", "codarti1", txtAux(1).Text, "T")
                            If EnPromocionOPrecioEspecial = "" Then EnPromocionOPrecioEspecial = "0"
                            If Val(EnPromocionOPrecioEspecial) > 0 Then
                                frmListado4.Opcion = 1
                                frmListado4.vCadena = txtAux(1).Text & "|"
                                frmListado4.Show vbModal
                            End If
                            EnPromocionOPrecioEspecial = ""
                            
                            
                        End If
                        
                        NumRegElim = Data2.Recordset!numlinea
                        TerminaBloquear
                        CargaTxtAux False, False
                        CargaGrid2 DataGrid1, Data2
                        ModificaLineas = 0
'                        PonerBotonCabecera True
                        BloquearTxt Text2(16), True
                        BloquearTxt Text2(17), True
                        
                        'AQUI
                        PosicionarData2
                        
                        PonerModo 2
                        

                        'BloquearTabs False
                        If Not Data1.Recordset.EOF Then _
                            Me.lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
                        If DataGrid1.Row >= 0 Then
                            DeseleccionaGrid DataGrid1
                           ' DataGrid1.Bookmark = 1
                        End If
                         
                        'Lanzaremos solo en cancelar 2021 Abril
                        'If HaModifEnLineas Then ComprobarPedidosClientesDesdeAlbProveedor Text1(0).Text, Text1(1).Text, Text1(4).Text

                    End If
                    Me.DataGrid1.Enabled = True
                End If
            End If
            CalcularDatosFactura 'rellenar campos pesta�a de totales
    End Select
    Screen.MousePointer = vbDefault

Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub PosicionarData2()
    On Error GoTo EPosicionarData2
    
    Data2.Recordset.Find "numlinea = " & NumRegElim
    If Data2.Recordset.EOF Then Data2.Recordset.MoveFirst
    NumRegElim = 0
    Exit Sub
EPosicionarData2:
    MuestraError Err.Number
End Sub



Private Function ComprobarCambioFecha() As Boolean
'Comprueba si se ha modificado el campo fecha de la cabecera.
'Ya que es clave primaria y se deberan cambiar tambien la fecha
'en tablas sliap y smoval
Dim RS As ADODB.Recordset
Dim SQL As String
Dim Izquierda As String, Derecha As String
Dim llis As Collection
Dim i As Integer
Dim b As Boolean


    If Data1.Recordset.EOF Then Exit Function

    
    If (CDate(Text1(1).Text) <> CDate(Data1.Recordset!FechaAlb)) Then
    'si ha modificado la fecha de albaran
        On Error GoTo EComprobar
        
        'seleccionar todas las lineas de ese albaran para actualizar la fecha (slialp)
        SQL = "SELECT * FROM " & NomTablaLineas & " WHERE numalbar=" & DBSet(Data1.Recordset!Numalbar, "T")
        SQL = SQL & " AND fechaalb=" & DBSet(Data1.Recordset!FechaAlb, "F")
        SQL = SQL & " AND codprove=" & Data1.Recordset!Codprove
        
        Set RS = New ADODB.Recordset
        RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        
        Set llis = New Collection
            
        'Nos guardamos todas las lineas con la modificacion de la fecha para
        'volverlas a insertar
        BACKUP_TablaIzquierda RS, Izquierda
        
        While Not RS.EOF
            BACKUP_Tabla RS, Derecha, "fechaalb", CStr(Text1(1).Text)
            llis.Add Derecha
            RS.MoveNext
        Wend
        
        RS.Close
       
        
        
        'Numeros de lotes fitosantiarios
        'Cambiamos en slotes
        If vParamAplic.ManipuladorFitosanitarios2 Then
            SQL = " Select * from " & NomTablaLineas & " WHERE numalbar = " & DBSet(Data1.Recordset!Numalbar, "T")
            SQL = SQL & " AND fechaalb=" & DBSet(Data1.Recordset!FechaAlb, "F")
            SQL = SQL & " AND codprove=" & Data1.Recordset!Codprove
            
            
            RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            While Not RS.EOF
                
                SQL = "UPDATE slotes SET fecentra = " & DBSet(Text1(1).Text, "F") & " WHERE "
                SQL = SQL & " codartic=" & DBSet(RS!codArtic, "T") & " AND numlotes=" & DBSet(RS!numlotes, "T") & " AND fecentra=" & DBSet(RS!FechaAlb, "F")
                conn.Execute SQL
                RS.MoveNext
            Wend
            RS.Close
            
        End If


        
        Set RS = Nothing
        
        
        
        
        
        
        'Eliminamos las lineas que tenia ese albaran (slialp) para volverlas a insertar con la fecha nueva
        SQL = "DELETE from slialp WHERE numalbar = " & DBSet(Data1.Recordset!Numalbar, "T")
        SQL = SQL & " AND fechaalb=" & DBSet(Data1.Recordset!FechaAlb, "F")
        SQL = SQL & " AND codprove=" & Data1.Recordset!Codprove
        conn.Execute SQL
        
        'Actualizamos la fecha en la cabecera (scaalp)
        SQL = "UPDATE scaalp SET fechaalb = " & DBSet(Text1(1).Text, "F")
        SQL = SQL & " WHERE numalbar = " & DBSet(Data1.Recordset!Numalbar, "T")
        SQL = SQL & " AND fechaalb=" & DBSet(Data1.Recordset!FechaAlb, "F")
        SQL = SQL & " AND codprove=" & Data1.Recordset!Codprove
        conn.Execute SQL
        
        
        'Agosto 2020.

            'Actualizamos la fecha en la tabla smoval
'            SQL = "UPDATE smoval SET fechamov=" & DBSet(Text1(1).Text, "F")
'            SQL = SQL & " WHERE document = " & DBSet(Data1.Recordset!Numalbar, "T")
'            SQL = SQL & " AND fechamov=" & DBSet(Data1.Recordset!FechaAlb, "F")
'            SQL = SQL & " AND codigope=" & Data1.Recordset!Codprove
'            SQL = SQL & " AND detamovi='" & CodTipoMov & "'"
'            conn.Execute SQL

        
        'Actualizar la fecha compra en los numeros de serie del albaran (si tiene articulos con num. serie)
        SQL = "UPDATE sserie SET fechacom=" & DBSet(Text1(1).Text, "F")
        SQL = SQL & " WHERE fechacom=" & DBSet(Data1.Recordset!FechaAlb, "F") & " AND "
        SQL = SQL & " numalbpr=" & DBSet(Data1.Recordset!Numalbar, "T")
        SQL = SQL & " AND codprove=" & Data1.Recordset!Codprove
        conn.Execute SQL
            
        
                
        'Volvemos a insertar las lineas con la fecha correcta (slialp)
        SQL = ""
        For i = 1 To llis.Count
            If (i Mod 10) = 0 Then
                SQL = SQL & CStr(llis(i)) & ","
                SQL = Mid(SQL, 1, Len(SQL) - 1) 'quitamos ultima coma
                SQL = "INSERT INTO " & NomTablaLineas & " " & Izquierda & " VALUES " & SQL & ";"
                conn.Execute SQL
                SQL = ""
            Else
                SQL = SQL & CStr(llis(i)) & ","
            End If
        Next i
        
        If SQL <> "" Then
            SQL = Mid(SQL, 1, Len(SQL) - 1) 'quitamos ultima coma
            SQL = "INSERT INTO " & NomTablaLineas & " " & Izquierda & " VALUES " & SQL & ";"
            conn.Execute SQL
            SQL = ""
        End If
        Set llis = Nothing
    End If
    b = True
    
EComprobar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "El campo fecha no se ha podido modificar", Err.Description
        b = False
    End If
    If b Then
        ComprobarCambioFecha = True
    Else
        ComprobarCambioFecha = False
    End If
End Function



Private Sub cmdAux_Click(Index As Integer)
    Select Case Index
        Case 0 'Busqueda de Cod. Almacen
            Set frmAlm = New frmAlmAlPropios
            frmAlm.DatosADevolverBusqueda = "0"
            frmAlm.Show vbModal
            Set frmAlm = Nothing
            PonerFoco txtAux(Index)
            
        Case 1 'Busqueda de Cod. Artic
            Set FrmArt = New frmBasico2
            'frmArt.DatosADevolverBusqueda3 = "@1@" 'Poner en modo b�squeda
            AyudaArticulos FrmArt, txtAux(1).Text
            Set FrmArt = Nothing
            PonerFoco txtAux(Index)
            
        Case 2 'COD. CENTRO COSTE
            If vEmpresa.TieneAnalitica Then
                'centro de coste
                AbrirForm_CentroCoste
                PonerFoco txtAux(9)
            End If
    End Select
    
End Sub


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
            TerminaBloquear
            CargaTxtAux False, False
            BloquearTxt Text2(16), True
            BloquearTxt Text2(17), True
            DataGrid1.Columns(5).Caption = "Articulo"
            If ModificaLineas = 1 Then 'INSERTAR
                ModificaLineas = 0
                DataGrid1.AllowAddNew = False
                If Not Data2.Recordset.EOF Then Data2.Recordset.MoveFirst
            End If
            lblF.Caption = ""
            
        
            PonerModo 2
            'BloquearTabs False
            If Not Data1.Recordset.EOF Then _
                Me.lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
            If DataGrid1.Row >= 0 Then
                DeseleccionaGrid DataGrid1
                DataGrid1.Bookmark = 1
            End If
            
            
            
            If HaModifEnLineas And ModificaLineas <> 2 Then ComprobarPedidosClientesDesdeAlbProveedor Text1(0).Text, Text1(1).Text, Text1(4).Text
            ModificaLineas = 0

            Me.DataGrid1.Enabled = True
    End Select
End Sub


Private Sub BotonAnyadir()
'A�adir registro en tabla de cabecera de Albaranes: scaalp (Cabecera)
Dim NomTraba As String

    LimpiarCampos 'Vac�a los TextBox
    'Poner los grid sin apuntar a nada
    LimpiarDataGrids
    
    'A�adiremos el boton de aceptar y demas objetos para insertar
    PonerModo 3

    'Poner el nombre del trabajador que esta conectado
    Text1(2).Text = PonerTrabajadorConectado(NomTraba)
    Text2(2).Text = NomTraba
    
    Text1(1).Text = Format(Now, "dd/mm/yyyy") 'Fecha Albaran
    Text1(30).Text = Format(Now, "dd/mm/yyyy") 'Fecha Albaran
    PonerFoco Text1(0)
End Sub


Private Sub BotonAnyadirLinea()

    PonerModo 5



    'Si no estaba modificando lineas salimos
    ' Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 2 Then Exit Sub
    HaModifEnLineas = True
    ModificaLineas = 1 'Ponemos Modo A�adir Linea
    'A�adiremos el boton de aceptar y demas objetos para insertar
    PonerBotonCabecera False
    lblIndicador.Caption = "INSERTAR"
    
    
    
    PonerUltAlmacen
    

    AnyadirLinea DataGrid1, Data2
    CargaTxtAux True, True
    'Poner el Almacen por defecto del Trabajador
    txtAux(0).Text = Format(AlmacenLineas, "000")
    'Campo Ampliacion Linea
    Text2(16).Text = ""
    Text2(17).Text = ""
     txtAux2(9).Text = ""
    BloquearTxt Text2(16), False
    BloquearTxt Text2(17), True
    
    
    ' ---- [20/10/2009] [LAURA]: a�adir campo centro de coste
    'si contab. analitica por trabajador traer su centro de coste
    If vEmpresa.TieneAnalitica And vParamAplic.ModoAnalitica = 0 Then
        txtAux(9).Text = DevuelveDesdeBDNew(conAri, "straba", "codccost", "codtraba", Text1(2).Text, "N")
        Me.txtAux2(9).Text = PonerNombreCCoste(Me.txtAux(9))
    End If
    
    
    
    
    PonerFoco txtAux(1)
    Me.DataGrid1.Enabled = False
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
        Text1(0).BackColor = vbLightBlue
    Else
        HacerBusqueda
        If Data1.Recordset.EOF Then
            Text1(kCampo).Text = ""
            Text1(kCampo).BackColor = vbLightBlue
            PonerFoco Text1(kCampo)
        End If
    End If
End Sub


Private Sub BotonVerTodos()
    If chkVistaPrevia.Value = 1 Then
        MandaBusquedaPrevia cadSelAlbaranes
    Else
        LimpiarCampos
        LimpiarDataGrids
        If cadSelAlbaranes = "" Then
            CadenaConsulta = "Select * from " & NombreTabla & " " & Ordenacion
        Else
            CadenaConsulta = "Select * from " & NombreTabla & " " & " WHERE " & cadSelAlbaranes & Ordenacion
        End If
        PonerCadenaBusqueda
    End If
End Sub


Private Sub Desplazamiento(Index As Integer)
'Botones de Desplazamiento de la Toolbar
'Para desplazarse por los registros de control Data
    If Data1.Recordset.EOF Then Exit Sub
    DesplazamientoData Data1, Index, True
    PonerCampos
End Sub


Private Sub BotonModificar()
Dim SQL As String
Dim DeVarios As Boolean

    On Error GoTo EModificar

    'A�adiremos el boton de aceptar y demas objetos para insertar
    PonerModo 4
    
    PonerFoco Text1(2)
    
    If EsDeVarios Then
        If Data1.Recordset.EOF Then Exit Sub
        SQL = " SELECT * FROM sprvar WHERE nifprove='" & Data1.Recordset!nifProve & "' FOR UPDATE "
        conn.Execute SQL
    End If
    
    'Si es Cliente de Varios no se pueden modificar sus datos
    DeVarios = EsProveedorVarios(Text1(4).Text)
    BloquearDatosProve (DeVarios)
    
EModificar:
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub BotonModificarLinea()
'Modificar una linea
Dim vWhere As String
'Dim cArt As CArticulo  'Si bloquearamos por el articulo

    On Error GoTo EModificarLinea

    PonerModo 5


    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Then Exit Sub '1= Insertar
    
    If Data2.Recordset.EOF Then Exit Sub
    HaModifEnLineas = True
    
    vWhere = ObtenerWhereCP(False) & " and numlinea=" & Data2.Recordset!numlinea
    vWhere = Replace(vWhere, NombreTabla, NomTablaLineas)
    
    If Not BloqueaRegistro(NomTablaLineas, vWhere) Then Exit Sub
    
    ModificaLineas = 2 'Modificar
    CargaTxtAux True, False
    'ModificaLineas = 2 'Modificar
    'A�adiremos el boton de aceptar y demas objetos para insertar
    Me.lblIndicador.Caption = "MODIFICAR"
    PonerBotonCabecera False
    BloquearTxt Text2(16), False 'Campo Ampliacion Linea
    BloquearTxt txtAux(2), True 'campo nombre articulo
    
    'bloquear el num_lote si el articulo es de una categoria q no lleva control
    'de n� de lote
'    BloquearTxt Text2(17), (DBLet(Data2.Recordset!numlotes, "T") = "")
'    Set cArt = New CArticulo
'    If cArt.LeerDatos(Data2.Recordset!codArtic) Then
'        BloquearTxt Text2(17), Not cArt.TieneNumLote
'    End If
'    Set cArt = Nothing
    BloquearTxt Text2(17), False
    
    
    PonerFoco txtAux(0)
    Me.DataGrid1.Enabled = False
    
EModificarLinea:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub BotonEliminar()
'Eliminar Registro de la Cabecera: Tabla de albaranes compras (scaalp)
' y los registros correspondientes de las tablas de lineas (slialp)
'al eliminar un albaran ademas habr� que restaurar valores:
' - actualizar stock en (salmac)
' - eliminar los movimientos que inserto el albaran en (smoval)
' - actualizar el ultprecio compra y ultima fecha compra en funcion del ult. movimiento ALC en smoval
' - reestablecer el precio medio ponderado
Dim cad As String

    On Error GoTo EEliminar

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
    
    
    If Not PuedeEliminar(False) Then Exit Sub
    
    cad = "Cabecera de Albaranes Compras" & vbCrLf
    cad = cad & "-------------------------------------------------" & vbCrLf & vbCrLf
    cad = cad & "Va a eliminar el Albaran:            "
    cad = cad & vbCrLf & "N�:  " & Text1(0).Text
    cad = cad & vbCrLf & "Fecha: " & Text1(1).Text
    cad = cad & vbCrLf & vbCrLf & " �Desea Eliminarlo? " & vbCrLf & vbCrLf
    cad = cad & "-------------------------------------------------"
    'Borramos
    If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then
        Screen.MousePointer = vbHourglass
    
        NumRegElim = Data1.Recordset.AbsolutePosition

        'Abrir frame de informes para pedir datos antes de grabar en el historico
        cadList = ""
        Set frmList = New frmListadoOfer
        frmList.OpcionListado = 80
        frmList.Show vbModal
        Set frmList = Nothing
        If cadList = "" Then Exit Sub
        
        If Not Eliminar() Then
            CargaGrid Me.DataGrid1, Me.Data2, True
            Exit Sub
        ElseIf SituarDataTrasEliminar(Data1, NumRegElim) Then
            PonerCampos
        Else
            LimpiarCampos
            'Poner los grid sin apuntar a nada
            LimpiarDataGrids
            PonerModo 0
        End If
    End If
    Screen.MousePointer = vbDefault
    
EEliminar:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Albaran", Err.Description
End Sub


Private Sub BotonEliminarLinea()
'Eliminar una linea De Mantenimiento. Tabla: slima1
Dim SQL As String

    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Or ModificaLineas = 2 Then Exit Sub '1= Insertar, 2=Modificar

    If Data2.Recordset.EOF Then Exit Sub
    
    If Not PuedeEliminar(True) Then Exit Sub
    
    
    
    HaModifEnLineas = True
    ModificaLineas = 3 'Eliminar
    SQL = "�Seguro que desea eliminar la l�nea de Albaran?     "
    SQL = SQL & vbCrLf & "NumLinea:  " & Data2.Recordset!numlinea & vbCrLf
    SQL = SQL & "Almacen:  " & Format(Data2.Recordset!codAlmac, "000")
    SQL = SQL & vbCrLf & "Art�culo:  " & Data2.Recordset!codArtic & " - " & Data2.Recordset!NomArtic
    
    If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        NumRegElim = Data2.Recordset.AbsolutePosition
        If EliminarLinea Then
            ModificaLineas = 0
            SituarDataTrasEliminar Data2, NumRegElim
            'CargaGrid2 DataGrid1, Data2
            CalcularDatosFactura
        End If
'        CancelaADODC
    End If
    PonerFocoBtn Me.cmdRegresar

EEliminarLinea:
        Screen.MousePointer = vbDefault
        If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Lineas Mantenimientos", Err.Description
End Sub


Private Sub cmdRegresar_Click()
'Este es el boton Cabecera
Dim cad As String

    'Quitar lineas y volver a la cabecera
    If Modo = 5 Then  'modo 5: Mantenimientos Lineas
        PonerModo 2
        'BloquearTabs False
        If Not Data1.Recordset.EOF Then _
            Me.lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
        If DataGrid1.Row >= 0 Then
            DeseleccionaGrid DataGrid1
            DataGrid1.Bookmark = 1
        End If
        
        If HaModifEnLineas Then ComprobarPedidosClientesDesdeAlbProveedor Text1(0).Text, Text1(1).Text, Text1(4).Text
        
        
    Else 'Se llama desde alg�n Prismatico de otro Form al Mantenimiento de Trabajadores
        If Data1.Recordset.EOF Then
            MsgBox "Ning�n registro devuelto.", vbExclamation
            Exit Sub
        End If
        cad = Data1.Recordset.Fields(0) & "|"
        cad = cad & Data1.Recordset.Fields(1) & "|"
        RaiseEvent DatoSeleccionado(cad)
        Unload Me
    End If
End Sub




Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim devuelve As String
Dim CadLote As String

    On Error GoTo Error1
    
    If Not Data2.Recordset.EOF And ModificaLineas <> 1 Then '1: Insertar
        'cadLote = "numlotes"
        'devuelve = DevuelveDesdeBDNew(conAri, NomTablaLineas, "ampliaci", "numalbar", Text1(0).Text, "T", cadLote, "numlinea", Data2.Recordset!numlinea, "N")
        'Poner descripcion de ampliacion lineas
        Text2(16).Text = DBLet(Data2.Recordset!Ampliaci, "T")
        'poner el numero de lote
        Text2(17).Text = DBLet(Data2.Recordset!numlotes, "T")
        
        '- centro de coste
        ' ---- [20/10/2009] [LAURA]: a�adir campo centro de coste familia
        If vEmpresa.TieneAnalitica Then
            Me.txtAux(9).Text = DBLet(Data2.Recordset!CodCCost, "T")
            Me.txtAux2(9).Text = PonerNombreCCoste(Me.txtAux(9))
        Else
            txtAux2(9).Text = ""
        End If
   
    End If
    
    If ModificaLineas = 1 Then Text2(16).Text = ""
    Exit Sub
    
Error1:
    MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    If PrimeraVezForm Then
        PrimeraVezForm = False
        
        'Viene de DblClick en frmAlmMovimArticulos y carga el form con los valores
        If hcoCodMovim <> "" And Not Data1.Recordset.EOF Then PonerCadenaBusqueda
        
        'Viene de click en VerAlbaranes en formulario de "Recepcion de Facturas compra"
        If cadSelAlbaranes <> "" And Not Data1.Recordset.EOF Then PonerCadenaBusqueda
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub Form_Load()
Dim i As Integer

    'Icono del formulario
    Me.Icon = frmPpal.Icon
    
    For i = 1 To imgBuscar.Count - 1
        imgBuscar(i).Picture = imgBuscar(0).Picture
    Next
    
    ' ICONITOS DE LA BARRA
    btnAnyadir = 5
    btnPrimero = 20
    Modo = 0
'    With Me.Toolbar1
'        .ImageList = frmPpal.imgListComun
'        .Buttons(1).Image = 1   'Bot�n Buscar
'        .Buttons(2).Image = 2   'Bot�n Todos
'        .Buttons(5).Image = 3   'Insertar Nuevo
'        .Buttons(6).Image = 4   'Modificar
'        .Buttons(7).Image = 5   'Borrar
'
'
'        .Buttons(9).Image = 45 'Mto Lineas Albaran
'        .Buttons(10).Image = 10 'Mto Lineas Albaran
'        .Buttons(12).Image = 32 'Pasar a hco pero sin mover la smoval ni precios ni "leches"
'        .Buttons(14).Image = 33 'N� Serie
'        .Buttons(15).Image = 40 'Imprimir etiquetas estanteria
'        .Buttons(16).Image = 16 'Imprimir Albaran proveedor (REA)
'
'        .Buttons(17).Image = 15  'Salir
'        .Buttons(btnPrimero).Image = 6  'Primero
'        .Buttons(btnPrimero + 1).Image = 7 'Anterior
'        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
'        .Buttons(btnPrimero + 3).Image = 9 '�ltimo
'    End With

    With Me.Toolbar1
        .HotImageList = frmPpal.imgListComun_OM2
        .DisabledImageList = frmPpal.imgListComun_BN2
        .ImageList = frmPpal.ImgListComun2
        .Buttons(1).Image = 3
        .Buttons(2).Image = 4
        .Buttons(3).Image = 5
        .Buttons(5).Image = 1
        .Buttons(6).Image = 2
        .Buttons(8).Image = 16
    End With

    With Me.Toolbar5
        .HotImageList = frmPpal.imgListComun_OM2
        .DisabledImageList = frmPpal.imgListComun_BN2
        .ImageList = frmPpal.ImgListComun2
        .Buttons(1).Image = 45 'Cambiar proveedor
        .Buttons(2).Image = 32 'Mover al hco
        .Buttons(3).Image = 47 '21 'Recepcionar
        .Buttons(4).Image = 42 '33 'Nro Series
        .Buttons(5).Image = 40 'Imprimir etiquetas de estanter�a
    End With
    
    ' desplazamiento
    With Me.ToolbarDes
        .HotImageList = frmPpal.imgListComun_OM2
        .DisabledImageList = frmPpal.imgListComun_BN2
        .ImageList = frmPpal.ImgListComun2
        .Buttons(1).Image = 6
        .Buttons(2).Image = 7
        .Buttons(3).Image = 8
        .Buttons(4).Image = 9
    End With
    
    For i = 0 To ToolAux.Count - 1
        With Me.ToolAux(i)
            .HotImageList = frmPpal.imgListComun_OM16
            .DisabledImageList = frmPpal.imgListComun_BN16
            .ImageList = frmPpal.imgListComun16
            .Buttons(1).Image = 3   'Insertar
            .Buttons(2).Image = 4   'Modificar
            .Buttons(3).Image = 5   'Borrar
        End With
    Next i



    Me.SSTab1.Tab = 0
    LimpiarCampos   'Limpia los campos TextBox
    cmdaux(2).Tag = "-1"
   
    'Sept 2010
    'Todos podran imprimirse un albaran
    'La imprimir solo es posible para albaranes a socios(REA)
    'Toolbar1.Buttons(12).visible = vParamAplic.IVA_REA > 0
    
    CodTipoMov = "ALC"
    VieneDeBuscar = False

    '## A mano
     Me.FrameHco.visible = EsHistorico
    
    If Not EsHistorico Then
        NombreTabla = "scaalp"
        NomTablaLineas = "slialp" 'Tabla lineas de Albaranes
        Me.Caption = "Albaranes Proveedores"
        Ordenacion = " ORDER BY numalbar, fechaalb,codprove "
    Else
        NombreTabla = "schalp"
        NomTablaLineas = "slhalp"
        CargarTagsHco Me, "scaalp", NombreTabla
        'Estos campos solo estan en la tabla del hist�rico
        Text1(22).Tag = "Fecha Eliminaci�n|F|N|||schalp|fechelim|dd/mm/yyyy|N|"
        Text1(23).Tag = "Trabajador Eliminaci�n|N|N|0|9999|schalp|trabelim|0000|N|"
        Text1(24).Tag = "Incidencia elim.|T|N|||schalp|codincid||N|"
        Me.Caption = "Hist�rico Albaranes Proveedores"
        Ordenacion = " ORDER BY numalbar,fechaalb,codprove "
    End If
    
         
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    'ASignamos un SQL al DATA1
    CadenaConsulta = "Select * from " & NombreTabla
    
    If hcoCodMovim <> "" Then
    'Se llama desde Dobleclick en frmAlmMovimArticulos
        CadenaConsulta = CadenaConsulta & " WHERE numalbar='" & hcoCodMovim & "' AND fechaalb= """ & Format(hcoFechaMovim, "yyyy-mm-dd") & """"
        CadenaConsulta = CadenaConsulta & " AND codprove=" & hcoCodProve
    ElseIf cadSelAlbaranes <> "" Then
        CadenaConsulta = CadenaConsulta & " WHERE " & cadSelAlbaranes
    Else
        CadenaConsulta = CadenaConsulta & " WHERE numalbar = -1"
    End If
    Data1.ConnectionString = conn
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
       
    'Cargar el DataGrid de lineas de Revisiones inicialmente a nada DATA2
    PrimeraVez = True
    PrimeraVezForm = True
    
    If hcoCodMovim = "" Then
        If DatosADevolverBusqueda = "" Then
            PonerModo 0
        Else
            PonerModo 1
            Text1(0).BackColor = vbLightBlue
        End If
        'Poner los grid sin apuntar a nada
        LimpiarDataGrids
    End If
    AlmacenLineas = -1
End Sub


Private Sub LimpiarCampos()
On Error Resume Next
    limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador.Caption = ""
    
    'Aqui va el especifico de cada form es
    '### a mano
    chkDocArchi.Value = 0
    Text3(0).Text = "BASE IMPONIBLE"
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
End Sub



Private Sub frmAlm_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Almacenes Propios
    txtAux(0).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000") 'Cod Almacen
End Sub

Private Sub frmArt_DatoSeleccionado(CadenaSeleccion As String)
'Mantenimiento de Articulos
    txtAux(1).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod Artic
End Sub

Private Sub frmB_DatoSeleccionado(CadenaSeleccion As String)
Dim cadB As String
Dim Aux As String
      
    If CadenaSeleccion <> "" Then
            cadB = ""
            Aux = ValorDevueltoFormGrid(Text1(0), CadenaSeleccion, 1)
            cadB = Aux
            Aux = ValorDevueltoFormGrid(Text1(1), CadenaSeleccion, 2)
            cadB = cadB & " and " & Aux
            Aux = ValorDevueltoFormGrid(Text1(4), CadenaSeleccion, 3)
            cadB = cadB & " and " & Aux
            
            CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
            PonerCadenaBusqueda
    End If
    Screen.MousePointer = vbDefault

End Sub

Private Sub frmCCos_DatoSeleccionado(CadenaDevuelta As String)
    If CadenaDevuelta <> "" Then
        Me.txtAux(9).Text = RecuperaValor(CadenaDevuelta, 1)
        Me.txtAux2(9).Text = RecuperaValor(CadenaDevuelta, 2)
    End If
End Sub

Private Sub frmCP_DatoSeleccionado(CadenaSeleccion As String)
'Formulario Mantenimiento C. Postales
Dim Indice As Byte
Dim devuelve As String

        Indice = 9
        Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'CPostal
        Text1(Indice + 1).Text = ObtenerPoblacion(Text1(Indice).Text, devuelve) 'Poblacion
        'provincia
        Text1(Indice + 2).Text = devuelve
End Sub


Private Sub frmF_Selec(vFecha As Date) 'Calendario Fechas
    Text1(CByte(Me.imgFecha(0).Tag)).Text = Format(vFecha, "dd/mm/yyyy")
End Sub


Private Sub frmFE_DatoSeleccionado(CadenaSeleccion As String)
    Text1(26).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod envio
    Text2(26).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom envio

End Sub

Private Sub frmFP_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Formas de Pago
Dim Indice As Byte
    Indice = 12
    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod Forma Pago
    FormateaCampo Text1(Indice)
    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Forma Pago
End Sub


Private Sub frmList_DatoSeleccionado(CadenaSeleccion As String)
'devuelve los datos necesarios para grabar en la tabla del historico al eliminar albaran
    cadList = ""
    cadList = DBSet(RecuperaValor(CadenaSeleccion, 1), "F") & " as fechelim,"
    cadList = cadList & RecuperaValor(CadenaSeleccion, 2) & " as trabelim,"
    cadList = cadList & DBSet(RecuperaValor(CadenaSeleccion, 3), "T") & " as codincid"
End Sub



Private Sub frmMen_DatoSeleccionado(CadenaSeleccion As String)
Dim Cant As Currency
Dim i As Byte
Dim cadSerie As String
Dim nSerie As CNumSerie

'si llegamos aqui hemos hecho un abono y vamos a eliminar el
'n� de serie de la tabla sserie del articulo que hemos devuelto.

    Cant = CCur(txtAux(3).Text)
    Cant = Abs(Cant)

    'Para cada valor empipado actualizar la tabla sserie
    On Error GoTo ErrorNSerie

    For i = 1 To Cant
        cadSerie = RecuperaValor(CadenaSeleccion, i + 1) 'Cod Forma Pago
        If cadSerie <> "" Then
            Set nSerie = New CNumSerie
            nSerie.numSerie = cadSerie
            nSerie.Articulo = RecuperaValor(CadenaSeleccion, 1)
            
            'como vamos a devolver esos n� serie de ese articulo
            'los eliminamos de la tabla sserie, ya no tenemos esos art�culos
            nSerie.EliminarNumSerie
            Set nSerie = Nothing
        End If
    Next i

ErrorNSerie:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Actualizar tabla N� Series", Err.Description
    End If
End Sub



Private Sub frmNSerie_CargarNumSeries()
'Cuando vuelve del formulario donde se han introducido los n� de Serie a cargar
'Insertar un registro en la tabla "sserie" para cada articulo

    'Estamos en COMPRAS
    If ModificaLineas = 4 Then
        'Viene de boton VErNumSeries de la toolbar, abre la ventana de cargar numSeries
        'y muestra los que tenga asignados el albaran
        CargarNumSeries
    Else
       'Viene de insertar N� de series al insertar una linea y pasa
       'los valores almacenados en la temporal a la sserie
       InsertarNumSeriesDeTMP
    End If
End Sub


Private Sub frmProve_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Proveedores
    Text1(4).Text = RecuperaValor(CadenaSeleccion, 1)  'Cod Prove
    FormateaCampo Text1(4)
End Sub

Private Sub frmPV_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Proveedores Varios
Dim Indice As Byte

    Indice = 6
    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'NIF
    Text1(Indice - 1).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom prove
    PonerDatosProveVario (Text1(Indice).Text)
End Sub

Private Sub frmT_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Trabajadores
'Dim Indice As Byte

'    Indice = 2
    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod Trabajador
    FormateaCampo Text1(Indice)
    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Trabajador
End Sub


Private Sub imgBuscar_Click(Index As Integer)
'Dim Indice As Byte

    If Index = 9 Then
        If Modo = 0 Then Exit Sub
    Else
        If Modo = 2 Or Modo = 0 Then Exit Sub
    End If
    Screen.MousePointer = vbHourglass

    Select Case Index
        Case 0 'Cod. Prove
            PonerFoco Text1(4)
            Set frmProve = New frmBasico2
'            frmProve.DatosADevolverBusqueda = "0"
'            frmProve.Show vbModal
            AyudaProveedores frmProve, Text1(4)
            Set frmProve = Nothing
            Indice = 4
            
        Case 1 'Realizada Por Trabajador
            Indice = 2
'            Set frmT = New frmAdmTrabajadores
'            frmT.DatosADevolverBusqueda = "0"
'            frmT.Show vbModal
            Set frmT = New frmBasico2
            AyudaTrabajadores frmT, Text1(Indice)
            Set frmT = Nothing
            
            
        Case 2 'Cod. Postal
            Indice = 9
            If Not Text1(Indice).Locked Then
                Set frmCP = New frmCPostal
                frmCP.DatosADevolverBusqueda = "0"
                frmCP.Show vbModal
                Set frmCP = Nothing
                
                VieneDeBuscar = True
            End If
            
        Case 3 'Forma de Pago
            Indice = 12
'            PonerFoco Text1(Indice)
'            Set frmFP = New frmFacFormasPago
'            frmFP.DatosADevolverBusqueda = "0"
'            frmFP.Show vbModal
            Set frmFP = New frmBasico2
            AyudaFormasPago frmFP, Text1(Indice)
            Set frmFP = Nothing
            PonerFoco Text1(Indice)
            
        Case 5 'NIF proveedor varios
            Set frmPV = New frmComProveV
            frmPV.DatosADevolverBusqueda = "0"
            frmPV.Show vbModal
            Set frmPV = Nothing
            Indice = 6
        Case 8
            Indice = 26
            PonerFoco Text1(Indice)
            Set frmFE = New frmFacFormasEnvio
            frmFE.DatosADevolverBusqueda = "0"
            frmFE.Show vbModal
            Set frmFE = Nothing
            
        Case 9
            '++
            Indice = 16
                CadenaDesdeOtroForm = Text2(16).Text
                frmFacClienteObser.Modificar = Modo = 5 And ModificaLineas > 0
                frmFacClienteObser.Text1 = CadenaDesdeOtroForm
                frmFacClienteObser.Show vbModal
                'Llevara DOS VALORES.
                'Si modifica y el texto
                If Modo = 5 And ModificaLineas > 0 Then
                    If RecuperaValor(CadenaDesdeOtroForm, 1) = "1" Then Text2(16).Text = Mid(CadenaDesdeOtroForm, 3)
                End If
                CadenaDesdeOtroForm = ""
        '++
        Case 4 ' trabajador pedido
            Indice = 21
            Set frmT = New frmBasico2
            AyudaTrabajadores frmT, Text1(Indice)
            Set frmT = Nothing
        Case 6 ' trabajador que elimino el albaran
            Indice = 23
            Set frmT = New frmBasico2
            AyudaTrabajadores frmT, Text1(Indice)
            Set frmT = Nothing
        
    End Select
    PonerFoco Text1(Indice)
    Screen.MousePointer = vbDefault
End Sub


Private Sub imgFecha_Click(Index As Integer) 'Abre calendario Fechas
Dim Indice As Byte

    If Modo = 2 Or Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass
    Set frmF = New frmCal
    frmF.Fecha = Now
    If Index = 0 Then
        Indice = 1 'fecalb
    ElseIf Index = 1 Then
        Indice = 25
    ElseIf Index = 3 Then
        Indice = 30
    Else
        Indice = 27
    End If
    Me.imgFecha(0).Tag = Indice
    PonerFormatoFecha Text1(Indice)
    If Text1(Indice).Text <> "" Then frmF.Fecha = CDate(Text1(Indice).Text)
    Screen.MousePointer = vbDefault
    frmF.Show vbModal
    Set frmF = Nothing
    PonerFoco Text1(Indice)
End Sub


Private Sub mnBuscar_Click()
    Me.SSTab1.Tab = 0
    BotonBuscar
End Sub


Private Sub mnEliminar_Click()
    If Modo = 5 Then 'Eliminar lineas de Pedido
         BotonEliminarLinea
    Else   'Eliminar Pedido
         BotonEliminar
         Screen.MousePointer = vbDefault
    End If
End Sub


Private Sub mnLineas_Click()
    BotonMtoLineas 0, "Albaranes"
End Sub


Private Sub mnModificar_Click()
    If Modo = 5 Then 'Modificar lineas
         BotonModificarLinea
    Else   'Modificar Cabecera Albaran
        If cadSelAlbaranes = "" Then
            If Not BLOQUEADesdeFormulario(Me) Then Exit Sub
        End If
        BotonModificar
    End If
End Sub


Private Sub mnNuevo_Click()
    If Modo = 5 Then 'A�adir lineas
         BotonAnyadirLinea
    Else 'A�adir Cabecera de Albaran
         Me.SSTab1.Tab = 0
         BotonAnyadir
    End If
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

'### A mano
'Los metodos del text tendran que estar
'Los descomentamos cuando esten puestos ya los controles
Private Sub Text1_GotFocus(Index As Integer)
    kCampo = Index
    If Index = 9 Then HaCambiadoCP = False 'CPostal
    ConseguirFoco Text1(Index), Modo
End Sub


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim Ind As Integer
Dim b As Boolean
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
    
    If KeyCode = 43 Or KeyCode = 107 Or KeyCode = 187 Then
        b = False
        If Text1(Index).Text = "" Then
            b = True
        Else
            If Text1(Index).SelLength = Len(Text1(Index).Text) Then b = True
        End If
        If b Then
                Ind = -1
                Select Case Index
                Case 2
                    Ind = 1

                Case 4
                    Ind = 0
                Case 6
                    Ind = 5
                Case 9
                    Ind = 2
                Case 12
                    Ind = 3
                Case 21
                    Ind = 4
                Case 26
                    Ind = 8

                End Select
                If Ind >= 0 Then
                    PulsadoMas2 = True
                    PulsarTeclaMas True, Ind
                End If
            End If
        End If
End Sub


Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
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
        
        
    If PulsadoMas2 Then
        'Para que cuando pulse el mas abra el form
        PulsadoMas2 = False
        Text1(Index).Text = ""
        Exit Sub
    End If
    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
          
    'Si queremos hacer algo ..
    Select Case Index
        Case 1, 25, 27, 30 'Fecha Albaran y fecha arhivo
                If Text1(Index).Text <> "" Then PonerFormatoFecha Text1(Index)
                
        Case 2 'Cod Trabajador
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "straba", "nomtraba", "codtraba")
            Else
                Text2(Index).Text = ""
            End If
            
        Case 4 'Cod. Proveedor
            If PonerFormatoEntero(Text1(Index)) Then
                If Modo = 1 Then 'Busqueda
                    'Text1(5).Text = PonerNombreDeCod(Text1(Index), conAri, "sprove", "nomprove")
                Else 'Si Insertar, recuperar datos de Tabla sprove
                    PonerDatosProveedor (Text1(Index).Text)
                End If
            Else
                LimpiarDatosProve
            End If
            
         Case 6 'NIF
            If Not EsDeVarios Or Modo <> 3 Then Exit Sub
            If Modo = 4 Then 'Modificar
                'si no se ha modificado el nif del cliente no hacer nada
                If Text1(6).Text = Data1.Recordset!nifProve Then
                    Exit Sub
                End If
            End If
            PonerDatosProveVario (Text1(Index).Text)
             
        Case 9 'Cod. Postal
            If Text1(Index).Locked Then Exit Sub
            If Text1(Index).Text = "" Then
                Text1(Index + 1).Text = ""
                Text1(Index + 2).Text = ""
                Exit Sub
            End If
            If (Not VieneDeBuscar) Or (VieneDeBuscar And HaCambiadoCP) Then
                Text1(Index + 1).Text = ObtenerPoblacion(Text1(Index).Text, devuelve)
                Text1(Index + 2).Text = devuelve
            End If
            VieneDeBuscar = False
            
        Case 12 'Forma de Pago
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "sforpa", "nomforpa")
            Else
                Text2(Index).Text = ""
            End If
            
        Case 13, 14 'Descuentos
            If PonerFormatoDecimal(Text1(Index), 4) Then 'Tipo 4: Decimal(4,2)
                If Modo = 4 Then CalcularDatosFactura
'--
'                If Index = 14 Then
'                    Me.SSTab1.Tab = 1
'                    PonerFoco Text1(15)
'                End If
'            Else
'                If Index = 14 And Text1(Index).Text = "" Then
'                    Me.SSTab1.Tab = 1
'                    PonerFoco Text1(15)
'                End If
            End If
        Case 19
            PonerFocoBtn Me.cmdAceptar
                
                    
        Case 26 'Codenvio
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "senvio", "nomenvio")
            Else
                Text2(Index).Text = ""
            End If
    End Select
End Sub


Private Sub HacerBusqueda()
Dim cadB As String

    cadB = ObtenerBusqueda(Me, False)
    If cadSelAlbaranes <> "" Then cadB = cadB & " AND " & cadSelAlbaranes
    If chkVistaPrevia = 1 Then
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
Dim tabla As String
Dim Titulo As String

'    'Llamamos a al form
'    '##A mano
'    cad = ""
'    cad = cad & ParaGrid(Text1(0), 20, "N� Albaran")
'    cad = cad & ParaGrid(Text1(1), 15, "Fecha Alb.")
'    cad = cad & ParaGrid(Text1(4), 15, "Provedor")
'    cad = cad & ParaGrid(Text1(5), 50, "Nombre Prov.")
'    tabla = NombreTabla
'    Titulo = "Albaranes"
'
'    If cad <> "" Then
'        Screen.MousePointer = vbHourglass
'        Set frmB = New frmBuscaGrid
'        frmB.vCampos = cad
'        frmB.vTabla = tabla
'        frmB.vSQL = cadB
'        HaDevueltoDatos = False
'        '###A mano
'        frmB.vDevuelve = "0|1|2|"
'        frmB.vTitulo = Titulo
'        frmB.vselElem = 0
'        frmB.vConexionGrid = conAri  'Conexi�n a BD: Ariges
'
'        frmB.Show vbModal
'        Set frmB = Nothing
'        'Si ha puesto valores y tenemos que es formulario de busqueda entonces
'        'tendremos que cerrar el form lanzando el evento
''        If HaDevueltoDatos Then
''''            If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
''''                cmdRegresar_Click
''        Else   'de ha devuelto datos, es decir NO ha devuelto datos
''            PonerFoco Text1(kCampo)
'        'End If
'    End If
'    Screen.MousePointer = vbDefault
    
    Set frmB = New frmBasico2
    AyudaAlbaranesCompra frmB, NombreTabla, Text1(0)
    Set frmB = Nothing
    
End Sub


Private Sub PonerCadenaBusqueda()
Screen.MousePointer = vbHourglass
On Error GoTo EEPonerBusq

    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        MsgBox "No hay ning�n registro en la tabla " & NombreTabla, vbInformation
        Screen.MousePointer = vbDefault
        If Modo = 1 Then
            PonerFoco Text1(0)
            Text1(0).BackColor = vbLightBlue
        End If
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
'Carga las Pesta�as con las tablas de lineas del Trabajador seleccionado para mostrar
On Error GoTo EPonerLineas

    Screen.MousePointer = vbHourglass

    'Datos de la tabla slipre
    CargaGrid DataGrid1, Data2, True

    BotonesToolBarAux

    If Data2.Recordset.EOF Then Text2(16).Text = ""
    
    Screen.MousePointer = vbDefault
    Exit Sub
EPonerLineas:
    MuestraError Err.Number, "PonerCamposLineas"
    PonerModo 2
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerCampos()
On Error Resume Next

    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma Me, Data1
    
    'Trabajador Albaran
    Text2(2).Text = PonerNombreDeCod(Text1(2), conAri, "straba", "nomtraba", "codtraba")
    Text2(12).Text = PonerNombreDeCod(Text1(12), conAri, "sforpa", "nomforpa")
    
    'Trabajador del Pedido
    Text2(21).Text = PonerNombreDeCod(Text1(21), conAri, "straba", "nomtraba", "codtraba")
    'Cod envio
    Text2(26).Text = PonerNombreDeCod(Text1(26), conAri, "senvio", "nomenvio")
    
    
    PonerCamposLineas 'Pone los datos de las tablas de lineas de Ofertas
    
    If EsHistorico Then
        'poner datos de eliminacion
        Text2(23).Text = PonerNombreDeCod(Text1(23), conAri, "straba", "nomtraba", "codtraba")
        Text2(24).Text = PonerNombreDeCod(Text1(24), conAri, "sincid", "nomincid", "codincid")
    End If
    
    CalcularDatosFactura 'rellenar campos pesta�a de totales
     
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    If Err.Number <> 0 Then Err.Clear
End Sub


'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
Private Sub PonerModo(Kmodo As Byte)
Dim i As Byte, NumReg As Byte
Dim b As Boolean

    On Error GoTo EPonerModo

    'Vuelvo a poner el lbl en la columna
    If Modo = 5 Then DataGrid1.Columns(5).Caption = "Art�culo"
    lblF.Caption = ""
    
    

    
    'Actualiza Iconos Insertar,Modificar,Eliminar
'    ActualizarToolbarGnral Me.Toolbar1, Modo, Kmodo, btnAnyadir
    
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
    
    '22 Sept 2010
    'El albaran lo puede imprimir en cualquier empresa
    'If vParamAplic.IVA_REA > 0 Then Toolbar1.Buttons(12).Enabled = b
    Toolbar5.Buttons(2).Enabled = b
    
    'Poner Flechas de desplazamiento visibles
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
'    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, NumReg
    DesplazamientoVisible b And Data1.Recordset.RecordCount > 1
        
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar adem�s limpia los campos Text1 y bloquea la clave primaria
    BloquearText1 Me, Modo
    
    'Campo N� Albaran siempre bloqueado, excepto si estamos en modo de busqueda
    BloquearTxt Text1(0), (Modo <> 1) And (Modo <> 3), True
    
    'La fecha de albaran es clave primaria pero dejamos modificarla
    BloquearTxt Text1(1), (Modo = 0 Or Modo = 2 Or Modo = 5)
    b = (Modo <> 1)
    'Bloquear los campos de Pedido, excepto en Busqueda
    BloquearTxt Text1(3), b
    BloquearTxt Text1(20), b
    BloquearTxt Text1(21), b
    
    
   
    'datos cliente siempre bloqueados hasta que sea de varios
    If Modo = 3 Then
        EsDeVarios = False
        BloquearDatosProve (EsDeVarios)
    End If
     
    '-----  Datos Totales de Factura siempre bloqueado
    For i = 33 To 50
        BloquearTxt Text3(i), True
    Next i
    'Campo B.Imp y Imp. IVA siempre en azul
    Text3(36).BackColor = &HFFFFC0
    Text3(46).BackColor = &HFFFFC0
    Text3(47).BackColor = &HFFFFC0
    Text3(48).BackColor = &HFFFFC0
    Text3(49).BackColor = &HC0C0FF    'Tatal factura
    Text3(50).BackColor = &HC0C0FF    'Tatal factura
    '---------------------------------------------------
          
    'Si no es modo lineas Boquear los TxtAux
    For i = 0 To txtAux.Count - 1
        BloquearTxt txtAux(i), (Modo <> 5)
    Next i
    BloquearTxt Text2(16), (Modo <> 5)
    
    
    '---------------------------------------------
    b = (Modo = 3 Or Modo = 4 Or Modo = 1)
    cmdCancelar.visible = b
    cmdAceptar.visible = b
    If cmdCancelar.visible Then cmdCancelar.Cancel = True
    chkDocArchi.Enabled = b
        
    
    
    
    For i = 0 To Me.imgFecha.Count - 1
'        Me.imgFecha(i).Enabled = b
        BloquearImg imgFecha(i), Not b
    Next i
    
    For i = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(i).Enabled = b
    Next i
    Me.imgBuscar(4).Enabled = (Modo = 1)
    Me.imgBuscar(0).Enabled = (Modo = 3 Or Modo = 1)
    imgBuscar(9).Enabled = True
    
    
    'Modo Linea de Albaranes. Campo Ampliacion Linea
  '  Me.Label1(35).visible = (Modo = 5)
  '  Me.Text2(16).visible = (Modo = 5)
  '  imgBuscar(9).visible = (Modo = 5)
    BloquearTxt Text2(16), True
    'Modo Linea de Albaranes. Campo num_lote
    Me.Label1(2).visible = (Modo = 5)
    Me.Text2(17).visible = (Modo = 5)
    BloquearTxt Text2(17), True
    
    ' ---- [20/10/2009] [LAURA] : a�adir del centro de coste
    Me.Label1(46).visible = (vEmpresa.TieneAnalitica) And (Modo = 5)
    Me.txtAux2(9).visible = (vEmpresa.TieneAnalitica) And (Modo = 5)
    BloquearTxt txtAux2(9), True
    
       
    Me.chkVistaPrevia.Enabled = (Modo <= 2)
       
    'Poner el tama�o de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos
    
    PonerModoOpcionesMenu (Modo) 'Activar opciones de menu seg�n modo
    PonerOpcionesMenu 'Activar opciones de menu seg�n nivel de permisos del usuario
    
    
    
    
EPonerModo:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub DesplazamientoVisible(bol As Boolean)
    FrameDesplazamiento.visible = bol
    FrameDesplazamiento.Enabled = bol
End Sub



Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de b�squeda o no
'para los campos que permitan introducir criterios m�s largos del tama�o del campo
    PonerLongCamposGnral Me, Modo, 1
End Sub


Private Function DatosOk() As Boolean
'Comprobar que los datos de la cabecera son correctos antes de Insertar o Modificar
'la cabecera del Pedido
Dim b As Boolean
Dim cad As String

    On Error GoTo EDatosOK

    DatosOk = False
       
    b = CompForm(Me, 1) 'Comprobar formato datos ok
    If Not b Then Exit Function
    
    If Abs(DateDiff("m", CDate(Text1(1).Text), CDate(Text1(30).Text))) > 3 Then
        MsgBox "La diferencia entre fecha albaran y entrada mercancia mayor 3 meses", vbExclamation
        If vUsu.Nivel > 1 Then b = False: Exit Function
    End If
    
    
    
    
    If Modo = 3 Then
        'Comprobare que este albaran/proveedor NO existe
        'a que es campo clave en scafpc1
        cad = " year(fechaalb)=" & Year(CDate(Text1(1).Text)) & " AND numalbar=" & DBSet(Text1(0).Text, "T") & " AND codprove"
        cad = DevuelveDesdeBD(conAri, "fechaalb", "scaalp", cad, Text1(4).Text, "N")
        
        If cad = "" Then
            cad = " year(fechaalb)=" & Year(CDate(Text1(1).Text)) & " AND numalbar=" & DBSet(Text1(0).Text, "T") & " AND codprove"
            cad = DevuelveDesdeBD(conAri, "fechaalb", "scafpa", cad, Text1(4).Text, "N")
            If cad <> "" Then cad = cad & " (Facturado)"
        End If
        
        If cad <> "" Then
            cad = Text1(5).Text & vbCrLf & "con fecha " & cad
            cad = "Ya existe el albaran: " & Text1(0).Text & " del proveedor " & vbCrLf & Text1(4).Text & " " & cad
            
            MsgBox cad, vbExclamation
            b = False
        End If
    End If
    
    
    'Febrero 208
    'LLegado aqui veremos las fechas son "razonables"
    If Modo = 3 And b Then
        cad = ""
        If CDate(Text1(1).Text) < vEmpresa.FechaIni Then
            cad = "menor "
        Else
            If CDate(Text1(1).Text) > DateAdd("yyyy", 1, vEmpresa.FechaFin) Then cad = "mayor"
        End If
        If cad <> "" Then
            cad = "Fecha " & cad & " que ejercicios.          " & Text1(1).Text & vbCrLf & vbCrLf & "�Continuar?"
            If MsgBox(cad, vbQuestion + vbYesNoCancel) <> vbYes Then b = False
        End If
    End If
        
        
    If Modo = 4 And b Then
        If vParamAplic.ManipuladorFitosanitarios2 Then
            If CDate(Text1(1).Text) <> Data1.Recordset!FechaAlb Then
                'Si tienen lotes asignados NO podremos seguir
                If Not PuedeEliminar(False) Then b = False
            End If
        End If
    End If
    
    DatosOk = b
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Function DatosOkLinea() As Boolean
Dim b As Boolean
Dim i As Byte
Dim cart As CArticulo
Dim Aux As String
Dim DiferenciaCantidad As Currency

    On Error GoTo EDatosOkLinea

    DatosOkLinea = False
    
    
    'Febrero 2010   Si han apretado Alt+A NO recalcula
    '----------------------------------------------------------------------------------
    'txtAux(8).Text = CalcularImporte(txtAux(3).Text, txtAux(4).Text, txtAux(6).Text, txtAux(7).Text, vParamAplic.TipoDtos)
    Aux = RecalculoImporteLineas(txtAux(3), txtAux(4), txtAux(5), txtAux(6), vParamAplic.TipoDtos)
    Aux = Format(Aux, FormatoImporte)
    If Aux <> txtAux(7).Text Then
        Aux = "Importe linea distinto calculado: " & Aux & "  <>  " & txtAux(7).Text & vbCrLf & vbCrLf & "�Continuar?"
        If MsgBox(Aux, vbQuestion + vbYesNo) = vbNo Then Exit Function
    End If
    

    
    
    
    b = True
    'Comprobar que los campos requeridos tengan valor
    For i = 0 To txtAux.Count - 1
        If txtAux(i).Text = "" Then
            If i = 9 And vEmpresa.TieneAnalitica = False Then
                'no hace nada pq puede ser nulo
            Else
                Screen.MousePointer = vbDefault
                MsgBox "El campo " & txtAux(i).Tag & " no puede ser nulo", vbExclamation
                b = False
                PonerFoco txtAux(i)
                Exit Function
            End If
        End If
    Next i
    
    
    'si el articulo tiene control de numero de lotes, el campo del lote ser� requerido
    Set cart = New CArticulo
    If cart.LeerDatos(txtAux(1).Text) Then
        If cart.TieneNumLote Then
            If Trim(Text2(17).Text) = "" Then
                b = False
                MsgBox "El n� de lote no puede ser nulo." & vbCrLf & vbCrLf & "El art�culo tiene control de lotes.", vbExclamation
                PonerFoco Text2(17)
            End If
            
            
            'Cuando lleve Registro fitosanitario entonces el modificar articulo "cantidad" sera SI no se ha vendido NADA
            '------------------------------------------------------------------------------
            If b And vParamAplic.ManipuladorFitosanitarios2 Then
            
            
                'Modificando
                If ModificaLineas = 2 Then
                    
                    cadList = ""
                    'Si cambia cantidad vendida o n� lote entonces comprobaremos
                    If Data2.Recordset!numlotes <> Me.Text2(17).Text Then
                        cadList = "LO"
                    ElseIf Data2.Recordset!cantidad <> ImporteFormateado(Me.txtAux(3).Text) Then
                        cadList = "OK"
                    End If
            
                    If cadList <> "" Then
                        
                        If cadList = "LO" Then
                            
                            cadList = "numlote =" & DBSet(Data2.Recordset!numlotes, "T") & " AND fecentra = " & DBSet(Data1.Recordset!FechaAlb, "F")
                            cadList = cadList & " AND codartic = " & DBSet(Data2.Recordset!codArtic, "T") & " AND 1"
                            cadList = DevuelveDesdeBD(conAri, "count(*)", "slialblotes", cadList, "1")
                            If Val(cadList) > 0 Then
                                MsgBox "El lote ya ha sido vendido", vbExclamation
                                b = False
                            Else
                                cadList = "numlote =" & DBSet(Data2.Recordset!numlotes, "T") & " AND fecentra = " & DBSet(Data1.Recordset!FechaAlb, "F")
                                cadList = cadList & " AND codartic = " & DBSet(Data2.Recordset!codArtic, "T") & " AND 1"
                                cadList = DevuelveDesdeBD(conAri, "count(*)", "slivenlotes", cadList, "1")
                                If Val(cadList) > 0 Then
                                    MsgBox "El lote esta siendo vendido", vbExclamation
                                    b = False
                                End If
                            End If
                        Else
                            'Es el mismo lote de venta. Vemos si cambiando la cantidad tiene bastante
                            DiferenciaCantidad = Data2.Recordset!cantidad - ImporteFormateado(Me.txtAux(3).Text)
                            If DiferenciaCantidad > 0 Then
                                'Hemos puesto menos cantidad de la que habiamos puesto
                                cadList = "numlotes =" & DBSet(Data2.Recordset!numlotes, "T") & " AND fecentra = " & DBSet(Data1.Recordset!FechaAlb, "F")
                                cadList = cadList & " AND codartic = " & DBSet(Data2.Recordset!codArtic, "T") & " AND 1"
                                cadList = DevuelveDesdeBD(conAri, "(canentra-vendida)", "slotes", cadList, "1")
                                If cadList = "" Then cadList = "0"
                                If DiferenciaCantidad > CCur(cadList) Then
                                      MsgBox "Cantidad disponible insuficiente", vbExclamation
                                      b = False
                                End If
                            End If
                        End If
                    End If
                
                
                Else
                    'Insertando linea
                    'No puede poner el mismo articulo, lote en el mismo albaran

                    cadList = "numlotes =" & DBSet(Text2(17).Text, "T") & " AND fechaalb = " & DBSet(Data1.Recordset!FechaAlb, "F")
                    cadList = cadList & " AND codartic = " & DBSet(txtAux(1).Text, "T") & " AND 1"
                    cadList = DevuelveDesdeBD(conAri, "count(*)", "slialp", cadList, "1")
                    If Val(cadList) > 0 Then
                        'NO PUEDE
                        MsgBox "Ya existe el mismo articulo-lote-albaran ", vbExclamation
                        b = False
                    End If
                End If ' manniulador
            End If 'ModificaLineas=2
            
            
            
        End If
    End If
    Set cart = Nothing
    
'    If Me.Text2(17).Locked = False Then
'        If Trim(Text2(17).Text) = "" Then
'            b = False
'            MsgBox "El n� de lote no puede ser nulo." & vbCrLf & vbCrLf & "El art�culo tiene control de lotes.", vbExclamation
'        End If
'    End If
    
        
    DatosOkLinea = b
EDatosOkLinea:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function



Private Sub Text2_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 40 Then 'campo num_lote y Flecha hacia abajo
        If Index = 16 And Text2(17).Locked Then PonerFocoBtn Me.cmdAceptar
        If Index = 17 Then PonerFocoBtn Me.cmdAceptar
    Else
        If Index <> 16 Then KEYdown KeyCode
    End If
End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then 'campo Amliacion Linea y ENTER
       If Index = 16 And Text2(17).Locked Then
            PonerFocoBtn Me.cmdAceptar
       ElseIf Index = 17 Then
            PonerFocoBtn Me.cmdAceptar
            
        Else
            KEYpress KeyAscii
        End If
    End If
End Sub

Private Sub Text2_LostFocus(Index As Integer)
    'quitamos los espacios en blanco
    Text2(Index).Text = Trim(Text2(Index).Text)
    
    If Index = 16 And (Text2(Index).Locked = False) Then Text2(Index).Text = UCase(Text2(Index).Text)
End Sub

Private Sub ToolAux_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            BotonAnyadirLinea
        Case 2
            BotonModificarLinea
        Case 3
            BotonEliminarLinea
        Case Else
    End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1  'Nuevo
            mnNuevo_Click
        Case 2  'Modificar
            mnModificar_Click
        Case 3  'Borrar
            mnEliminar_Click
        Case 5  'Buscar
            mnBuscar_Click
        Case 6  'Todos
            mnVerTodos_Click
        Case 8
            'Imprimir SOLO si lleva REA
            Imprimir
    End Select
End Sub

Private Sub Imprimir()
  
        'Es la impresion de los albaranes de socios

        CadenaDesdeOtroForm = "|||"
        If Text1(0).Text <> "" Then
            
            'vOY A CARGAR LOS DATOS
            CadenaDesdeOtroForm = Text1(0).Text & "|" & Text1(1).Text & "|" & Text1(4).Text & "|" & Text1(5).Text & "|"
        End If
        frmListado2.Opcion = 10
        frmListado2.Show vbModal
    
End Sub


Private Sub ImpirmirEtiqEsta()
    If Modo <> 2 Then Exit Sub
    If Me.Data2.Recordset.EOF Then Exit Sub
    
    frmListado.OpcionListado = 513
    frmListado.CadTag = "Albaran: " & Text1(0).Text & " de " & Text1(1).Text & ": " & Text1(4).Text & "-" & Text1(5).Text
    frmListado.NumCod = Data2.RecordSource
    frmListado.Show vbModal
    
End Sub


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub

    
Private Function InsertarLinea(numlinea As String) As Boolean
'Inserta un registro en la tabla de lineas de Albaranes: slialb
'OUT -> NumLinea: devuelve el N� de linea que acaba de insertar
Dim SQL As String
Dim b As Boolean
Dim vCStock As CStock
Dim vArtic As CArticulo
Dim MenError As String
Dim DentroTRANS As Boolean
Dim ImpReciclado As Single
    
    InsertarLinea = False
    SQL = ""
    DentroTRANS = False
    
    'Conseguir el siguiente numero de linea
    SQL = ObtenerWhereCP(False)
    SQL = Replace(SQL, NombreTabla, NomTablaLineas)
    numlinea = SugerirCodigoSiguienteStr(NomTablaLineas, "numlinea", SQL)
    Me.cmdaux(0).Tag = numlinea
    
    Set vCStock = New CStock
    If Not InicializarCStock(vCStock, "E", numlinea) Then Exit Function
    
    vCStock.ComprobarFechaInventario True, ""
    
    
    If DatosOkLinea() Then 'Lineas de Albaranes Proveedor
         
        
        'Inserta en tabla "slialp"
        SQL = "INSERT INTO " & NomTablaLineas
        SQL = SQL & " (numalbar, fechaalb, codprove, numlinea, codalmac, codartic, nomartic, ampliaci, cantidad, precioar, dtoline1, dtoline2, importel,numlotes,codccost) "
        SQL = SQL & "VALUES (" & DBSet(Text1(0).Text, "T") & ", " & DBSet(Text1(1).Text, "F") & ", " & Val(Text1(4).Text) & ", " & numlinea & ", " & Val(txtAux(0).Text) & ","
        SQL = SQL & DBSet(txtAux(1).Text, "T") & ", " & DBSet(txtAux(2).Text, "T") & ", " & DBSet(Text2(16).Text, "T") & ", "
        SQL = SQL & DBSet(txtAux(3).Text, "N") & ", "
        SQL = SQL & DBSet(txtAux(4).Text, "S") & ", " & DBSet(txtAux(5).Text, "N") & ", "
        SQL = SQL & DBSet(txtAux(6).Text, "N") & ", "
        SQL = SQL & DBSet(txtAux(7).Text, "N") & ", " & DBSet(Text2(17).Text, "T") & ","
        SQL = SQL & DBSet(txtAux(9).Text, "T", "S") 'centro coste
        SQL = SQL & ");"
     Else
        Set vCStock = Nothing
        Exit Function
     End If
    
    If SQL <> "" Then
        On Error GoTo EInsertarLinea
        conn.BeginTrans
        DentroTRANS = True
        
        MenError = "Insertando lineas Albaran Compras"
        conn.Execute SQL
        
        
        '==== LAURA 20/09/2006
        'Realizar estas actualizaciones antes de modificar el stock del almacen
        MenError = "Actualizar ult. fecha compra"
        '-- Actualizar en la tabla sartic el ult precio de compra y la ult. fecha compra
        Set vArtic = New CArticulo
        vArtic.Codigo = txtAux(1).Text
        
        
        'Si es negativo NO actualizo PUC
        If CCur(txtAux(3).Text) > 0 Then
            'Laura 19/12/2006: calcular precio_ult_compra con el precio con descuentos, ed. importe/cantidad, en lugar de con el precio
            'b = vArtic.ActualizarUltFechaCompra(Text1(1).Text, txtAux(4).Text)
            b = vArtic.ActualizarUltFechaCompra(Text1(1).Text, CStr(Round2(CCur(txtAux(7).Text) / CCur(txtAux(3).Text), 4)))
        Else
            b = True
        End If
                
        'Actualizar en la tabla sartic el precio medio ponderado
        If CCur(txtAux(3).Text) <> 0 Then
            MenError = "Actualizar precio medio ponderado"
            'Laura 19/12/2006: calcular precio_ult_compra con el precio con descuentos, ed. importe/cantidad, en lugar de con el precio
            'If b Then b = vArtic.ActualizarPrecioMedPond(CCur(txtAux(3).Text), CCur(txtAux(4).Text))
            If b Then b = vArtic.ActualizarPrecioMedPond(CCur(txtAux(3).Text), Round2(CCur(txtAux(7).Text) / CCur(txtAux(3).Text), 4))
            Set vArtic = Nothing
            '====
        End If
        
        'en actualizar stock comprobamos si el articulo tiene control de stock
        If b Then
            MenError = "Actualizando Stocks"
            b = vCStock.ActualizarStock
            
            
            vCStock.ComprobarFechaInventario True, ""  'Dejo seguir
    
            
        End If
        
        
        If b Then
            'si el articulo tiene control de numero de lotes, insertar en la tabla slotes
            If Me.Text2(17).Locked = False Then
                'si ya existe la linea aumentamos la cantidad entrada
                SQL = " codartic=" & DBSet(txtAux(1).Text, "T") & " AND numlotes=" & DBSet(Text2(17).Text, "T") & " AND fecentra"
                SQL = DevuelveDesdeBD(conAri, "canentra", "slotes", SQL, DBSet(Text1(1).Text, "F"), "")
                If SQL = "" Then SQL = "0"
                If CCur(SQL) <> 0 Then
                    
                    If DBSet(txtAux(3).Text, "N") < 0 Then
                        SQL = ""
                    Else
                        SQL = "+"
                    End If
                    SQL = "UPDATE slotes SET canentra=canentra  " & SQL & DBSet(txtAux(3).Text, "N")
                    
                    SQL = SQL & " WHERE " & " codartic=" & DBSet(txtAux(1).Text, "T") & " AND numlotes=" & DBSet(Text2(17).Text, "T") & " AND fecentra=" & DBSet(Text1(1).Text, "F")
                    conn.Execute SQL
                Else
                    SQL = "INSERT INTO slotes (codartic,numlotes,fecentra,canentra,canasign) VALUES ("
                    SQL = SQL & DBSet(txtAux(1).Text, "T") & ", " & DBSet(Text2(17).Text, "T") & ", "
                    'fecha entrada, cantidad entrada y cantidad asignada
                    SQL = SQL & DBSet(Text1(1).Text, "F") & "," & DBSet(txtAux(3).Text, "N") & ","
                    
                    If DBSet(txtAux(3).Text, "N") < 0 Then
                        'Es un abono. La cantidad "vendida" es la misma que la "cantidad"
                        ' ya que de este material NO se vende
                        SQL = SQL & "0)"     'SQL = SQL & DBSet(txtAux(3).Text, "N") & ")"
                    Else
                        SQL = SQL & "0)"
                    End If
                    conn.Execute SQL
                End If
            End If
        End If
        
        
        
        
        'Articulo reciclado
        If b Then
          If vParamAplic.ArtReciclado <> "" Then
                
                If ArticuloConTasaReciclado(txtAux(1).Text, ImpReciclado) Then
                     
                    MenError = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", vParamAplic.ArtReciclado, "T")
                  
                    SQL = "INSERT INTO " & NomTablaLineas
                    SQL = SQL & " (numalbar, fechaalb, codprove, numlinea, codalmac, codartic, nomartic, ampliaci, cantidad, precioar, dtoline1, dtoline2, importel,numlotes,codccost) "
                    SQL = SQL & "VALUES (" & DBSet(Text1(0).Text, "T") & ", " & DBSet(Text1(1).Text, "F") & ", " & Val(Text1(4).Text) & ", " & numlinea + 1 & ", " & Val(txtAux(0).Text) & ","
                    SQL = SQL & DBSet(vParamAplic.ArtReciclado, "T") & ", " & DBSet(MenError, "T") & ",null, "
                    SQL = SQL & DBSet(txtAux(3).Text, "N") & ", "
                    SQL = SQL & DBSet(ImpReciclado, "S") & ", 0,0,"
                    ImpReciclado = ImporteFormateado(txtAux(3).Text) * ImpReciclado
                    ImpReciclado = Round2(ImpReciclado, 2)
                    SQL = SQL & DBSet(ImpReciclado, "N") & ",null,null);"
                    MenError = "Art. reciclado"
                    conn.Execute SQL
                End If
            End If
        End If
        
        
        
    End If
    
    Set vCStock = Nothing
    
EInsertarLinea:
    If Err.Number <> 0 Then b = False
    
    If b Then
        If DentroTRANS Then conn.CommitTrans
        AlmacenLineas = CInt(txtAux(0).Text)
        InsertarLinea = True
    Else
        If DentroTRANS Then conn.RollbackTrans
        InsertarLinea = False
        MuestraError Err.Number, "Insertar Lineas Albaran" & vbCrLf & MenError & vbCrLf, Err.Description
    End If
End Function


Private Function ModificarLinea() As Boolean
'Modifica un registro en la tabla de lineas de Albaran: slialb
Dim SQL As String, vWhere As String
Dim vCStock As CStock
Dim vArtic As CArticulo
Dim b As Boolean
Dim MenError As String
Dim dentroTRANSAC As Boolean
Dim cadNumLote As String
Dim cLote As CNumLote
Dim ImpReciclado As Single


    On Error GoTo EModificarLinea

    ModificarLinea = False
    SQL = ""
    dentroTRANSAC = False
    
    Set vCStock = New CStock
    If Not InicializarCStock(vCStock, "E") Then Exit Function
    
    
    vCStock.ComprobarFechaInventario True, ""  'Dejo seguir
    
    Set vArtic = New CArticulo
    If Not vArtic.LeerDatos(txtAux(1).Text) Then Exit Function
    

'    If DatosOkLinea() Then
    'sql para actualizar la linea del albaran compras
    SQL = "UPDATE " & NomTablaLineas & " Set codalmac = " & txtAux(0).Text & ", codartic=" & DBSet(txtAux(1).Text, "T") & ", "
    SQL = SQL & "nomartic=" & DBSet(txtAux(2).Text, "T") & ", ampliaci=" & DBSet(Text2(16).Text, "T", "S") & ", "
    SQL = SQL & "cantidad= " & DBSet(txtAux(3).Text, "N") & ", "
    SQL = SQL & "precioar=" & DBSet(txtAux(4).Text, "S") & ", " 'precio
    SQL = SQL & "dtoline1= " & DBSet(txtAux(5).Text, "N") & ", dtoline2= " & DBSet(txtAux(6).Text, "N") & ", "
    SQL = SQL & "importel= " & DBSet(txtAux(7).Text, "N") & ", "
    SQL = SQL & "numlotes=" & DBSet(Text2(17).Text, "T", "S") & ","
    SQL = SQL & "codccost=" & DBSet(txtAux(9).Text, "T", "S")
    
    vWhere = ObtenerWhereCP(True) & " AND numlinea=" & Data2.Recordset!numlinea
    vWhere = Replace(vWhere, NombreTabla, NomTablaLineas)
    SQL = SQL & vWhere

    If SQL <> "" Then
        dentroTRANSAC = True
        conn.BeginTrans
            
        MenError = "Actualizando Lineas Albaran Compras"
        conn.Execute SQL
            
            
        '==== Laura 20/09/2006, antes de actualizar el stock
        ' deshacer el precio medio ponderado y luego calcularlo otra vez con los nuevos valores
        MenError = "Recalcular precio medio ponderado del articulo."
        '-- Laura 18/12/2006: calcular precio_med_pond con el precio aplicandole el descuento, ed. importe/cantidad.
        'b = vArtic.ReestablecerPrecioMedPon(CCur(Data2.Recordset!Cantidad), CCur(Data2.Recordset!precioar))
        b = vArtic.ReestablecerPrecioMedPon(CCur(Data2.Recordset!cantidad), CCur(Data2.Recordset!ImporteL) / CCur(Data2.Recordset!cantidad))
        
        '-- Laura 18/12/2006: calcular precio_med_pond con el precio aplicandole el descuento, ed. importe/cantidad.
        'If b Then b = vArtic.ActualizarPrecioMedPond(CCur(txtAux(3).Text), CCur(txtAux(4).Text), CCur(Data2.Recordset!Cantidad))
        If b Then b = vArtic.ActualizarPrecioMedPond(CCur(txtAux(3).Text), Round2(CCur(txtAux(7).Text) / CCur(txtAux(3).Text), 4), CCur(Data2.Recordset!cantidad))
        
        'Actualizar ultima fecha de compra del articulo
        If b Then
            'Noacutalizamos si cantidad negativa
            If CCur(txtAux(3).Text) > 0 Then
                MenError = "Actualizando ult. fecha compra"
                '-- Laura 18/12/2006: actualizar precio_ult_compra con el precio aplicandole el descuento, ed. importe/cantidad.
                'b = vArtic.ActualizarUltFechaCompra(Text1(1).Text, txtAux(4).Text)
                b = vArtic.ActualizarUltFechaCompra(Text1(1).Text, Round2(CCur(txtAux(7).Text) / CCur(txtAux(3).Text), 4))
            End If
        End If
        '====
            
            
        'Actualizar Stocks de los articulos y movimientos
        '===================================================
        If b Then
            MenError = "Actualizando stocks y movimientos almacen"
            'si no se ha modificado el almacen reestablecemos cantidad y precio
            If CInt(Data2.Recordset!codAlmac) = CInt(txtAux(0).Text) Then
'                MenError = "Actualizando Stocks"
                b = vCStock.ModificarStock(Data2.Recordset!cantidad)
            Else
                'deshacer el movimiento para el almacen anterior y devolver stock
                b = InicializarCStock(vCStock, "S") 'movim. de salida
                If b Then b = vCStock.DevolverStock2
                            
                'Insertar el movimiento para el nuevo almacen y actualizar stock
                b = InicializarCStock(vCStock, "E") 'mov. de entrada
                If b Then b = vCStock.ActualizarStock
            End If
        End If
                

        
        '=== CONTROL N� DE LOTES DEL ARTICULO
        '===============================================
        If b Then
            'comprobar si el art�culo que modificamos tiene control de lotes
            MenError = "Actualizando N� Lote."
            If vArtic.TieneNumLote Then
                    'si no existe en la tabla slotes lo a�adimos sino lo modificamos
                    SQL = "SELECT COUNT(*) FROM slotes "
                    SQL = SQL & " WHERE codartic=" & DBSet(Data2.Recordset!codArtic, "T") & " AND numlotes=" & DBSet(Text2(17).Text, "T")
                    SQL = SQL & " AND fecentra=" & DBSet(Data2.Recordset!FechaAlb, "F")
                    If RegistrosAListar(SQL) > 0 Then
                        'actualizar la cantidad de entrada de la tabla slotes
                        SQL = "UPDATE slotes SET canentra=" & DBSet(txtAux(3).Text, "N")  '- CSng(Me.Data2.Recordset!cantidad), "N")
                        SQL = SQL & " WHERE codartic=" & DBSet(Data2.Recordset!codArtic, "T") & " AND numlotes=" & DBSet(Data2.Recordset!numlotes, "T") & " AND fecentra=" & DBSet(Data2.Recordset!FechaAlb, "F")
                        conn.Execute SQL
                    ElseIf Text2(17).Text <> "" Then
                        'SI NO EXISTE LO INSERTAMOS
                        SQL = "INSERT INTO slotes (codartic,numlotes,fecentra,canentra,canasign) VALUES ("
                        SQL = SQL & DBSet(Data2.Recordset!codArtic, "T") & "," & DBSet(Text2(17).Text, "T") & "," & DBSet(Data2.Recordset!FechaAlb, "F") & ","
                        SQL = SQL & DBSet(txtAux(3).Text, "N") & ",0)"
                        conn.Execute SQL
                    End If
                                
                    'SI HEMOS MODIFICADO EL N� DE LOTE
                    'DESCONTAMOS LA CANTIDAD DE LA LINEA DE LA VIEJA
                    'Y SI ES CERO LA BORRAMOS
                    If Text2(17).Text <> CStr(DBLet(Data2.Recordset!numlotes, "T")) Then
                        If Not IsNull(Data2.Recordset!numlotes) Then
                            If DBLet(Data2.Recordset!numlotes, "T") <> "" Then
                                'actualizar la cantidad de entrada de la tabla slotes
                                SQL = "UPDATE slotes SET canentra=canentra - " & DBSet(txtAux(3).Text, "N")
                                SQL = SQL & " WHERE codartic=" & DBSet(Data2.Recordset!codArtic, "T") & " AND numlotes=" & DBSet(Data2.Recordset!numlotes, "T") & " AND fecentra=" & DBSet(Data2.Recordset!FechaAlb, "F")
                                conn.Execute SQL
                                'borrar si
                                SQL = "DELETE FROM slotes "
                                SQL = SQL & " WHERE codartic=" & DBSet(Data2.Recordset!codArtic, "T") & " AND numlotes=" & DBSet(Data2.Recordset!numlotes, "T") & " AND fecentra=" & DBSet(Data2.Recordset!FechaAlb, "F")
                                SQL = SQL & " AND canentra=0"
                                conn.Execute SQL
                            End If
                        End If
                    End If
            End If
        End If
                
            
        
        'Articulo reciclado
        If b Then
            If vParamAplic.ArtReciclado <> "" Then
                
                If ArticuloConTasaReciclado(txtAux(1).Text, ImpReciclado) Then
                                           
                     'Si el articulo siguiente es PV entoces lo updatearemos
                     SQL = Replace(ObtenerWhereCP(False), NombreTabla, NomTablaLineas) & " AND numlinea"
                     NumRegElim = Val(DBLet(Data2.Recordset!numlinea, "N")) + 1
                     SQL = DevuelveDesdeBD(conAri, "codartic", "slialp", SQL, CStr(NumRegElim))
                     'En SQL tengo el codarti de la linea SIGUIENTE
                     'SI es punto verde de parametros, supondremos que esta vinculado con la linea que estamos modificando
                     If SQL = vParamAplic.ArtReciclado Then
                     
                          SQL = "UPDATE " & NomTablaLineas & " SET "
                          SQL = SQL & "cantidad= " & DBSet(txtAux(3).Text, "N") & ", "
                          SQL = SQL & "precioar= " & DBSet(ImpReciclado, "N") & ", " 'precio
                          ImpReciclado = ImporteFormateado(txtAux(3).Text) * ImpReciclado
                          SQL = SQL & "importel= " & DBSet(ImpReciclado, "N")  'Importe
                          'WHERE
                          SQL = SQL & Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas) & " AND numlinea=" & NumRegElim
                          conn.Execute SQL
                    End If  'linea siguiente con codarti=puntoverde
                End If
            End If
        End If
            
            
        If b Then
            conn.CommitTrans
        Else
            conn.RollbackTrans
        End If
        ModificarLinea = b
    End If
        
    
    Set vCStock = Nothing
    Set vArtic = Nothing
    Exit Function
    
EModificarLinea:
    If dentroTRANSAC Then conn.RollbackTrans
    If Not vArtic Is Nothing Then Set vArtic = Nothing
    If Not vCStock Is Nothing Then Set vCStock = Nothing
    ModificarLinea = False
    MuestraError Err.Number, "Modificar Lineas Albaran" & vbCrLf & MenError & vbCrLf & Err.Description
End Function





Private Sub PonerBotonCabecera(b As Boolean)
'Pone el boton de Regresar a la Cabecera si pasamos a MAntenimiento de Lineas
'o Pone los botones de Aceptar y cancelar en Insert,update o delete lineas
On Error Resume Next

    Me.cmdAceptar.visible = Not b
    Me.cmdCancelar.visible = Not b
    Me.cmdRegresar.visible = b
    Me.cmdRegresar.Caption = "Cabecera"
    If cmdRegresar.visible Then
        cmdRegresar.Cancel = True
    Else
        cmdCancelar.Cancel = True
    End If
    If b Then
        Me.lblIndicador.Caption = "L�neas " & TituloLinea
        PonerFocoBtn Me.cmdRegresar
    End If
    'Habilitar las opciones correctas del menu segun Modo
'    PonerModoOpcionesMenu (Modo)
'    PonerOpcionesMenu 'Habilitar las opciones correctas del menu segun Nivel de Acceso
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub CargaGrid(ByRef vDataGrid As DataGrid, ByRef vData As Adodc, enlaza As Boolean)
Dim b As Boolean
Dim SQL As String

On Error GoTo ECargaGrid

    b = DataGrid1.Enabled
    
    SQL = MontaSQLCarga(enlaza)
    CargaGridGnral vDataGrid, vData, SQL, PrimeraVez
    
    CargaGrid2 vDataGrid, vData
    
    
    b = (Modo = 5) And (ModificaLineas = 1 Or ModificaLineas = 2)
    vDataGrid.Enabled = Not b
    If Modo = 2 Then vDataGrid.Enabled = True
    PrimeraVez = False
    
    DataGrid1.ScrollBars = dbgAutomatic
    Exit Sub
    
ECargaGrid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub


Private Sub CargaGrid2(ByRef vDataGrid As DataGrid, ByRef vData As Adodc)
Dim i As Byte
On Error GoTo ECargaGrid

    vData.Refresh

    vDataGrid.Columns(0).visible = False
    vDataGrid.Columns(1).visible = False
    vDataGrid.Columns(2).visible = False
    vDataGrid.Columns(3).visible = False
    
    Select Case vDataGrid.Name
        Case "DataGrid1" 'Cod. Almacen
            i = 4
            vDataGrid.Columns(i).Caption = "Alm."
            vDataGrid.Columns(i).Width = 500 + 100
            vDataGrid.Columns(i).NumberFormat = "000"
                
            i = i + 1
            vDataGrid.Columns(i).Caption = "Art�culo"
            vDataGrid.Columns(i).Width = 1700 + 600
            i = i + 1
            vDataGrid.Columns(i).Caption = "Descripci�n"
            vDataGrid.Columns(i).Width = 3400 + 1000
            If Not vEmpresa.TieneAnalitica Then vDataGrid.Columns(i).Width = vDataGrid.Columns(i).Width + 660
            
            i = i + 1
            vDataGrid.Columns(i).visible = False
            i = i + 1
            vDataGrid.Columns(i).Caption = "Cantidad"
            vDataGrid.Columns(i).Width = 850 + 300
            vDataGrid.Columns(i).Alignment = dbgRight
            vDataGrid.Columns(i).NumberFormat = FormatoImporte
            
            i = i + 1
            vDataGrid.Columns(i).Caption = "Precio"
            vDataGrid.Columns(i).Width = 1140 + 300
            vDataGrid.Columns(i).Alignment = dbgRight
            vDataGrid.Columns(i).NumberFormat = FormatoPrecio2
                
            i = i + 1
            vDataGrid.Columns(i).Caption = "Dto.1"
            vDataGrid.Columns(i).Width = 550 + 200
            vDataGrid.Columns(i).Alignment = dbgRight
            vDataGrid.Columns(i).NumberFormat = FormatoDescuento
            
            i = i + 1
            vDataGrid.Columns(i).Caption = "Dto.2"
            vDataGrid.Columns(i).Width = 550 + 200
            vDataGrid.Columns(i).Alignment = dbgRight
            vDataGrid.Columns(i).NumberFormat = FormatoDescuento
                
            i = i + 1
            vDataGrid.Columns(i).Caption = "Importe"
            vDataGrid.Columns(i).Width = 1080 + 450
            vDataGrid.Columns(i).Alignment = dbgRight
            vDataGrid.Columns(i).NumberFormat = FormatoImporte
            
            i = i + 1
            vDataGrid.Columns(i).visible = False 'numlote
            
            i = i + 1
            vDataGrid.Columns(i).Caption = "IVA"
            vDataGrid.Columns(i).Width = 390 + 50
            vDataGrid.Columns(i).Alignment = dbgRight
            vDataGrid.Columns(i).NumberFormat = "# "
            
            i = i + 1
            If vEmpresa.TieneAnalitica Then
                vDataGrid.Columns(i).Caption = "CCost"
                vDataGrid.Columns(i).Width = 660
            Else
                vDataGrid.Columns(i).visible = False 'codccost
            End If
            vDataGrid.Columns(i + 1).visible = False 'ampliaci
            vDataGrid.Columns(i + 2).visible = False 'numlote
            
            
    End Select

    For i = 0 To vDataGrid.Columns.Count - 1
        vDataGrid.Columns(i).Locked = True
        vDataGrid.Columns(i).AllowSizing = False
    Next i
    '++
    vDataGrid.HoldFields
    vDataGrid.RowHeight = 350
    
    
    Exit Sub
    
ECargaGrid:
    MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub


'Esta funcion sustituye a LlamaLineas
Private Sub CargaTxtAux(visible As Boolean, limpiar As Boolean)
'IN: visible: si es true ponerlos visibles en la posici�n adecuada
'    limpiar: si es true vaciar los txtAux
Dim alto As Single
Dim i As Byte

    If Not visible Then
        'Fijamos el alto (ponerlo en la parte inferior del form)
        For i = 0 To txtAux.Count - 1 'TextBox
            txtAux(i).Top = 290
            txtAux(i).visible = visible
        Next i
        cmdaux(0).visible = visible
        cmdaux(1).visible = visible
        cmdaux(2).visible = visible
    Else
        If limpiar Then 'Vaciar los textBox (Vamos a Insertar)
            DeseleccionaGrid DataGrid1
            For i = 0 To txtAux.Count - 1
                txtAux(i).Text = ""
                BloquearTxt txtAux(i), False
            Next i
        Else 'Vamos a modificar
            For i = 0 To txtAux.Count - 1
                If i < 3 Then 'campos anteriores a ampliacion linea (ampliaci)
                    txtAux(i).Text = DataGrid1.Columns(i + 4).Text
                '## LAURA 19/06/2008
                ElseIf i < 8 Then
                    txtAux(i).Text = DataGrid1.Columns(i + 5).Text
                Else
                    txtAux(i).Text = DataGrid1.Columns(i + 6).Text
                End If
                '##
                txtAux(i).Locked = False
            Next i
        End If
               
        'El campo Importe es calculado y lo bloqueamos.
        'David. Febrero 2009.  NO bloqueamos el importe. Para que puedan ajustar los valores
        'BloquearTxt txtAux(7), True
    
    
        '#Laura 15/11/2006
        'no se puede modificar el almacen y el articulo pq no elimina bien de smoval
        'y no reestablece stock si se cambia el articulo (REVISAR!!!)
'        BloquearTxt txtAux(0), (ModificaLineas = 2) 'codalmac
        BloquearTxt txtAux(1), (ModificaLineas = 2) 'codartic
'        Me.cmdAux(0).Enabled = (ModificaLineas <> 2)
        Me.cmdaux(1).Enabled = (ModificaLineas <> 2)
        '#
    
    
        '## LAURA 19/06/2008
        '   A�adimos columna de IVA siempre bloqueada
        BloquearTxt txtAux(8), True
        '##
    
        ' ---- [20/10/2009] [LAURA] : a�adir centro de coste
        BloquearTxt txtAux(9), Not (vEmpresa.TieneAnalitica And vParamAplic.ModoAnalitica = 2)
        Me.cmdaux(2).Enabled = Not txtAux(9).Locked
        Me.cmdaux(2).visible = Me.cmdaux(2).Enabled
        ' ----
    

        'Fijamos altura(Height) y posici�n Top
        '-------------------------------
        alto = ObtenerAlto(DataGrid1, 30)  '-- antes 20
        
        For i = 0 To txtAux.Count - 1
            txtAux(i).Top = alto
            txtAux(i).Height = DataGrid1.RowHeight
        Next i
        cmdaux(0).Top = alto
        cmdaux(1).Top = alto
        cmdaux(2).Top = alto
        cmdaux(0).Height = DataGrid1.RowHeight
        cmdaux(1).Height = DataGrid1.RowHeight
        cmdaux(2).Height = DataGrid1.RowHeight
        
        'Fijamos anchura y posicion Left
        '--------------------------------
        'Cod. Almac
        txtAux(0).Left = DataGrid1.Left + 330
        txtAux(0).Width = DataGrid1.Columns(4).Width - 160
        cmdaux(0).Left = txtAux(0).Left + txtAux(0).Width - 40
        'Cod Artic
        txtAux(1).Left = cmdaux(0).Left + cmdaux(0).Width + 20
        txtAux(1).Width = DataGrid1.Columns(5).Width - 160
        cmdaux(1).Left = txtAux(1).Left + txtAux(1).Width - 50
        'Nom Artic
        txtAux(2).Left = cmdaux(1).Left + cmdaux(1).Width
        txtAux(2).Width = DataGrid1.Columns(6).Width - 10
        'Cantidad
        txtAux(3).Left = txtAux(2).Left + txtAux(2).Width + 10
        txtAux(3).Width = DataGrid1.Columns(8).Width - 10
        'Precio, Dto1, Dto2, Precio
        For i = 4 To 7
            txtAux(i).Left = txtAux(i - 1).Left + txtAux(i - 1).Width + 10
            txtAux(i).Width = DataGrid1.Columns(i + 5).Width - 10
        Next i
        
        '## LAURA 19/06/2008
        txtAux(8).Left = txtAux(7).Left + txtAux(7).Width + 10
        txtAux(8).Width = DataGrid1.Columns(14).Width - 10
        '##
        
        ' ---- [20/10/2009] [LAURA] : a�adir el centro de coste
        txtAux(9).Left = txtAux(8).Left + txtAux(8).Width + 10
        txtAux(9).Width = DataGrid1.Columns(15).Width - 10
        cmdaux(2).Left = txtAux(9).Left + txtAux(9).Width - cmdaux(2).Width
        ' ----
        
        'Los ponemos Visibles o No
        '--------------------------
        For i = 0 To txtAux.Count - 2
            txtAux(i).visible = visible
        Next i
        txtAux(9).visible = visible And vEmpresa.TieneAnalitica
        cmdaux(0).visible = visible
        cmdaux(1).visible = visible
    End If
End Sub


Private Sub Toolbar5_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            'Cambiando en smoval
            ModificarProveedor
        Case 2
            'A hco sin tocar stocks ni smoval ni precios ni leches en vinagre
            EliminarSinStocks
        Case 3
            RecepcionarAlbaran
        Case 4 'N� Series
            BotonNSeries
        Case 5
            ImpirmirEtiqEsta
    End Select

End Sub

Private Sub ToolbarDes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Desplazamiento (Button.Index)
End Sub

Private Sub txtAux_GotFocus(Index As Integer)
Dim cadkey As Integer

    cadkey = ObtenerCadKey(kCampo, Index)
    kCampo = Index
    ConseguirFocoLin txtAux(Index), cadkey
    
    If Index = 3 Or Index = 4 Or Index = 1 Then
        
        If Index = 3 Then
            lblF.Caption = "Ver articulo"
        ElseIf Index = 4 Then
            lblF.Caption = "Ver precio"
        Else
            lblF.Caption = "EAN"
        End If
    Else
        lblF.Caption = ""
    End If
    
    
End Sub

Private Sub TxtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    
    ' ---- [02/11/2009] [LAURA] : al pulsar F2 para abrir articulos en la solapa Documentos|Pedidos
    If KeyCode = 113 Then
        If Index = 3 Then AbrirForm_Articulos
    
        If Index = 1 Then Me.DataGrid1.Columns(5).Caption = "EAN"
        If Index = 4 And txtAux(1).Text <> "" Then
                frmListadoPrecios.Opcion = 0
                frmListadoPrecios.CadenaPasoDatos = txtAux(1).Text & "|" & Text1(4).Text & "|"
                frmListadoPrecios.Show vbModal
        End If
    ' ----
    ElseIf KeyCode = 43 Or KeyCode = 107 Or KeyCode = 187 Then
                If Index < 2 Or Index = 9 Then  'Para los que tienen busqueda
                    If Modo = 5 And ModificaLineas = 1 Then
                        If txtAux(Index).Text = "" Then
                            PulsadoMas2 = True
                            KeyCode = 0
                
                            PulsarTeclaMas False, Index
                        End If
                    End If
                End If
           
        
    
    
    ElseIf Not (Index = 0 And KeyCode = 38) Then
        KEYdown KeyCode
    End If
    
End Sub


Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub txtAux_LostFocus(Index As Integer)
Dim devuelve As String
Dim TipoDto As Byte
Dim b As Boolean
Dim bLote As Boolean
Dim okArticulo As Boolean

    
    
    
    If PulsadoMas2 Then
        'Para que cuando pulse el mas abra el form
        PulsadoMas2 = False
        txtAux(Index).Text = ""
        Exit Sub
    End If
    
    If Not PerderFocoGnralLineas(txtAux(Index), ModificaLineas) Then Exit Sub
    
    
    Select Case Index
        Case 0 'Cod Almacen
            'Comprobar que existe el almacen
            devuelve = PonerAlmacen(txtAux(Index).Text)
            txtAux(Index).Text = devuelve
            If devuelve = "" Then PonerFoco txtAux(Index)
    
        Case 1 'Cod. ARTICULO
            If txtAux(1).Text = "" Then
                txtAux(2).Text = ""
                Exit Sub
            End If
            If txtAux(0).Text = "" Then
                MsgBox "Debe seleccionar un almacen.", vbInformation
                PonerFoco txtAux(0)
                Exit Sub
            End If
            
            
            If Me.DataGrid1.Columns(5).Caption = "EAN" Then
                'Ha pulsado F2, para meter, en lugar del codigo del articulo, el EAN
                okArticulo = PonerArticuloEan(txtAux(1), txtAux(2), txtAux(0).Text, CodTipoMov, ModificaLineas, , bLote, devuelve)
            Else
                okArticulo = PonerArticulo(txtAux(1), txtAux(2), txtAux(0).Text, CodTipoMov, ModificaLineas, , bLote, devuelve)
            End If
            If okArticulo Then
                '---- [20/10/2009] [LAURA] : a�adir centro de coste
                If Not vEmpresa.TieneAnalitica Then
                    txtAux(9).Text = ""
                ElseIf vParamAplic.ModoAnalitica = 1 Then
                    txtAux(9).Text = devuelve
                    Me.txtAux2(9).Text = PonerNombreCCoste(Me.txtAux(9))
                End If
                '----
            
            
                BloquearTxt Text2(17), Not bLote
                
                '## LAURA 19/06/2008
                'obtener el cod. iva del articulo
                txtAux(8).Text = DevuelveDesdeBDNew(conAri, "sartic", "codigiva", "codartic", txtAux(1).Text, "T")
                
                '##
                
                b = (Me.ActiveControl.Name = "txtAux")
                If b Then b = (Me.ActiveControl.Index = 0)
                
                If Not b Then
'                    If txtAux(2).Locked Then PonerFoco txtAux(3)
                Else
                    PonerFoco txtAux(0)
                End If
            Else
                PonerFoco txtAux(Index)
            End If
            
        Case 2 'Desc. Articulo
            If txtAux(Index).Locked = False Then txtAux(Index).Text = UCase(txtAux(Index).Text)
            
        Case 3 'CANTIDAD
            If PonerFormatoDecimal(txtAux(Index), 1) Then  'Tipo 1: Decimal(12,2)
                If (Modo = 5 And ModificaLineas = 1) Then 'Modo Insertar en Mto Lineas
                    'Obtener el precio correspondiente y los descuentos
                    ObtenerPrecioCompra
                End If
            End If

        Case 4 'Precio
            PonerFormatoDecimal_Single txtAux(Index), 9 'Tipo 9: COnstante
        Case 5, 6 'Descuentos
            PonerFormatoDecimal txtAux(Index), 4 'Tipo 4: Decimal(4,2)
        Case 7 'Importe Linea
            If txtAux(Index).Text <> "" Then
                If Not PonerFormatoDecimal(txtAux(Index), 1) Then  'Tipo 3: Decimal(12,2)
                    If ModificaLineas = 2 Then
                        'Ponemos el importe que tenia
                        txtAux(Index).Text = DataGrid1.Columns(12).Text
                    Else
                        txtAux(Index).Text = "0.00"
                    End If
                End If
            End If
            
        Case 9 'COD. CENTRO COSTE
            ' ---- [20/10/2009] [LAURA]: a�adir centro de coste a la linea
            If txtAux(Index).Text = "" Then
                 txtAux2(Index).Text = ""
            ElseIf vEmpresa.TieneAnalitica Then
                'centro de coste
                ' ---- [20/10/2009] [LAURA]: a�adir campo centro de coste familia
                Me.txtAux2(Index).Text = PonerNombreCCoste(Me.txtAux(Index))
            End If
    End Select
    
    
     If (Index = 3 Or Index = 4 Or Index = 5 Or Index = 6) Then
'        If Trim(TxtAux(3).Text) = "" Or Trim(TxtAux(4).Text) = "" Then Exit Sub
'        If Trim(TxtAux(6).Text) = "" Or Trim(TxtAux(7).Text) = "" Then Exit Sub
        If txtAux(1).Text = "" Then Exit Sub
        TipoDto = DevuelveDesdeBDNew(conAri, "sprove", "tipodtos", "codprove", Text1(4).Text, "N")
        txtAux(7).Text = CalcularImporteSng(txtAux(3).Text, txtAux(4).Text, txtAux(5).Text, txtAux(6).Text, TipoDto)
        PonerFormatoDecimal txtAux(7), 1
    End If
End Sub



Private Sub ObtenerPrecioCompra()
Dim vPrecio As CPreciosCom
Dim cad As String
Dim Aux2 As String
    On Error GoTo EPrecios
    
    Set vPrecio = New CPreciosCom
    If vPrecio.Leer(txtAux(1).Text, Text1(4).Text) Then
        If vPrecio.ComprobarCantidad(CInt(txtAux(3).Text)) Then
            txtAux(4).Text = vPrecio.ObtenerPrecio(Text1(1).Text)
            txtAux(5).Text = vPrecio.Descuento1
            txtAux(6).Text = vPrecio.Descuento2
        Else
            PonerFoco txtAux(3)
            Exit Sub
        End If
    Else
        'Obtener el ult. precio de compra de ese articulo (sartic)
       
        cad = DevuelveDesdeBDNew(conAri, "sartic", "preciouc", "codartic", txtAux(1).Text, "T")
        If cad <> "" Then txtAux(4).Text = cad
        
            vPrecio.CodigoArtic = txtAux(1).Text
            vPrecio.CodigoProve = Text1(4).Text
            cad = vPrecio.ObtenerDescuentos2(Text1(1).Text, Aux2)
            If cad = "" Then cad = "0"
            txtAux(5).Text = cad
            If Aux2 = "" Then Aux2 = "0"
            txtAux(6).Text = Aux2
            
            'txtAux(5).Text = "0"
            'txtAux(6).Text = "0"

    End If
    PonerFormatoDecimal_Single txtAux(4), 9
    PonerFormatoDecimal txtAux(5), 4
    PonerFormatoDecimal txtAux(6), 4
    
    Set vPrecio = Nothing
    
EPrecios:
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub BotonMtoLineas(numTab As Integer, cad As String)
        Me.SSTab1.Tab = numTab
        TituloLinea = cad
        ModificaLineas = 0
        If Data2.Recordset.EOF Then Text2(16).Text = ""
        PonerModo 5
        PonerBotonCabecera True
        AlmacenLineas = -1
        HaModifEnLineas = False
        PonerUltAlmacen
End Sub


Private Function Eliminar() As Boolean
Dim vWhere As String
Dim b As Boolean
Dim SQL As String

    On Error GoTo FinEliminar

        conn.BeginTrans
        vWhere = " " & ObtenerWhereCP(False)
                
        
        'Reestablecer el stock en la tabla salmac a partir de todas las lineas del albaran
        'Eliminar los movimientos de smoval
        b = ReestablecerStock(vWhere)
        

        If b Then
                        'Hay que eliminar los lotes
            If vParamAplic.ManipuladorFitosanitarios2 Then EliminarEnSlotes
                    
        End If
        
        If b Then
            'Pasar los datos al historico de albaranes de compra y borrarlos de albaranes
            'scaalp --> schalp
            'slialp --> slhalp
            b = ActualizarElTraspaso("", vWhere, CodTipoMov, cadList)
            
            'Eliminar los numeros de serie del albaran sino estan vendidos a ningun cliente
            If b Then
                SQL = "DELETE FROM sserie WHERE numalbpr=" & DBSet(Data1.Recordset!Numalbar, "T")
                SQL = SQL & " AND fechacom=" & DBSet(Data1.Recordset!FechaAlb, "F")
                SQL = SQL & " AND codprove=" & Data1.Recordset!Codprove
                SQL = SQL & " AND (isnull(numfactu) and isnull(numalbar))"
                conn.Execute SQL
            End If
            
        End If
        
FinEliminar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar Albaran Compras", Err.Description
        b = False
    End If
    If Not b Then
        conn.RollbackTrans
        Eliminar = False
    Else
        conn.CommitTrans
        Eliminar = True
    End If
End Function


Private Sub LimpiarDataGrids()
'Pone los Grids sin datos, apuntando a ning�n registro
On Error Resume Next
    CargaGrid DataGrid1, Data2, False
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub PosicionarData()
'Dim Indicador As String
'Dim vWhere As String

    If Not Data1.Recordset.EOF Then
        'Hay datos en el Data1 bien porque se ha hecho VerTodos o una Busqueda
         'vWhere = "(" & ObtenerWhereCP(False) & ")"
         
         'SEPT 2010
         'NO parece ir muy fino
         'Voy a cambiar
         Data1.Refresh
         
         'If SituarDataMULTI(Data1, vWhere, Indicador) Then
         If SituarData Then
             PonerModo 2
            
        Else
             LimpiarCampos
             'Poner los grid sin apuntar a nada
             LimpiarDataGrids
             PonerModo 0
         End If
    Else
        'El Data esta vacio, desde el modo de inicio se pulsa Insertar
        CadenaConsulta = "Select * from " & NombreTabla & ObtenerWhereCP(True) & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Function ObtenerWhereCP(conW As Boolean) As String
Dim SQL As String
On Error Resume Next
    
    SQL = ""
    If conW Then SQL = " WHERE "
    SQL = SQL & NombreTabla & ".numalbar= " & DBSet(Text1(0).Text, "T") & " and " & NombreTabla & ".fechaalb='" & Format(Text1(1).Text, FormatoFecha)
    SQL = SQL & "' and " & NombreTabla & ".codprove=" & Val(Text1(4).Text)
    
    ObtenerWhereCP = SQL
End Function


Private Function MontaSQLCarga(enlaza As Boolean) As String
'--------------------------------------------------------------------
' MontaSQlCarga:
'   Bas�ndose en la informaci�n proporcionada por el vector de campos
'   crea un SQl para ejecutar una consulta sobre la base de datos que los
'   devuelva.
' Si ENLAZA -> Enlaza con el data1
'           -> Si no lo cargamos sin enlazar a ningun campo
'--------------------------------------------------------------------
Dim SQL As String
    
    SQL = "SELECT numalbar,fechaalb," & NomTablaLineas & ".codprove, numlinea, codalmac, " & NomTablaLineas & ".codartic,"
    SQL = SQL & NomTablaLineas & ".nomartic, ampliaci, cantidad, precioar, dtoline1, dtoline2, importel, numlotes,codigiva,codccost "
    
    'Octubre 2015
    'Hay que cargar todos los datos del "foragrid" de abajo. Si no esta estamos leyendo DOS veces
    'peeeero, no se muestran en el grid
    SQL = SQL & ",ampliaci, numlotes "
    
    SQL = SQL & " FROM " & NomTablaLineas
    'Para que tanto el hco como el nomal apunte a slialp
    
    SQL = SQL & " inner join sartic on " & NomTablaLineas & ".codartic = sartic.codartic"
    If enlaza Then
        SQL = SQL & " " & ObtenerWhereCP(True)
    Else
        SQL = SQL & " WHERE numalbar = -1"
    End If
    SQL = SQL & " Order by numalbar, fechaalb, codprove, numlinea"
    SQL = Replace(SQL, NombreTabla, NomTablaLineas)
    MontaSQLCarga = SQL
End Function


Private Sub PonerModoOpcionesMenu(Modo As Byte)
'Activas unas Opciones de Menu y Toolbar seg�n el modo en que estemos
Dim b As Boolean
Dim bAux As Boolean
Dim i As Integer
   
        b = (Modo = 2) Or (Modo = 5 And ModificaLineas = 0)
        'Insertar
        Toolbar1.Buttons(1).Enabled = (b Or Modo = 0) And (cadSelAlbaranes = "" Or (cadSelAlbaranes <> "" And Modo = 5)) And Not EsHistorico
        Me.mnNuevo.Enabled = (b Or Modo = 0) And (cadSelAlbaranes = "" Or (cadSelAlbaranes <> "" And Modo = 5)) And Not EsHistorico
        'Modificar
        Toolbar1.Buttons(2).Enabled = b And Not EsHistorico
        Me.mnModificar.Enabled = b And Not EsHistorico
        'eliminar
        'Toolbar1.Buttons(7).Enabled = b And cadSelAlbaranes = "" And Not EsHistorico
        'Me.mnEliminar.Enabled = b And cadSelAlbaranes = "" And Not EsHistorico
        
        'No permito borrar
        If b Then
            'Si modo=2 NO dejare que borre
            If Modo = 2 And cadSelAlbaranes <> "" Then b = False
        End If
        Toolbar1.Buttons(3).Enabled = b And Not EsHistorico
        Me.mnEliminar.Enabled = Toolbar1.Buttons(7).Enabled
            
        b = (Modo = 2) And Not EsHistorico
        'Mantenimiento lineas
        
'--      Toolbar1.Buttons(10).Enabled = (Modo = 2)
'        Me.mnLineas.Enabled = (Modo = 2)
        Toolbar5.Buttons(1).Enabled = b
        Toolbar5.Buttons(2).Enabled = b
        Toolbar5.Buttons(3).Enabled = b
        Toolbar5.Buttons(4).Enabled = b
        
        b = (Modo >= 3) Or Modo = 1
        'Buscar
        Toolbar1.Buttons(5).Enabled = (Not b)
        Me.mnBuscar.Enabled = (Not b)
        'Ver Todos
        Toolbar1.Buttons(6).Enabled = (Not b)
        Me.mnVerTodos.Enabled = (Not b)
        
        BotonesToolBarAux
End Sub



Private Sub BotonesToolBarAux()
Dim b As Boolean
    
    b = (Modo = 2) And Not EsHistorico
    
    ToolAux(0).Buttons(1).Enabled = b
    If b Then
    
        If Data2.Recordset Is Nothing Then
            b = False
        Else
            b = Me.Data2.Recordset.RecordCount
        End If
    End If
    ToolAux(0).Buttons(2).Enabled = b
    ToolAux(0).Buttons(3).Enabled = b

        

End Sub



Private Function InsertarAlbaran(vSQL As String) As Boolean
Dim MenError As String
Dim devuelve As String
Dim bol As Boolean

    On Error GoTo EInsertarOferta
    
    bol = False
    
    'Aqui empieza transaccion
    conn.BeginTrans
    MenError = "Error al insertar en la tabla Cabecera de Albaranes (" & NombreTabla & ")."
    conn.Execute vSQL, , adCmdText
    
    'Actualizar los datos del proveedor si es de varios
    If EsDeVarios Then
        MenError = "Modificando datos proveedor varios."
        bol = ActualizarProveVarios(Text1(4).Text, Text1(6).Text)
    End If
    
    
    'Actualizar el campo fecha ult.compra(fechamov) en la tabla proveedores (sprove)
    devuelve = DevuelveDesdeBDNew(conAri, "sprove", "fechamov", "codprove", Text1(4).Text, "N")
    If (devuelve = "") Then devuelve = "01/01/1900"
    If CDate(Text1(1).Text) > CDate(devuelve) Then
        vSQL = "UPDATE sprove SET fechamov=" & DBSet(Text1(1).Text, "F")
        vSQL = vSQL & " WHERE codprove=" & Text1(4).Text
        conn.Execute vSQL, , adCmdText
    End If
    bol = True
    
    
EInsertarOferta:
    If Err.Number <> 0 Then
        MenError = "Insertando Albaran." & vbCrLf & "----------------------------" & vbCrLf & MenError
        MuestraError Err.Number, MenError, Err.Description
        bol = False
    End If
    If bol Then
        conn.CommitTrans
        InsertarAlbaran = True
    Else
        conn.RollbackTrans
        InsertarAlbaran = False
    End If
End Function


Private Sub LimpiarDatosProve()
Dim i As Byte

    For i = 4 To 14
        Text1(i).Text = ""
    Next i
End Sub
    

Private Function EliminarLinea() As Boolean
Dim vCStock As CStock
Dim cLote As CNumLote
Dim cart As CArticulo
Dim SQL As String
Dim b As Boolean
Dim Aux As String
Dim ImpReciclado As Single

    EliminarLinea = False
    
    
    'Inicilizar la clase para Actualizar los stocks
    Set vCStock = New CStock
    If Not InicializarCStock(vCStock, "S") Then Exit Function
    
    vCStock.ComprobarFechaInventario True, ""  'Dejo seguir
    
    
    '==== Laura: 20/09/2006
    'Inicializar la clase para actualizar precio medio ponderado del Articulo
    Set cart = New CArticulo
    If Not cart.LeerDatos(vCStock.codArtic) Then Exit Function
    '====
    
    On Error GoTo EEliminarLinea
    conn.BeginTrans
    
    'Eliminar las lineas de la tabla "sserie", insertadas para la linea del albaran a eliminar
    SQL = " WHERE  numalbpr= " & DBSet(Text1(0).Text, "T") & " and fechacom='" & Format(Text1(1).Text, FormatoFecha)
    SQL = SQL & "' and codprove=" & Val(Text1(4).Text) & " AND numline2=" & Data2.Recordset!numlinea
    conn.Execute "Delete from sserie " & SQL
    
        
    'Si tiene tasa recilcado y la siguiente linea es la de reciclado me la cargo
        'Tasa recilcaje
    If vParamAplic.ArtReciclado <> "" Then
        If ArticuloConTasaReciclado(CStr(Data2.Recordset!codArtic), ImpReciclado) Then

                SQL = ObtenerWhereCP(False) & " and numlinea>" & Data2.Recordset!numlinea & " AND 1 "
                SQL = Replace(SQL, NombreTabla, NomTablaLineas)
                'Vere si la siguiente linea es de tasa de reciclado me la cargo tb
                Aux = "numlinea"
                SQL = DevuelveDesdeBD(conAri, "codartic", NomTablaLineas, SQL, "1", "N", Aux)

                If SQL = vParamAplic.ArtReciclado Then
                    'la borro tb
                    SQL = ObtenerWhereCP(True) & " and numlinea=" & Aux
                    SQL = Replace(SQL, NombreTabla, NomTablaLineas)
                    SQL = "DELETE FROM " & NomTablaLineas & SQL
                    conn.Execute SQL
                End If



        End If
    End If
    
    
    
    'Construir la SQL para eliminar la linea de la tabla "slialb"
    SQL = ObtenerWhereCP(True) & " and numlinea=" & Data2.Recordset!numlinea
    SQL = Replace(SQL, NombreTabla, NomTablaLineas)
    SQL = "Delete from " & NomTablaLineas & SQL
    conn.Execute SQL 'Eliminar linea
    
    '==== Laura: 20/09/2006
    'reestablecer el precio medio ponderado,
    'debe calcularse antes de reestablecer el stock
    '-- Laura 19/12/2006: calcular el precio medio ponderado con precio con los descuentos ( importe/cantidad)
    'cArt.ReestablecerPrecioMedPon vCStock.Cantidad, CCur(Data2.Recordset!precioar)
    If CCur(Data2.Recordset!cantidad) <> 0 Then cart.ReestablecerPrecioMedPon vCStock.cantidad, Round2(CCur(Data2.Recordset!ImporteL) / CCur(Data2.Recordset!cantidad), 4)
    Set cart = Nothing
    '====
    
    b = vCStock.DevolverStock2
    Set vCStock = Nothing
    
    'Si el articulo tiene control de lotes eliminar la cantidad eliminada
    'si la linea se queda con cero borrarla.
    If b Then
        If Not IsNull(Data2.Recordset!numlotes) Then
            Set cLote = New CNumLote
            If cLote.LeerDatos(CStr(Data2.Recordset!codArtic), CStr(Data2.Recordset!numlotes), CStr(Data2.Recordset!FechaAlb)) Then
                b = cLote.Eliminar(CSng(Data2.Recordset!cantidad))
            
            End If
            Set cLote = Nothing
        End If
    End If
    

EEliminarLinea:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar Linea Albaran " & vbCrLf & Err.Description
        b = False
    End If
    If b Then
        conn.CommitTrans
        EliminarLinea = True
    Else
        conn.RollbackTrans
         EliminarLinea = False
    End If
End Function


Private Function InicializarCStock(ByRef vCStock As CStock, TipoM As String, Optional numlinea As String) As Boolean
'On Error Resume Next
On Error Resume Next

    vCStock.tipoMov = TipoM 'Movimiento de Entrada o Salida
    vCStock.DetaMov = CodTipoMov '"ALC=Albaran de Compra"
    
    'Agosto 2020
    vCStock.FechaMov = Text1(30).Text   'Text1(1).Text
    
    
    vCStock.Trabajador = CLng(Text1(4).Text) 'En smoval guardamos el Proveedor
    vCStock.Documento = Text1(0).Text
    
    If ModificaLineas = 1 Or (ModificaLineas = 2 And TipoM = "E") Then '1=Insertar, 2=Modificar
        vCStock.codArtic = txtAux(1).Text
        vCStock.codAlmac = CInt(txtAux(0).Text)
        vCStock.cantidad = CSng(ComprobarCero(txtAux(3).Text))
        vCStock.Importe = CCur(ComprobarCero(txtAux(7).Text))
    Else
        vCStock.codArtic = Data2.Recordset!codArtic
        vCStock.codAlmac = CInt(Data2.Recordset!codAlmac)
        vCStock.cantidad = CSng(Data2.Recordset!cantidad)
        vCStock.Importe = CCur(Data2.Recordset!ImporteL)
    End If
    If ModificaLineas = 1 Then
         vCStock.LineaDocu = CInt(ComprobarCero(numlinea))
    Else
        vCStock.LineaDocu = CInt(Data2.Recordset!numlinea)
    End If
    If Err.Number <> 0 Then
        MsgBox "No se han podido inicializar la clase para actualizar Stock", vbExclamation
        InicializarCStock = False
    Else
        InicializarCStock = True
    End If
End Function


Private Function ReestablecerStock(cadSel As String) As Boolean
Dim vCStock As CStock
Dim cart As CArticulo
Dim cLote As CNumLote
Dim b As Boolean
Dim RS As ADODB.Recordset
Dim SQL As String

    On Error GoTo ERestablecer
    
    SQL = "SELECT * FROM " & NomTablaLineas & " WHERE " & Replace(cadSel, NombreTabla, NomTablaLineas)
    SQL = SQL & " ORDER BY numlinea desc "
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    b = True
    While (Not RS.EOF) And b
        'para cada linea de albaran deshacemos movimientos y precios medios ponderados
        Set vCStock = New CStock
           If InicializarCStock(vCStock, "S", RS!numlinea) Then
                'estos valores hay q leerlos del RS y no del data2
                 vCStock.codArtic = RS!codArtic
                 vCStock.codAlmac = CInt(RS!codAlmac)
                 vCStock.cantidad = CSng(RS!cantidad)
                 vCStock.Importe = CCur(RS!ImporteL)
                 vCStock.LineaDocu = RS!numlinea
           
                '==== Laura 20/09/2006
                'antes de actualizar el stock reestablecer el precio medio ponderado del articulo
                Set cart = New CArticulo
                If cart.LeerDatos(vCStock.codArtic) Then
                    'Laura 19/12/2006: Calcular precio medio pond. con precio con los descuentos (importe/cantidad)
                    'If Not cArt.ReestablecerPrecioMedPon(CCur(vCStock.Cantidad), CCur(RS!precioar)) Then b = False
                    If Not cart.ReestablecerPrecioMedPon(CCur(vCStock.cantidad), Round2(vCStock.Importe / vCStock.cantidad, 4)) Then b = False
                    
                    'Si el articulo tiene control de lotes eliminar la cantidad eliminada
                    'si la linea se queda con cero borrarla.
                    If b Then
                        If cart.TieneNumLote Then
                            Set cLote = New CNumLote
                            If cLote.LeerDatos(cart.Codigo, CStr(DBLet(RS!numlotes, "T")), CStr(RS!FechaAlb)) Then
                                b = cLote.Eliminar(vCStock.cantidad)
                            
                            End If
                            Set cLote = Nothing
                        End If
                    End If
                End If
                Set cart = Nothing
                '====
                
                
                'Actualiza el stock en salmac y borra de smoval
                'Para cada linea de albaran reestablecer el stock. Como era Mov. de Entrada
                'en Almacen ahora lo tiene que borrar(S).
                If b Then
                    If Not vCStock.DevolverStock2() Then b = False
                End If
           Else
               b = False
           End If
'           Data2.Recordset.MoveNext
           Set vCStock = Nothing
    
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    
    
    '#### ANTES DEL 20/09/2006
    
'    b = True
    
'    If Not Data2.Recordset.EOF Then
''       Data2.Refresh
'       Data2.Recordset.MoveFirst
'
'       'Para cada linea de albaran reestablecer el stock. Como era Mov. de Entrada
'       'en Almacen ahora lo tiene que borrar(S).
'       While (Not Data2.Recordset.EOF) And b
'           Set vCStock = New CStock
'           If InicializarCStock(vCStock, "S", Data2.Recordset!numlinea) Then
'                '==== Laura 20/09/2006
'                'antes de actualizar el stock reestablecer el precio medio ponderado del articulo
'                Set cArt = New CArticulo
'                If cArt.LeerDatos(vCStock.codArtic) Then
'                    If Not cArt.ReestablecerPrecioMedPon(CCur(vCStock.Cantidad), CCur(Data2.Recordset!precioar)) Then b = False
'                End If
'                Set cArt = Nothing
'
'               'Actualiza el stock en salmac y borra de smoval
'                If b Then
'                    If Not vCStock.DevolverStock() Then b = False
'                End If
'           Else
'               b = False
'           End If
'           Data2.Recordset.MoveNext
'           Set vCStock = Nothing
'       Wend
'    End If

ERestablecer:
    If Err.Number <> 0 Then b = False
    If Not b Then
        ReestablecerStock = False
        MuestraError Err.Number, "Reestablecer stock.", Err.Description
    Else
        ReestablecerStock = True
    End If
End Function




Private Function ReestablecerUltFecCompra() As Boolean
Dim cart As CArticulo
Dim SQL As String
Dim b As Boolean

    On Error GoTo ERestCompra
    
    b = True
    
'    select distinct codartic from slialp
'where numalbar=2100045 and fechaalb='2006-09-15' and codprove=21
    
    
'    If Not Data2.Recordset.EOF Then
'       Data2.Refresh
'       Data2.Recordset.MoveFirst
'
'       'Para cada articulo del albaran reestablecer la fecha ultima compra
'       'y el precio ultima compra
'
'       While (Not Data2.Recordset.EOF) And b
'           Set vCStock = New CStock
'           If InicializarCStock(vCStock, "S", Data2.Recordset!numlinea) Then
'               'Actualiza el stock en salmac y borra de smoval
'               If Not vCStock.DevolverStock() Then b = False
'           Else
'               b = False
'           End If
'           Data2.Recordset.MoveNext
'           Set vCStock = Nothing
'       Wend
'    End If
    
    
    
    
    ReestablecerUltFecCompra = b
    
    
ERestCompra:
    
'    If Not b Then
        ReestablecerUltFecCompra = False
'    Else
'        ReestablecerUltFecCompra = True
'    End If
End Function





'Private Function ReestablecerPrecioMedPon() As Boolean
''reestablecer el valor del precio medio ponderado
''Dim vCStock As CStock
'Dim b As Boolean
'
'    On Error GoTo EResPMP
'
'    b = True
''    If Not Data2.Recordset.EOF Then
''       Data2.Refresh
''       Data2.Recordset.MoveFirst
''
''       'Para cada linea de albaran reestablecer el stock. Como era Mov. de Entrada
''       'en Almacen ahora lo tiene que borrar(S).
''       While (Not Data2.Recordset.EOF) And b
''           Set vCStock = New CStock
''           If InicializarCStock(vCStock, "S", Data2.Recordset!numlinea) Then
''               'Actualiza el stock en salmac y borra de smoval
''               If Not vCStock.DevolverStock() Then b = False
''           Else
''               b = False
''           End If
''           Data2.Recordset.MoveNext
''           Set vCStock = Nothing
''       Wend
''    End If
'    ReestablecerPrecioMedPon = b
'    Exit Function
'
'EResPMP:
''    If Not b Then
'        ReestablecerPrecioMedPon = False
''    Else
''        ReestablecerPrecioMedPon = True
''    End If
'End Function



Private Sub InsertarCabecera()
Dim SQL As String

    SQL = CadenaInsertarDesdeForm(Me)
    If SQL <> "" Then
        If InsertarAlbaran(SQL) Then
'                            PosicionarData
            CadenaConsulta = "Select * from " & NombreTabla & ObtenerWhereCP(True) & Ordenacion
            PonerCadenaBusqueda
            'Ponerse en Modo Insertar Lineas
            BotonMtoLineas 1, "Albaranes"
            BotonAnyadirLinea
        End If
    End If
    Me.SSTab1.Tab = 0
End Sub


Private Sub BotonNSeries()
Dim cadWhere As String, SQL As String
Dim RSLineas As ADODB.Recordset

    ModificaLineas = 4

    cadWhere = " WHERE numalbar=" & DBSet(Text1(0).Text, "T")
    cadWhere = cadWhere & " and fechaalb=" & DBSet(Text1(1).Text, "F")
    cadWhere = cadWhere & " and slialp.codprove=" & Text1(4).Text
    
    'Seleccionamos aquellas lineas de albaran que tienen N� de Serie
    SQL = "SELECT numlinea, slialp.codartic, sum(cantidad) as cantidad "
    SQL = SQL & " FROM slialp INNER JOIN sartic on slialp.codartic=sartic.codartic "
    SQL = SQL & cadWhere & " And nseriesn = 1 "
    SQL = SQL & " GROUP BY numlinea,codartic ORDER BY Codartic "

    Set RSLineas = New ADODB.Recordset
    RSLineas.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not RSLineas.EOF Then
        'Abre el formulario de pedir n� serie al comprarlos
        'pero mostrando los n� de serie ya asignados para poder modificarlos
        PedirNSeries RSLineas
    Else
        MsgBox "No hay ninguna linea de Articulo con Control de N� Serie", vbInformation
    End If
    RSLineas.Close
    Set RSLineas = Nothing
    ModificaLineas = 0
    DescargarDatosTMPNumSeries ("tmpnseries")
End Sub


Private Sub PedirNSeries(ByRef RS As ADODB.Recordset)
Dim RSseries As ADODB.Recordset
Dim SQL As String
Dim linea As Integer

    On Error GoTo EPedirNSeries

        'Inicializo la tabla temporal de los num.serie
        PedirNSeriesGnral RS, False
        
        RS.MoveFirst
        While Not RS.EOF
            linea = 0
            'Cargar los N� de serie asignados al albaran
            SQL = "SELECT numserie, codartic FROM sserie "
            SQL = SQL & " WHERE numalbpr=" & DBSet(Text1(0).Text, "T")
            SQL = SQL & " and fechacom='" & Format(Text1(1).Text, FormatoFecha) & "'"
            SQL = SQL & " and codprove=" & Text1(4).Text & " and numline2=" & RS!numlinea
            SQL = SQL & " ORDER BY codartic "
            
            Set RSseries = New ADODB.Recordset
            RSseries.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not RSseries.EOF
                linea = linea + 1
                SQL = "UPDATE tmpnseries SET numserie=" & DBSet(RSseries!numSerie, "T")
                SQL = SQL & " WHERE codusu=" & vUsu.Codigo & " and codartic=" & DBSet(RS!codArtic, "T")
                SQL = SQL & " and numlinealb=" & RS!numlinea
                SQL = SQL & " and numlinea=" & linea
                conn.Execute SQL
                RSseries.MoveNext
            Wend
            RS.MoveNext
        Wend
        RSseries.Close
        Set RSseries = Nothing
        
        SQL = "select count(*) from tmpnseries Where codusu=" & vUsu.Codigo
        If RegistrosAListar(SQL) > 0 Then
            Set frmNSerie = New frmRepCargarNSerie
            frmNSerie.DeVentas = False 'Se llama desde Alb. de compras
            frmNSerie.NumAlb = Text1(0).Text
            frmNSerie.Show vbModal
            Set frmNSerie = Nothing
            Espera 0.2
        Else
            MsgBox "No hay n� de serie asignados a ese albaran", vbInformation
        End If
        
EPedirNSeries:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub CargarNumSeries()
'Insertar un registro en la tabla "sserie" por cada uno de los
'N� de Serie introducidos en la Tabla Temporal o actualizalo
Dim SQL As String
Dim b As Boolean

    On Error GoTo ECargar
    conn.BeginTrans
    
    'Borrar todos los N� de Serie asignados a ese albaran de compra
    'y que no tienen asignado ya un albaran de venta
    SQL = "DELETE FROM sserie "
    SQL = SQL & " WHERE codprove=" & Val(Text1(4).Text) & " and numalbpr=" & DBSet(Text1(0).Text, "T")
    SQL = SQL & " and fechacom='" & Format(Text1(1).Text, FormatoFecha) & "'"
    SQL = SQL & " and (isnull(numalbar) and isnull(numfactu))"
    conn.Execute SQL
    
    'Si algun N� serie tenia asignado albaran venta y no lo pude borrar entonces limpiamos
    'los campos del albaran de compra
    SQL = "UPDATE sserie SET codprove=" & ValorNulo & ", numalbpr=" & ValorNulo & ", fechacom="
    SQL = SQL & ValorNulo & ", numline2=" & ValorNulo
    SQL = SQL & " WHERE codprove=" & Val(Text1(4).Text) & " and numalbpr=" & DBSet(Text1(0).Text, "T")
    SQL = SQL & " and fechacom='" & Format(Text1(1).Text, FormatoFecha) & "'"
    conn.Execute SQL
    
    b = InsertarNumSeriesDeTMP
    
   
ECargar:
    If Err.Number <> 0 Then b = False
    If b Then
        conn.CommitTrans
    Else
        conn.RollbackTrans
    End If
    
End Sub


Private Function InsertarNumSeriesDeTMP() As Boolean
'Inserta en la tabla sserie todos los n� de serie q se han cargado en la temporal
Dim SQL As String
Dim Numalbar As String
Dim b As Boolean
Dim RStmp As ADODB.Recordset
Dim nSerie As CNumSerie

    On Error GoTo EInsertarNSeries

    'Inicializamos el objeto n� de serie con los valores comunes a todos
    Set nSerie = New CNumSerie
    nSerie.Proveedor = CInt(Text1(4).Text)
    nSerie.NumAlbProve = Text1(0).Text
    nSerie.fechacom = Text1(1).Text
    
    
    'Recuperar los N� Serie de ese articulo cargados en la Temporal
    'Seleccionar los n� de serie cargados en la temporal: tmpnseries
    SQL = "SELECT * FROM tmpnseries WHERE codusu=" & vUsu.Codigo
    SQL = SQL & " ORDER BY codartic, numlinealb, numlinea "
    
    Set RStmp = New ADODB.Recordset
    RStmp.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
                
    b = True
    While Not RStmp.EOF And b
        nSerie.numSerie = RStmp!numSerie
        nSerie.Articulo = RStmp!codArtic
        nSerie.NumLinAlbPr = RStmp!numlinealb
        
        'obtenemos los dias de garantia del articulo
        SQL = DevuelveDesdeBDNew(conAri, "sartic", "garantia", "codartic", RStmp!codArtic, "T")
        'fin garantia= fecha albaran + dias de garantia
        nSerie.FinGarantia = CStr(CDate(Text1(1).Text) + CInt(ComprobarCero(SQL)))
    
        'Comprobar si existe en la tabla sserie ese n� de serie
        Numalbar = "numalbpr" 'N� albaran de Venta prove
        SQL = DevuelveDesdeBDNew(conAri, "sserie", "numserie", "numserie", RStmp!numSerie, "T", Numalbar, "codartic", RStmp!codArtic, "T")
        If SQL <> "" Then
            If Numalbar = "" Then 'ya existe el n� serie y actualizamos ya que no esta asignado a ningun albaran
                b = nSerie.ActualizarNumSerie(False)
            End If
        Else
            b = nSerie.InsertarNumSerie
        End If
        
'        b = InsertarNSerie(RStmp!NumSerie, RStmp!codArtic, RStmp!NumLinealb)
        RStmp.MoveNext
    Wend
    RStmp.Close
    Set RStmp = Nothing
    
    Set nSerie = Nothing
    
EInsertarNSeries:
    If Err.Number <> 0 Then b = False
    If Not b Then
        InsertarNumSeriesDeTMP = False
    Else
        InsertarNumSeriesDeTMP = True
    End If
End Function




Private Sub PonerDatosProveedor(Codprove As String, Optional nifProve As String)
Dim vProve As CProveedor
Dim Observaciones As String
    
    On Error GoTo EPonerDatos
    
    If Codprove = "" Then
        LimpiarDatosProve
        Exit Sub
    End If

    Set vProve = New CProveedor
    'si se ha modificado el proveedor volver a cargar los datos
    If vProve.Existe(Codprove) Then
        If vProve.LeerDatos(Codprove) Then
            'NUEVO. Situacion proveedor
            If vProve.ProveedorBloqueado Then
                LimpiarDatosProve
                Set vProve = Nothing
                PonerFoco Text1(4)
                Exit Sub
            End If
            
            EsDeVarios = vProve.DeVarios
            BloquearDatosProve (EsDeVarios)
        
            If Modo = 4 And EsDeVarios Then 'Modificar
                'si no se ha modificado el proveedor no hacer nada
                If CLng(Text1(4).Text) = CLng(Data1.Recordset!Codprove) Then
                    Set vProve = Nothing
                    Exit Sub
                End If
            End If
        
            Text1(4).Text = vProve.Codigo
            FormateaCampo Text1(4)
            If (Modo = 3) Or (Modo = 4) Then
                Text1(5).Text = vProve.Nombre  'Nom prove
                Text1(8).Text = vProve.Domicilio
                Text1(9).Text = vProve.CPostal
                Text1(10).Text = vProve.Poblacion
                Text1(11).Text = vProve.Provincia
                Text1(6).Text = vProve.NIF
                Text1(7).Text = DBLet(vProve.TfnoAdmon, "T")
            End If
            
            If Modo = 3 Then 'insertar
                Text1(12).Text = vProve.ForPago
                Text2(12).Text = PonerNombreDeCod(Text1(12), conAri, "sforpa", "nomforpa")
                Text1(13).Text = Format(vProve.DtoPPago, FormatoDescuento)
                Text1(14).Text = Format(vProve.DtoGnral, FormatoDescuento)
            End If


            'Solo insertando
            If Modo = 3 Then vProve.PonerObservaciones Text1(15), Text1(16), Text1(17), Text1(18), Text1(19)


            Observaciones = DBLet(vProve.Observaciones)
            If Observaciones <> "" Then
                MsgBox Observaciones, vbInformation, "Observaciones del proveedor"
            End If
        End If
    Else
        LimpiarDatosProve
        If Modo = 3 Then PonerFoco Text1(4)
        
    End If
    Set vProve = Nothing

EPonerDatos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poner Datos Proveedor", Err.Description
End Sub


Private Sub PonerDatosProveVario(nifProve As String)
'Poner el los campos Text el valor del cliente
Dim vProve As CProveedor
Dim b As Boolean
   
    If nifProve = "" Then Exit Sub
   
    Set vProve = New CProveedor
    b = vProve.LeerDatosProveVario(nifProve)
    If b Then
        Text1(5).Text = vProve.Nombre   'Nom proveedor
        Text1(8).Text = vProve.Domicilio
        Text1(9).Text = vProve.CPostal
        Text1(10).Text = vProve.Poblacion
        Text1(11).Text = vProve.Provincia
        Text1(7).Text = DBLet(vProve.TfnoAdmon, "T")
    End If
    Set vProve = Nothing
End Sub



Private Sub BloquearDatosProve(bol As Boolean)
Dim i As Byte

    'bloquear/desbloquear campos de datos segun sea de varios o no
    If Modo <> 5 Then
        Me.imgBuscar(5).visible = bol 'NIF
        Me.imgBuscar(5).Enabled = bol 'NIF
        Me.imgBuscar(2).Enabled = bol 'poblacion
        
        For i = 5 To 11 'si no es de varios no se pueden modificar los datos
            BloquearTxt Text1(i), Not bol
        Next i
    End If
End Sub


Private Function ActualizarProveVarios(Prove As String, NIF As String) As Boolean
Dim vProve As CProveedor

    On Error GoTo EActualizarCV

    ActualizarProveVarios = False
    
    Set vProve = New CProveedor
    If EsProveedorVarios(Prove) Then
        vProve.NIF = NIF
        vProve.Nombre = Text1(5).Text
        vProve.Domicilio = Text1(8).Text
        vProve.CPostal = Text1(9).Text
        vProve.Poblacion = Text1(10).Text
        vProve.Provincia = Text1(11).Text
        vProve.TfnoAdmon = Text1(7).Text
        vProve.ActualizarProveV (NIF)
    End If
    Set vProve = Nothing
    
    ActualizarProveVarios = True
    
EActualizarCV:
    If Err.Number <> 0 Then
        ActualizarProveVarios = False
    Else
        ActualizarProveVarios = True
    End If
End Function



Private Sub CalcularDatosFactura()
Dim i As Byte
Dim cadWhere As String
Dim vFactu As CFacturaCom

    'Limpiar en el form los datos calculados de la factura
    'y volvemos a recalcular
    For i = 33 To 50
         Text3(i).Text = ""
    Next i
    
    cadWhere = ObtenerWhereCP(False)
    
    Set vFactu = New CFacturaCom
    vFactu.DtoPPago = CCur(ComprobarCero(Text1(13).Text))
    vFactu.DtoGnral = CCur(ComprobarCero(Text1(14).Text))
    vFactu.FijarTipoIvaProveedor Val(Text1(4).Text)
    If vFactu.CalcularDatosFactura2(cadWhere, NombreTabla, NomTablaLineas, CDate(Text1(1).Text), False) Then
        Text3(33).Text = vFactu.BrutoFac
        Text3(34).Text = vFactu.ImpPPago
        Text3(35).Text = vFactu.ImpGnral
        Text3(36).Text = vFactu.BaseImp
        Text3(37).Text = QuitarCero(vFactu.TipoIVA1)
        Text3(38).Text = QuitarCero(vFactu.TipoIVA2)
        Text3(39).Text = QuitarCero(vFactu.TipoIVA3)
        Text3(40).Text = vFactu.PorceIVA1
        Text3(41).Text = vFactu.PorceIVA2
        Text3(42).Text = vFactu.PorceIVA3
        Text3(43).Text = vFactu.BaseIVA1
        Text3(44).Text = vFactu.BaseIVA2
        Text3(45).Text = vFactu.BaseIVA3
        Text3(46).Text = vFactu.ImpIVA1
        Text3(47).Text = vFactu.ImpIVA2
        Text3(48).Text = vFactu.ImpIVA3
        Text3(49).Text = vFactu.TotalFac
        Text3(50).Text = vFactu.BaseImp
        
        FormatoDatosTotales
    Else
        MuestraError Err.Number, "Calculando Totales", Err.Description
    End If
    Set vFactu = Nothing
End Sub



Private Sub FormatoDatosTotales()
Dim i As Byte

    For i = 33 To 36
        If i = 34 Or i = 35 Then Text3(i).Text = QuitarCero(Text3(i).Text)
        Text3(i).Text = Format(Text3(i).Text, FormatoImporte)
    Next i
    
    'Desglose B.Imponible por IVA
    For i = 43 To 45
        If Text3(i).Text <> "" Then
             If CSng(Text3(i).Text) = 0 And Text3(i - 6).Text = "" Then
                Text3(i).Text = QuitarCero(Text3(i).Text)
                Text3(i - 3).Text = QuitarCero(Text3(i - 3).Text)
                Text3(i - 6).Text = QuitarCero(Text3(i - 6).Text)
                Text3(i + 3).Text = QuitarCero(Text3(i + 3).Text)
            Else
                Text3(i).Text = Format(Text3(i).Text, FormatoImporte)
                Text3(i - 3) = Format(Text3(i - 3).Text, FormatoDescuento)
    '            Text3(i - 6) = Format(Text3(i - 6).Text, "000")
                Text3(i + 3).Text = Format(Text3(i + 3).Text, FormatoImporte)
            End If
        End If
    Next i
    
    'Formatear el total de Factura
    Text3(49).Text = Format(Text3(49).Text, FormatoImporte)
    Text3(50).Text = Format(Text3(50).Text, FormatoImporte)
End Sub




Private Sub ComprobarNumSeries(numlinea As String)
'Comprobamos para una linea de Albaran si el articulo tiene control de n� de serie
'y procedemos
Dim SQL As String
Dim cadW As String
Dim RSLineas As ADODB.Recordset
'Dim Mostrar As Boolean 'Indica si vamos a pedir num series o a mostrarlos
'Dim cant As Integer 'cantidad que vamos a insertar


    'si la cantidad es >0 pedimos n� serie articulos comprados
    'si la cantidad es <0 mostramos los n� serie para devolver (ABONOS)
                    
    SQL = DevuelveDesdeBDNew(conAri, "sartic", "nseriesn", "codartic", txtAux(1).Text, "T")
    
    If SQL = "1" Then 'Si el Articulo tiene control de n� de serie
'        If Modo = 5 Then
'            If ModificaLineas = 1 Then 'INSERTAR linea
'                If CCur(txtAux(3).Text) > 0 Then 'cantidad linea
'                    Mostrar = False
'                    cant = CSng(txtAux(3).Text)
'                ElseIf CCur(txtAux(3).Text) < 0 Then 'cantidad linea
'                    'Es un ABONO
'                    'cantidad es < 0 (es un abono, devolvemos el articulo comprado)
'                    Mostrar = True
'                End If
'
'            ElseIf ModificaLineas = 2 Then 'MODIFICAR linea
'                'comprobar que la cantidad introducida se ha modificado
'                If CSng(txtAux(3).Text) <> CSng(Data2.Recordset!Cantidad) Then
'                    cant = CSng(txtAux(3).Text) - CSng(Data2.Recordset!Cantidad)
'                    If cant > 0 Then 'a�adir nuevos num serie
'                        Mostrar = False
'                    ElseIf cant < 0 Then 'mostrar num serie y quitar el que toca
'                        Mostrar = True
'                    End If
'                Else
'                    Exit Sub
'                End If
'            End If
'        End If
        
        
        
        If CCur(txtAux(3).Text) > 0 Then 'cantidad
'        If Mostrar = False Then
                SQL = "El Articulo tiene control de N� de Serie." & vbCrLf & vbCrLf
                SQL = SQL & "Introduzca los N� de Serie"
                If ModificaLineas = 2 Then
                    SQL = SQL & " que se han a�adido"
                End If
                MsgBox SQL & "." & vbCrLf, vbInformation
                'Cargar la tabla temporal con tantas filas como cantidad de Articulo
                'Para introducir el N� de Serie
                DescargarDatosTMPNumSeries "tmpnseries"
                CargarDatosTMPNumSeries "tmpnseries", txtAux(1).Text, CInt(txtAux(3).Text), numlinea
                'Visualizar en pantalla el Grid, y rellenar los N� Serie
                ModificaLineas = 0
                Set frmNSerie = New frmRepCargarNSerie
                frmNSerie.DeVentas = False
                frmNSerie.NumAlb = ""
                frmNSerie.Show vbModal
                Set frmNSerie = Nothing
                
        Else   'cantidad es < 0 (es un ABONO, devolvemos el articulo comprado)
           
            'Comprobar que efectivamente estan en tabla sserie los N�Serie del Articulo
            ' y que no esten asignados ya a otro albaran de venta
            SQL = " select distinct count(numserie) from sserie "
            cadW = " WHERE codartic=" & DBSet(txtAux(1).Text, "T")
            cadW = cadW & " and codprove=" & Text1(4).Text
            cadW = cadW & " and (numalbar='' or isnull(numalbar))"
            SQL = SQL & cadW
            
            If RegistrosAListar(SQL) > 0 Then 'Hay N� de Serie para elegir
                'mostrar los n� de serie de ese proveedor que no esten vendidos y selecccionar
                'el que vamos a devolver
                'Seleccionamos aquellas lineas de albaran que tienen N� de Serie
                SQL = "SELECT codartic, sum(cantidad) as cantidad, numlinea "
                SQL = SQL & " FROM " & NomTablaLineas
                
                cadW = " WHERE numalbar=" & DBSet(Text1(0).Text, "T") & " and "
                cadW = cadW & " fechaalb=" & DBSet(Text1(1).Text, "F")
                cadW = cadW & " and codprove= " & Text1(4).Text & " and numlinea=" & numlinea

                SQL = SQL & cadW
                SQL = SQL & " GROUP BY codartic ORDER BY Codartic "

                Set RSLineas = New ADODB.Recordset
                RSLineas.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

                MostrarNSeries RSLineas
                RSLineas.Close
                Set RSLineas = Nothing
            End If
        End If
    End If
End Sub



Private Sub MostrarNSeries(ByRef RSLineas As ADODB.Recordset)
'Si los N� de serie se introdujeron en ALBARAN COMPRAS se muestran
'los N� de serie de los articulos comprados y se seleccionamos
'los que vamos a devolver (Para ABONOS)
Dim SQL As String
Dim Campos As String
On Error GoTo EMostrarNSeries

    SQL = MostrarNSeriesGnral(RSLineas, Campos)
    SQL = SQL & " and sserie.codprove=" & Text1(4).Text
    
    Set frmMen = New frmMensajes
    frmMen.cadWhere = SQL
    frmMen.OpcionMensaje = 4 'N� Series Articulo
    frmMen.vCampos = Campos
    frmMen.Show vbModal
    Set frmMen = Nothing
    
EMostrarNSeries:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Function ModificarCabAlbaran() As Boolean
Dim b As Boolean
Dim MenError As String

    On Error GoTo EModificaAlb

    conn.BeginTrans

    MenError = "Modificando fecha de albaran en tablas relacionadas."
    b = ComprobarCambioFecha
                
    If b Then
        If (CDate(Text1(30).Text) <> CDate(Data1.Recordset!fentrada)) Then
            'Actualizamos la fecha en la tabla smoval
            MenError = "UPDATE smoval SET fechamov=" & DBSet(Text1(30).Text, "F")
            MenError = MenError & " WHERE document = " & DBSet(Data1.Recordset!Numalbar, "T")
            MenError = MenError & " AND fechamov=" & DBSet(Data1.Recordset!fentrada, "F")
            MenError = MenError & " AND codigope=" & Data1.Recordset!Codprove
            MenError = MenError & " AND detamovi='" & CodTipoMov & "'"
            If Not ejecutar(MenError, True) Then b = False
                
        End If
    End If
    If b Then
        MenError = "Modificando el albaran (scaalb)."
        b = ModificaDesdeFormulario(Me, 1)
        
        If b Then
            'Actualizar los datos del Proveedor si es de varios
            MenError = "Actualizando proveedor de varios."
            b = ActualizarProveVarios(Text1(4).Text, Text1(6).Text)
        End If
    End If

EModificaAlb:
    If Err.Number <> 0 Then b = False
    If b Then
        conn.CommitTrans
    Else
        conn.RollbackTrans
        MsgBox "Error Modificando el albaran." & vbCrLf & MenError, vbExclamation
    End If
    ModificarCabAlbaran = b
    Espera 0.2
End Function




'Private Function ArticuloTieneMargen() As Boolean
'Dim cad As String
'
'    'Comprobar que el art�culo tiene margen comercial
'    cad = DevuelveDesdeBDNew(conAri, "sartic", "margecom", "codartic", txtAux(1).Text, "T")
'    If cad = "" Then
'        cad = "NO SE HAN PODIDO ACTUALIZAR LOS PRECIOS." & vbCrLf
'        cad = cad & "El art�culo no tiene margen comercial para calcular nuevos precios."
'        MsgBox cad, vbExclamation
'        ArticuloTieneMargen = False
'        Exit Function
'    End If
'
'
'    'comprobar que las tarifas tienen margen comercial
'    cad = "SELECT count(*)"
'    cad = cad & " FROM slista INNER JOIN starif ON slista.codlista = starif.codlista "
'    cad = cad & " WHERE slista.codartic=" & DBSet(txtAux(1).Text, "T") & " AND  isnull(margecom)"
'    If RegistrosAListar(cad) > 0 Then
'        cad = "NO SE HAN PODIDO ACTUALIZAR LOS PRECIOS." & vbCrLf
'        cad = cad & "El art�culo tiene tarifas sin %PVP necesario para calcular nuevos precios."
'        MsgBox cad, vbExclamation
'        ArticuloTieneMargen = False
'        Exit Function
'    End If
'
'    ArticuloTieneMargen = True
'
'End Function


Private Sub AbrirForm_CentroCoste()
'    Screen.MousePointer = vbHourglass
'    cmdAux(2).Tag = "2"
'
'    Set frmB = New frmBuscaGrid
'    If vParamAplic.ContabilidadNueva Then
'        frmB.vCampos = "Codigo|ccoste|codccost|T||20�Descripci�n|ccoste|nomccost|T||70�"
'        frmB.vTabla = "ccoste"
'    Else
'        frmB.vCampos = "Codigo|cabccost|codccost|T||20�Descripci�n|cabccost|nomccost|T||70�"
'        frmB.vTabla = "cabccost"
'    End If
'    frmB.vSQL = ""
'    HaDevueltoDatos = False
'    '###A mano
'    frmB.vDevuelve = "0|1|"
'    frmB.vTitulo = "Centros de coste"
'    frmB.vselElem = 0
'    frmB.vConexionGrid = conConta
'
'    frmB.Show vbModal
'    Set frmB = Nothing
'    cmdAux(2).Tag = "-1"

    Set frmCCos = New frmBasico2
    AyudaCentroCoste frmCCos, txtAux(9)
    Set frmCCos = Nothing
    
End Sub


' ---- [02/11/2009] [LAURA] : al pulsar F2 para abrir articulos en la solapa Documentos|Pedidos
Private Sub AbrirForm_Articulos()
    If Trim(txtAux(1).Text) = "" Then Exit Sub
    
'    Set frmArt = New frmAlmArticulos
'    frmArt.DeConsulta = True
'    frmArt.DatosADevolverBusqueda3 = "::" & Trim(txtAux(1).Text)  'DevNombreSQL(Data2.Recordset!codarti1)
'    frmArt.parNumTAb = 6
'    frmArt.Show vbModal
'    Set frmArt = Nothing
    
    
'    frmAlmArticulos.DeConsulta = True
'    frmAlmArticulos.DatosADevolverBusqueda = "::" & Trim(txtAux(1).Text)  'DevNombreSQL(Data2.Recordset!codarti1)
'    frmAlmArticulos.parNumTAb = 6
'    frmAlmArticulos.Show vbModal
'    Set frmAlmArticulos = Nothing
'
    
    
    frmAlmArticulosGr.DeConsulta = True
    frmAlmArticulosGr.DatosADevolverBusqueda = "::" & Trim(txtAux(1).Text)  'DevNombreSQL(Data2.Recordset!codarti1)
    frmAlmArticulosGr.parNumTAb = 6
    frmAlmArticulosGr.Show vbModal
    Set frmAlmArticulosGr = Nothing
    
    
    
    
    
End Sub
' -----


Private Sub ModificarProveedor()
Dim OK As Boolean
    OK = True
    If EsHistorico Then
        OK = False
    Else
            If Modo = 2 Then
                If Data1.Recordset Is Nothing Then
                    OK = False
                Else
                    If Data1.Recordset.EOF Then
                       OK = False
                    Else
                        'data1.Recordset!esdevarios
                        If EsDeVarios Then
                            MsgBox "Proveedor de VARIOS", vbExclamation
                            OK = False
                        End If
                    End If
                End If
            Else
                OK = False
            End If
    End If
    If OK Then
        If vUsu.Nivel > 1 Then
            MsgBox "usuario sin permiso", vbExclamation
            OK = False
        End If
    End If
    If Not OK Then Exit Sub
    

    
    CadenaDesdeOtroForm = ""
    frmListado2.Opcion = 18
    frmListado2.Show vbModal
    If CadenaDesdeOtroForm = "" Then Exit Sub
    'Si es el mismo no hago nada
    If (CLng(CadenaDesdeOtroForm)) = CLng(Text1(4).Text) Then
        MsgBox "Mismo proveedor", vbExclamation
        Exit Sub
    End If
    

        
        
     Screen.MousePointer = vbHourglass
    'pedimos el nuevo proveedor
    Set miRsAux = New ADODB.Recordset
    conn.BeginTrans
    OK = HacerUpdatesCodProve(CLng(CadenaDesdeOtroForm))
    If OK Then
        'Situamos
        conn.CommitTrans
    Else
        conn.RollbackTrans
    End If
    If OK Then
        'Sitauremos
        CadenaDesdeOtroForm = " numalbar = " & DBSet(Text1(0).Text, "T") & " AND fechaalb = " & DBSet(Text1(1).Text, "F") & " AND codprove = " & CadenaDesdeOtroForm
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadenaDesdeOtroForm & " " & Ordenacion
        PonerCadenaBusqueda
        
    End If
    Screen.MousePointer = vbDefault
    
    
    
End Sub




Private Function HacerUpdatesCodProve(NuevoProve As Long) As Boolean
Dim CadenaLineas As String
Dim J As Integer
Dim SQL As String
Dim vPr As CProveedor
        
        On Error GoTo EHacerUpdatesCodProve
        HacerUpdatesCodProve = False
        
        Set vPr = New CProveedor
        If Not vPr.LeerDatos(CStr(NuevoProve)) Then
            Set vPr = Nothing
            Exit Function
        End If
        
            
        
        
        SQL = "Select "
        SQL = SQL & "fechaalb,numlotes,numalbar,codartic,ampliaci,nomartic,numlinea,codalmac,cantidad,precioar,"
        SQL = SQL & "dtoline1,dtoline2,importel,codprove,codccost"
        SQL = SQL & " FROM slialp"
        SQL = SQL & " WHERE numalbar = " & DBSet(Text1(0).Text, "T") & " AND fechaalb = " & DBSet(Text1(1).Text, "F") & " AND codprove = " & Text1(4).Text
        miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        CadenaLineas = ""
        While Not miRsAux.EOF
            SQL = ", ('" & Format(miRsAux!FechaAlb, FormatoFecha) & "'"
            'Texto
            For J = 1 To 5
                If IsNull(miRsAux.Fields(J)) Then
                    SQL = SQL & ",NULL"
                Else
                    SQL = SQL & ",'" & DevNombreSQL(miRsAux.Fields(J)) & "'"
                End If
            Next J
            'Numero
                        'Texto
            For J = 6 To 12
                If IsNull(miRsAux.Fields(J)) Then
                    SQL = SQL & ",NULL"
                Else
                    SQL = SQL & "," & TransformaComasPuntos(CStr(miRsAux.Fields(J)))
                End If
            Next J
            'Nuevo proveedor
            SQL = SQL & "," & NuevoProve
            
            'Abril 2011.  codccost
            SQL = SQL & "," & DBSet(miRsAux!CodCCost, "T", "S")
            
            CadenaLineas = CadenaLineas & SQL & ")"

            'Sig
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        'Borramos las lineas
        If CadenaLineas <> "" Then
                
            SQL = "DELETE FROM slialp WHERE numalbar = " & DBSet(Text1(0).Text, "T") & " AND fechaalb = " & DBSet(Text1(1).Text, "F") & " AND codprove = " & Text1(4).Text
            conn.Execute SQL
                    
            SQL = "INSERT INTO slialp ("
            SQL = SQL & "fechaalb,numlotes,numalbar,codartic,ampliaci,nomartic,numlinea,codalmac,cantidad,precioar,"
            SQL = SQL & "dtoline1,dtoline2,importel,codprove,codccost"
            'Quito la primara coma
            CadenaLineas = Mid(CadenaLineas, 2)
            SQL = SQL & ") VALUES " & CadenaLineas
            CadenaLineas = SQL
        
        End If
        
        'ACtualizamos
        'Busco los datos del proveedor
        SQL = "UPDATE scaalp SET codprove = " & NuevoProve
        'Resto de datos del proveedor: nomprove,domprove,codpobla,pobprove,proprove,nifprove,telprove
        SQL = SQL & ", nomprove=" & DBSet(vPr.Nombre, "T")
        SQL = SQL & ",domprove=" & DBSet(vPr.Domicilio, "T")
        SQL = SQL & ",codpobla=" & DBSet(vPr.CPostal, "T")
        SQL = SQL & ",pobprove=" & DBSet(vPr.Poblacion, "T")
        SQL = SQL & ",proprove=" & DBSet(vPr.Provincia, "T")
        SQL = SQL & ",nifprove=" & DBSet(vPr.NIF, "T")
        SQL = SQL & ",telprove=" & DBSet(vPr.TfnoAdmon, "T", "S")
        SQL = SQL & " WHERE"
        SQL = SQL & " numalbar = " & DBSet(Text1(0).Text, "T") & " AND fechaalb = " & DBSet(Text1(1).Text, "F") & " AND codprove = " & Text1(4).Text
        conn.Execute SQL
        Set vPr = Nothing
        If CadenaLineas <> "" Then
                'meto las lineas con el nuevo proveedor
                conn.Execute CadenaLineas
                
                'UPDATEO las tablas de smoval
                SQL = "UPDATE smoval SET codigope = " & NuevoProve
                SQL = SQL & " WHERE detamovi='ALC' AND "
                SQL = SQL & " document = " & DBSet(Text1(0).Text, "T") & " AND fechamov = " & DBSet(Text1(1).Text, "F") & " AND codigope = " & Text1(4).Text
                conn.Execute SQL
                
                
        End If
        
        
        HacerUpdatesCodProve = True
        Exit Function
EHacerUpdatesCodProve:
    MuestraError Err.Number, Err.Description
End Function


Private Sub PonerUltAlmacen()
Dim C As String
    
       If Not Data2.Recordset.EOF Then
            C = ObtenerWhereCP(True)
            C = Replace(C, NombreTabla, NomTablaLineas)
            AlmacenLineas = DevuelveUltimoAlmacen(NomTablaLineas, C)
        Else
            AlmacenLineas = -1
       End If
            
       If AlmacenLineas < 0 Then
            'No hay datos todavia
            '                                                                trabajador
            C = DevuelveDesdeBDNew(conAri, "straba", "codalmac", "codtraba", Text1(2).Text, "N")
            If C <> "" Then AlmacenLineas = Val(C)
        End If
    
End Sub





'Nuevo. Cuando pulse MAS (y es el primer carcater abre el prismatico asociado)
Private Sub PulsarTeclaMas(InsertandoCabecera As Boolean, Index As Integer)

    If InsertandoCabecera Then
        If imgBuscar(Index).visible Then imgBuscar_Click Index
        
    Else
        'Lineas
        If Index = 9 Then Index = 2
        cmdAux_Click Index
        
        
    End If
        
End Sub


Private Function SituarData() As Boolean
    On Error GoTo ES
    SituarData = False
    '
     'DBSet(Text1(0).Text, "T") & " and " & NombreTabla & ".fechaalb='" & Format(Text1(1).Text, FormatoFecha)
    'SQL = SQL & "' and " & NombreTabla & ".codprove=" & Val(Text1(4).Text)
    
    While Not Data1.Recordset.EOF
        If Val(Data1.Recordset!Codprove) = Val(Text1(4).Text) Then
            If Data1.Recordset!Numalbar = Text1(0).Text Then
                If Format(Data1.Recordset!FechaAlb, "dd/mm/yyyy") = Text1(1).Text Then
                    Me.lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
                    SituarData = True
                    Exit Function
                End If
            End If
        End If
        Data1.Recordset.MoveNext
    Wend
    
    'Si llega aqui no lo ha encontrado
    Exit Function
ES:
    MuestraError Err.Number
End Function


Private Sub EliminarSinStocks()
Dim vWhere As String

    cadList = String(50, "*") & vbCrLf
    cadList = cadList & "Desea llevar a hist�rico el albar�n....."
    cadList = cadList & vbCrLf & "N�:  " & Text1(0).Text
    cadList = cadList & vbCrLf & "Fecha: " & Text1(1).Text
    cadList = cadList & vbCrLf & vbCrLf & " �Continuar? "
    cadList = cadList & vbCrLf & vbCrLf & String(50, "*") & vbCrLf

    If MsgBox(cadList, vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
        Screen.MousePointer = vbHourglass
    
        NumRegElim = Data1.Recordset.AbsolutePosition

        'Abrir frame de informes para pedir datos antes de grabar en el historico
        cadList = ""
        Set frmList = New frmListadoOfer
        frmList.OpcionListado = 80
        frmList.Show vbModal
        Set frmList = Nothing
        If cadList = "" Then Exit Sub
        
        
        vWhere = ObtenerWhereCP(False)
        
        If ActualizarElTraspaso("", vWhere, CodTipoMov, cadList) Then
            CargaGrid Me.DataGrid1, Me.Data2, True
            Exit Sub
        ElseIf SituarDataTrasEliminar(Data1, NumRegElim) Then
            PonerCampos
        Else
            LimpiarCampos
            'Poner los grid sin apuntar a nada
            LimpiarDataGrids
            PonerModo 0
        End If

    End If
End Sub

Private Sub InsertarSubComponentes()
Dim SQL As String
Dim TipoDto As Byte
Dim J As Integer
Dim cantidad As Currency
Dim numlinea As Integer
Dim vCStock As CStock
'Si el articulo es de conjuntos, preguntara si quiere insertar la lineas de los conjuntos
   
        SQL = DevuelveDesdeBD(conAri, "conjunto", "sartic", "codartic", txtAux(1).Text, "T")
        If SQL = "1" Then
        
        
            'SI!!!!!!, es de conjuntos
            If MsgBox("Articulo con componentes. Desea insertar las lineas?", vbQuestion + vbYesNo) <> vbYes Then Exit Sub
            

            SQL = DevuelveDesdeBDNew(conAri, "sprove", "tipodtos", "codprove", Text1(4).Text, "N")
            TipoDto = CByte(SQL)
            cantidad = ImporteFormateado(txtAux(3).Text)
            
            SQL = ObtenerWhereCP(False)
            SQL = Replace(SQL, NombreTabla, NomTablaLineas)
            numlinea = SugerirCodigoSiguienteStr(NomTablaLineas, "numlinea", SQL)
                 
   
   
            SQL = "Select sarti1.*,nomartic from sarti1,sartic where sarti1.codarti1=sartic.codartic and sarti1.codartic=" & DBSet(txtAux(1).Text, "T")
            Set miRsAux = New ADODB.Recordset
            'miRsAux.Open SQL, conn, adOpenKeyset, adLockPessimistic, adCmdText
            miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            SQL = ""
            While Not miRsAux.EOF
                'Limpiamos todo menos el almacen y el CC si lo tuviera
                For J = 1 To 8
                    txtAux(J).Text = ""
                Next
                
                
                Text2(16).Text = ""
                Text2(17).Text = ""
                txtAux(1).Text = miRsAux!codarti1
                txtAux(2).Text = miRsAux!NomArtic
                'Cantidad es la cantidad de la linea ppal * la del escandallo
                txtAux(3).Text = cantidad * miRsAux!cantidad
            
                ObtenerPrecioCompra
            
                
                txtAux(7).Text = CalcularImporteSng(txtAux(3).Text, txtAux(4).Text, txtAux(5).Text, txtAux(6).Text, TipoDto)
            
            
                Set vCStock = New CStock
                If InicializarCStock(vCStock, "E", CStr(numlinea)) Then
                
            
            
                    SQL = "INSERT INTO " & NomTablaLineas
                    SQL = SQL & " (numalbar, fechaalb, codprove, numlinea, codalmac, codartic, nomartic, ampliaci, cantidad, precioar, dtoline1, dtoline2, importel,numlotes,codccost) "
                    SQL = SQL & "VALUES (" & DBSet(Text1(0).Text, "T") & ", " & DBSet(Text1(1).Text, "F") & ", " & Val(Text1(4).Text) & ", " & numlinea & ", " & Val(txtAux(0).Text) & ","
                    SQL = SQL & DBSet(txtAux(1).Text, "T") & ", " & DBSet(txtAux(2).Text, "T") & ", " & DBSet(Text2(16).Text, "T") & ", "
                    SQL = SQL & DBSet(txtAux(3).Text, "N") & ", "
                    SQL = SQL & DBSet(txtAux(4).Text, "S") & ", " & DBSet(txtAux(5).Text, "N") & ", "
                    SQL = SQL & DBSet(txtAux(6).Text, "N") & ", "
                    SQL = SQL & DBSet(txtAux(7).Text, "N") & ", " & DBSet(Text2(17).Text, "T") & ","
                    SQL = SQL & DBSet(txtAux(9).Text, "T", "S") 'centro coste
                    SQL = SQL & ");"
                
                    If Not ejecutar(SQL, True) Then
                        MsgBox "Error a�adiendo el componente: " & miRsAux!NomArtic, vbExclamation
                    
                    Else
                        numlinea = numlinea + 1
                        If Not vCStock.ActualizarStock Then MsgBox "Error actualizando stock componentes: " & miRsAux!NomArtic, vbExclamation
                    End If
                Else
                    MsgBox "Error stock: " & miRsAux!NomArtic, vbExclamation
                End If
                    
                    
                Set vCStock = Nothing
                miRsAux.MoveNext
            Wend
            miRsAux.Close
            Set miRsAux = Nothing
        End If
        
 
    
End Sub



Private Function PuedeEliminar(EsLinea As Boolean) As Boolean
Dim RN As ADODB.Recordset
Dim Aux As String

    PuedeEliminar = True
    
    
    If Not vParamAplic.ManipuladorFitosanitarios2 Then Exit Function
    
    If EsLinea Then
        If DBLet(Data2.Recordset!numlotes, "T") = "" Then Exit Function
    End If
    
    cadList = ObtenerWhereCP(True)
    cadList = Replace(cadList, NombreTabla, NomTablaLineas)
    cadList = "Select codartic,numlotes,fechaalb from " & NomTablaLineas & cadList & " AND numlotes "
    If EsLinea Then
        cadList = cadList & " = " & DBSet(Data2.Recordset!numlotes, "T")
    Else
        cadList = cadList & " <> ''"
    End If
    
    Set miRsAux = New ADODB.Recordset
    Set RN = New ADODB.Recordset
    miRsAux.Open cadList, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    cadList = ""
    While Not miRsAux.EOF
        
        Aux = "SELECT * FROM slotes WHERE "
        Aux = Aux & " codartic=" & DBSet(miRsAux!codArtic, "T") & " AND numlotes=" & DBSet(miRsAux!numlotes, "T") & " AND fecentra=" & DBSet(miRsAux!FechaAlb, "F")
        RN.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If RN.EOF Then
            MsgBox "Lote no encontrado: " & miRsAux!codArtic & "    " & miRsAux!numlotes & "     " & miRsAux!FechaAlb, vbExclamation
        Else
            If RN!vendida > 0 Then
                cadList = cadList & "Ya vendida  -> " & miRsAux!codArtic & "    " & miRsAux!numlotes & vbCrLf
            Else
                'Vemos si existe LOTES realacionados en ventas, aunque NO deberia pasar
                Aux = "numlote =" & DBSet(miRsAux!numlotes, "T") & " AND fecentra = " & DBSet(miRsAux!FechaAlb, "F")
                Aux = Aux & " AND codartic = " & DBSet(miRsAux!codArtic, "T") & " AND 1"
                Aux = DevuelveDesdeBD(conAri, "count(*)", "slialblotes", Aux, "1")
                If Val(Aux) > 0 Then
                    cadList = cadList & "Lotes asignados  -> " & miRsAux!codArtic & "    " & miRsAux!numlotes & vbCrLf
                Else
                    Aux = "numlote =" & DBSet(miRsAux!numlotes, "T") & " AND fecentra = " & DBSet(miRsAux!FechaAlb, "F")
                    Aux = Aux & " AND codartic = " & DBSet(miRsAux!codArtic, "T") & " AND 1"
                    Aux = DevuelveDesdeBD(conAri, "count(*)", "slivenlotes", Aux, "1")
                    If Val(Aux) > 0 Then cadList = cadList & "En venta(I)  -> " & miRsAux!codArtic & "    " & miRsAux!numlotes & vbCrLf
                End If
            End If
        End If
        RN.Close
        miRsAux.MoveNext
    Wend
    miRsAux.Close
        
        
    If cadList <> "" Then
        cadList = "Error en lotes. " & vbCrLf & vbCrLf & cadList
        MsgBox cadList, vbExclamation
        PuedeEliminar = False
    End If
    
    Set miRsAux = Nothing
    Set RN = Nothing
End Function


Private Sub EliminarEnSlotes()
Dim cad As String
    Set miRsAux = New ADODB.Recordset
    cad = ObtenerWhereCP(True)
    cad = Replace(cad, NombreTabla, NomTablaLineas)
    cad = "Select codartic,numlotes,fechaalb from " & NomTablaLineas & cad & " AND numlotes  <> ''"
    miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
   
    While Not miRsAux.EOF
                
        cad = "DELETE FROM slotes WHERE "
        cad = cad & " codartic=" & DBSet(miRsAux!codArtic, "T") & " AND numlotes=" & DBSet(miRsAux!numlotes, "T") & " AND fecentra=" & DBSet(miRsAux!FechaAlb, "F")
        conn.Execute cad
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
End Sub



Private Sub RecepcionarAlbaran()
Dim SQL As String

    If Modo <> 2 Then Exit Sub
    If EsHistorico Or vUsu.Nivel > 1 Then Exit Sub
    
    'Si teiene lineas
    If Data2.Recordset.EOF Then
        MsgBox "Sin lineas", vbExclamation
        Exit Sub
    End If
    
    
    'Lanazaremos la recepccion de facvtura , cargando estos datos
    SQL = ObtenerWhereCP(False)
        
    frmComFacturarGR.CadenaAlbaran = SQL
    frmComFacturarGR.Codprove = Val(Text1(4).Text)
    frmComFacturarGR.Show vbModal
    
    SQL = ObtenerWhereCP(False)
    SQL = DevuelveDesdeBD(conAri, "numalbar", "scaalp", SQL & " AND 1", "1")
    If SQL = "" Then
        'Ha eliminado el registro
        NumRegElim = Data1.Recordset.AbsolutePosition
        If SituarDataTrasEliminar(Data1, NumRegElim) Then
            PonerCampos
        Else
            LimpiarCampos
            'Poner los grid sin apuntar a nada
            LimpiarDataGrids
            PonerModo 0
        End If
    End If
    
    
End Sub