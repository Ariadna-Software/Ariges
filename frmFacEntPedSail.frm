VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFacEntPedSail 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pedidos Clientes"
   ClientHeight    =   7815
   ClientLeft      =   45
   ClientTop       =   4335
   ClientWidth     =   16125
   Icon            =   "frmFacEntPedSail.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   16125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   120
      TabIndex        =   82
      Top             =   410
      Width           =   15855
      Begin VB.CheckBox chkServirCom 
         Caption         =   "Servir completo"
         Enabled         =   0   'False
         Height          =   240
         Left            =   4080
         TabIndex        =   4
         Tag             =   "Servir completo|N|N|||scaped|servcomp||N|"
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   5
         Left            =   11640
         MaxLength       =   60
         TabIndex        =   9
         Tag             =   "Nombre Cliente|T|N|||scaped|nomclien||N|"
         Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwwwwwwwwww aq"
         Top             =   480
         Width           =   3870
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   4
         Left            =   10845
         MaxLength       =   30
         TabIndex        =   8
         Tag             =   "Cod. Cliente|N|N|||scaped|codclien|000000|N|"
         Text            =   "Text1"
         Top             =   480
         Width           =   780
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   3
         Left            =   10845
         MaxLength       =   30
         TabIndex        =   7
         Tag             =   "Realizada Por|N|N|0|9999|scaped|codtraba|0000|N|"
         Text            =   "Text1"
         Top             =   130
         Width           =   660
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   3
         Left            =   11550
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   88
         Text            =   "Text2"
         Top             =   130
         Width           =   3975
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   1
         Left            =   1080
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Fecha Pedido|F|N|||scaped|fecpedcl|dd/mm/yyyy|N|"
         Top             =   360
         Width           =   1065
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000013&
         Height          =   315
         Index           =   0
         Left            =   120
         MaxLength       =   7
         TabIndex        =   0
         Tag             =   "Nº Pedido|N|S|0||scaped|numpedcl|0000000|S|"
         Text            =   "Text1 7"
         Top             =   360
         Width           =   885
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   2
         Left            =   2280
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "Fecha Entrega|F|N|||scaped|fecentre|dd/mm/yyyy|N|"
         Top             =   360
         Width           =   1065
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   18
         Left            =   3360
         MaxLength       =   10
         TabIndex        =   3
         Tag             =   "Semana Entrega|N|N|0|52|scaped|sementre|0|N|"
         Top             =   360
         Width           =   465
      End
      Begin VB.CheckBox chkVisadoRes 
         Caption         =   "Visado Responsable"
         Height          =   240
         Left            =   7680
         TabIndex        =   6
         Tag             =   "Visado Responsable|N|N|||scaped|visadore||N|"
         Top             =   360
         Width           =   1815
      End
      Begin VB.CheckBox chkRestoPed 
         Caption         =   "Resto de Pedido"
         Enabled         =   0   'False
         Height          =   240
         Left            =   5880
         TabIndex        =   5
         Tag             =   "Resto de Pedido|N|N|||scaped|restoped||N|"
         Top             =   355
         Width           =   1575
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   10545
         Picture         =   "frmFacEntPedSail.frx":000C
         ToolTipText     =   "Buscar cliente"
         Top             =   510
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente"
         Height          =   255
         Index           =   0
         Left            =   9720
         TabIndex        =   89
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Realiz. Por"
         Height          =   255
         Index           =   21
         Left            =   9720
         TabIndex        =   87
         Top             =   135
         Width           =   780
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   10545
         Picture         =   "frmFacEntPedSail.frx":010E
         ToolTipText     =   "Buscar trabajador"
         Top             =   165
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "F. Pedido"
         Height          =   255
         Index           =   14
         Left            =   1080
         TabIndex        =   86
         Top             =   165
         Width           =   735
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   1800
         Picture         =   "frmFacEntPedSail.frx":0210
         ToolTipText     =   "Buscar fecha"
         Top             =   120
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Pedido"
         Height          =   255
         Index           =   50
         Left            =   120
         TabIndex        =   85
         Top             =   165
         Width           =   735
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   1
         Left            =   3120
         Picture         =   "frmFacEntPedSail.frx":029B
         ToolTipText     =   "Buscar fecha"
         Top             =   120
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "F. Entrega"
         Height          =   255
         Index           =   51
         Left            =   2280
         TabIndex        =   84
         Top             =   165
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Semana"
         Height          =   255
         Index           =   8
         Left            =   3480
         TabIndex        =   83
         Top             =   165
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Height          =   475
      Index           =   0
      Left            =   120
      TabIndex        =   37
      Top             =   7320
      Width           =   2175
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   240
         Left            =   240
         TabIndex        =   38
         Top             =   180
         Width           =   1755
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   14880
      TabIndex        =   35
      Top             =   7320
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   13680
      TabIndex        =   55
      Top             =   7320
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   4680
      Top             =   7320
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
      TabIndex        =   39
      Top             =   0
      Width           =   16125
      _ExtentX        =   28443
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   24
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
            Object.ToolTipText     =   "Lineas Pedido"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Generar Albaran"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Facturar"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir Pedido"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir Orden Instal."
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Confirmación de entrega"
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Último"
            ImageIndex      =   9
         EndProperty
      EndProperty
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFC0&
         Height          =   285
         Index           =   56
         Left            =   10080
         MaxLength       =   15
         TabIndex        =   122
         Text            =   "Text1 7"
         Top             =   80
         Width           =   1530
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   285
         Index           =   0
         Left            =   8520
         MaxLength       =   15
         TabIndex        =   121
         Text            =   "TOTAL"
         Top             =   100
         Width           =   1490
      End
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   7320
         TabIndex        =   40
         Top             =   120
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc Data2 
      Height          =   330
      Left            =   3480
      Top             =   7320
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
      Height          =   5940
      Left            =   120
      TabIndex        =   41
      Tag             =   "Fecha Oferta|F|N|||scapre|fecentre|dd/mm/yyyy|N|"
      Top             =   1275
      Width           =   15855
      _ExtentX        =   27966
      _ExtentY        =   10478
      _Version        =   393216
      Style           =   1
      TabsPerRow      =   4
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Datos básicos"
      TabPicture(0)   =   "frmFacEntPedSail.frx":0326
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(6)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(35)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "imgBuscar2(13)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "imgBuscar2(12)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "imgBuscar2(11)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(27)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1(28)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "DataGrid1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtAux(0)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtAux(1)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtAux(3)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtAux(4)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtAux(6)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtAux(7)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtAux(8)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtAux(2)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "cmdAux(0)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "cmdAux(1)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "FrameCliente"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtAux(5)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtAux(9)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txtAux(10)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txtAux(11)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Text2(16)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "txtAux2(11)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "txtAux2(12)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "txtAux(12)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "txtAux2(13)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "txtAux(13)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).ControlCount=   29
      TabCaption(1)   =   "Otros Datos"
      TabPicture(1)   =   "frmFacEntPedSail.frx":0342
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Text1(33)"
      Tab(1).Control(1)=   "Text1(29)"
      Tab(1).Control(2)=   "Text1(30)"
      Tab(1).Control(3)=   "FrameHco"
      Tab(1).Control(4)=   "Text1(25)"
      Tab(1).Control(5)=   "Text1(24)"
      Tab(1).Control(6)=   "Text1(23)"
      Tab(1).Control(7)=   "Text1(22)"
      Tab(1).Control(8)=   "Text1(21)"
      Tab(1).Control(9)=   "Text1(20)"
      Tab(1).Control(10)=   "Text1(19)"
      Tab(1).Control(11)=   "Label1(29)"
      Tab(1).Control(12)=   "Label1(18)"
      Tab(1).Control(13)=   "Label1(5)"
      Tab(1).Control(14)=   "Label1(3)"
      Tab(1).Control(15)=   "Label1(45)"
      Tab(1).ControlCount=   16
      TabCaption(2)   =   "Totales"
      TabPicture(2)   =   "frmFacEntPedSail.frx":035E
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FrameFactura"
      Tab(2).ControlCount=   1
      Begin VB.TextBox Text1 
         Height          =   1005
         Index           =   33
         Left            =   -70080
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   34
         Tag             =   "Obs CRM|T|S|||scaped|observacrm|||"
         Top             =   4755
         Width           =   8805
      End
      Begin VB.TextBox txtAux 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   13
         Left            =   11400
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   53
         Text            =   "codc"
         Top             =   4680
         Width           =   615
      End
      Begin VB.TextBox txtAux2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   13
         Left            =   12000
         Locked          =   -1  'True
         MaxLength       =   60
         TabIndex        =   149
         Text            =   "nom ccoste"
         Top             =   4680
         Width           =   3525
      End
      Begin VB.TextBox txtAux 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   12
         Left            =   11400
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   52
         Text            =   "codc"
         Top             =   4080
         Width           =   615
      End
      Begin VB.TextBox txtAux2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   12
         Left            =   12000
         Locked          =   -1  'True
         MaxLength       =   60
         TabIndex        =   148
         Text            =   "nom ccoste"
         Top             =   4080
         Width           =   3525
      End
      Begin VB.TextBox txtAux2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   11
         Left            =   12120
         Locked          =   -1  'True
         MaxLength       =   60
         TabIndex        =   145
         Text            =   "nom ccoste"
         Top             =   5400
         Visible         =   0   'False
         Width           =   3525
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   795
         Index           =   16
         Left            =   11400
         Locked          =   -1  'True
         MaxLength       =   60
         MultiLine       =   -1  'True
         TabIndex        =   51
         Text            =   "frmFacEntPedSail.frx":037A
         Top             =   3000
         Width           =   4245
      End
      Begin VB.TextBox txtAux 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   11
         Left            =   11400
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   54
         Text            =   "codc"
         Top             =   5400
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   10
         Left            =   14640
         MaxLength       =   15
         TabIndex        =   56
         Text            =   "numlote"
         Top             =   2640
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   9
         Left            =   7440
         MaxLength       =   5
         TabIndex        =   47
         Tag             =   "Bultos"
         Text            =   "12345"
         Top             =   4080
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Height          =   280
         Index           =   29
         Left            =   -70080
         MaxLength       =   80
         TabIndex        =   32
         Tag             =   "Observación pedido 1|T|S|||scaped|observap1||N|"
         Top             =   3840
         Width           =   8805
      End
      Begin VB.TextBox Text1 
         Height          =   280
         Index           =   30
         Left            =   -70080
         MaxLength       =   80
         TabIndex        =   33
         Tag             =   "Observación pedido 2|T|S|||scaped|observap2||N|"
         Top             =   4200
         Width           =   8805
      End
      Begin VB.Frame FrameHco 
         Height          =   1400
         Left            =   -70080
         TabIndex        =   123
         Top             =   480
         Width           =   5775
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   26
            Left            =   1560
            MaxLength       =   10
            TabIndex        =   128
            Top             =   200
            Width           =   1185
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   27
            Left            =   1560
            MaxLength       =   30
            TabIndex        =   127
            Text            =   "Text1"
            Top             =   570
            Width           =   660
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   27
            Left            =   2235
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   126
            Text            =   "Text2"
            Top             =   570
            Width           =   3285
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   28
            Left            =   1560
            MaxLength       =   30
            TabIndex        =   125
            Text            =   "Text1"
            Top             =   940
            Width           =   660
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   28
            Left            =   2235
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   124
            Text            =   "Text2"
            Top             =   940
            Width           =   3285
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha Eliminación"
            Height          =   255
            Index           =   37
            Left            =   120
            TabIndex        =   131
            Top             =   200
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Trabajador"
            Height          =   255
            Index           =   38
            Left            =   120
            TabIndex        =   130
            Top             =   570
            Width           =   825
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   7
            Left            =   1080
            Picture         =   "frmFacEntPedSail.frx":03B7
            ToolTipText     =   "Buscar trabajador"
            Top             =   570
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Incidencia"
            Height          =   255
            Index           =   40
            Left            =   120
            TabIndex        =   129
            Top             =   940
            Width           =   825
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   8
            Left            =   1080
            Picture         =   "frmFacEntPedSail.frx":04B9
            ToolTipText     =   "Buscar incidencia"
            Top             =   940
            Width           =   240
         End
      End
      Begin VB.Frame FrameFactura 
         Height          =   3300
         Left            =   -72360
         TabIndex        =   90
         Top             =   1200
         Width           =   10575
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   49
            Left            =   6600
            MaxLength       =   5
            TabIndex        =   138
            Text            =   "Text1 7"
            Top             =   1440
            Width           =   525
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   52
            Left            =   7200
            MaxLength       =   15
            TabIndex        =   137
            Text            =   "Text1 7"
            Top             =   1440
            Width           =   1245
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   50
            Left            =   6600
            MaxLength       =   5
            TabIndex        =   136
            Text            =   "Text1 7"
            Top             =   1755
            Width           =   525
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   53
            Left            =   7200
            MaxLength       =   15
            TabIndex        =   135
            Text            =   "Text1 7"
            Top             =   1755
            Width           =   1245
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   51
            Left            =   6600
            MaxLength       =   5
            TabIndex        =   134
            Text            =   "Text1 7"
            Top             =   2085
            Width           =   525
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   54
            Left            =   7200
            MaxLength       =   15
            TabIndex        =   133
            Text            =   "Text1 7"
            Top             =   2085
            Width           =   1245
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   55
            Left            =   6600
            MaxLength       =   15
            TabIndex        =   107
            Text            =   "Text1 7"
            Top             =   2760
            Width           =   1845
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   48
            Left            =   4920
            MaxLength       =   15
            TabIndex        =   106
            Text            =   "Text1 7"
            Top             =   2085
            Width           =   1485
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   42
            Left            =   4320
            MaxLength       =   5
            TabIndex        =   105
            Text            =   "Text1 7"
            Top             =   2085
            Width           =   525
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   39
            Left            =   1920
            MaxLength       =   4
            TabIndex        =   104
            Text            =   "Text1 7"
            Top             =   2085
            Width           =   525
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   45
            Left            =   2640
            MaxLength       =   15
            TabIndex        =   103
            Text            =   "Text1 7"
            Top             =   2085
            Width           =   1485
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   47
            Left            =   4920
            MaxLength       =   15
            TabIndex        =   102
            Text            =   "Text1 7"
            Top             =   1755
            Width           =   1485
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   41
            Left            =   4320
            MaxLength       =   5
            TabIndex        =   101
            Text            =   "Text1 7"
            Top             =   1755
            Width           =   525
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   38
            Left            =   1920
            MaxLength       =   4
            TabIndex        =   100
            Text            =   "Text1 7"
            Top             =   1755
            Width           =   525
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   44
            Left            =   2640
            MaxLength       =   15
            TabIndex        =   99
            Text            =   "Text1 7"
            Top             =   1755
            Width           =   1485
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   46
            Left            =   4920
            MaxLength       =   15
            TabIndex        =   98
            Text            =   "Text1 7"
            Top             =   1440
            Width           =   1485
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   40
            Left            =   4320
            MaxLength       =   5
            TabIndex        =   97
            Text            =   "Text1 7"
            Top             =   1440
            Width           =   525
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   37
            Left            =   1920
            MaxLength       =   4
            TabIndex        =   96
            Text            =   "Text1 7"
            Top             =   1440
            Width           =   525
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   43
            Left            =   2640
            MaxLength       =   15
            TabIndex        =   95
            Text            =   "Text1 7"
            Top             =   1440
            Width           =   1485
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   36
            Left            =   5760
            MaxLength       =   15
            TabIndex        =   94
            Text            =   "Text1 7"
            Top             =   555
            Width           =   1485
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   35
            Left            =   3960
            MaxLength       =   15
            TabIndex        =   93
            Text            =   "Text1 7"
            Top             =   555
            Width           =   1365
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   34
            Left            =   2160
            MaxLength       =   15
            TabIndex        =   92
            Text            =   "Text1 7"
            Top             =   555
            Width           =   1365
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   33
            Left            =   240
            MaxLength       =   15
            TabIndex        =   91
            Text            =   "Text1 7"
            Top             =   555
            Width           =   1485
         End
         Begin VB.Label Label1 
            Caption         =   "Imp. RE"
            Height          =   255
            Index           =   22
            Left            =   7440
            TabIndex        =   140
            Top             =   1200
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "% RE"
            Height          =   255
            Index           =   48
            Left            =   6600
            TabIndex        =   139
            Top             =   1200
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "Cod. IVA"
            Height          =   255
            Index           =   42
            Left            =   1920
            TabIndex        =   120
            Top             =   1200
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "% IVA"
            Height          =   255
            Index           =   41
            Left            =   4320
            TabIndex        =   119
            Top             =   1230
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "TOTAL PEDIDO"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   255
            Index           =   39
            Left            =   4800
            TabIndex        =   118
            Top             =   2760
            Width           =   1530
         End
         Begin VB.Label Label1 
            Caption         =   "+"
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
            Index           =   36
            Left            =   11880
            TabIndex        =   117
            Top             =   2160
            Width           =   135
         End
         Begin VB.Line Line1 
            X1              =   1800
            X2              =   8520
            Y1              =   1080
            Y2              =   1080
         End
         Begin VB.Label Label1 
            Caption         =   "Imp. IVA"
            Height          =   255
            Index           =   33
            Left            =   4920
            TabIndex        =   116
            Top             =   1230
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "="
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
            Index           =   32
            Left            =   5520
            TabIndex        =   115
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
            Left            =   3720
            TabIndex        =   114
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
            Index           =   30
            Left            =   1920
            TabIndex        =   113
            Top             =   480
            Width           =   135
         End
         Begin VB.Label Label1 
            Caption         =   "Base Imponible"
            Height          =   255
            Index           =   2
            Left            =   5760
            TabIndex        =   112
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Imp. Dto Gn"
            Height          =   255
            Index           =   12
            Left            =   3960
            TabIndex        =   111
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Imp. Dto PP"
            Height          =   255
            Index           =   11
            Left            =   2160
            TabIndex        =   110
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Imp. Bruto"
            Height          =   255
            Index           =   10
            Left            =   240
            TabIndex        =   109
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Base Imponible"
            Height          =   255
            Index           =   9
            Left            =   2760
            TabIndex        =   108
            Top             =   1200
            Width           =   1215
         End
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   25
         Left            =   -72480
         MaxLength       =   10
         TabIndex        =   79
         Tag             =   "Fecha Oferta|F|S|||scaped|fecofert|dd/mm/yyyy|N|"
         Top             =   840
         Width           =   1185
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   24
         Left            =   -73560
         MaxLength       =   7
         TabIndex        =   78
         Tag             =   "Nº Oferta|N|S|||scaped|numofert|0000000|N|"
         Text            =   "Text1 7"
         Top             =   840
         Width           =   885
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   5
         Left            =   8880
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   57
         Tag             =   "Descuento 1"
         Text            =   "OF"
         Top             =   4080
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Frame FrameCliente 
         Height          =   2190
         Left            =   200
         TabIndex        =   62
         Top             =   310
         Width           =   15495
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   32
            Left            =   9240
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   146
            Text            =   "Text2"
            Top             =   480
            Width           =   2415
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   32
            Left            =   7320
            MaxLength       =   30
            TabIndex        =   18
            Tag             =   "Actuacion|T|S|||scaped|actuacion|||"
            Text            =   "Text1"
            Top             =   480
            Width           =   1860
         End
         Begin VB.CheckBox chkEnviadaConfir 
            Caption         =   "Enviado e-mail confir."
            Enabled         =   0   'False
            Height          =   240
            Left            =   12120
            TabIndex        =   24
            Tag             =   "Enviado e-mail confirmación|N|N|||scaped|envconfir||N|"
            Top             =   1320
            Width           =   2055
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   31
            Left            =   10680
            MaxLength       =   40
            TabIndex        =   26
            Tag             =   "E-mail confirmación|T|S|||scaped|mailconfir||N|"
            Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwwwww aqteter"
            Top             =   1800
            Width           =   4650
         End
         Begin VB.CheckBox chkRecogeClien 
            Caption         =   "Recoge cliente"
            Enabled         =   0   'False
            Height          =   240
            Left            =   7200
            TabIndex        =   25
            Tag             =   "Recoge cliente|N|N|||scaped|recogecl||N|"
            Top             =   1800
            Width           =   1455
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   12
            Left            =   7875
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   75
            Tag             =   "Direccion/Dpto.|T|S|||scaped|nomdirec||N|"
            Text            =   "Text2"
            Top             =   165
            Width           =   3765
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   12
            Left            =   7290
            MaxLength       =   30
            TabIndex        =   17
            Tag             =   "Direccion/Dpto.|N|S|0|999|scaped|coddirec|000|N|"
            Text            =   "Text1"
            Top             =   165
            Width           =   540
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   11
            Left            =   1170
            MaxLength       =   30
            TabIndex        =   15
            Tag             =   "Provincia|T|N|||scaped|proclien||N|"
            Text            =   "Text1 Text1 Text1 Text1 Text22"
            Top             =   1290
            Width           =   2565
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   9
            Left            =   1170
            MaxLength       =   6
            TabIndex        =   13
            Tag             =   "CPostal|T|N|||scaped|codpobla||N|"
            Text            =   "Text15"
            Top             =   867
            Width           =   630
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   10
            Left            =   1820
            MaxLength       =   30
            TabIndex        =   14
            Tag             =   "Población|T|N|||scaped|pobclien||N|"
            Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwaq"
            Top             =   867
            Width           =   3525
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   7
            Left            =   3375
            MaxLength       =   20
            TabIndex        =   11
            Tag             =   "teléfono Cliente|T|S|||scaped|telclien||N|"
            Text            =   "12345678911234567899"
            Top             =   165
            Width           =   1965
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   6
            Left            =   1200
            MaxLength       =   15
            TabIndex        =   10
            Tag             =   "NIF Cliente|T|N|||scaped|nifclien||N|"
            Text            =   "123456789"
            Top             =   155
            Width           =   1110
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   13
            Left            =   1170
            MaxLength       =   255
            TabIndex        =   16
            Tag             =   "Referencia Cliente|T|S|||scaped|referenc||N|"
            Text            =   "Text1 Text1 Text1 Te"
            Top             =   1800
            Width           =   5445
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   17
            Left            =   7290
            MaxLength       =   30
            TabIndex        =   19
            Tag             =   "Cod. Agente|N|N|0|9999|scaped|codagent|0000|N|"
            Text            =   "Text1"
            Top             =   870
            Width           =   540
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   17
            Left            =   7875
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   69
            Text            =   "Text2"
            Top             =   870
            Width           =   3765
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   14
            Left            =   7290
            MaxLength       =   30
            TabIndex        =   20
            Tag             =   "Forma de Pago|N|N|0|999|scaped|codforpa|000|N|"
            Text            =   "Text1"
            Top             =   1230
            Width           =   540
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   14
            Left            =   7875
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   64
            Text            =   "Text2"
            Top             =   1230
            Width           =   3750
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   15
            Left            =   13560
            MaxLength       =   7
            TabIndex        =   21
            Tag             =   "Descuento P.Pago|N|N|0|99.90|scaped|dtoppago|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   120
            Width           =   540
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   16
            Left            =   13560
            MaxLength       =   7
            TabIndex        =   22
            Tag             =   "Descuento General|N|N|0|99.90|scaped|dtognral|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   480
            Width           =   540
         End
         Begin VB.ComboBox cboFacturacion 
            Height          =   315
            Left            =   13560
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Tag             =   "Tipo Facturación|N|N|||scaped|tipofact||N|"
            Top             =   840
            Width           =   1820
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   8
            Left            =   1170
            MaxLength       =   60
            TabIndex        =   12
            Tag             =   "Domicilio|T|N|||scaped|domclien||N|"
            Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwwwww aq"
            Top             =   516
            Width           =   4170
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   9
            Left            =   7005
            Picture         =   "frmFacEntPedSail.frx":05BB
            ToolTipText     =   "Buscar forma de pago"
            Top             =   480
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   4
            Left            =   7005
            Picture         =   "frmFacEntPedSail.frx":06BD
            ToolTipText     =   "Buscar forma de pago"
            Top             =   1260
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Actuación"
            Height          =   255
            Index           =   24
            Left            =   6000
            TabIndex        =   147
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "E-mail confirmación"
            Height          =   255
            Index           =   23
            Left            =   9120
            TabIndex        =   141
            Top             =   1800
            Width           =   1455
         End
         Begin VB.Image imgBuscar 
            Enabled         =   0   'False
            Height          =   240
            Index           =   6
            Left            =   900
            Picture         =   "frmFacEntPedSail.frx":07BF
            ToolTipText     =   "Buscar población"
            Top             =   880
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Direc."
            Height          =   255
            Index           =   1
            Left            =   6000
            TabIndex        =   77
            Top             =   165
            Width           =   495
         End
         Begin VB.Image imgBuscar 
            Enabled         =   0   'False
            Height          =   240
            Index           =   2
            Left            =   7005
            Picture         =   "frmFacEntPedSail.frx":08C1
            ToolTipText     =   "Buscar direc./dpto"
            Top             =   180
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Provincia"
            Height          =   255
            Index           =   17
            Left            =   120
            TabIndex        =   76
            Top             =   1290
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Población"
            Height          =   255
            Index           =   16
            Left            =   120
            TabIndex        =   74
            Top             =   867
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Teléfono"
            Height          =   255
            Index           =   19
            Left            =   2565
            TabIndex        =   73
            Top             =   165
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "NIF"
            Height          =   255
            Index           =   20
            Left            =   120
            TabIndex        =   72
            Top             =   165
            Width           =   615
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   1
            Left            =   900
            Picture         =   "frmFacEntPedSail.frx":09C3
            ToolTipText     =   "Buscar cliente varios"
            Top             =   180
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Ref. Cliente"
            Height          =   255
            Index           =   13
            Left            =   60
            TabIndex        =   71
            Top             =   1800
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Agente"
            Height          =   255
            Index           =   34
            Left            =   6000
            TabIndex        =   70
            Top             =   870
            Width           =   615
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   5
            Left            =   7005
            Picture         =   "frmFacEntPedSail.frx":0AC5
            ToolTipText     =   "Buscar agente"
            Top             =   885
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Forma Pago"
            Height          =   255
            Index           =   15
            Left            =   6000
            TabIndex        =   68
            Top             =   1230
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Dto. P. Pago"
            Height          =   255
            Index           =   25
            Left            =   12120
            TabIndex        =   67
            Top             =   165
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Dto. Gral"
            Height          =   255
            Index           =   26
            Left            =   12120
            TabIndex        =   66
            Top             =   510
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo Facturac."
            Height          =   255
            Index           =   4
            Left            =   12120
            TabIndex        =   65
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Domicilio"
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   63
            Top             =   516
            Width           =   735
         End
      End
      Begin VB.CommandButton cmdAux 
         Appearance      =   0  'Flat
         Caption         =   "+"
         Height          =   315
         Index           =   1
         Left            =   2640
         TabIndex        =   61
         ToolTipText     =   "Buscar artículo"
         Top             =   4080
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.CommandButton cmdAux 
         Appearance      =   0  'Flat
         Caption         =   "+"
         Height          =   315
         Index           =   0
         Left            =   960
         TabIndex        =   60
         ToolTipText     =   "Buscar almacen"
         Top             =   4080
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   2
         Left            =   2880
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   45
         Tag             =   "Nombre Artículo"
         Text            =   "nomArtic"
         Top             =   4080
         Visible         =   0   'False
         Width           =   3285
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   8
         Left            =   9360
         MaxLength       =   12
         TabIndex        =   58
         Tag             =   "Importe"
         Text            =   "Importe"
         Top             =   4560
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   7
         Left            =   10080
         MaxLength       =   30
         TabIndex        =   50
         Tag             =   "Descuento 2"
         Text            =   "Dto2"
         Top             =   4080
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   6
         Left            =   9360
         MaxLength       =   5
         TabIndex        =   49
         Tag             =   "Descuento 1"
         Text            =   "Dto1"
         Top             =   4080
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   4
         Left            =   7920
         MaxLength       =   12
         TabIndex        =   48
         Tag             =   "Precio"
         Text            =   "123,456.7879"
         Top             =   3960
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   3
         Left            =   6120
         MaxLength       =   16
         TabIndex        =   46
         Tag             =   "Cantidad"
         Text            =   "1,234,567,891.25"
         Top             =   4080
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   1
         Left            =   1200
         MaxLength       =   18
         TabIndex        =   44
         Tag             =   "Código Artículo"
         Text            =   "Artic Artic Artic5"
         Top             =   4020
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   0
         Left            =   360
         MaxLength       =   15
         TabIndex        =   43
         Tag             =   "Código Almacen"
         Text            =   "codalmac"
         Top             =   4020
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   280
         Index           =   23
         Left            =   -70080
         MaxLength       =   80
         TabIndex        =   31
         Tag             =   "Observación 5|T|S|||scaped|observa05||N|"
         Top             =   3240
         Width           =   8805
      End
      Begin VB.TextBox Text1 
         Height          =   280
         Index           =   22
         Left            =   -70080
         MaxLength       =   80
         TabIndex        =   30
         Tag             =   "Observación 4|T|S|||scaped|observa04||N|"
         Top             =   2955
         Width           =   8805
      End
      Begin VB.TextBox Text1 
         Height          =   280
         Index           =   21
         Left            =   -70080
         MaxLength       =   80
         TabIndex        =   29
         Tag             =   "Observación 3|T|S|||scaped|observa03||N|"
         Top             =   2670
         Width           =   8805
      End
      Begin VB.TextBox Text1 
         Height          =   280
         Index           =   20
         Left            =   -70080
         MaxLength       =   80
         TabIndex        =   28
         Tag             =   "Observación 2|T|S|||scaped|observa02||N|"
         Top             =   2385
         Width           =   8805
      End
      Begin VB.TextBox Text1 
         Height          =   280
         Index           =   19
         Left            =   -70080
         MaxLength       =   80
         TabIndex        =   27
         Tag             =   "Observación 1|T|S|||scaped|observa01||N|"
         Top             =   2100
         Width           =   8805
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmFacEntPedSail.frx":0BC7
         Height          =   3000
         Left            =   195
         TabIndex        =   59
         Top             =   2760
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   5292
         _Version        =   393216
         AllowUpdate     =   -1  'True
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
      Begin VB.Label Label1 
         Caption         =   "Observaciones del Pedido"
         Height          =   255
         Index           =   29
         Left            =   -72720
         TabIndex        =   152
         Top             =   4800
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "T.O."
         Height          =   255
         Index           =   28
         Left            =   11400
         TabIndex        =   151
         Top             =   3840
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Capitulo"
         Height          =   255
         Index           =   27
         Left            =   11400
         TabIndex        =   150
         Top             =   4440
         Width           =   735
      End
      Begin VB.Image imgBuscar2 
         Height          =   240
         Index           =   11
         Left            =   12360
         Picture         =   "frmFacEntPedSail.frx":0BDC
         ToolTipText     =   "Buscar población"
         Top             =   5160
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgBuscar2 
         Height          =   240
         Index           =   12
         Left            =   12000
         Picture         =   "frmFacEntPedSail.frx":0CDE
         ToolTipText     =   "Buscar población"
         Top             =   3840
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgBuscar2 
         Height          =   240
         Index           =   13
         Left            =   12120
         Picture         =   "frmFacEntPedSail.frx":0DE0
         ToolTipText     =   "Buscar población"
         Top             =   4440
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Ampliación Línea"
         Height          =   255
         Index           =   35
         Left            =   11400
         TabIndex        =   144
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Centro coste"
         Height          =   255
         Index           =   6
         Left            =   11400
         TabIndex        =   143
         Top             =   5160
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Observaciones del Pedido"
         Height          =   255
         Index           =   18
         Left            =   -72720
         TabIndex        =   132
         Top             =   3840
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Oferta"
         Height          =   255
         Index           =   5
         Left            =   -72480
         TabIndex        =   81
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Oferta"
         Height          =   255
         Index           =   3
         Left            =   -73560
         TabIndex        =   80
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Observaciones"
         Height          =   255
         Index           =   45
         Left            =   -72720
         TabIndex        =   42
         Top             =   2160
         Width           =   2415
      End
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   14880
      TabIndex        =   36
      Top             =   7320
      Visible         =   0   'False
      Width           =   1035
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
      Left            =   8880
      TabIndex        =   142
      Top             =   7440
      Width           =   3615
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
      Begin VB.Menu mnBarra6 
         Caption         =   "-"
      End
      Begin VB.Menu mnLineas 
         Caption         =   "&Lineas"
         HelpContextID   =   2
         Shortcut        =   ^L
      End
      Begin VB.Menu mnGenAlbaran 
         Caption         =   "&Generar Albaran"
         HelpContextID   =   2
         Shortcut        =   ^G
      End
      Begin VB.Menu mnGeneraFactura 
         Caption         =   "Generar factura"
         Shortcut        =   ^Q
      End
      Begin VB.Menu mnBarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnImprimir 
         Caption         =   "&Imprimir"
         Begin VB.Menu mnImpPedido 
            Caption         =   "&Pedido"
            Shortcut        =   ^P
         End
         Begin VB.Menu mnImpOrde 
            Caption         =   "&Orden Instalación"
            Shortcut        =   ^O
         End
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
Attribute VB_Name = "frmFacEntPedSail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ---- MODIFICACIONES ------
' ---- [04/11/2009] [LAURA] : Añadir botón para enviar informe confirmacion entrega del Pedido
' --------------------------

Public DatosADevolverBusqueda2 As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado2(CadenaSeleccion As String)

Public EsHistorico As Boolean 'Si es true abrir el formulario con la tabla de
                              'de historico schped, y solo en modo de consulta


Private WithEvents frmB As frmBuscaGrid  'Form para busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmF As frmCal 'Form Calendario Fecha
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmCP As frmCPostal 'Codigos Postales
Attribute frmCP.VB_VarHelpID = -1

Private WithEvents frmC As frmFacClientes3 'Form Mto Clientes
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmCV As frmFacClientesV  'Form Mto Clientes Varios
Attribute frmCV.VB_VarHelpID = -1
Private WithEvents frmFP As frmFacFormasPago 'Form Mto Formas de Pago
Attribute frmFP.VB_VarHelpID = -1
Private WithEvents frmT As frmAdmTrabajadores  'Form Mto Trabajadores
Attribute frmT.VB_VarHelpID = -1
Private WithEvents frmA As frmFacAgentesCom   'Form Mto Agentes
Attribute frmA.VB_VarHelpID = -1
Private WithEvents frmAlm As frmAlmAlPropios   'Form Almacenes Propios
Attribute frmAlm.VB_VarHelpID = -1
Private WithEvents FrmArt As frmAlmArticulos   'Form Articulos
Attribute FrmArt.VB_VarHelpID = -1

Private WithEvents frmList As frmListadoPed 'Listados para Pedidos (pasar pedido a albaran)
Attribute frmList.VB_VarHelpID = -1
Private WithEvents frmList2 As frmListadoOfer  'Listados para pedir datos para grabar en historico
Attribute frmList2.VB_VarHelpID = -1
Private WithEvents frmNSerie As frmRepCargarNSerie  'Form Cargar nº Series
Attribute frmNSerie.VB_VarHelpID = -1
Private WithEvents frmMen As frmMensajes  'Form Mensajes para mostrar los Nº serie y elegir
Attribute frmMen.VB_VarHelpID = -1

Private WithEvents frmOT As frmObraOT
Attribute frmOT.VB_VarHelpID = -1
Private WithEvents frmOC As frmObraCapitulo
Attribute frmOC.VB_VarHelpID = -1
Private WithEvents frmAc As frmObraActua
Attribute frmAc.VB_VarHelpID = -1
Private WithEvents FrmArtEul As frmAlmArticuEUL
Attribute FrmArtEul.VB_VarHelpID = -1

Private Modo As Byte
'-----------------------------
'Se distinguen varios modos
'   0.-  Formulario limpio sin nungun campo rellenado
'   1.-  Preparando para hacer la busqueda
'   2.-  Ya tenemos registros y los vamos a recorrer
'        y podemos editarlos Edicion del campo
'   3.-  Insercion de nuevo registro
'   4.-  Modificar
'   5.-  Mantenimiento Lineas
'   6.- Cargar cantidad servidas al Generar Albaran no completo (Pedido --> Albaran)
'-------------------------------------------------------------------------


Private ModificaLineas As Byte
'1.- Añadir,  2.- Modificar,  3.- Borrar,  0.-Pasar control a Lineas

Dim TituloLinea As String 'Descripcion de la linea que estamos en Mantenimiento

Dim primeravez As Boolean

Dim EsCabecera As Boolean
'Para saber en MandaBusquedaPrevia si busca en la tabla scapla o en la tabla sdirec

Dim CodTipoMov As String
'Codigo tipo de movimiento en función del valor en la tabla de parámetros: stipom

Dim EsDeVarios As Boolean
'Si el cliente mostrado es de Varios o No

'SQL de la tabla principal del formulario
Private CadenaConsulta As String
Private CadenaSQL As String 'Para crear consulta de Generar Albaran a partir del Pedido

Private Ordenacion As String
Private NombreTabla As String  'Nombre de la tabla de Cabecera
Private NomTablaLineas As String 'Nombre de la Tabla de lineas
Private kCampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean

Dim btnAnyadir As Byte
'Variable que indica el número del Boton  Anyadir en la Toolbar1
Dim btnPrimero As Byte
'Variable que indica el número del Boton  PrimerRegistro en la Toolbar1


Private VieneDeBuscar As Boolean
'Para cuando devuelve dos poblaciones con el mismo codigo Postal. Si viene de pulsar prismatico
'de busqueda poner el valor de poblacion seleccionado y no volver a recuperar de la Base de Datos

Private HaCambiadoCP As Boolean
'Para saber si tras haber vuelto de prismaticos ha cambiado el valor del CPostal

Dim PorCaja As Boolean
'Para Saber si se ha salido con precio caja y hay que calcular el importe de la
'linea aplicando el precio de la caja. Si PorCaja=false se aplicaca el precio de unidad

Dim Precio As String 'Precio de la linea de Articulo

Dim ImprimeAlb As Boolean 'Para saber cuando vuelve de Generar ALbaran si se ha solicitado Imprimir Albaran o no
Dim ImprimeEtiq As Boolean
Dim ImprimeHojaExp As Boolean 'Imprime hoja de Expedicion
Dim FechaAlb As String 'Para cuando vuelve de pedir datos para Generar Albaran, saber la fecha que se introdujo

'Cambiamos este valor para indicar si es
'no solo a mostrador, sino tb ALE-ALO-ALR-ALV
Dim EsAMostrador2 As String

Dim gridCargado As Boolean 'Saber si el grid esta cargado cuando se ejecuta DataGrid1_RowColChange

Dim AlbCompleto As Boolean 'Si se va a servir el Pedido Completo (slialb.cantidad=sliped.cantidad)
                            'o se va a servir una parte (slialb.cantidad=sliped.servidas)

Dim CtaBancoPropi As String 'Cuando facturamos el pedido directamente, para saber la caja

Dim txtAnterior As String   'Para que no realice las acciones en el lost_focus si NO ha cambiado nada


Dim ClienteConTasaReciclado As Boolean  'Cuando pasamos a las lineas pondremos esta variab

' ---- [15/09/2009] (LAURA)
'Dim ElArticulo As String   'para la sdesca
' ----

' Tipo fontenas
Dim KilosAnteriores As Currency
Dim RutaCliente As Integer
Dim ZonaCliente As Integer

Dim LineaIntercalar As Integer 'NO reutilizar

Dim PulsadoMas2 As Boolean

Dim CadenaArticulosEULER As String

'================================================================================

Private Sub cboFacturacion_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
    
'    If KeyAscii = 13 Then 'ENTER
'        Me.SSTab1.Tab = 1
'        PonerFoco Text1(19)
'    End If
End Sub


Private Sub chkServirCom_KeyPress(KeyAscii As Integer)
     KEYpress KeyAscii
End Sub

Private Sub chkVisadoRes_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()
Dim SQL As String
Dim PrimeraLin As Boolean 'Si se inserta la primera linea no esta creado el datagrid1 entonces llamar
                          ' a DataGrid, sino llamar solo a DataGrid2

    Screen.MousePointer = vbHourglass
    On Error GoTo Error1

    Select Case Modo
        Case 1  'BUSQUEDA
            HacerBusqueda
        Case 3 'INSERTAR Cabecera Pedido
            If DatosOk Then InsertarCabecera
        Case 4  'MODIFICAR Cabecera Pedido
            If DatosOk Then
                ActualizarClienteVarios Text1(4).Text, Text1(6).Text
                If ModificaDesdeFormulario(Me, 1) Then
                    TerminaBloquear
                    'Por si acaso ha cambioado coiddirec
                    UpdateaNomDirec
                    
                    PosicionarData
                    
                End If
            End If
            
         Case 5 'InsertarModificar LINEA
            'Actualizar el registro en la tabla de lineas 'sliped'
            If ModificaLineas = 1 Then 'INSERTAR lineas Pedidos
                PrimeraLin = False
                If Data2.Recordset.EOF = True Then PrimeraLin = True
                If InsertarLinea Then
                    If PrimeraLin Then
                        CargaGrid DataGrid1, Data2, True
                    Else
                        CargaGrid2 DataGrid1, Data2
                    End If
                    If LineaIntercalar > 0 Then
                        'HA intercalado la linea. Ponemos luego en normal
                        Me.DataGrid1.Enabled = True
                        DataGrid1.AllowAddNew = False
                        NumRegElim = LineaIntercalar
                        CargaTxtAux2 False, False
                        CargaGrid2 DataGrid1, Data2
                        PosicionarData2
                        ModificaLineas = 0
                        PonerBotonCabecera True
                        
                    
                    Else
                        BotonAnyadirLinea False
                    End If
                End If
            ElseIf ModificaLineas = 2 Then 'MODIFICAR lineas
                If ModificarLinea Then
                    TerminaBloquear
                    NumRegElim = Data2.Recordset!numlinea
                    CargaTxtAux2 False, False
                    CargaGrid2 DataGrid1, Data2
                    PosicionarData2
                    ModificaLineas = 0
                    PonerBotonCabecera True
                    
                End If
                Me.DataGrid1.Enabled = True
            End If
            CalcularDatosFactura
            
        Case 6 'PASAR Pedido a ALBARAN
            'Comprobar que la cantidad a servir es mayor que cero
             SQL = "SELECT SUM(servidas) as servidas from sliped WHERE "
             SQL = SQL & Replace(ObtenerWhereCP, NombreTabla, NomTablaLineas)
                          
             If RegistrosAListar(SQL) = 0 Then 'No hay cantidad en linea para el Albaran
                SQL = "La cantidad total a servir en el Albaran es cero." & vbCrLf
                SQL = SQL & vbCrLf & "Introduzca la cantidad a servir."
                MsgBox SQL, vbExclamation
                PonerFoco txtaux(3)
             Else
                If SePuedeServirPedido Then GenerarAlbaran False
             End If
    End Select
    Screen.MousePointer = vbDefault
    
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub cmdAux_Click(index As Integer)
    Select Case index
        Case 0 'Busqueda de Cod. Almacen
            Set frmAlm = New frmAlmAlPropios
            frmAlm.DatosADevolverBusqueda = "0"
            frmAlm.Show vbModal
            Set frmAlm = Nothing
            PonerFoco txtaux(index)
            
        Case 1 'Busqueda de Cod. Artic
            
            If InstalacionEsEulerTaxco Then
                'EULER  As
                Set FrmArtEul = New frmAlmArticuEUL
                'frmArt.DatosADevolverBusqueda3 = "@1@" 'Poner en modo busqueda
                FrmArtEul.FechaDoc = CDate(Text1(1).Text)
                FrmArtEul.Codprove = -1
                FrmArtEul.DesdeVentas = True
                FrmArtEul.Show vbModal
                Set FrmArtEul = Nothing
            
            Else
        
        
                Set FrmArt = New frmAlmArticulos
                FrmArt.DatosADevolverBusqueda = "@1@" 'Poner en Modo Busqueda
                FrmArt.DeConsulta = True
                FrmArt.Show vbModal
                Set FrmArt = Nothing
            End If
            
            
            
            PonerFoco txtaux(index)
            
        Case 2 'COD. CENTRO DE COSTE
            If vEmpresa.TieneAnalitica Then
                EsCabecera = False
                'centro de coste
                AbrirForm_CentroCoste
                PonerFoco txtaux(11)
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
            CargaTxtAux2 False, False
            
            If ModificaLineas = 1 Then 'INSERTAR
                ModificaLineas = 0
                DataGrid1.AllowAddNew = False
                If Not Data2.Recordset.EOF Then Data2.Recordset.MoveFirst
            End If
            PonerForaGrid
            ModificaLineas = 0
            LineaIntercalar = 0
            PonerBotonCabecera True
            Me.DataGrid1.Enabled = True
            
        Case 6 'Insertar servidas en Generar Albaran (Pedido --> Albaran)
            InicializarServidas
            PonerModo 2
            CargaTxtAuxServidas False, False
            CargaGrid DataGrid1, Data2, True, False
            PonerForaGrid
    End Select
End Sub


Private Sub BotonAnyadir()
'Añadir registro en tabla de cabecera de Pedidos: scaped (Cabecera)
Dim NomTraba As String

    LimpiarCampos 'Vacía los TextBox
    'Poner los grid sin apuntar a nada
    LimpiarDataGrids
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 3


    If InstalacionEsEulerTaxco Then
        chkVisadoRes.Value = 1
        chkServirCom.Value = 1
    End If

    'Poner el nombre del trabajador que esta conectado
    Text1(3).Text = PonerTrabajadorConectado(NomTraba)
    Text2(3).Text = NomTraba

    Text1(1).Text = Format(Now, "dd/mm/yyyy") 'Fecha Oferta
    PonerFoco Text1(1)
End Sub


Private Sub BotonAnyadirLinea(Intercalando As Boolean)
    'Si no estaba modificando lineas salimos
    ' Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 2 Then Exit Sub
       
    ModificaLineas = 1 'Ponemos Modo Añadir Linea
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerBotonCabecera False
    SSTab1.Tab = 0
    
    If Intercalando Then
        lblIndicador.Caption = "** INTERCALAR **"
        If Not Data2.Recordset.EOF Then
            LineaIntercalar = Data2.Recordset!numlinea
        End If
    Else
        lblIndicador.Caption = "INSERTAR"
    End If
    
    AnyadirLinea DataGrid1, Data2
    CargaTxtAux2 True, True
    
    'Poner el Almacen por defecto del Trabajador
    txtaux(0).Text = DevuelveDesdeBDNew(conAri, "straba", "codalmac", "codtraba", Text1(3).Text, "N")
    If txtaux(0).Text <> "" Then txtaux(0).Text = Format(txtaux(0).Text, "000")
    
    
    
    ' ---- [20/10/2009] [LAURA]: añadir campo centro de coste
    'si contab. analitica por trabajador traer su centro de coste
    If vEmpresa.TieneAnalitica And vParamAplic.ModoAnalitica = 0 Then
        txtaux(11).Text = DevuelveDesdeBDNew(conAri, "straba", "codccost", "codtraba", Text1(3).Text, "N")
        Me.txtAux2(11).Text = PonerNombreCCoste(Me.txtaux(11))
    Else
        Me.txtAux2(11).Text = ""
    End If
    If Intercalando Then
        txtaux(0).BackColor = vbRed
    Else
        txtaux(0).BackColor = vbWhite
    End If
    
    
    PonerFoco txtaux(1)
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
'    LimpiarCampos
    If chkVistaPrevia.Value = 1 Then
        EsCabecera = True
        MandaBusquedaPrevia ""
    Else
        LimpiarCampos
        LimpiarDataGrids
        CadenaConsulta = "Select * from " & NombreTabla & " " & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Sub Desplazamiento(index As Integer)
'Botones de Desplazamiento de la Toolbar
'Para desplazarse por los registros de control Data
    If Data1.Recordset.EOF Then Exit Sub
    DesplazamientoData Data1, index
    PonerCampos
End Sub


Private Sub BotonModificar()
'Prepara el Form para Modificar la cabecera de Pedidos (tabla: scaped)
Dim DeVarios As Boolean

    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 4
    PonerFoco Text1(1)
        
    'Si es Cliente de Varios no se pueden modificar sus datos
    DeVarios = EsClienteVarios(Text1(4).Text)
    BloquearDatosCliente (DeVarios)
End Sub


Private Sub BotonModificarLinea()
'Prepara el Form para Modificar una linea de Pedido (tabla: sliped)
Dim vWhere As String

    On Error GoTo EModificarLinea

    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Then Exit Sub '1= Insertar
    If Data2.Recordset.EOF Then Exit Sub
    
    vWhere = ObtenerWhereCP & " and numlinea=" & Data2.Recordset!numlinea
    vWhere = Replace(vWhere, NombreTabla, NomTablaLineas)
    If Not BloqueaRegistro(NomTablaLineas, vWhere) Then Exit Sub
    
    CargaTxtAux2 True, False
    ModificaLineas = 2 'Modificar
    'Añadiremos el boton de aceptar y demas objetos para insertar
    Me.lblIndicador.Caption = "MODIFICAR"
    PonerBotonCabecera False
    BloquearTxt txtaux(2), True 'campo nombre articulo
    PonerFoco txtaux(0)
    Me.DataGrid1.Enabled = False
    Exit Sub
    
EModificarLinea:
    MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub BotonEliminar()
'Eliminar Registro de la Cabecera: Tabla de Pedidos (scaped)
' y los registros correspondientes de las tablas de lineas (sliped)
Dim cad As String

    On Error GoTo EEliminar

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub

    cad = "Cabecera de Pedidos." & vbCrLf
    cad = cad & "----------------------------------" & vbCrLf & vbCrLf
    cad = cad & "Va a eliminar el Pedido:            "
    cad = cad & vbCrLf & "Nº:  " & Format(Text1(0).Text, "0000000")
    cad = cad & vbCrLf & "Cliente:  " & Format(Text1(4).Text, "000000") & " - " & Text1(5).Text
    cad = cad & vbCrLf & vbCrLf & "¿Desea Eliminarlo? "
    
    Screen.MousePointer = vbHourglass
    
    'Borramos
    If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        NumRegElim = Data1.Recordset.AbsolutePosition
        
        'Abrir frame de informes para pedir datos antes de grabar en el historico
        CadenaSQL = ""
        Set frmList2 = New frmListadoOfer
        frmList2.OpcionListado = 81
        frmList2.Show vbModal
        Set frmList2 = Nothing
        If CadenaSQL = "" Then Exit Sub
        
        If Not Eliminar() Then Exit Sub
        PosicionarDataTrasEliminar
    End If
    Screen.MousePointer = vbDefault
    
EEliminar:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Pedido", Err.Description
End Sub


Private Sub BotonEliminarLinea()
'Eliminar una linea Del Pedido. (Tabla: sliped)
Dim SQL As String

    On Error GoTo EEliminarLinea

    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Or ModificaLineas = 2 Then Exit Sub '1= Insertar, 2=Modificar

    If Data2.Recordset.EOF Then Exit Sub
            
    ModificaLineas = 3 'Eliminar
    SQL = "¿Seguro que desea eliminar la línea del Pedido?     "
    SQL = SQL & vbCrLf & "NumLinea:  " & Data2.Recordset!numlinea & vbCrLf
    SQL = SQL & "Almacen:  " & Format(Data2.Recordset!codAlmac, "000")
    SQL = SQL & vbCrLf & "Artículo:  " & Data2.Recordset!codArtic & " - " & Data2.Recordset!NomArtic
    
    ' ---- [15/09/2009] (LAURA)
'    ElArticulo = Data2.Recordset!codArtic
    ' ----
    
    If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        NumRegElim = Data2.Recordset.AbsolutePosition
        SQL = "Delete from " & NomTablaLineas & " WHERE " & Replace(ObtenerWhereCP, NombreTabla, NomTablaLineas)
        SQL = SQL & " and numlinea=" & Data2.Recordset!numlinea
        conn.Execute SQL
        
        
        ' ---- [15/09/2009] (LAURA)
'        DescuentosCantidad ElArticulo
        ' ----
        
        ModificaLineas = 0
        CargaGrid2 DataGrid1, Data2
'        SituarDataTrasEliminar Data2, NumRegElim
        SituarDataPosicion Me.Data2, NumRegElim, SQL
        CalcularDatosFactura
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
Dim Port As Integer      'Port: para saber si ha metido/Modificado el articulo de portes


    'Quitar lineas y volver a la cabecera
    If Modo = 5 Then  'modo 5: Mantenimientos Lineas
    
        If vParamAplic.TipoPortes = 1 Then
            'Si lleva portes haremos varias cosas
            Port = HacerAccionesPortes
            CargaGrid DataGrid1, Data2, True
            Set miRsAux = Nothing
        End If
    
        ' ---- [15/09/2009] (LAURA)
        DescuentosCantidad ""
        ' ----
    
        PonerModo 2
        'BloquearTabs False
        Me.lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
        If DataGrid1.Row >= 0 Then
            If Port = 0 Then
                DeseleccionaGrid DataGrid1
                'DataGrid1.Bookmark = 1
            Else
                Data2.Recordset.MoveLast  'El ultimo es el porte
            End If
        End If
        cmdCancelar.Cancel = True
    Else 'Se llama desde algún Prismatico de otro Form al Mantenimiento de Trabajadores
        If Data1.Recordset.EOF Then
            MsgBox "Ningún registro devuelto.", vbExclamation
            Exit Sub
        End If
        'cad = Data1.Recordset.Fields(0) & "|"
        'cad = cad & Data1.Recordset.Fields(1) & "|"
        cad = Data1.Recordset.Fields(0)
        RaiseEvent DatoSeleccionado2(cad)
        Unload Me
    End If
End Sub


Private Sub DataGrid1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Ayuda de Etiqueta de precio de salida de la Función de Precios
    If Data2.Recordset.EOF Then Exit Sub
    If (Modo = 2) Or (Modo = 5 And ModificaLineas = 0) Then
        If X > 7750 And X < 8000 Then
            Select Case DataGrid1.Columns(8).Value
                Case "P": Me.DataGrid1.ToolTipText = "P: Promoción"
                Case "E": Me.DataGrid1.ToolTipText = "E: Precio Especial"
                Case "T": Me.DataGrid1.ToolTipText = "T: Tarifa Artículo"
                Case "A": Me.DataGrid1.ToolTipText = "A: Precio Artículo"
                Case "M": Me.DataGrid1.ToolTipText = "M: Manual"
                Case Else
                    Me.DataGrid1.ToolTipText = ""
            End Select
        Else
            Me.DataGrid1.ToolTipText = ""
        End If
    End If
End Sub


Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'Dim devuelve As String

    On Error GoTo Error1

    If Modo = 6 And gridCargado Then '6: Pasar Pedido a Albaran no Completo (Introducir las servidas)
        CargaTxtAuxServidas True, True
        txtaux(3).Text = Data2.Recordset!servidas
        ' ---- [28/09/2009] (LAURA) : añadimos esta linea para el formato
        PonerFormatoDecimal txtaux(3), 1
        ' ----
        txtaux(9).Text = Data2.Recordset!bultosser
    End If
    
    'If Modo = 5 Then 'Poner el valor al camp ampliacion linea '5: modo lineas
        
        If Not Data2.Recordset.EOF And ModificaLineas <> 1 Then '1: Insertar
        
            PonerForaGrid
        
'            devuelve = DevuelveDesdeBDNew(conAri, NomTablaLineas, "ampliaci", "numpedcl", Text1(0).Text, "N", , "numlinea", Data2.Recordset!numlinea, "N")
'            'Poner descripcion de ampliacion lineas
'            Text2(16).Text = devuelve
'
            '- centro de coste
'            ' ---- [20/10/2009] [LAURA]: añadir campo centro de coste familia

                If vEmpresa.TieneAnalitica Then
                    Me.txtaux(11).Text = DBLet(Data2.Recordset!CodCCost, "T")
                    Me.txtAux2(11).Text = PonerNombreCCoste(Me.txtaux(11))
'                Else
'                    txtAux2(11).Text = ""
'                End If
                End If
            
        Else
            Text2(16).Text = ""
            txtAux2(11).Text = ""
        End If
    'End If
    Exit Sub
    
Error1:
    MuestraError Err.Number, "", Err.Description
End Sub


Private Sub Form_activate()
    If Me.Tag <> "" Then
        Me.Tag = ""
        PonerCampos
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
    'Icono del formulario
    Me.Icon = frmPpal.Icon

    ' ICONITOS DE LA BARRA
    btnAnyadir = 5
    btnPrimero = 21
    With Me.Toolbar1
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 1   'Botón Buscar
        .Buttons(2).Image = 2   'Botón Todos
        .Buttons(5).Image = 3   'Insertar Nuevo
        .Buttons(6).Image = 4   'Modificar
        .Buttons(7).Image = 5   'Borrar
        .Buttons(10).Image = 10 'Mto Lineas Ofertas
        .Buttons(11).Image = 26 'Generar Albaran
        
        'Enero08
        .Buttons(12).Image = 42 'Generar factura desde el pedido
        
        .Buttons(14).Image = 16 'Imprimir Pedido
        .Buttons(15).Image = 27 'Imprimir Orden Instalacion
        .Buttons(16).Image = 40 'confirmación de entrega
        
        .Buttons(18).Image = 15  'Salir
        .Buttons(btnPrimero).Image = 6  'Primero
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
        .Buttons(btnPrimero + 3).Image = 9 'Último
    End With
    
    Me.SSTab1.Tab = 0
    LimpiarCampos   'Limpia los campos TextBox
'    cmdAux(0).Tag = "-1"
    
    CargarComboFacturacion
    CodTipoMov = "PEV"
    VieneDeBuscar = False
   
    'Comprobar si es Departamento o Direccion

    Me.Label1(1).Caption = DevuelveTextoDepto(True)
    
    'Lbl obs crm
    If vParamAplic.TieneCRM Then
        Label1(29).Caption = "Observaciones CRM"
    Else
        Label1(29).Caption = "Observaciones internas"
    End If
    '## A mano
    Me.FrameHco.visible = EsHistorico
    chkEnviadaConfir.visible = Not EsHistorico
    Label1(23).visible = Not EsHistorico
    Text1(31).visible = Not EsHistorico
    
    
    If Not EsHistorico Then
        NombreTabla = "scaped"
        NomTablaLineas = "sliped" 'Tabla lineas de Pedido
        Me.Caption = "Pedidos Clientes"
        Ordenacion = " ORDER BY numpedcl "
    Else
        NombreTabla = "schped"
        NomTablaLineas = "slhped"
        CargarTagsHco Me, "scaped", NombreTabla
        'Estos campos solo estan en la tabla del histórico
        Text1(26).Tag = "Fecha Eliminación|F|N|||schped|fechelim|dd/mm/yyyy|N|"
        Text1(27).Tag = "Trabajador Eliminación|N|N|0|9999|schped|trabelim|0000|N|"
        Text1(28).Tag = "Incidencia elim.|T|N|||schped|codincid||N|"
        Me.Caption = "Histórico Pedidos Clientes"
        Ordenacion = " ORDER BY numpedcl,fecpedcl "
    End If
        
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    
    If DatosADevolverBusqueda2 = "" Then
        CodTipoMov = "-1"
    Else
        CodTipoMov = DatosADevolverBusqueda2
    End If
    
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = conn
    Data1.RecordSource = "Select * from " & NombreTabla & " where numpedcl=" & CodTipoMov
    Data1.Refresh
    
    
    Label1(6).visible = vEmpresa.TieneAnalitica
    txtaux(11).visible = vEmpresa.TieneAnalitica
    txtAux2(11).visible = vEmpresa.TieneAnalitica
    
    If vParamAplic.NumeroInstalacion = vbTaxco And vUsu.Nivel2 = 2 Then
        txtaux(4).TabIndex = 200
         txtaux(6).TabIndex = 201
         txtaux(7).TabIndex = 202
    End If
    
    
    Me.Tag = "" 'Para que no carge los datos
    If DatosADevolverBusqueda2 = "" Then
        PonerModo 0
    Else
        If Data1.Recordset.EOF Then
            PonerModo 1
            Text1(0).BackColor = vbYellow
        Else
            Me.Tag = "P" 'Para que en el activate ponga los campos
            PonerModo 2
        End If
    End If
    CodTipoMov = "PEV"
    'Cargar el DataGrid de lineas de Revisiones inicialmente a nada DATA2
    primeravez = True
    
    'Poner los grid sin apuntar a nada
    LimpiarDataGrids
End Sub


Private Sub LimpiarCampos()
On Error Resume Next

    limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador.Caption = ""
    'Aqui va el especifico de cada form es
    '### a mano
    Me.cboFacturacion.ListIndex = -1
    Me.chkVisadoRes.Value = 0
    Me.chkRestoPed.Value = 0
    Me.chkServirCom.Value = 0
    
    Text3(0).Text = "BASE IMP."
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
End Sub


Private Sub frmA_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Agentes
Dim Indice As Byte
    Indice = 17
    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod agente
    FormateaCampo Text1(Indice)
    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nombre agente
End Sub

Private Sub frmAc_DatoSeleccionado(CadenaSeleccion As String)
    Text1(32).Text = RecuperaValor(CadenaSeleccion, 3)
    Text2(32).Text = RecuperaValor(CadenaSeleccion, 4) & "  " & RecuperaValor(CadenaSeleccion, 5)
End Sub

Private Sub frmAlm_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Almacenes Propios
    txtaux(0).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000") 'Cod Almacen
End Sub

Private Sub frmArt_DatoSeleccionado(CadenaSeleccion As String)
'Mantenimiento de Articulos
    txtaux(1).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod Artic
End Sub


Private Sub FrmArtEul_DatoSeleccionado(CadenaSeleccion As String)
    txtaux(1).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod Artic
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
Dim CadB As String
Dim Aux As String
      
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        If EsCabecera Then 'Llama desde VerTodos del Form
            CadB = ""
            Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
            CadB = Aux
            If EsHistorico Then
                Aux = ValorDevueltoFormGrid(Text1(1), CadenaDevuelta, 2)
                CadB = CadB & " and " & Aux
            End If
            CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
            PonerCadenaBusqueda
            Text1(0).Text = Format(RecuperaValor(CadenaDevuelta, 1), "0000000")
            
        ElseIf Val(cmdAux(0).Tag) > 0 Then
            'Llama desde boton busqueda centros de coste
            ' ---- [20/10/2009] [LAURA]: añadir campo centro de coste familia
            Me.txtaux(11).Text = RecuperaValor(CadenaDevuelta, 1)
            Me.txtAux2(11).Text = PonerNombreCCoste(Me.txtaux(11))
            
        Else 'Llama desde Prismatico Direcciones/Departamentos
            Text1(12).Text = Format(RecuperaValor(CadenaDevuelta, 1), "000")
            Text2(12).Text = RecuperaValor(CadenaDevuelta, 2)
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub frmC_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Clientes
    Text1(4).Text = RecuperaValor(CadenaSeleccion, 1)  'Cod Clien
    FormateaCampo Text1(4)
    HaDevueltoDatos = True
End Sub


Private Sub frmCP_DatoSeleccionado(CadenaSeleccion As String)
'Formulario Mantenimiento C. Postales
Dim Indice As Byte
Dim devuelve As String

    Indice = 9
    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'CPostal
     'Poblacion
    Text1(Indice + 1).Text = ObtenerPoblacion(Text1(Indice).Text, devuelve)
    'provincia
    Text1(Indice + 2).Text = devuelve
End Sub


Private Sub frmCV_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Clientes Varios
Dim Indice As Byte

    Indice = 6
    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'NIF
    Text1(Indice - 1).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Clien
    PonerDatosClienteVario (Text1(Indice).Text)
End Sub


Private Sub frmF_Selec(vFecha As Date) 'Calendario Fechas
Dim Indice As Byte
    Indice = CByte(Me.imgFecha(0).Tag) + 1
    Text1(Indice).Text = Format(vFecha, "dd/mm/yyyy")
End Sub


Private Sub frmFP_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Formas de Pago
Dim Indice As Byte
    Indice = 14
    Text1(Indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000") 'Cod Forma Pago
    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Forma Pago
End Sub


Private Sub frmList_DatoSeleccionado(CadenaSeleccion As String)
 'Cuando pasa de Pedido -> Albaran
'Aqui devuelve los valores que se introducen desde el Form de Listado de Pedido
'para generar el Albaran
Dim vSQL As String

    'Construimos parte de la SQL para insertar en tabla de Albaranes(scaalb)
    FechaAlb = RecuperaValor(CadenaSeleccion, 4)
    vSQL = ""
    vSQL = " '" & Format(FechaAlb, FormatoFecha) & "' as fechaalb, " 'Fecha Albaran
    vSQL = vSQL & "0 as factursn, " 'facturar s/n
    vSQL = vSQL & "codclien, nomclien, domclien, codpobla, pobclien, proclien, nifclien, "
    vSQL = vSQL & "telclien, coddirec, nomdirec, referenc,  "
    vSQL = vSQL & RecuperaValor(CadenaSeleccion, 1) & " as codtraba, " 'Trabajador de Albaran
    vSQL = vSQL & " codtraba as codtrab1, "
    vSQL = vSQL & RecuperaValor(CadenaSeleccion, 2) & " as codtrab2, " 'Material Preparado por
    vSQL = vSQL & "codagent, codforpa, "
    vSQL = vSQL & RecuperaValor(CadenaSeleccion, 3) & " as codenvio, " 'Cod Envio
    vSQL = vSQL & "dtoppago, dtognral, tipofact, observa01, observa02, observa03, observa04, observa05, "
    vSQL = vSQL & "numofert, fecofert, "  'Nº Oferta, fecha de la Oferta
    vSQL = vSQL & Text1(0).Text & " as numpedcl, '" 'Nº Pedido
    vSQL = vSQL & Format(Text1(1).Text, FormatoFecha) & "' as fecpedcl, '" 'Fecha Pedido
    vSQL = vSQL & Format(Text1(2).Text, FormatoFecha) & "' as fecentre, " 'Fecha Prevista Entrega
    vSQL = vSQL & Text1(18).Text & " as sementre " 'Semana entrega Pedido
    CadenaSQL = vSQL
    
    'Se almacena aqui si el usuario quiere imprimir el Albaran tras generarlo
    ImprimeAlb = CBool(RecuperaValor(CadenaSeleccion, 5))
    ImprimeEtiq = CBool(RecuperaValor(CadenaSeleccion, 6))
    ImprimeHojaExp = CBool(RecuperaValor(CadenaSeleccion, 7))
    
    'zona envio
    'en el 9 va la zona de envio. Para SALI no la utilizamos... de momento
    'RecuperaValor(CadenaSeleccion, 7))
    
    'Solo para la facturacion
    CtaBancoPropi = RecuperaValor(CadenaSeleccion, 8)
    
    'NUNCA SON A MOSTRADOR
    vSQL = RecuperaValor(CadenaSeleccion, 10)
    'EsAMostrador = vSQL = "1"
    
    'QUE TIPO DE ALBARAN
    EsAMostrador2 = RecuperaValor(CadenaSeleccion, 11)
    If EsAMostrador2 = "" Then EsAMostrador2 = "ALV" 'nunca deberia pasar
End Sub


Private Sub frmList2_DatoSeleccionado(CadenaSeleccion As String)
'devuelve los datos necesarios para grabar en la tabla cabecera del historico
    CadenaSQL = ""
    CadenaSQL = DBSet(RecuperaValor(CadenaSeleccion, 1), "F") & " as fechelim,"
    CadenaSQL = CadenaSQL & RecuperaValor(CadenaSeleccion, 2) & " as trabelim,"
    CadenaSQL = CadenaSQL & DBSet(RecuperaValor(CadenaSeleccion, 3), "T") & " as codincid"
End Sub


Private Sub frmMen_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mensaje de Nº de Serie disponibles
'En cadena seleccion estan concatenados los seleccionados
Dim I As Integer, J As Integer, K As Integer
Dim nSerie As String
Dim SQL As String
Dim devuelve As String
Dim cadSel As String
Dim codArtic As String
Dim RS As ADODB.Recordset
Dim Contador As Integer
Dim numSerie As CNumSerie

    On Error GoTo ErrorNSerie
    
    'Para cada articulo (separado por ., obtener los nº de serie empipados
    I = 0
    J = I + 1
    I = InStr(J, CadenaSeleccion, "·")
    
    While I > 0
        cadSel = Mid(CadenaSeleccion, J, I - J)
        
        'Para cada valor empipado actualizar la tabla sserie
        K = InStr(1, cadSel, "|")
        If K > 0 Then
            codArtic = Mid(cadSel, 1, K - 1) 'El primero es el codartic
            cadSel = Mid(cadSel, K + 1, Len(cadSel)) 'Los Nº de serie
            SQL = "select codartic, cantidad, numlinea from slialb "
            SQL = SQL & " WHERE codtipom='ALV' and numalbar= " & Me.cmdAux(1).Tag & " and codartic=" & DBSet(codArtic, "T")
            Set RS = New ADODB.Recordset
            RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        
            K = InStr(1, cadSel, "|")
            Contador = RS!cantidad
            While K > 0
                nSerie = Mid(cadSel, 1, K - 1)
                cadSel = Mid(cadSel, K + 1, Len(cadSel))
                
                If Contador = 0 Then
                    RS.MoveNext
                    If Not RS.EOF Then Contador = RS!cantidad
                End If
                If Contador > 0 Then
                    'Actualizar la tabla sserie
                    Set numSerie = New CNumSerie
                    numSerie.Cliente = Val(Text1(4).Text)
                    numSerie.DirDpto = Text1(12).Text
                    numSerie.tipoMov = "ALV"
                    'Obtenemos la fecha del albaran insertado
                    devuelve = DevuelveDesdeBDNew(conAri, "scaalb", "fechaalb", "codtipom", "ALV", "T", , "numalbar", Me.cmdAux(1).Tag, "N")
                    numSerie.FechaVta = devuelve
                    numSerie.ObtenFechaFinGarantia codArtic, devuelve

                    numSerie.NumAlbaran = Me.cmdAux(1).Tag
                    numSerie.NumLinAlb = ComprobarCero(RS!numlinea)
                    numSerie.Articulo = codArtic
                    numSerie.numSerie = nSerie
                    
                    numSerie.ActualizarNumSerie (True)
                    
                    Set numSerie = Nothing
                End If
                Contador = Contador - 1
                K = InStr(1, cadSel, "|")
            Wend
            RS.Close
            Set RS = Nothing
        End If
        J = I + 1
        I = InStr(J, CadenaSeleccion, "·")
    Wend
    
ErrorNSerie:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Actualizar tabla Nº Series", Err.Description
        MsgBox "No se cargaron correctamente los Nº de Serie.", vbExclamation
    End If
End Sub


Private Sub frmNSerie_CargarNumSeries()
'Insertar un registro en la tabla "sserie" por cada uno de los
'Nº de Serie introducidos en la Tabla Temporal
Dim RStmp As ADODB.Recordset
Dim RSalb As ADODB.Recordset
Dim SQL As String
Dim I As Byte

    On Error GoTo EInsertar
    
    SQL = "SELECT slialb.codartic, numlinea, cantidad "
    SQL = SQL & " FROM slialb INNER JOIN sartic on slialb.codartic=sartic.codartic "
    SQL = SQL & " WHERE (codtipom='ALV' and numalbar=" & Me.cmdAux(1).Tag
    SQL = SQL & " And nseriesn = 1) ORDER BY codartic, numlinea "

    Set RSalb = New ADODB.Recordset
    RSalb.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not RSalb.EOF 'Para cada linea del ALbaran
        'Recuperar los Nº Serie de ese articulo cargados en la Temporal
        'Seleccionar los nº de serie cargados en la temporal: tmpnseries
        SQL = "SELECT * FROM tmpnseries WHERE codusu=" & vUsu.Codigo
        SQL = SQL & " AND codartic=" & DBSet(RSalb!codArtic, "T")
        SQL = SQL & " ORDER BY codartic, numlinea "
        Set RStmp = New ADODB.Recordset
        RStmp.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        'If Not RStmp.EOF Then RStmp.MoveFirst
        'Intentar asignar un Nº serie al total de cantidad del articulo
        
        For I = 1 To RSalb!cantidad
            If Not RStmp.EOF Then
                InsertarNSerie RStmp!numSerie, RStmp!codArtic, RSalb!numlinea, DBLet(RStmp!nummante, "T")
                RStmp.MoveNext
            End If
        Next I
        RStmp.Close
        Set RStmp = Nothing
        RSalb.MoveNext
    Wend
    RSalb.Close
    Set RSalb = Nothing
    
EInsertar:
    If Err.Number <> 0 Then MuestraError Err.Number, "Insertando Nº Serie", Err.Description
End Sub


Private Sub frmOC_DatoSeleccionado(CadenaSeleccion As String)
    FechaAlb = CadenaSeleccion
End Sub

Private Sub frmOT_DatoSeleccionado(CadenaSeleccion As String)
    FechaAlb = CadenaSeleccion
End Sub

Private Sub frmT_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Trabajadores
Dim Indice As Byte

    Indice = 3
    Text1(Indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000") 'Cod Trabajador
    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Trabajador
End Sub


Private Sub imgBuscar_Click(index As Integer)
Dim Indice As Byte

    If Modo = 2 Or Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass
    TerminaBloquear

    Select Case index
        Case 0 'Cod. Cliente
            HaDevueltoDatos = False
            PonerFoco Text1(4)
            Indice = 4
            Set frmC = New frmFacClientes3
            frmC.DatosADevolverBusqueda = "0"
            If Not IsNumeric(Text1(4).Text) Then Text1(4).Text = ""
            frmC.Show vbModal
            Set frmC = Nothing
            If HaDevueltoDatos Then
                txtAnterior = ""
                Text1_LostFocus 4
                txtAnterior = Text1(4).Text
            End If
        Case 1 'NIF para cliente de Varios
            Set frmCV = New frmFacClientesV
            frmCV.DatosADevolverBusqueda = "0"
            frmCV.Show vbModal
            Set frmCV = Nothing
            Indice = 6
            
        Case 2 'Cod. Direc.
            'Mostrar las Direc. o Dptos del cliente seleccionado
            If Trim(Text1(4).Text) = "" Then
                MsgBox "Debe seleccionar un cliente.", vbInformation
                Screen.MousePointer = vbDefault
                Exit Sub
            Else
                EsCabecera = False
                MandaBusquedaPrevia " codclien= " & Val(Text1(4).Text)
                Indice = 12
            End If
            
        Case 3 'Realizada Por Trabajador
            Indice = 3
            Set frmT = New frmAdmTrabajadores
            frmT.DatosADevolverBusqueda = "0"
            frmT.Show vbModal
            Set frmT = Nothing
            
        Case 4 'Forma de Pago
            Indice = 14
            PonerFoco Text1(Indice)
            Set frmFP = New frmFacFormasPago
            frmFP.DatosADevolverBusqueda = "0"
            frmFP.Show vbModal
            Set frmFP = Nothing
            
        Case 5 'Agente
            Indice = 17
            PonerFoco Text1(Indice)
            Set frmA = New frmFacAgentesCom
            frmA.DatosADevolverBusqueda = "0"
            frmA.Show vbModal
            Set frmA = Nothing
            
        Case 6 'Cod. Postal
            Set frmCP = New frmCPostal
            frmCP.DatosADevolverBusqueda = "0"
            frmCP.Show vbModal
            Set frmCP = Nothing
            Indice = 9
            VieneDeBuscar = True
            
        Case 9 'masl
             If Trim(Text1(4).Text) = "" Then
                MsgBox "Debe seleccionar un cliente.", vbInformation
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            If Trim(Text1(12).Text) = "" Then
                MsgBox "Debe seleccionar un obra para el cliente.", vbInformation
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            EsCabecera = False
            Set frmAc = New frmObraActua
            frmAc.DatosADevolverBusqueda = Text1(4).Text & "|" & Text1(12).Text & "|"
            frmAc.Show vbModal
            Set frmAc = Nothing
            Indice = 9
            'MandaBusquedaPrevia "codclien = " & Text1(4).Text & " AND coddirec = " & Text1(12).Text, False
            
       
            
            
            
    End Select
    If index <> 9 Then
        PonerFoco Text1(Indice)
    Else
        PonerFoco Text1(32)
    End If
    Screen.MousePointer = vbDefault
    
    If Modo = 4 Then
         If Not BLOQUEADesdeFormulario(Me) Then cmdCancelar_Click
    End If
End Sub


Private Sub imgBuscar2_Click(index As Integer)
    If Modo <> 5 Then Exit Sub
    Screen.MousePointer = vbHourglass
    FechaAlb = ""
    Select Case index
    Case 11
    
    
        EsCabecera = False
        cmdAux(0).Tag = "9"
        
        Set frmB = New frmBuscaGrid
        FechaAlb = "cabccost"
        If vParamAplic.ContabilidadNueva Then FechaAlb = "ccoste"
        frmB.vCampos = "Codigo|" & FechaAlb & "|codccost|T||20·Descripción|" & FechaAlb & "|nomccost|T||70·"
        frmB.vTabla = FechaAlb
        FechaAlb = ""
        frmB.vSQL = ""
        HaDevueltoDatos = False
        '###A mano
        frmB.vDevuelve = "0|1|"
        frmB.vTitulo = "Centros de coste"
        frmB.vselElem = 0
        frmB.vConexionGrid = conConta
        
        frmB.Show vbModal
        Set frmB = Nothing
        cmdAux(0).Tag = "-1"

    
    
    Case 12
        Set frmOT = New frmObraOT
        frmOT.DatosADevolverBusqueda = "0|1|"
        frmOT.Show vbModal
        Set frmOT = Nothing
        If FechaAlb <> "" Then
            txtaux(12).Text = RecuperaValor(FechaAlb, 1)
            txtAux2(12).Text = RecuperaValor(FechaAlb, 2)
            txtAnterior = txtaux(12)
            PonerFoco txtaux(12)
        End If
    
    Case 13
        Set frmOC = New frmObraCapitulo
        frmOC.DatosADevolverBusqueda = "0|1|"
        frmOC.Show vbModal
        Set frmOC = Nothing
        If FechaAlb <> "" Then
            txtaux(13).Text = RecuperaValor(FechaAlb, 1)
            txtAux2(13).Text = RecuperaValor(FechaAlb, 2)
            txtAnterior = txtaux(13)
            PonerFoco txtaux(13)
        End If
    
    End Select
    FechaAlb = ""
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub imgFecha_Click(index As Integer) 'Abre calendario Fechas
Dim Indice As Byte

   If Modo = 2 Or Modo = 0 Then Exit Sub
   Screen.MousePointer = vbHourglass
   Set frmF = New frmCal
   frmF.Fecha = Now
   Indice = index + 1
   Me.imgFecha(0).Tag = index
   
   PonerFormatoFecha Text1(Indice)
   If Text1(Indice).Text <> "" Then frmF.Fecha = CDate(Text1(Indice).Text)

   Screen.MousePointer = vbDefault
   frmF.Show vbModal
   Set frmF = Nothing
   Text1_LostFocus CInt(Indice)
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
    End If
End Sub


Private Sub mnGenAlbaran_Click()
'Pasar una Pedido a Albaran
Dim Resp As Byte
Dim b As Boolean
Dim cadMen As String

    'Comprobar que hay un Pedido seleccionado
    If Not ComprobarOpcionTraspaso(False) Then Exit Sub
    
    
    'si no se va a servir completo preguntar como se quiere servir si completo o no
    If Me.chkServirCom = 0 Then
        'Preguntar si se sirve el pedido completo o no
        Resp = MsgBox("¿Servir el pedido completo?", vbYesNoCancel)
        If Resp = vbCancel Then Exit Sub
    
        If Resp = vbYes Then
            AlbCompleto = True 'SERVIR COMPLETO
        Else
            AlbCompleto = False
        End If
    Else
        AlbCompleto = True
    End If
        
    If AlbCompleto Then 'SERVIR COMPLETO
        Screen.MousePointer = vbHourglass
        'comprobar si hay control de stock si se puede servir el pedido
        b = SePuedeServirPedido
        
        If b Then 'Hay suficiente stock
            'Si hay stock generar albaran completo
            GenerarAlbaran False
        Else
            Screen.MousePointer = vbDefault
            'Si no se puede servir mostrar mensaje detallando y bloquear
            cadMen = "No hay suficiente Stock para servir el Pedido. "
            cadMen = cadMen & vbCrLf & "¿Desea Ver Detalle?"
            If MsgBox(cadMen, vbYesNo, "Contol de Stock") = vbYes Then
                'ANTES 01/12/08
                'frmMensajes.cadWHERE = " WHERE numpedcl = " & Text1(0).Text & " "   'And sfamia.instalac = 0 "
                'ahora
                frmMensajes.cadWhere = " WHERE numpedcl = " & Text1(0).Text & " and ctrstock=1 "
                frmMensajes.vCampos = NomTablaLineas
                frmMensajes.OpcionMensaje = 2 'Articulos sin Stock
                frmMensajes.Show vbModal
            End If
            Exit Sub
        End If
        
    Else 'SERVIR INCOMPLETO
        AlbCompleto = False
        'Si no se va a servir completo Mostrar lineas para que se indiquen las Servidas
        MsgBox "Introduzca la cantidad  a servir para cada línea.", vbInformation
        Modo = 6
        gridCargado = False
        Me.cmdAceptar.visible = True
        Me.cmdCancelar.visible = True
        PonerModoOpcionesMenu Modo
        CargaGrid DataGrid1, Data2, True, True
        CargaTxtAuxServidas True, True
        PonerFoco txtaux(3)
        primeravez = True
    End If
End Sub


Private Function ComprobarOpcionTraspaso(Factura As Boolean) As Boolean

    ComprobarOpcionTraspaso = False
    
   'Comprobar que hay un Pedido seleccionado
    If Text1(0).Text = "" Then Exit Function
    
    CtaBancoPropi = "- No tiene lineas el pedido" & vbCrLf
    If Not (Data2.Recordset Is Nothing) Then
        If Data2.Recordset.RecordCount > 0 Then CtaBancoPropi = ""
    End If
    
    
    'Comprobar que el Pedido esta visado por el Responsable
    If Me.chkVisadoRes = 0 Then CtaBancoPropi = CtaBancoPropi & "- El pedido debe tener el Visado del Responsable." & vbCrLf
        
    
    'si no se va a servir completo preguntar como se quiere servir si completo o no
    If Factura Then
        If Me.chkServirCom = 0 Then CtaBancoPropi = CtaBancoPropi & "-Solo se facturan drectamente pedidos completos" & vbCrLf
    End If
        
        
    If CtaBancoPropi <> "" Then
        CtaBancoPropi = "Faltan campos: " & vbCrLf & vbCrLf & CtaBancoPropi
        MsgBox CtaBancoPropi, vbExclamation
        CtaBancoPropi = ""
        Exit Function
    End If
        
        
    '17 Diciembre 2010
    If EsClienteBloqueado(Text1(4).Text, False, False) Then Exit Function
        
        
    If vEmpresa.TieneAnalitica Then
        'Todas las lineas deben
        CtaBancoPropi = DevuelveDesdeBD(conAri, "count(*)", "sliped", "codccost is null AND numpedcl", Text1(0).Text)
        kCampo = 0
        If CtaBancoPropi <> "" Then
            If CtaBancoPropi > 0 Then
                kCampo = 1
                CtaBancoPropi = "Lineas sin asignar centro de coste: " & CtaBancoPropi
                MsgBox CtaBancoPropi, vbExclamation
            End If
        End If
        CtaBancoPropi = ""
        If kCampo = 1 Then Exit Function
        
    End If
    
    
    
    
    
    
    'Llegado aqui: bien
    ComprobarOpcionTraspaso = True
End Function


Private Sub mnGeneraFactura_Click()
Dim b As Boolean

   'Comprobaciones iniciales
   '----------------------------------------------------------------------------
   If Not ComprobarOpcionTraspaso(True) Then Exit Sub
    Screen.MousePointer = vbHourglass
    
    'Solo se generan albarenes completos
    AlbCompleto = True
    
    'comprobar si hay control de stock si se puede servir el pedido
    b = SePuedeServirPedido
        
    If b Then 'Hay suficiente stock
        'Si hay stock generar albaran completo
        GenerarAlbaran True
    Else
        Screen.MousePointer = vbDefault
        'Si no se puede servir mostrar mensaje detallando y bloquear
        TituloLinea = "No hay suficiente Stock para servir el Pedido. "
        TituloLinea = TituloLinea & vbCrLf & "¿Desea Ver Detalle?"
        If MsgBox(TituloLinea, vbYesNo, "Contol de Stock") = vbYes Then
            frmMensajes.cadWhere = " WHERE numpedcl = " & Text1(0).Text & " And sfamia.instalac = 0 "
            frmMensajes.vCampos = NomTablaLineas
            frmMensajes.OpcionMensaje = 2 'Articulos sin Stock
            frmMensajes.Show vbModal
        End If
        TituloLinea = ""
    End If
End Sub


Private Sub mnImpOrde_Click()
'Impreme la Orden de Instalacion de un pedido
Dim cadFormula As String, cadParam As String
Dim devuelve As String, nomDocu As String
Dim numParam As Byte

    'Comprobar que hay un pedido seleccionado
    If Text1(0).Text = "" Then
        MsgBox "No hay ningún Pedido seleccionado.", vbInformation
        Exit Sub
    End If

'    'Comprobar que algun Articulo pertenece a la familia de Instalaciones
'    If Not PedidoConInstalaciones Then
'        MsgBox "El Pedido no tiene ningún Artículo que sea Instalación.", vbInformation
'        Exit Sub
'    End If

    '=======================================================================
    '=============== FORMULA    ============================================
    cadFormula = ""
    cadParam = ""
    numParam = 0
    
    If Text1(0).Text <> "" Then 'Seleccionar el Pedido
        devuelve = "{" & NombreTabla & ".numpedcl}=" & Val(Text1(0).Text)
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
    End If
    
    'Seleccionar solo las lineas de Articulos que son de una familia que es Instalacion
    'Devuelve = "{sfamia.instalac}=1"
    'If Not AnyadirAFormula(cadFormula, Devuelve) Then Exit Sub
    
    If Not PonerParamRPT2(9, cadParam, numParam, nomDocu, pImprimeDirecto, pPdfRpt, pRptvMultiInforme) Then Exit Sub

    With frmImprimir
        .NombreRPT = nomDocu
        .NombrePDF = pPdfRpt
        .SeleccionaRPTCodigo = pRptvMultiInforme
        .FormulaSeleccion = cadFormula
        .OtrosParametros = cadParam
        .NumeroParametros = numParam
        .SoloImprimir = False
        .EnvioEMail = False
        .Opcion = 39
        .Titulo = ""
        .Show vbModal
    End With
End Sub


Private Sub mnImpPedido_Click()
'Imprime un Pedido
       frmListadoOfer.NumCod = Text1(0).Text   'Nº de Pedido
       frmListadoOfer.codClien = Text1(4).Text 'cliente del pedido
       If EsHistorico Then
            frmListadoOfer.FecEntre = Text1(1).Text   'Fecha de Pedido
            AbrirListadoOfer (239) '239: Informe de Pedidos (Historico)
       Else
            AbrirListadoOfer (38) '38: Informe de Pedidos
       End If
End Sub





' ---- [04/11/2009] [LAURA] : Añadir botón para enviar informe confirmacion entrega del Pedido
Private Sub mnConfirmacion_Click()
'Enviar confirmacion de entrega

    'Comprobar que hay un pedido seleccionado
    If Text1(0).Text = "" Then
        MsgBox "No hay ningún Pedido seleccionado.", vbInformation
        Exit Sub
    End If

    'Debe estar visado el responsable
    If Me.Data1.Recordset.EOF Then Exit Sub
    If Not CInt(Data1.Recordset!visadore) = 1 Then
        MsgBox "Debe estar visado por el responsable.", vbInformation
        Exit Sub
    End If


    If CInt(Me.Data1.Recordset!envconfir) = 1 Then
        If Not MsgBox("Ya se ha enviado una confirmación de entrega." & vbCrLf & "¿Desea continuar?", vbQuestion + vbYesNo) = vbYes Then
            Exit Sub
        End If
    End If



    'Imprime una confirmacion entrega Pedido
    frmListadoOfer.NumCod = Text1(0).Text   'Nº de Pedido
    frmListadoOfer.FecEntre = Text1(1).Text  'fecha del pedido

    AbrirListadoOfer (238) '38: Informe confirmacion entrega de Pedidos

End Sub

' ----







Private Sub mnLineas_Click()
    BotonMtoLineas 0, "Pedidos"
End Sub


Private Sub mnModificar_Click()
    If Modo = 5 Then 'Modificar lineas
         BotonModificarLinea
    Else   'Modificar Pedido
         If BLOQUEADesdeFormulario(Me) Then BotonModificar
    End If
End Sub


Private Sub mnNuevo_Click()
    If Modo = 5 Then 'Añadir lineas
         BotonAnyadirLinea False
    Else 'Añadir Cabecera de Pedidos
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


'Private Sub SSTab1_Click(PreviousTab As Integer)
'    Me.Label1(35).visible = Me.SSTab1.Tab = 0
'    Me.Text2(16).visible = Me.SSTab1.Tab = 0
'    Me.Label1(6).visible = (Modo = 5) And (vEmpresa.TieneAnalitica) And SSTab1.Tab = 0
'    Me.txtAux2(11).visible = (Modo = 5) And (vEmpresa.TieneAnalitica) And Me.SSTab1.Tab = 0
'
'End Sub

Private Sub Text1_Change(index As Integer)
    If index = 9 Then HaCambiadoCP = True
End Sub

'### A mano
'Los metodos del text tendran que estar
'Los descomentamos cuando esten puestos ya los controles
Private Sub Text1_GotFocus(index As Integer)
    txtAnterior = Text1(index).Text
    kCampo = index
    If index = 9 Then HaCambiadoCP = False 'CPostal
    ConseguirFoco Text1(index), Modo
End Sub


Private Sub Text1_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
Dim Ind As Integer
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
    
        
    If KeyCode = 43 Or KeyCode = 107 Or KeyCode = 187 Then
        If Text1(index).Text = "" Then
            Ind = -1
            Select Case index
            Case 3
                Ind = 3
            Case 4
                Ind = 0
            Case 6
                Ind = 1
            Case 9
                Ind = 6
            Case 12
                Ind = 2
            Case 17
                Ind = 5
            Case 14
                Ind = 4
            End Select
            If Ind >= 0 Then
                PulsadoMas2 = True
                PulsarTeclaMas True, Ind
            End If
        End If
    End If
    
End Sub


Private Sub Text1_KeyPress(index As Integer, KeyAscii As Integer)
    If index = 31 And KeyAscii = 13 Then 'ENTER
        Me.SSTab1.Tab = 1
        PonerFoco Text1(19)
    Else
        KEYpress KeyAscii
    End If
End Sub


'----------------------------------------------------------------
'----------------------------------------------------------------
' Cunado el campo de texto pierde el enfoque
' Es especifico de cada formulario y en el podremos controlar
' lo que queramos, desde formatear un campo si asi lo deseamos
' hasta pedir que nos devuelva los datos de la empresa
'----------------------------------------------------------------
'----------------------------------------------------------------
Private Sub Text1_LostFocus(index As Integer)
Dim devuelve As String
        
    'Han pulsado el mas
    If PulsadoMas2 Then
        'Para que cuando pulse el mas abra el form
        PulsadoMas2 = False
        Text1(index).Text = ""
        Exit Sub
    End If
        
        
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    If Not PerderFocoGnral(Text1(index), Modo) Then Exit Sub
       
    
    If txtAnterior = Text1(index).Text And Text1(index).Text <> "" Then Exit Sub
    
    'Si queremos hacer algo ..
    Select Case index
        Case 1, 2 'Fecha Oferta, Fecha Entrega
            If Text1(index).Text = "" Then Exit Sub
            PonerFormatoFecha Text1(index)
            
            If index = 2 And Text1(index).Text <> "" Then 'Fecha Entrega
                'Comprobar que es posterior a la del pedido
                If Not EsFechaIgualPosterior(Text1(1).Text, Text1(2).Text, True, "La Fecha de Entrega debe ser posterior a la Fecha del Pedido.") Then
                    Text1(index).Text = ""
                    PonerFoco Text1(index)
                    Exit Sub
                End If
                'Obtener la semana de Entrega
                Text1(18).Text = CalculaSemana(CDate(Text1(2).Text))
            End If
            
        Case 3 'Cod Vendedor
            If PonerFormatoEntero(Text1(index)) Then
                Text2(index).Text = PonerNombreDeCod(Text1(index), conAri, "straba", "nomtraba")
            Else
                Text2(index).Text = ""
            End If
            
        Case 4 'Cod. Cliente
            If PonerFormatoEntero(Text1(index)) Then
                If Modo = 1 Then 'Modo=1 Busqueda
                    Text1(5).Text = PonerNombreDeCod(Text1(index), conAri, "sclien", "nomclien")
                Else 'Insertando
                    PonerDatosCliente2 (Text1(index).Text)
                    If Text1(index).Text = "" Then PonerFoco Text1(index)
                End If
            Else
                LimpiarDatosCliente
            End If
            
        Case 6 'NIF
'            If Not EsDeVarios Then Exit Sub
'            If Modo = 4 Then 'Modificar
'                'si no se ha modificado el nif del cliente no hacer nada
'                If Text1(6).Text = Data1.Recordset!nifClien Then
'                    Exit Sub
'                End If
'            End If
'            PonerDatosClienteVario (Text1(Index).Text)
             
        Case 9 'Cod. Postal
            If Text1(index).Locked Then Exit Sub
            If Text1(index).Text = "" Then
                Text1(index + 1).Text = ""
                Text1(index + 2).Text = ""
                Exit Sub
            End If
            If (Not VieneDeBuscar) Or (VieneDeBuscar And HaCambiadoCP) Then
                 Text1(index + 1).Text = ObtenerPoblacion(Text1(index).Text, devuelve)
                 Text1(index + 2).Text = devuelve
            End If
            VieneDeBuscar = False
            
        Case 12 'Cod. Direc
            If Text1(index).Text = "" Then
                Text2(12).Text = ""
'                Exit Sub
            Else
                Text1(index).Text = Format(Text1(index).Text, "000")
            End If
            If Modo = 1 Then
                If Not IsNumeric(Text1(index).Text) Then
                    MsgBox "Campo numerico", vbExclamation
                    PonerFoco Text1(index)
                End If
                Exit Sub
            End If

            If PonerDptoEnCliente Then
                'Comprobar que el cliente seleccionada tiene esa direccion
                'Comprobar que el cliente tiene mantenimientos en esa direc/dpto
                devuelve = DevuelveDesdeBDNew(conAri, "scaman", "nummante", "codclien", Text1(4).Text, "N", , "coddirec", Text1(12).Text, "N")
                If devuelve <> "" Then MsgBox "El cliente tiene Mantenimientos.", vbInformation
            ElseIf Text1(index) <> "" Then
                Text2(12).Text = ""
                PonerFoco Text1(index)
            End If
            
        Case 13 'Referencia Obligatoria
            If Trim(Text1(4).Text) <> "" Then ComprobarRefObligatoria
            '-------------------------------------------------------------------------
        Case 14 'Forma de Pago
            If PonerFormatoEntero(Text1(index)) Then
                Text2(index).Text = PonerNombreDeCod(Text1(index), conAri, "sforpa", "nomforpa")
            Else
                Text2(index).Text = ""
            End If
            
        Case 15, 16 'Descuentos
            If PonerFormatoDecimal(Text1(index), 4) Then  'Tipo 4: Decimal(4,2)
                If Modo = 4 Then CalcularDatosFactura
            End If
        
        Case 17 'Cod. Agente
            If PonerFormatoEntero(Text1(index)) Then
                Text2(index).Text = PonerNombreDeCod(Text1(index), conAri, "sagent", "nomagent")
            Else
                Text2(index).Text = ""
            End If
        Case 32
            'Actuacion
            If Text1(32).Text = "" Then
                Text2(32).Text = ""
            Else
                PonerCampoActuacion
                If Text1(32).Text = "" Then PonerFoco Text1(32)
            End If
    End Select
End Sub


Private Sub HacerBusqueda()
Dim CadB As String

    CadB = ObtenerBusqueda(Me, False)
    
    If chkVistaPrevia = 1 Then
        EsCabecera = True
        MandaBusquedaPrevia CadB
    ElseIf CadB <> "" Then
        'Se muestran en el mismo form
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Sub MandaBusquedaPrevia(CadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim cad As String
Dim tabla As String
Dim Titulo As String
Dim Desc As String, devuelve As String
    'Llamamos a al form
    '##A mano
    cad = ""
    If EsCabecera Then
        cad = cad & ParaGrid(Text1(0), 15, "Nº Pedido")
        cad = cad & ParaGrid(Text1(1), 20, "Fecha Ped.")
        cad = cad & ParaGrid(Text1(4), 15, "Cliente")
        cad = cad & ParaGrid(Text1(5), 50, "Nombre Cliente")
        tabla = NombreTabla
        If EsHistorico Then
            Titulo = "Histórico de Pedidos"
            devuelve = "0|1|"
        Else
            Titulo = "Pedidos"
            devuelve = "0|"
        End If
        
    Else
        If vParamAplic.HayDeparNuevo = 1 Then
            Titulo = "Dptos Cliente: "
            Desc = "Dpto."
        ElseIf vParamAplic.HayDeparNuevo = 0 Then
            Titulo = "Direc. Cliente: "
            Desc = "Direc."
        Else
            Titulo = "Obra Cliente: "
            Desc = "Obra"
        End If
        Titulo = Titulo & Text1(4).Text & " - " & Text1(5).Text
        cad = cad & "Cod. " & Desc & "|sdirec|coddirec|N||15·"
        cad = cad & "Desc. " & Desc & "|sdirec|nomdirec|T||35·"
        tabla = "sdirec"
        devuelve = "0|1|"
    End If
    
           
    If cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = cad
        frmB.vTabla = tabla
        frmB.vSQL = CadB
        HaDevueltoDatos = False
        '###A mano
'        frmB.vDevuelve = "0|1|"
        frmB.vDevuelve = devuelve
        frmB.vTitulo = Titulo
        frmB.vselElem = 0
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
        If Modo = 1 Then
            Me.cboFacturacion.ListIndex = -1
            PonerFoco Text1(kCampo)
'            Text1(0).BackColor = vbYellow
        End If
        Exit Sub
    Else
        Data1.Recordset.MoveFirst
        PonerModo 2
        PonerCampos
        DataGrid1_RowColChange 0, 0
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

    'Datos de la tabla slipre
    CargaGrid DataGrid1, Data2, True

    PonerForaGrid
    
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
  
    'Poner Nombre del Trabajador
    Text2(3).Text = PonerNombreDeCod(Text1(3), conAri, "straba", "nomtraba")
    'Poner Desc. del Dpto/Direc.
    Text2(12).Text = DevuelveDesdeBDNew(conAri, "sdirec", "nomdirec", "codclien", Text1(4).Text, "N", , "coddirec", Text1(12).Text, "N")
    'Poner el Nombre del Agente
    Text2(17).Text = PonerNombreDeCod(Text1(17), conAri, "sagent", "nomagent")
    'Poner la Desc. de la Forma de Pago
    Text2(14).Text = PonerNombreDeCod(Text1(14), conAri, "sforpa", "nomforpa")
       
    PonerCamposLineas 'Pone los datos de las tablas de lineas de Ofertas
    
    If EsHistorico Then
        'poner datos de eliminacion
        Text2(27).Text = PonerNombreDeCod(Text1(27), conAri, "straba", "nomtraba", "codtraba")
        Text2(28).Text = PonerNombreDeCod(Text1(28), conAri, "sincid", "nomincid", "codincid")
    End If
    
    PonerCampoActuacion
    
    
    CalcularDatosFactura
    
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
    
    lblF.Caption = ""
    
    'Actualiza Iconos Insertar,Modificar,Eliminar
    ActualizarToolbarGnral Me.Toolbar1, Modo, Kmodo, btnAnyadir
    
    Modo = Kmodo
    PonerIndicador lblIndicador, Modo
    
    If Modo = 6 Then Me.lblIndicador.Caption = "Insertar Cant. Servidas"
    
    'Modo 2. Hay datos y estamos visualizandolos
    '=========================================
    b = (Modo = 2)
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda2 <> "" Then
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
    BloquearText1 Me, Modo
    'Campo Numero de Albaran siempre bloqueado, excepto si estamos en modo de busqueda
    b = (Modo <> 1)
    BloquearTxt Text1(0), b, True
    'Bloquear los campos de Oferta
    BloquearTxt Text1(24), b
    BloquearTxt Text1(25), b


    'Campo Semana Se calcula automat., siempre bloqueado
    'BloquearTxt Text1(18), True
    
    '-----  Datos Totales de Factura siempre bloqueado
    For I = 33 To 56
        BloquearTxt Text3(I), True
    Next I
    'Campo B.Imp y Imp. IVA siempre en azul
    Text3(36).BackColor = &HFFFFC0
    For I = 46 To 48
        Text3(I).BackColor = &HFFFFC0
        Text3(I + 6).BackColor = &HFFFFC0
    Next I
    'Campos total Factura en verde
    Text3(55).BackColor = &HC0FFC0
    Text3(56).BackColor = &HC0FFC0    'Tatal factura
    '---------------------------------------------------
    
    
    b = (Modo = 3) Or (Modo = 4) Or (Modo = 1)
    Me.cboFacturacion.Enabled = b
    Me.chkVisadoRes.Enabled = b
    Me.chkServirCom.Enabled = b
    Me.chkRecogeClien.Enabled = b
    Me.chkEnviadaConfir.Enabled = b
    
    
    'Si no es modo lineas Boquear los TxtAux
    For I = 0 To 8
        BloquearTxt txtaux(I), (Modo <> 5)
    Next I
 
    
    
    '---------------------------------------------
    b = (Modo <> 0 And Modo <> 2 And Modo <> 5)
    cmdCancelar.visible = b
    cmdAceptar.visible = b
    
    For I = 0 To Me.imgFecha.Count - 1
        Me.imgFecha(I).Enabled = b
    Next I
    
    For I = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(I).Enabled = b
    Next I
    Me.imgBuscar(1).visible = False
           
    
    'usuarios sin permiso 2 (taller en taxco) no puede tocar la forma de pago
    If vUsu.Nivel2 = 2 And b Then
        BloquearTxt Text1(14), True
        imgBuscar(4).Enabled = False
    End If
    
    
    
    ' ---- [20/10/2009] [LAURA] : añadir del centro de coste
  '  SSTab1_Click 0
  '  BloquearTxt txtAux2(11), True
  '  BloquearTxt Text2(16), True
    
       
    Me.chkVistaPrevia.Enabled = (Modo <= 2)
    
    'Poner el tamaño de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos
       
    PonerModoOpcionesMenu (Modo) 'Activar opciones de menu según modo
    PonerOpcionesMenu 'Activar opciones de menu según nivel de permisos del usuario
    Exit Sub
    
EPonerModo:
    MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
    PonerLongCamposGnral Me, Modo, 1
End Sub


Private Function DatosOk() As Boolean
'Comprueba si los datos de la cabecera son correctos antes de Insertar o Modificar el
'Pedido
Dim b As Boolean
Dim devuelve As String

    On Error GoTo EDatosOK

    DatosOk = False
    b = CompForm(Me, 1) 'Comprobar formato datos ok
    If Not b Then Exit Function
    
    'Comprobar que la Fecha Entrega es posterior a la del pedido
    If Not EsFechaIgualPosterior(Text1(1).Text, Text1(2).Text, True, "La Fecha de Entrega debe ser posterior a la Fecha del Pedido.") Then Exit Function
    
    'Comprobar si la referencia del cliente es obligatoria que tenga valor
    If Trim(Text1(4).Text) <> "" Then
        devuelve = DevuelveDesdeBDNew(conAri, "sclien", "referobl", "codclien", Text1(4).Text, "N")
        If devuelve = "1" And Text1(13).Text = "" Then 'Referencia Obligatoria
            MsgBox "La Referencia del Cliente es Obligatoria.", vbInformation
            PonerFoco Text1(13)
            b = False
        End If
    End If
    If Not b Then Exit Function
          
    DatosOk = b
    
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Function DatosOkLinea() As Boolean
'Comprueba si los datos de una linea son correctos antes de Insertar o Modificar
'una linea del Pedido
Dim b As Boolean
Dim I As Byte
Dim vArtic As CArticulo
Dim Aux As String

    On Error GoTo EDatosOkLinea

    DatosOkLinea = False
    
    
    'Febrero 2010   Si han apretado Alt+A NO recalcula
    '----------------------------------------------------------------------------------
    'txtAux(8).Text = CalcularImporte(txtAux(3).Text, txtAux(4).Text, txtAux(6).Text, txtAux(7).Text, vParamAplic.TipoDtos)
    Aux = RecalculoImporteLineas(txtaux(3), txtaux(4), txtaux(6), txtaux(7), vParamAplic.TipoDtos)
    Aux = Format(Aux, FormatoImporte)
    If Aux <> txtaux(8).Text Then txtaux(8).Text = Aux
    

    
    
    
    b = True
    'Comprobar que los campos NOT NULL tienen valor
    For I = 0 To txtaux.Count - 3  'los dos ultimos(12,13) campos pueden ser nulos y uno mas pq empieza en el 0
        'Debug.Print i & ": " & txtAux(i).Text
        If txtaux(I).Text = "" And I <> 10 Then
            If I = 11 And vEmpresa.TieneAnalitica = False Then
                'puede ser nulo
            Else
                
                MsgBox "El campo " & txtaux(I).Tag & " no puede ser nulo", vbExclamation
                b = False
                PonerFoco txtaux(I)
                Exit Function
            End If
        End If
    Next I
        
    'Comprobar que existe el articulo en el almacen seleccionado
    Set vArtic = New CArticulo
    vArtic.Codigo = txtaux(1).Text
    If Not vArtic.ExisteEnAlmacen(txtaux(0).Text) Then
        b = False
        PonerFoco txtaux(1)
    End If
    Set vArtic = Nothing
    
    
    
    DatosOkLinea = b

EDatosOkLinea:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Sub Text2_GotFocus(index As Integer)
    If index = 16 Then
        'Campo observaciones. NO, repito NO, se selecciona todo
        If Text2(index).Text <> "" Then
            Text2(index).Text = Text2(index).Text & " "
            Text2(index).SelStart = Len(Text2(index).Text)
        End If
    End If
End Sub

Private Sub Text2_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
    If index = 16 And KeyCode = 40 Then 'campo Ampliacion linea y Flecha hacia abajo
       ' PonerFocoBtn Me.cmdAceptar
    Else
        KEYdown KeyCode
    End If
End Sub


Private Sub Text2_KeyPress(index As Integer, KeyAscii As Integer)
    'campo Ampliación linea y ENTER
    KEYpress KeyAscii
   ' If Index = 16 And KeyAscii = 13 Then PonerFocoBtn Me.cmdAceptar
End Sub

Private Sub Text2_LostFocus(index As Integer)
    If index = 16 And (Text2(index).Locked = False) Then Text2(index).Text = UCase(Text2(index).Text)
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.index
        Case 1: mnBuscar_Click  'Buscar
        Case 2: mnVerTodos_Click  'Todos
            
        Case 5: mnNuevo_Click  'Nuevo
        Case 6: mnModificar_Click  'Modificar
        Case 7: mnEliminar_Click  'Borrar
            
        Case 10: mnLineas_Click  'Lineas
        Case 11:
                If Modo = 5 Then
                    'Insertar intercalando
                    BotonAnyadirLinea True
                Else
                    mnGenAlbaran_Click 'Generar Albaran
                End If
        
        
        
        Case 12: mnGeneraFactura_Click 'Genera la factura directamente
        
        
            
        Case 14: mnImpPedido_Click 'Imprimir Pedido
        Case 15: mnImpOrde_Click 'Imprimir Orden Instalacion
        ' ---- [04/11/2009] [LAURA] : Añadir botón para enviar informe confirmacion entrega del Pedido
        Case 16: mnConfirmacion_Click 'confirmacion de entrega
        ' ----
        
        Case 18: mnSalir_Click    'Salir
            
        Case btnPrimero To btnPrimero + 3 'Flechas Desplazamiento
            Desplazamiento (Button.index - btnPrimero)
    End Select
End Sub


Private Sub PonerOpcionesMenu()
Dim J As Byte

    PonerOpcionesMenuGeneral Me
      
    J = Val(Me.mnGenAlbaran.HelpContextID)
    If J < vUsu.Nivel Then Me.mnGenAlbaran.Enabled = False
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim Cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, Cerrar
    If Cerrar Then Unload Me
End Sub
    
    
Private Function InsertarLinea() As Boolean
'Inserta un registro en la tabla de lineas de Pedido: slipre
Dim SQL As String
Dim ImpReciclado As Single
Dim numlinea As String, vWhere As String
Dim ImporSIGAUS As String
    On Error GoTo EInsertarLinea

    InsertarLinea = False
    SQL = ""

    If DatosOkLinea() Then 'Lineas de Pedidos
         vWhere = Replace(ObtenerWhereCP, NombreTabla, NomTablaLineas)
         If LineaIntercalar = 0 Then
            'INSERCION NORMAL
            numlinea = SugerirCodigoSiguienteStr(NomTablaLineas, "numlinea", vWhere)
         
    
        Else
            

            SQL = "UPDATE " & NomTablaLineas & " SET numlinea=numlinea + 1 WHERE " & vWhere & " and numlinea >= " & LineaIntercalar
            SQL = SQL & " order by numlinea desc " 'Para que empieza por las ultimas
            conn.Execute SQL
            numlinea = LineaIntercalar
        End If
 
       
        'Construir la sentencia SQL
'        vWhere = ObtenerWhereCP
        SQL = "INSERT INTO " & NomTablaLineas
        SQL = SQL & "(numpedcl,numlinea, codalmac, codartic, nomartic, ampliaci, cantidad, servidas, numbultos,precioar, dtoline1, dtoline2, importel, origpre,numlote,codccost,codtipor,codcapit) "
        SQL = SQL & "VALUES (" & Val(Text1(0).Text) & ", " & numlinea & ", " & Val(txtaux(0).Text) & ","
        SQL = SQL & DBSet(txtaux(1).Text, "T") & ", " & DBSet(txtaux(2).Text, "T") & ", " & DBSet(Text2(16).Text, "T") & ", "
        SQL = SQL & DBSet(txtaux(3).Text, "N") & ", 0," & DBSet(txtaux(9).Text, "N") & ", "
        SQL = SQL & DBSet(txtaux(4).Text, "N") & ", " & DBSet(txtaux(6).Text, "N") & ", "
        SQL = SQL & DBSet(txtaux(7).Text, "N") & ", " 'Dto2
        SQL = SQL & DBSet(txtaux(8).Text, "N") & ", "
        '- origpre, numlote
        SQL = SQL & DBSet(txtaux(5).Text, "T") & "," & DBSet(txtaux(10).Text, "T", "S") & ","
        '- codccost
        SQL = SQL & DBSet(UCase(txtaux(11).Text), "T", "S") & ","
        'SAIL
        SQL = SQL & DBSet(UCase(txtaux(12).Text), "T", "S") & ","
        SQL = SQL & DBSet(UCase(txtaux(13).Text), "N", "S") & ")"
        
    End If
    
    If SQL <> "" Then
        conn.Execute SQL
        InsertarLinea = True
        
        ' ---- [15/09/2009] (LAURA)
'        ElArticulo = txtAux(1).Text
'        DescuentosCantidad ElArticulo
        ' ----
        
        If ClienteConTasaReciclado Then
            If ArticuloConTasaReciclado(txtaux(1).Text, ImpReciclado) Then
                'Insertamos la linea del reciclado
                ImporSIGAUS = "preciove"
                vWhere = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", vParamAplic.ArtReciclado, "T", ImporSIGAUS)
                If vParamAplic.NumeroInstalacion = vbTaxco Then
                    If ImporSIGAUS = "" Then ImporSIGAUS = "0"
                    ImpReciclado = CCur(ImporSIGAUS)
                End If
                SQL = "INSERT INTO " & NomTablaLineas
                SQL = SQL & "(numpedcl,numlinea, codalmac, codartic, nomartic, ampliaci, cantidad, servidas, precioar,"
                SQL = SQL & "dtoline1, dtoline2, importel, origpre) "
                SQL = SQL & "VALUES (" & Val(Text1(0).Text) & ", " & numlinea + 1 & ", " & Val(txtaux(0).Text) & ","
                SQL = SQL & DBSet(vParamAplic.ArtReciclado, "T") & "," & DBSet(vWhere, "T") & ", Null, "
                SQL = SQL & DBSet(txtaux(3).Text, "N") & ", 0," 'Cantidad. La misma
                SQL = SQL & DBSet(ImpReciclado, "N") & ",0,0,"
                'Importe linea
                ImpReciclado = ImporteFormateado(txtaux(3).Text) * ImpReciclado
                SQL = SQL & DBSet(ImpReciclado, "N") & ", 'A')"
                conn.Execute SQL
                    
                
            End If
        End If
        
    End If
    Exit Function
    
EInsertarLinea:
    MuestraError Err.Number, "Insertar Lineas Pedido" & vbCrLf & Err.Description
End Function


Private Function ModificarLinea() As Boolean
'Modifica un registro en la tabla de lineas de Pedido: sliped
Dim SQL As String

    On Error GoTo EModificarLinea

    ModificarLinea = False
    SQL = ""
    
    If DatosOkLinea() Then
        'Creamos la sentencia SQL
        SQL = "UPDATE " & NomTablaLineas & " Set codalmac = " & txtaux(0).Text & ", codartic=" & DBSet(txtaux(1).Text, "T") & ", "
        SQL = SQL & "nomartic=" & DBSet(txtaux(2).Text, "T") & ", ampliaci=" & DBSet(Text2(16).Text, "T") & ", "
        SQL = SQL & " cantidad = " & DBSet(txtaux(3).Text, "N") & ", "
        SQL = SQL & " numbultos = " & DBSet(txtaux(9).Text, "N") & ", "
        SQL = SQL & " precioar = " & DBSet(txtaux(4).Text, "N") & ", "
        SQL = SQL & "dtoline1= " & DBSet(txtaux(6).Text, "N") & ", dtoline2= " & DBSet(txtaux(7).Text, "N") & ", "
        SQL = SQL & "importel= " & DBSet(txtaux(8).Text, "N") & ","
        SQL = SQL & "origpre=" & DBSet(txtaux(5).Text, "T") & ","
        SQL = SQL & "numlote=" & DBSet(txtaux(10).Text, "T", "S") & ","
        SQL = SQL & "codccost=" & DBSet(UCase(txtaux(11).Text), "T", "S") & ","
        'SAIL
        SQL = SQL & "codcapit=" & DBSet(txtaux(13).Text, "T", "S") & ","
        SQL = SQL & "codtipor=" & DBSet(UCase(txtaux(12).Text), "T", "S")
        SQL = SQL & " WHERE " & Replace(ObtenerWhereCP, NombreTabla, NomTablaLineas) & " AND numlinea=" & Data2.Recordset!numlinea
    End If
    
    If SQL <> "" Then
        conn.Execute SQL
        ModificarLinea = True
        
        
        ' ---- [15/09/2009] (LAURA)
'        If txtAux(1).Text = Data2.Recordset!codArtic Then
'            ElArticulo = Data2.Recordset!codArtic
'        Else
'            'Son distintos. Que recalcule todo
'            ElArticulo = ""
'        End If
'        DescuentosCantidad ElArticulo
        ' ----
        
    End If
    Exit Function
    
EModificarLinea:
    MuestraError Err.Number, "Modificar Lineas Pedido" & vbCrLf & Err.Description
End Function


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
        cmdRegresar.Cancel = True
    Else
        cmdCancelar.Cancel = True
    End If

    
    'Habilitar las opciones correctas del menu según Modo
    PonerModoOpcionesMenu (Modo)
    PonerOpcionesMenu 'Habilitar las opciones correctas del menu según Nivel de Acceso
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub CargaGrid(ByRef vDataGrid As DataGrid, ByRef vData As Adodc, enlaza As Boolean, Optional conServidas As Boolean)
'IN: enlaza= si carga el grid con valores de la tabla o lo muestra vacio si no enlaza
'    conServidas=si enlaza, se muestra la columna de servidas solo cuando se va a generar el Albaran no completo
Dim b As Boolean
Dim SQL As String

    On Error GoTo ECargaGrid

    b = DataGrid1.Enabled
    
    SQL = MontaSQLCarga(enlaza, conServidas)
    CargaGridGnral vDataGrid, vData, SQL, primeravez
    
'    If PrimeraVez Or conServidas Then
    If conServidas Then
        vDataGrid.ClearFields
        vDataGrid.ReBind
        vDataGrid.Refresh
    End If
    
    CargaGrid2 vDataGrid, vData, conServidas
    vDataGrid.ScrollBars = dbgAutomatic
    
    b = (Modo = 5) And (ModificaLineas = 1 Or ModificaLineas = 2) '5:Modo Mto Lineas (Insertando o Modificando linea)
    vDataGrid.Enabled = Not b
    primeravez = False
    gridCargado = True
    Exit Sub
    
ECargaGrid:
    MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub


Private Sub CargaGrid2(ByRef vDataGrid As DataGrid, ByRef vData As Adodc, Optional conServidas As Boolean)
Dim I As Byte

    On Error GoTo ECargaGrid

    vData.Refresh
    vDataGrid.Columns(0).visible = False
    vDataGrid.Columns(1).visible = False

    Select Case vDataGrid.Name
        Case "DataGrid1" 'Cod. Almacen
            vDataGrid.Columns(2).Caption = "Alm."
            If conServidas Then
                vDataGrid.Columns(2).Width = 450
            Else
                vDataGrid.Columns(2).Width = 500
            End If
            vDataGrid.Columns(2).NumberFormat = "000"
                
            vDataGrid.Columns(3).Caption = "Articulo"
            If conServidas Then
                vDataGrid.Columns(3).Width = 1500
            Else
                vDataGrid.Columns(3).Width = 1600
            End If
                
            vDataGrid.Columns(4).Caption = "Desc. Artículo"
            If conServidas Then
                vDataGrid.Columns(4).Width = 3000
            Else
                vDataGrid.Columns(4).Width = 3100
            End If
                
            vDataGrid.Columns(5).visible = False
            
            vDataGrid.Columns(6).Caption = "Cantidad"
            vDataGrid.Columns(6).Width = 860
            vDataGrid.Columns(6).Alignment = dbgRight
            vDataGrid.Columns(6).NumberFormat = FormatoImporte
            
            If conServidas Then
                'Cargar el grid con la columna de cantidad servida
                vDataGrid.Columns(7).Caption = "Servidas"
                vDataGrid.Columns(7).Width = 800
                vDataGrid.Columns(7).Alignment = dbgRight
                vDataGrid.Columns(7).NumberFormat = FormatoImporte
                I = 8
            Else
                I = 7
            End If
                
            
            vDataGrid.Columns(I).Caption = "Bultos"
            vDataGrid.Columns(I).Width = 620
            vDataGrid.Columns(I).Alignment = dbgRight

'                vDataGrid.Columns(i).NumberFormat = FormatoPrecio
                
            I = I + 1
            vDataGrid.Columns(I).Caption = "Precio"
            vDataGrid.Columns(I).Width = 950
            vDataGrid.Columns(I).Alignment = dbgRight
            vDataGrid.Columns(I).NumberFormat = FormatoPrecio
            
            I = I + 1
            vDataGrid.Columns(I).Caption = "OP"
            vDataGrid.Columns(I).Width = 340
            vDataGrid.Columns(I).Alignment = dbgCenter
                
            I = I + 1
            vDataGrid.Columns(I).Caption = "Dto.1"
'            If conServidas Then
                vDataGrid.Columns(I).Width = 540
'            Else
'                vDataGrid.Columns(i).Width = 560
'            End If
            vDataGrid.Columns(I).Alignment = dbgRight
            vDataGrid.Columns(I).NumberFormat = FormatoDescuento
                
            I = I + 1
            vDataGrid.Columns(I).Caption = "Dto.2"
'            If conServidas Then
                vDataGrid.Columns(I).Width = 550
'            Else
'                vDataGrid.Columns(i).Width = 560
'            End If
            vDataGrid.Columns(I).Alignment = dbgRight
            vDataGrid.Columns(I).NumberFormat = FormatoDescuento
            
            I = I + 1
            vDataGrid.Columns(I).Caption = "Importe"
            If conServidas Then
                vDataGrid.Columns(I).Width = 1050
            ElseIf vEmpresa.TieneAnalitica Then
                vDataGrid.Columns(I).Width = 1100
            Else
                vDataGrid.Columns(I).Width = 1250
            End If
            vDataGrid.Columns(I).Alignment = dbgRight
            vDataGrid.Columns(I).NumberFormat = FormatoImporte
            
            
            If vEmpresa.TieneAnalitica And conServidas = False Then
                I = I + 1
'                vDataGrid.Columns(I).Caption = "CCost"
'                vDataGrid.Columns(I).Width = 640
                vDataGrid.Columns(I).visible = False 'centro de coste
            Else
                I = I + 1
                vDataGrid.Columns(I).visible = False 'centro de coste
            End If
            
            
            'YA NO HAY LOTE
'            i = i + 1
'            vDataGrid.Columns(i).Caption = "Nº Lote"
'            If conServidas Then
'                vDataGrid.Columns(i).Width = 1220
'            Else
'                vDataGrid.Columns(i).Width = 1280
'            End If
            I = I + 1
            vDataGrid.Columns(I).visible = False 'codtipor
            I = I + 1
            vDataGrid.Columns(I).visible = False 'codcapit
            I = I + 1
            vDataGrid.Columns(I).visible = False 'ampliaci
            
'            vDataGrid.Columns(i).Alignment = dbgRight
'            vDataGrid.Columns(i).NumberFormat = FormatoImporte
            
    End Select

    For I = 0 To vDataGrid.Columns.Count - 1
        vDataGrid.Columns(I).Locked = True
        vDataGrid.Columns(I).AllowSizing = False
    Next I
    vDataGrid.HoldFields
    Exit Sub
    
ECargaGrid:
    MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub


'Esta funcion sustituye a LlamaLineas
Private Sub CargaTxtAux2(visible As Boolean, limpiar As Boolean)
'IN: visible: si es true ponerlos visibles en la posición adecuada
'    limpiar: si es true vaciar los txtAux
Dim alto As Single
Dim I As Byte

    On Error Resume Next

    If Not visible Then
        'Fijamos el alto (ponerlo en la parte inferior del form)
        For I = 0 To 9 'TextBox
            txtaux(I).Top = 290
            txtaux(I).visible = visible
        Next I
        cmdAux(0).visible = visible
        cmdAux(1).visible = visible
        
       ' cmdAux(2).visible = visible
    Else
        If limpiar Then 'Vaciar los textBox (Vamos a Insertar)
            DeseleccionaGrid DataGrid1
            For I = 0 To txtaux.Count - 1
                txtaux(I).Text = ""
                BloquearTxt txtaux(I), False
            Next I
            
      
            'HAY QUE LIMPIAR LOS CC
            For I = 11 To 13
                txtAux2(I).Text = ""
            Next
        Else 'Vamos a modificar
            For I = 0 To txtaux.Count - 1
                If I < 3 Then
                    txtaux(I).Text = DataGrid1.Columns(I + 2).Text
                ElseIf I = 3 Then
                    txtaux(I).Text = DataGrid1.Columns(I + 3).Text
                ElseIf I >= 4 And I < 9 Then
                    txtaux(I).Text = DataGrid1.Columns(I + 4).Text
                ElseIf I = 9 Then
                    txtaux(I).Text = DataGrid1.Columns(7).Text
                ElseIf I = 10 Then
                    txtaux(I).Text = DataGrid1.Columns(I + 4).Text
                ElseIf I = 11 Then
                    txtaux(I).Text = DataGrid1.Columns(I + 2).Text
                End If
                txtaux(I).Locked = False
            Next I
        End If
               
        'El Campo de Origen del precio se actualiza por programa al modificar el precio
        BloquearTxt txtaux(5), True
        'El campo Importe es calculado y lo bloqueamos.
        BloquearTxt txtaux(8), True
        'El campo Nº Bultos es calculado y lo bloqueamos.
        BloquearTxt txtaux(9), True

        ' ---- [20/10/2009] [LAURA] : añadir centro de coste
        If txtaux(11).visible Then
            BloquearTxt txtaux(11), Not (vEmpresa.TieneAnalitica And vParamAplic.ModoAnalitica = 2)
            Me.cmdAux(2).Enabled = Not txtaux(11).Locked
        End If
        'Me.cmdAux(2).visible = Me.cmdAux(2).Enabled
        ' ----





        'Fijamos altura(Height) y posición Top
        '-------------------------------
        alto = ObtenerAlto(DataGrid1, 10)
        
        For I = 0 To 9
            txtaux(I).Top = alto
            txtaux(I).Height = DataGrid1.RowHeight
        Next I
        
        cmdAux(0).Top = alto
        cmdAux(1).Top = alto
        cmdAux(2).Top = alto
        cmdAux(0).Height = DataGrid1.RowHeight
        cmdAux(1).Height = DataGrid1.RowHeight
        'cmdAux(2).Height = DataGrid1.RowHeight
        
        'Fijamos anchura y posicion Left
        '--------------------------------
        'Cod. Almac
        txtaux(0).Left = DataGrid1.Left + 330
        txtaux(0).Width = DataGrid1.Columns(2).Width - 160
        cmdAux(0).Left = txtaux(0).Left + txtaux(0).Width - 40
        'Cod Artic
        txtaux(1).Left = cmdAux(0).Left + cmdAux(0).Width + 20
        txtaux(1).Width = DataGrid1.Columns(3).Width - 160
        cmdAux(1).Left = txtaux(1).Left + txtaux(1).Width - 50
        'Nom Artic
        txtaux(2).Left = cmdAux(1).Left + cmdAux(1).Width
        txtaux(2).Width = DataGrid1.Columns(4).Width - 10
        'Cantidad
        txtaux(3).Left = txtaux(2).Left + txtaux(2).Width + 10
        txtaux(3).Width = DataGrid1.Columns(6).Width - 10
        'Bultos
        txtaux(9).Left = txtaux(3).Left + txtaux(3).Width + 10
        txtaux(9).Width = DataGrid1.Columns(7).Width - 10
        'Precio
        txtaux(4).Left = txtaux(9).Left + txtaux(9).Width + 10
        txtaux(4).Width = DataGrid1.Columns(8).Width - 10
        
        'OP,Dto1, Dto2, Importe
        For I = 5 To 8
            txtaux(I).Left = txtaux(I - 1).Left + txtaux(I - 1).Width + 10
            txtaux(I).Width = DataGrid1.Columns(I + 4).Width - 10
        Next I
        
        
        'Los ponemos Visibles o No
        '--------------------------
        For I = 0 To 9
           txtaux(I).visible = visible
        Next I
        If vUsu.Nivel2 = 2 And visible Then
            For I = 4 To 8
                BloquearTxt txtaux(I), True
            Next
        End If
        
        
        
        
        cmdAux(0).visible = visible
        cmdAux(1).visible = visible
    End If

    For I = 12 To 13
        BloquearTxt txtaux(I), Not visible
        Me.imgBuscar2(I).visible = visible
    Next
   
    BloquearTxt Text2(16), Not visible
    
    If vEmpresa.TieneAnalitica Then
        BloquearTxt txtaux(11), Not visible
        Me.imgBuscar2(11).visible = visible
    End If
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub CargaTxtAuxServidas(visible As Boolean, limpiar As Boolean)
'IN: visible: si es true ponerlos visibles en la posición adecuada
'    limpiar: si es true vaciar los txtAux
'Carga el TxtAux(3) con el campo servidas de la tabla sliped
Dim alto As Single
Dim I As Byte, i2 As Byte

    On Error Resume Next

    I = 3
    i2 = 9
    If Not visible Then
        'Fijamos el alto (ponerlo en la parte inferior del form)
        txtaux(I).Top = 290
        txtaux(I).visible = visible
        txtaux(I).BackColor = vbWhite
        
        txtaux(i2).Top = 290
        txtaux(i2).visible = visible
        txtaux(i2).BackColor = vbWhite
        
    Else
        If limpiar Then 'Vaciar los textBox (Vamos a Insertar)
            DeseleccionaGrid DataGrid1
            txtaux(I).Text = ""
            BloquearTxt txtaux(I), False
            txtaux(I).BackColor = &H80000013
            
            txtaux(i2).Text = ""
            BloquearTxt txtaux(i2), False
            txtaux(i2).BackColor = &H80000013
        End If
      
        'Fijamos altura(Height) y posición Top
        '-------------------------------------
        If DataGrid1.Row < 0 Then
            alto = DataGrid1.Top + 230
        Else
            alto = DataGrid1.Top + DataGrid1.RowTop(DataGrid1.Row) + 10
        End If
        
        txtaux(I).Top = alto
        txtaux(I).Height = DataGrid1.RowHeight
        
        txtaux(i2).Top = alto
        txtaux(i2).Height = DataGrid1.RowHeight
        
        'Fijamos anchura y posicion Left
        '--------------------------------
        'Cantidad servida
        alto = DataGrid1.Left + 330 + DataGrid1.Columns(2).Width + DataGrid1.Columns(3).Width
        alto = alto + DataGrid1.Columns(4).Width + DataGrid1.Columns(6).Width
        txtaux(I).Left = alto + 10
        txtaux(I).Width = DataGrid1.Columns(7).Width - 30
        
        txtaux(i2).Left = alto + 10 + DataGrid1.Columns(7).Width
        txtaux(i2).Width = DataGrid1.Columns(8).Width - 30
        
        'Los ponemos Visibles o No
        '--------------------------
        txtaux(I).visible = visible
        txtaux(i2).visible = visible
        If kCampo = 3 Or kCampo = 9 Then
            PonerFoco txtaux(kCampo)
        Else
            PonerFoco txtaux(I)
        End If
    End If
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub TxtAux_Change(index As Integer)
    If index = 4 And ModificaLineas = 2 Then 'Precio y Modo Modificar Lineas
        txtaux(5).Text = "M"
    End If
End Sub

Private Sub txtAux_GotFocus(index As Integer)
Dim cadkey As Integer

    If Modo >= 5 Then cadkey = ObtenerCadKey(kCampo, index)
    kCampo = index
    
    If index = 16 Then
        'Campo observaciones. NO, repito NO, se selecciona todo
        If txtaux(index).Text <> "" Then
            txtaux(index).Text = txtaux(index).Text & " "
            txtaux(index).SelStart = Len(txtaux(index).Text)
        End If
    Else
        ConseguirFocoLin txtaux(index), cadkey
    End If
    LabelAyudatxtAux index, lblF
    
End Sub





Private Sub TxtAux_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.

    If Modo <> 6 Then 'Modo6: Pasar de Pedido a Albaran
    
        ' ---- [02/11/2009] [LAURA] : al pulsar F2 para abrir articulos en la solapa Documentos|Pedidos
        If KeyCode = 113 Then
           AccionesF2 index
        ' ----
    
        ElseIf KeyCode = 43 Or KeyCode = 107 Or KeyCode = 187 Then
            If index < 2 Or index = 11 Then  'Para los que tienen busqueda
                If Modo = 5 And ModificaLineas = 1 Then
                    If txtaux(index).Text = "" Then
                        PulsadoMas2 = True
                        KeyCode = 0
                
                        PulsarTeclaMas False, index
                    End If
                End If
             End If
        
    
    
        ElseIf Not (index = 0 And KeyCode = 38) Then
            KEYdown KeyCode
        End If
        
    Else 'Modo lineas
        Select Case KeyCode
            Case 38 'Desplazamieto Fecha Hacia Arriba
                    If DataGrid1.Row > 0 Then
                        DataGrid1.Row = DataGrid1.Row - 1
                        CargaTxtAuxServidas True, True
                    Else
                        PonerFoco txtaux(3)
                    End If
                    txtaux(3).Text = Data2.Recordset!servidas
                    txtaux(9).Text = Data2.Recordset!bultosser
                    ConseguirFocoLin txtaux(3)

            Case 40 'Desplazamiento Flecha Hacia Abajo
'                    If DataGrid1.Row < Data2.Recordset.RecordCount - 1 Then
                    PonerServidas index
'                    MoverSigRegisros
'                    If Data2.Recordset.AbsolutePosition <= Data2.Recordset.RecordCount - 1 Then
'                        DataGrid1.Row = DataGrid1.Row + 1
'                        CargaTxtAuxServidas True, True
'                    Else
'                        PonerFocoBtn Me.cmdAceptar
'                    End If
'                    txtAux(3).Text = Data2.Recordset!servidas
'                    ConseguirFocoLin txtAux(3)
        End Select
    End If
End Sub


Private Sub AccionesF2(index As Integer)
    If index = 3 Then
        AbrirForm_Articulos txtaux(1).Text
    Else
        If index = 4 Then
            AbrirConsultaPrecio2 Text1(4).Text, txtaux(1).Text, Text1(1).Text, ""
        End If
            
    End If
End Sub

Private Sub MoverSigRegistro()
    On Error GoTo EMover
    
    If Data2.Recordset.EOF Then Exit Sub
    If Data2.Recordset.AbsolutePosition <= Data2.Recordset.RecordCount - 1 Then
        DataGrid1.Row = DataGrid1.Row + 1
        CargaTxtAuxServidas True, True
    Else
        PonerFocoBtn Me.cmdAceptar
    End If
    txtaux(3).Text = Data2.Recordset!servidas
    txtaux(9).Text = Data2.Recordset!bultosser
    ConseguirFocoLin txtaux(3)
    Exit Sub
    
EMover:
    MuestraError Err.Description, "Mover registro.", Err.Description
End Sub


Private Sub txtAux_KeyPress(index As Integer, KeyAscii As Integer)
    
    If Modo <> 6 Then
        KEYpress KeyAscii
    Else 'Modo 6: Pasar el Pedido a Albaran
        If KeyAscii = 13 Then 'ENTER
            If index = 3 Or index = 9 Then
                
                PonerServidas index
            End If
'            ConseguirFoco txtAux(3), Modo
        End If
    End If
End Sub


Private Sub txtAux_LostFocus(index As Integer)
Dim devuelve As String, cadMen As String
'Dim codTarif As String
Dim CPrecioFact As CPreciosFact
Dim vCStock As CStock
Dim NumCajas As Integer, RestoUnid As Integer
Dim OrigP As String 'De donde viene el precio
Dim b As Boolean
Dim codCC As String
    
    If PulsadoMas2 Then
        'Para que cuando pulse el mas abra el form
        PulsadoMas2 = False
        txtaux(index).Text = ""
        Exit Sub
    End If


    If Not PerderFocoGnralLineas(txtaux(index), ModificaLineas) Then Exit Sub
    
    Select Case index
        Case 0 'Cod Almacen
            'Comprobar que existe el almacen
            devuelve = PonerAlmacen(txtaux(index).Text)
            txtaux(index).Text = devuelve
            If devuelve = "" Then PonerFoco txtaux(index)

        Case 1 'Cod. Articulo
            If txtaux(1).Text = "" Then 'Cod Artic
                txtaux(2).Text = "" 'Nom Artic
                Exit Sub
            End If
            If txtaux(0).Text = "" Then 'Cod Almacen
                MsgBox "Debe seleccionar un almacen.", vbInformation
                PonerFoco txtaux(0)
                Exit Sub
            End If

            devuelve = ""
            If ModificaLineas = 2 Then
                If Not Data2.Recordset.EOF Then devuelve = Data2.Recordset!codArtic
            End If
            
            If PonerArticulo(txtaux(1), txtaux(2), txtaux(0).Text, CodTipoMov, ModificaLineas, devuelve, , codCC) Then
                
                If devuelve <> txtaux(1).Text Then
                    'ha cambiado el articulo
                    Me.txtaux(3).Text = ""
                    Me.txtaux(4).Text = ""
                    Me.txtaux(5).Text = ""
                    Me.txtaux(6).Text = ""
                    Me.txtaux(7).Text = ""
                    Me.txtaux(9).Text = ""
                End If
                
            
            
                '---- [20/10/2009] [LAURA] : añadir centro de coste
                If Not vEmpresa.TieneAnalitica Then
                    txtaux(11).Text = ""
                ElseIf vParamAplic.ModoAnalitica = 1 Then 'Por familia
                    txtaux(11).Text = codCC
                    Me.txtAux2(11).Text = PonerNombreCCoste(Me.txtaux(11))
                End If
                '----
            
            
                b = (Me.ActiveControl.Name = "txtAux")
                If b Then b = (Me.ActiveControl.index = 0)
                
                If Not b Then
'                    If txtAux(2).Locked Then PonerFoco txtAux(3)
                Else
                    PonerFoco txtaux(0)
                End If
                
                
                If Text2(16).Text = "" Then _
                    Text2(16).Text = DevuelveDesdeBD(conAri, "txtauxdocumento", "sartic", "codartic", txtaux(1).Text, "T")

                
                
            Else
                PonerFoco txtaux(index)
            End If
            
        Case 2 'desc Articulo
            If txtaux(index).Locked = False Then txtaux(index).Text = UCase(txtaux(index).Text)
            
        Case 3 'CANTIDAD
            If PonerFormatoDecimal(txtaux(index), 1) Then  'Tipo 1: Decimal(12,2)
                If Modo = 5 Then 'Mantenimiento lineas
                    'Comprobar si hay suficiente stock
                    Set vCStock = New CStock
                    If Not InicializarCStock(vCStock, "S") Then Exit Sub
                    If vCStock.MueveStock Then
                        If Not vCStock.MoverStock(False, False) Then
                            Set vCStock = Nothing
                            Exit Sub
                        End If
                    End If
                    
                    
                    b = False
                    If Modo = 5 Then
                        'Comprobar si el articulo se vende por cajas antes de entrar a la función
                        devuelve = DevuelveDesdeBDNew(conAri, "sartic", "unicajas", "codartic", txtaux(1).Text, "T")
                    
                        If devuelve <> "" Then
                            '- obtener el nº bultos: cantidad/unids.caja
                            txtaux(9).Text = CalcularNumBultos2(CCur(txtaux(3).Text), CInt(devuelve))
                        End If
                    
                        If ModificaLineas = 1 Then 'insertar linea
                            b = True
                        ElseIf ModificaLineas = 2 Then 'modificar linea
                            If Data2.Recordset!codArtic <> txtaux(1).Text Then
                                b = True
                            Else
                                If CStr(DBLet(Data2.Recordset!origpre, "T")) <> "M" Then b = True
                            End If
                        End If
                    End If
                    
                    If b Then 'Modo Insertar en Mto Lineas
                        'Obtener el precio correspondiente y los descuentos
                        'Comprobar si el articulo se vende por cajas antes de entrar a la función
'                        devuelve = DevuelveDesdeBDNew(conAri, "sartic", "unicajas", "codartic", txtAux(1).Text, "T")
                        If devuelve <> "" Then
'                            '- obtener el nº bultos: cantidad/unids.caja
'                            txtAux(9).Text = CalcularNumBultos(CCur(txtAux(3).Text), CInt(devuelve))
                        
                            Set CPrecioFact = New CPreciosFact
                            
                            If vParamAplic.CajasCompletas Then
                                NumCajas = CPrecioFact.ObtenerNumCajas(vCStock.cantidad, devuelve)
                                RestoUnid = CInt(vCStock.cantidad) - NumCajas * CInt(devuelve)
                            Else
                                NumCajas = 0
                                RestoUnid = 0
                            End If
                            'Obtenemos la Tarifa del Cliente
                            'codTarif = DevuelveDesdeBDNew(conAri, "sclien", "codtarif", "codclien", Text1(4).Text, "N")
                            'CPrecioFact.CodigoLista = codTarif
                            CPrecioFact.CodigoArtic = vCStock.codArtic
                            CPrecioFact.CodigoClien = Text1(4).Text
                            CPrecioFact.FijarTarifaActividad
                            
                            PorCaja = (NumCajas > 0)
                            Precio = CPrecioFact.ObtenerPrecio(PorCaja, Text1(1).Text, OrigP, "")
                            'Si PorCaja vuelve de ObtenerPrecio a false se calcula con precio unidad aunque NumCajas>0
                            'Ya que a regresado con pvp del Articulo
                            If PorCaja And NumCajas > 0 And RestoUnid > 0 Then
                                cadMen = "El Artículo puede venderse por Cajas (" & devuelve & "uds. por Caja)." & vbCrLf
                                cadMen = cadMen & vbCrLf & "Inserte dos Lineas:   "
                                cadMen = cadMen & vbCrLf & "   Linea 1:  " & NumCajas * CInt(devuelve) & " uds a Precio Caja"
                                cadMen = cadMen & vbCrLf & "   Linea 2:  " & CInt(vCStock.cantidad) - NumCajas * CInt(devuelve) & " uds a Precio Unidad"
                                MsgBox cadMen, vbInformation
                                PonerFoco txtaux(index)
                            Else
                                If (txtaux(4).Text = "") Or (txtaux(4).Text <> "" And ModificaLineas = 2 And b) Then
                                    txtaux(4).Text = Precio
                                    txtaux(5).Text = OrigP 'De donde viene el precio
                                End If
                                PonerFormatoDecimal txtaux(4), 2
                                If txtaux(6).Text = "" Then txtaux(6).Text = CPrecioFact.Descuento1
                                PonerFormatoDecimal txtaux(6), 4
                                If txtaux(7).Text = "" Then txtaux(7).Text = CPrecioFact.Descuento2
                                PonerFormatoDecimal txtaux(7), 4
                            End If
    '                        ConseguirFoco txtAux(Index + 1), Modo
                            Set CPrecioFact = Nothing
                        End If
                    End If
                    If vUsu.Nivel2 <> 2 Then
                        PonerFoco Text2(16)
                    Else
                        ConseguirFocoLin txtaux(4)
                    End If
    '            End If
                Set vCStock = Nothing
            End If
        End If
            
        Case 4 'PRECIO
             If txtaux(index).Text <> "" Then
                PonerFormatoDecimal txtaux(index), 2 'Tipo 2: Decimal(10,4)
                If ModificaLineas = 1 Then
                    'Precio=valor devuelto por la funcion de precios
                    If CSng(txtaux(index).Text) <> CSng(ComprobarCero(Precio)) Then txtaux(5).Text = "M"
                End If
            End If
            
        Case 6, 7 'Descuentos
            PonerFormatoDecimal txtaux(index), 4 'Tipo 4: Decimal(4,2)
        Case 8 'Importe Linea
            PonerFormatoDecimal txtaux(index), 1 'Tipo 3: Decimal(12,2)
        Case 9
            
        Case 11 'COD. CENTRO COSTE
            ' ---- [20/10/2009] [LAURA]: añadir centro de coste a la linea
            If txtaux(index).Text = "" Then
                 txtAux2(index).Text = ""
            ElseIf vEmpresa.TieneAnalitica Then
                'centro de coste
                ' ---- [20/10/2009] [LAURA]: añadir campo centro de coste familia
                Me.txtAux2(index).Text = PonerNombreCCoste(Me.txtaux(index))
                
            End If
            
   
        Case 12, 13
            PonerDatosNuevosLineaAlbaran True, index
    End Select
    
    If Modo = 5 Then 'Modo Lineas
         If (index = 3 Or index = 4 Or index = 6 Or index = 7) Then 'Cant., Precio, dto1, dto2
            If txtaux(1).Text = "" Then Exit Sub 'Cod artic
            txtaux(8).Text = CalcularImporte(txtaux(3).Text, txtaux(4).Text, txtaux(6).Text, txtaux(7).Text, vParamAplic.TipoDtos)
            PonerFormatoDecimal txtaux(8), 1
        End If
    End If
End Sub


Private Sub BotonMtoLineas(numTab As Integer, cad As String)
        Me.SSTab1.Tab = numTab
        TituloLinea = cad
        ModificaLineas = 0
        LineaIntercalar = 0
        If vParamAplic.ArtReciclado <> "" Then
            If vParamAplic.NumeroInstalacion = vbTaxco Then
                ClienteConTasaReciclado = True
            Else
                ClienteConTasaReciclado = Val(DevuelveDesdeBD(conAri, "tasareciclado", "sclien", "codclien", Text1(4).Text)) = 1
            End If
        Else
            ClienteConTasaReciclado = False
        End If
                
        If vParamAplic.TipoPortes = 1 Then KilosAnteriores = SumaKilosLineas
        
        PonerModo 5
        PonerBotonCabecera True
End Sub


Private Function Eliminar() As Boolean
Dim b As Boolean
Dim SQL As String
Dim MenError As String
Dim vTipoMov As CTiposMov

    On Error GoTo FinEliminar

        conn.BeginTrans
        SQL = ObtenerWhereCP
        
        'CadenaSQL: datos introducidos en el form de eliminacion
        b = ActualizarElTraspaso(MenError, SQL, CodTipoMov, CadenaSQL)

        If b Then
            'Devolvemos contador, si no estamos actualizando
            Set vTipoMov = New CTiposMov
            b = vTipoMov.DevolverContador(CodTipoMov, Data1.Recordset.Fields(0).Value)
            Set vTipoMov = Nothing
        Else
            MsgBox MenError, vbExclamation
        End If
FinEliminar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar Pedido" & vbCrLf & MenError, Err.Description
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
'Pone los Grids sin datos, apuntando a ningún registro
On Error Resume Next
    CargaGrid DataGrid1, Data2, False
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub PosicionarData()
'Despues de hacer refresh del Data, volver a situar el Data en el registro que estaba
Dim Indicador As String
Dim vWhere As String

    If Not Data1.Recordset.EOF Then
        'Hay datos en el Data1 bien porque se ha hecho VerTodos o una Busqueda
         vWhere = Replace(ObtenerWhereCP, NombreTabla & ".", "")
         If SituarData(Data1, vWhere, Indicador) Then
             PonerModo 2
             PonerCampos
             lblIndicador.Caption = Indicador
        Else
             LimpiarCampos
             'Poner los grid sin apuntar a nada
             LimpiarDataGrids
             PonerModo 0
         End If
    Else
        'El Data esta vacio, desde el modo de inicio se pulsa Insertar
        CadenaConsulta = "Select * from " & NombreTabla & " WHERE " & ObtenerWhereCP & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Sub PosicionarDataTrasEliminar()
'Despues Eliminar y hacer refresh del Data, situar el Data en el registro siguiente
    If SituarDataTrasEliminar(Data1, NumRegElim) Then
        PonerCampos
    Else
        LimpiarCampos
        LimpiarDataGrids
        PonerModo 0
    End If
End Sub


Private Function ObtenerWhereCP() As String
'Obtiene la where de la Clave Primaria de la tabla de Cabecera: scaped
Dim SQL As String

    On Error Resume Next
    
    SQL = NombreTabla & ".numpedcl= " & Val(Text1(0).Text)
    If EsHistorico Then SQL = SQL & " AND " & NomTablaLineas & ".fecpedcl=" & DBSet(Text1(1).Text, "F")
    ObtenerWhereCP = SQL
    
    If Err.Number <> 0 Then Err.Clear
End Function


Private Function MontaSQLCarga(enlaza As Boolean, Optional conServidas As Boolean) As String
'--------------------------------------------------------------------
' MontaSQlCarga:
'   Basándose en la información proporcionada por el vector de campos
'   crea un SQl para ejecutar una consulta sobre la base de datos que los
'   devuelva.
' Si ENLAZA -> Enlaza con el data2
'           -> Si no lo cargamos sin enlazar a ningun campo
'--------------------------------------------------------------------
Dim SQL As String
    
    SQL = "SELECT numpedcl, numlinea, codalmac, codartic, nomartic, ampliaci, cantidad, "
    If conServidas Then
        SQL = SQL & "servidas,bultosser,"
    Else
        SQL = SQL & "numbultos,"
    End If
    SQL = SQL & "precioar, origpre, dtoline1, dtoline2,importel,codccost"
    'QUITO EL klote
    ',numlote"
    SQL = SQL & ",codtipor ,codcapit,ampliaci"
    
    SQL = SQL & " FROM " & NomTablaLineas
    If enlaza Then
        SQL = SQL & " WHERE " & Replace(ObtenerWhereCP, NombreTabla, NomTablaLineas)
        If EsHistorico Then SQL = SQL & " and fecpedcl='" & Format(Text1(1).Text, FormatoFecha) & "'"
    Else
        SQL = SQL & " WHERE numpedcl = -1"
    End If
    SQL = SQL & " Order by numpedcl, numlinea"
    MontaSQLCarga = SQL
End Function


Private Sub PonerModoOpcionesMenu(Modo As Byte)
'Activas unas Opciones de Menu y Toolbar según el Modo en que estemos
Dim b As Boolean

        b = (Modo = 2) Or (Modo = 5 And ModificaLineas = 0)
        'Me.mnOpciones.Enabled = (b Or Modo = 0)
        'Insertar
        Toolbar1.Buttons(5).Enabled = (b Or Modo = 0) And Not EsHistorico
        Me.mnNuevo.Enabled = (b Or Modo = 0) And Not EsHistorico
        'Modificar
        Toolbar1.Buttons(6).Enabled = b And Not EsHistorico
        Me.mnModificar.Enabled = b And Not EsHistorico
        'eliminar
        Toolbar1.Buttons(7).Enabled = b And Not EsHistorico
        Me.mnEliminar.Enabled = b And Not EsHistorico
            
        b = (Modo = 2)
        'Mantenimiento lineas
        Toolbar1.Buttons(10).Enabled = b And Not EsHistorico
        Me.mnLineas.Enabled = b And Not EsHistorico
        
        
        
        
  
        

        Toolbar1.Buttons(12).Enabled = b And Not EsHistorico
        Me.mnGeneraFactura.Enabled = b And Not EsHistorico
        'Generar Albaran desde Pedido  o insertar intercalando
        
        If Modo = 5 Then
            Toolbar1.Buttons(11).Image = 34 '.Buttons(11).Image = 26
            Toolbar1.Buttons(11).ToolTipText = "Insertar intercalando"
            b = (ModificaLineas = 0)
        Else
            'b=modo=2
            b = b And Not EsHistorico
            Toolbar1.Buttons(11).Image = 26   '26
            Toolbar1.Buttons(11).ToolTipText = "Generar albarán"
        End If
        Toolbar1.Buttons(11).Enabled = b
        Me.mnGenAlbaran.Enabled = b And Modo <> 5
        
        
        'Imprimir orden de instalacion
        Me.Toolbar1.Buttons(15).Enabled = Not EsHistorico
        Me.mnImpOrde.Enabled = Not EsHistorico
        
        b = (Modo >= 3) Or Modo = 1
        'Buscar
        Toolbar1.Buttons(1).Enabled = Not b
        Me.mnBuscar.Enabled = Not b
        'Ver Todos
        Toolbar1.Buttons(2).Enabled = Not b
        Me.mnVerTodos.Enabled = Not b
End Sub


Private Sub CargarComboFacturacion()
'### Combo Tipo Facturación
'Cargaremos el combo, o bien desde una tabla o con valores fijos o como
'se quiera, la cuestion es cargarlo
' El estilo del combo debe de ser 2 - Dropdown List
' Si queremos que este ordenado, o lo ordenamos por la sentencia sql
' o marcamos la opcion sorted del combo
'0-Factura Colectiva, 1-Factura x Albaran

    cboFacturacion.Clear
    cboFacturacion.AddItem "Factura Colectiva"
    cboFacturacion.ItemData(cboFacturacion.NewIndex) = 0

    cboFacturacion.AddItem "Factura x Albaran"
    cboFacturacion.ItemData(cboFacturacion.NewIndex) = 1

End Sub


Private Function InsertarPedido(vSQL As String, vTipoMov As CTiposMov) As Boolean
'Insertar la Cabecera de un Pedido, tabla: scaped
Dim MenError As String
Dim bol As Boolean, Existe As Boolean
Dim cambiaSQL As Boolean

    On Error GoTo EInsertarOferta
    
    bol = True
    
    cambiaSQL = False
    'Comprobar si mientras tanto se incremento el contador de Pedidos
    'para ello vemos si existe un Pedido con ese contador y si existe lo incrementamos
    Do
        MenError = DevuelveDesdeBDNew(conAri, NombreTabla, "numpedcl", "numpedcl", Text1(0).Text, "N")
        If MenError <> "" Then
            'Ya existe el contador incrementarlo
            Existe = True
            vTipoMov.IncrementarContador (CodTipoMov)
            Text1(0).Text = vTipoMov.ConseguirContador(CodTipoMov)
            cambiaSQL = True
        Else
            Existe = False
        End If
    Loop Until Not Existe
    If cambiaSQL Then vSQL = CadenaInsertarDesdeForm(Me)
    
    
    'Aqui empieza transaccion
    conn.BeginTrans
    MenError = "Insertando en la tabla Cabecera de Pedidos (" & NombreTabla & ")."
    conn.Execute vSQL, , adCmdText
    
    'Actualizar los datos del cliente si es de varios
    If EsDeVarios Then
        MenError = "Actualizando el Cliente de Varios (sclvar)."
        bol = ActualizarClienteVarios(Text1(4).Text, Text1(6).Text)
    End If
    
    MenError = "Actualizando el contador del Pedido."
'    bol = vTipoMov.IncrementarContador("REG")
    vTipoMov.IncrementarContador (CodTipoMov)

EInsertarOferta:
        If Err.Number <> 0 Then
            MenError = "Insertando Pedido." & vbCrLf & "----------------------------" & vbCrLf & MenError
            MuestraError Err.Number, MenError, Err.Description
            bol = False
        End If
        If bol Then
            conn.CommitTrans
            InsertarPedido = True
        Else
            conn.RollbackTrans
            InsertarPedido = False
        End If
End Function


Private Sub LimpiarDatosCliente()
'Limpia los campos del Form con datos del cliente
Dim I As Byte

    For I = 4 To 13
        Text1(I).Text = ""
    Next I
    If Modo = 3 Then
        For I = 14 To 17
            Text1(I).Text = ""
        Next I
        Text2(12).Text = ""
        Text2(14).Text = ""
        Text2(17).Text = ""
'        Text2(8).Text = ""
        Me.cboFacturacion.ListIndex = -1
    End If
End Sub
    

Private Function PedidoConInstalaciones() As Boolean
'Comprobar si en las lineas del Pedido hay algun articulo que sea Instalacion
'Si no hay niguna linea que sea instalacion no se imprimira la Orden de Instalacion
Dim SQL As String
Dim RS As ADODB.Recordset

    On Error GoTo EInstalac

    PedidoConInstalaciones = False
    SQL = "SELECT sliped.codartic, sliped.numlinea,scaped.numpedcl, sfamia.instalac "
    SQL = SQL & " FROM ((sliped INNER JOIN scaped ON sliped.numpedcl=scaped.numpedcl) "
    SQL = SQL & " INNER JOIN sartic ON sliped.codartic=sartic.codartic) INNER JOIN "
    SQL = SQL & " sfamia ON sartic.codfamia=sfamia.codfamia "
    SQL = SQL & " WHERE scaped.numpedcl = " & Val(Text1(0).Text) & " And sfamia.instalac = 1"
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If RS.EOF Then
        PedidoConInstalaciones = False
    Else
        PedidoConInstalaciones = True
    End If
    RS.Close
    Set RS = Nothing
    
EInstalac:
    If Err.Number <> 0 Then MuestraError Err.Number, "Comprobar si hay Articulos que son Instalaciones.", Err.Description
End Function


Private Function InicializarCStockAlbar(ByRef vCStock As CStock, TipoM As String, Optional numlinea As String, Optional ByRef RS As ADODB.Recordset) As Boolean
'Para comprobar stock al pasar de Pedido a Albaran de Venta
On Error Resume Next
    
    vCStock.tipoMov = TipoM
    
    'vCStock.DetaMov = "ALV"
    'If EsAMostrador Then vCStock.DetaMov = "ALM"
    vCStock.DetaMov = EsAMostrador2
    
    vCStock.Trabajador = CLng(Text1(4).Text) 'En codigope ponemos el Cliente
    vCStock.Documento = Text1(0).Text
    vCStock.codArtic = RS!codArtic
    vCStock.codAlmac = CInt(RS!codAlmac)
    
    If AlbCompleto Then
        vCStock.cantidad = CSng(RS!cantidad)
        If RS.Fields.Count > 3 Then 'Si no se selecciona el campo importe de la tabla es que solo vamos a comprobar stock y no se necesita
            vCStock.Importe = CCur(RS!ImporteL)
        End If
    Else
        vCStock.cantidad = CSng(RS!servidas)
        'Si se va a Insertar en alguna linea obtener el importe
        'Si solo vamos a comprobar stock no hace falta el importe
        If RS.Fields.Count > 4 Then
            vCStock.Importe = CCur(CalcularImporte(RS!servidas, RS!precioar, RS!dtoline1, RS!dtoline2, vParamAplic.TipoDtos))
        End If
    End If
    
    vCStock.LineaDocu = CInt(ComprobarCero(numlinea))
    
    If Err.Number <> 0 Then
        MsgBox "No se han podido inicializar la clase para actualizar Stock", vbExclamation
        InicializarCStockAlbar = False
    Else
        InicializarCStockAlbar = True
    End If
End Function


Private Function PasarPedidoAAlbaran(vSQL As String, NumAlb As String) As Boolean
'IN -> vSQL: cadena para el Select con los datos obtenidos en frmList
'OUT -> numAlb: Nº de Albaran de Venta que se ha insertado
Dim bol As Boolean
Dim MenError As String
Dim devuelve As String
Dim SQL As String
Dim RS As ADODB.Recordset
Dim cCli As CCliente

    On Error GoTo EGenPedido

    bol = False
        
    'Aqui empieza transaccion
    conn.BeginTrans
    
    'Insertar en tablas de Albaranes el Pedido (scaalb, slialb)
    bol = InsertarAlbaran(vSQL, MenError, NumAlb)
    
    'Actualizar Stock en salmac, e introducir movimiento en smoval
    If bol Then
        MenError = "Error al insertar movimientos de stock."
        bol = InsertarMovStock(NumAlb)
    End If
    
    If bol Then
        If AlbCompleto Then  'Si se inserta Albaran
            'Borrar el Pedido de las tablas de Pedidos (scaped, sliped)
            MenError = "Eliminar pedido."
            bol = EliminarPedido(CLng(Text1(0).Text))
        Else
            'Actualizar la cantidad=cantidad-servidas y servidas= 0 en sliped
            bol = ActualizarPedido()
            'Marcar Resto de pedido: restoped=1
            If bol Then bol = ActualizarCabPedido
        End If
        
        If bol Then
            'Actualizar el campo fechamov (ult. movimiento) de la tabla de clientes (sclien)
            'si la fecha es posterior a la que tiene
            Set cCli = New CCliente
            If cCli.LeerDatos(Text1(4).Text) Then
                bol = cCli.ActualizaUltFecMovim(FechaAlb)
            Else
                bol = False
            End If
            Set cCli = Nothing
            
'            devuelve = DevuelveDesdeBDNew(conAri, "sclien", "fechamov", "codclien", Text1(4).Text, "N")
'            If CDate(FechaAlb) > CDate(devuelve) Then
'                MenError = "Actualizando Fecha Movimiento del Cliente."
'                devuelve = "UPDATE sclien SET fechamov=" & DBSet(FechaAlb, "F")
'                devuelve = devuelve & " WHERE codclien=" & Text1(4).Text
'                Conn.Execute devuelve, , adCmdText
'            End If
        End If
    End If
    
    If bol Then
    
        'DAVID. LO SACO DE AQUI.
        'Si no quiere meter los numeros que no los meta, que le den
        'Comprobar si Hay Nº SERIE en compras, si hay Mostrar los Nº Serie y seleccionar
        'sino, pedir los Nº de serie de aquellos articulos que lo requieran
        'ComprobarNSeriesLineas (NumAlb)
            
            
            
            
            
        If Not AlbCompleto Then
            'Eliminar las filas del pedido que se servieron completas (sliped)
            SQL = "DELETE FROM sliped WHERE " & Replace(ObtenerWhereCP, NombreTabla, NomTablaLineas) & " AND cantidad=0"
            conn.Execute SQL
            
            'Comprobar que si no quedan lineas en el pedido se elimine la cabecera del pedido
            SQL = "select codalmac,codartic FROM sliped WHERE numpedcl=" & Text1(0).Text
            Set RS = New ADODB.Recordset
            RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If RS.EOF Then 'No hay lineas de pedido --> Eliminar la cabecera
                SQL = "DELETE FROM " & NombreTabla & " WHERE numpedcl=" & Text1(0).Text
                conn.Execute SQL
            End If
            RS.Close
            Set RS = Nothing
        End If
    End If
    
    
EGenPedido:
    If Err.Number <> 0 Or Not bol Then
        MenError = "Pasando Pedido a Albaran." & vbCrLf & "----------------------------" & vbCrLf & MenError
        MuestraError Err.Number, MenError, Err.Description
        bol = False
    End If
    If bol Then
        conn.CommitTrans
    Else
        conn.RollbackTrans
    End If
    PasarPedidoAAlbaran = bol
End Function



Private Function InsertarAlbaran(vSQL As String, MenError As String, NumAlb As String) As Boolean
'Devuelve el mensaje de error si se produce
Dim bol As Boolean, Existe As Boolean
Dim devuelve As String
Dim vTipoMov As CTiposMov 'Clase Tipo Movimiento
Dim codtipom As String
Dim codtipomAUX  As String
Dim ObtenerContador As Boolean

    On Error GoTo EInsertarAlbaran
    
    bol = False
    InsertarAlbaran = bol
    
    'Obtener el Contador de PEDIDO
    'codtipom = "ALV"
    'If EsAMostrador Then codtipom = "ALM"
    codtipom = EsAMostrador2
    
    
    ObtenerContador = True   'Obtener contador
    codtipomAUX = codtipom
    If InstalacionEsEulerTaxco Then
        If codtipom = "ALR" Then
       
            
            'Si el trabajador es de Valencia sera los ALR, si es de EUSAKADI seran CAR
            devuelve = DevuelveDesdeBD(conAri, "codalmac", "straba", "codtraba", Text1(3).Text)
            If devuelve = "10" Then codtipomAUX = "CAR"
        
        End If
    End If
    
    
    
    
    Set vTipoMov = New CTiposMov
    If vTipoMov.Leer(codtipomAUX) Then
        'Comprobar si mientras tanto se incremento el contador de Pedidos
        'para ello vemos si existe una oferta con ese contador y si existe la incrementamos
        Do
            NumAlb = vTipoMov.ConseguirContador(codtipomAUX)
            devuelve = DevuelveDesdeBDNew(conAri, "scaalb", "numalbar", "codtipom", codtipom, "T", , "numalbar", NumAlb, "N")
            If devuelve <> "" Then
                'Ya existe el contador incrementarlo
                Existe = True
                vTipoMov.IncrementarContador (codtipomAUX)
                NumAlb = vTipoMov.ConseguirContador(codtipomAUX)
            Else
                Existe = False
            End If
        Loop Until Not Existe
            
    Else 'No existe el tipo de Movimiento
        Set vTipoMov = Nothing
        Exit Function
    End If
    
    'Acabar la sql con el contador seleccionado
    devuelve = vSQL
    vSQL = "INSERT INTO scaalb (codtipom,numalbar,fechaalb,factursn,codclien,nomclien,domclien,codpobla,pobclien,proclien,nifclien,telclien,"
    vSQL = vSQL & "coddirec,nomdirec,referenc,codtraba,codtrab1,codtrab2,codagent,codforpa,codenvio,dtoppago,dtognral,tipofact,"
    vSQL = vSQL & "observa01,observa02,observa03,observa04,observa05,numofert,fecofert,numpedcl,fecpedcl,fecentre,sementre,actuacion,fechaaux) "
    vSQL = vSQL & "SELECT '" & codtipom & "' as codtipom, " & NumAlb & " as numalbar, " & devuelve
    'SAIL
    vSQL = vSQL & ",actuacion"
    If InstalacionEsEulerTaxco Then
        vSQL = vSQL & "," & DBSet(FechaAlb, "F", "N")
    Else
        vSQL = vSQL & ",null"
    End If
    vSQL = vSQL & " FROM " & NombreTabla & " WHERE numpedcl=" & Text1(0).Text

    'Insertar Cabecera
    MenError = "Error al insertar en la tabla Cabecera de Albaranes (scaalb )."
    conn.Execute vSQL, , adCmdText
    
    'Insertar Lineas Pedido
    MenError = "Error al insertar en la tabla Lineas de Albaran (slialb)."
    If Not InsertarLineasAlbaran(codtipom, NumAlb) Then Exit Function
    
    
    
    
    'EN EULER
    If InstalacionEsEulerTaxco Then
        
        devuelve = ""
        If Not AlbCompleto Then
            CadenaArticulosEULER = Mid(CadenaArticulosEULER, 2)
            devuelve = " AND codartic IN (" & CadenaArticulosEULER & ")"
            CadenaArticulosEULER = ""
        End If
        vSQL = "UPDATE slippr set codtipomV = " & DBSet(codtipom, "T")
        vSQL = vSQL & " , numalbarV  =" & DBSet(NumAlb, "T")
        vSQL = vSQL & " , fechaalbV  =" & DBSet(FechaAlb, "F")
        vSQL = vSQL & " WHERE codclien=" & Text1(4).Text & " AND numpedV =" & Text1(0).Text & " AND codtipomV is null"
        vSQL = vSQL & devuelve
        conn.Execute vSQL
        
        vSQL = "UPDATE slialp set codtipomV = " & DBSet(vTipoMov.TipoMovimiento, "T")
        vSQL = vSQL & " , numalbarV  =" & DBSet(NumAlb, "T")
        vSQL = vSQL & " , fechaalbV  =" & DBSet(FechaAlb, "F")
        vSQL = vSQL & " WHERE codclien=" & Text1(4).Text & " AND numpedV =" & Text1(0).Text & " AND codtipomV is null"
        vSQL = vSQL & devuelve
        conn.Execute vSQL
        
        
        vSQL = "UPDATE slifpc set codtipomV = " & DBSet(vTipoMov.TipoMovimiento, "T")
        vSQL = vSQL & " , numalbarV  =" & DBSet(NumAlb, "T")
        vSQL = vSQL & " , fechaalbV  =" & DBSet(FechaAlb, "F")
        vSQL = vSQL & " WHERE codclien=" & Text1(4).Text & " AND numpedV =" & Text1(0).Text & " AND codtipomV is null"
        vSQL = vSQL & devuelve
        conn.Execute vSQL
        
    End If
    MenError = "Error al actualizar el contador del ALbaran."
    vTipoMov.IncrementarContador (codtipom)
    Set vTipoMov = Nothing
    bol = True
    
EInsertarAlbaran:
        If Err.Number <> 0 Then bol = False
        InsertarAlbaran = bol
        
End Function


Private Function InsertarLineasAlbaran(TipoM As String, NumAlb As String) As Boolean
'Inserta en la tabla de lineas de albaran (slialb)
Dim SQL As String
Dim RS As ADODB.Recordset
Dim ImpLinea As String
Dim NumBulto As String

    On Error Resume Next

    'ENERO 2008.   codprove en slialb para traza de proveedores en lineas

    If AlbCompleto Then
        'Insertar en la tabla de Pedido, los registros seleccionados de la tabla de Ofertas
        SQL = ""
        SQL = "SELECT '" & TipoM & "', " & NumAlb & " as numalbar, numlinea, codalmac,"
        SQL = SQL & NomTablaLineas & ".codartic, " & NomTablaLineas & ".nomartic, ampliaci, "
        SQL = SQL & "cantidad, numbultos,precioar, dtoline1, dtoline2, importel, origpre"
        'traza
        SQL = SQL & ",codprove,numlote,codccost"
        SQL = SQL & " FROM " & NomTablaLineas & ",sartic WHERE " & NomTablaLineas & ".codartic = sartic.codartic"
        SQL = SQL & " AND numpedcl=" & Text1(0).Text
        SQL = "INSERT INTO slialb (codtipom,numalbar,numlinea,codalmac,codartic,nomartic,ampliaci,cantidad,numbultos,precioar,dtoline1,dtoline2,importel,origpre,codproveX,numlote,codccost) " & SQL
        conn.Execute SQL
    Else
        
        CadenaArticulosEULER = ""  'En euler, para cuando actualice los pedidos de proveedor
        SQL = "select sliped.*,codprove from sliped,sartic WHERE  sliped.codartic=sartic.codartic "
        SQL = SQL & " AND " & Replace(ObtenerWhereCP, NombreTabla, NomTablaLineas)
        SQL = SQL & " AND servidas>0"
        Set RS = New ADODB.Recordset
        RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        While Not RS.EOF 'Para cada linea de pedido insertar una de albaran si servidas >0
            If RS!servidas > 0 Then
                ImpLinea = CalcularImporte(RS!servidas, RS!precioar, RS!dtoline1, RS!dtoline2, vParamAplic.TipoDtos)
'                NumBulto = DevuelveDesdeBDNew(conAri, "sartic", "unicajas", "codartic", RS!codArtic, "T")
'                NumBulto = CalcularNumBultos(RS!servidas, CInt(NumBulto))
                
                SQL = "INSERT INTO slialb (codtipom,numalbar,numlinea,codalmac,codartic,nomartic,ampliaci,"
                SQL = SQL & "cantidad,numbultos,precioar,dtoline1,dtoline2,importel,origpre,codprovex,numlote,codccost) "
                SQL = SQL & " VALUES('" & TipoM & "', " & NumAlb & ", " & RS!numlinea & " , "
                SQL = SQL & RS!codAlmac & ", " & DBSet(RS!codArtic, "T") & ", " & DBSet(RS!NomArtic, "T") & ", " & DBSet(RS!Ampliaci, "T") & ", "
                SQL = SQL & DBSet(RS!servidas, "N") & ", " & DBSet(RS!bultosser, "N") & ", "
                SQL = SQL & DBSet(RS!precioar, "N") & ", " & DBSet(RS!dtoline1, "N") & ", " & DBSet(RS!dtoline2, "N") & ", "
                SQL = SQL & DBSet(ImpLinea, "N") & ", " & DBSet(RS!origpre, "T") & "," & RS!Codprove & "," & DBSet(RS!numLote, "T") & ","
                SQL = SQL & DBSet(RS!CodCCost, "T", "S") & ")"
                conn.Execute SQL
                
                
                CadenaArticulosEULER = CadenaArticulosEULER & ", " & DBSet(RS!codArtic, "T")
                
            End If
            RS.MoveNext
        Wend
        RS.Close
        Set RS = Nothing
    End If
    
    If Err.Number <> 0 Then
         'Hay error , almacenamos y salimos
        InsertarLineasAlbaran = False
    Else
        InsertarLineasAlbaran = True
    End If
End Function


Private Function EliminarPedido(numPed As Long) As Boolean
'Eliminar las lineas y la Cabecera de un Pedido. Tablas: scaped, sliped
Dim SQL As String

    On Error GoTo EEliminarPed

     SQL = " WHERE  numpedcl=" & numPed

    'Lineas de Pedido
    conn.Execute "Delete from " & NomTablaLineas & SQL

    'Cabecera
    conn.Execute "Delete from " & NombreTabla & SQL

EEliminarPed:
    If Err.Number <> 0 Then
        EliminarPedido = False
    Else
        EliminarPedido = True
    End If
End Function


Private Function ActualizarPedido() As Boolean
'Actualiza la tabla de lineas de pedido (sliped)
'cantidad=cantidad-servidas y servidas=0
Dim SQL As String
Dim RS As ADODB.Recordset
Dim ImpLinea As String
Dim NumBultos As String
    
    On Error GoTo EActPedido
    
    SQL = "select codalmac, codartic, cantidad,servidas,numbultos, precioar,dtoline1,dtoline2,numpedcl,numlinea from sliped "
    SQL = SQL & " WHERE " & Replace(ObtenerWhereCP, NombreTabla, NomTablaLineas)
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not RS.EOF 'Para cada linea
        ImpLinea = CalcularImporte(RS!cantidad - RS!servidas, RS!precioar, RS!dtoline1, RS!dtoline2, vParamAplic.TipoDtos)
        NumBultos = DevuelveDesdeBDNew(conAri, "sartic", "unicajas", "codartic", RS!codArtic, "T")
        NumBultos = CalcularNumBultos2(RS!cantidad - RS!servidas, CInt(NumBultos))
        SQL = "UPDATE sliped SET cantidad=cantidad-servidas, servidas=0, importel=" & DBSet(ImpLinea, "N")
        SQL = SQL & ", numbultos=" & DBSet(NumBultos, "N") & ",bultosser=0"
        SQL = SQL & " WHERE codalmac=" & RS!codAlmac & " AND codartic=" & DBSet(RS!codArtic, "T")
        'Para que no cambie los importes. Abril 2008
        SQL = SQL & " AND numpedcl= " & RS!NumPedcl & " AND numlinea=" & RS!numlinea
        conn.Execute SQL
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing

EActPedido:
    If Err.Number <> 0 Then
        ActualizarPedido = False
    Else
        ActualizarPedido = True
    End If
End Function


Private Function ActualizarCabPedido() As Boolean
Dim SQL As String

    On Error Resume Next

    SQL = "UPDATE scaped SET restoped=1 " & " WHERE " & ObtenerWhereCP
    conn.Execute SQL
    
    If Err.Number <> 0 Then
        ActualizarCabPedido = False
    Else
        ActualizarCabPedido = True
    End If
End Function


Private Function InsertarMovStock(NumAlb As String) As Boolean
Dim vCStock As CStock
Dim b As Boolean
Dim RS As ADODB.Recordset
Dim SQL As String

    On Error Resume Next

    InsertarMovStock = False
    
    Set vCStock = New CStock
    b = True
    
    SQL = "select * from sliped WHERE " & Replace(ObtenerWhereCP, NombreTabla, NomTablaLineas)
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    vCStock.FechaMov = FechaAlb
    
    'para cada linea del Pedido Insertar en smoval y Actualizar Stock en salmac
    While (Not RS.EOF) And b
        'si hay control de stock
'        SQL = DevuelveDesdeBDNew(conAri, "sartic", "ctrstock", "codartic", RS!codartic, "T")
'        If Val(SQL) = 1 Then
            If Not InicializarCStockAlbar(vCStock, "S", CStr(RS!numlinea), RS) Then Exit Function

            'vCStock.Documento = numAlb
            vCStock.Documento = Format(NumAlb, "0000000")
            If vCStock.cantidad <> 0 Then
                'en actualizar stock comprobamos si el articulo tiene control de stock
                    b = vCStock.ActualizarStock(False, False)
            End If
'        End If
        RS.MoveNext
    Wend
    Set vCStock = Nothing
    RS.Close
    Set RS = Nothing
    
    InsertarMovStock = b
    
End Function


Private Sub ImprimirAlbaran(Opcion As Integer, Numalbar As String)
Dim cadFormula As String
Dim cadParam As String
Dim numParam As Byte
Dim cadSelect As String 'select para insertar en tabla temporal
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim NomTabla As String
Dim clien As String
    
    cadFormula = ""
    cadParam = ""
    cadSelect = ""
    numParam = 0
    NomTabla = "scaalb"
   
    '===================================================
    '============ PARAMETROS ===========================
    If (Opcion = 45) Then indRPT = 10 'Albaran Clientes
    If Not PonerParamRPT2(indRPT, cadParam, numParam, nomDocu, pImprimeDirecto, pPdfRpt, pRptvMultiInforme) Then
        Exit Sub
    End If

    'Añadir el codigo de usuario como parametro para link con tabla Temporal (tmptiposiva) en el Report
    'tabla temporal para el calculo del bruto total para cada tipo de IVA
    cadParam = cadParam & "pCodUsu=" & vUsu.Codigo & "|"
    numParam = numParam + 1
    
    'Nombre fichero .rpt a Imprimir
    frmImprimir.NombreRPT = nomDocu
    frmImprimir.NombrePDF = pPdfRpt
    frmImprimir.SeleccionaRPTCodigo = pRptvMultiInforme
    '===================================================
    '================= FORMULA =========================
    'Cadena para seleccion Nº de Albaran
    '---------------------------------------------------
    If Numalbar <> "" Then
        'Cod Tipo Movimiento
        devuelve = "{" & NomTabla & ".codtipom}='" & EsAMostrador2 & "'" 'Val(txtCodigo(0).Text)
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
        'Nº Albaran
        devuelve = "{" & NomTabla & ".numalbar}=" & Numalbar
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
        'select para insertar en tabla temporal
        cadSelect = QuitarCaracterACadena(cadFormula, "{")
        cadSelect = QuitarCaracterACadena(cadSelect, "}")
    End If
   
    '=========================================================================

    'Aqui sabemos que valor tiene CodClien y añadimos a los parametros el tipo de IVA
    'que se aplica a ese cliente
    devuelve = DevuelveDesdeBDNew(conAri, "scaalb", "codclien", "codtipom", "ALV", "T", , "numalbar", Numalbar, "N")
    If devuelve <> "" Then
        clien = "albarcon"   'Albaran valorado
        devuelve = DevuelveDesdeBDNew(conAri, "sclien", "tipoiva", "codclien", devuelve, "N", clien)
        If devuelve <> "" Then
            cadParam = cadParam & "pTipoIVA=" & devuelve & "|"
            numParam = numParam + 1
        End If
        If clien = "" Or clien = "albarcon" Then clien = "0"
        ' 0 "Todo"
        ' 1 "Cantidad y Precio"
        ' 2 "Cantidad"
        cadParam = cadParam & "Albarcon=" & clien & "|"
        numParam = numParam + 1
    End If
     
     
     With frmImprimir
            .FormulaSeleccion = cadFormula
            .OtrosParametros = cadParam
            .NumeroParametros = numParam
            .SoloImprimir = False
            .EnvioEMail = False
            .Opcion = Opcion
            .Titulo = "Albaran de Cliente"
            .ConSubInforme = True
            .Show vbModal
    End With
End Sub


Private Function InicializarCStock(ByRef vCStock As CStock, TipoM As String, Optional numlinea As String) As Boolean
On Error Resume Next

    vCStock.tipoMov = TipoM
    If Modo = 6 Then 'Pasar Pedido a Albaran
        vCStock.DetaMov = "ALV"
    Else
        vCStock.DetaMov = CodTipoMov
    End If
    
    vCStock.Trabajador = CLng(Text1(4).Text) 'ponemos el cliente del pedido
    vCStock.Documento = Text1(0).Text 'Nº Pedido
    vCStock.FechaMov = Text1(1).Text
    
    If ModificaLineas = 1 Or ModificaLineas = 2 Then '1=Insertar, 2=Modificar
        vCStock.codArtic = txtaux(1).Text
        vCStock.codAlmac = CInt(txtaux(0).Text)
        vCStock.cantidad = CSng(ComprobarCero(txtaux(3).Text))
        vCStock.Importe = CCur(ComprobarCero(txtaux(8).Text))
    Else
        vCStock.codArtic = Data2.Recordset!codArtic
        vCStock.codAlmac = CInt(Data2.Recordset!codAlmac)
        If Modo = 6 Then 'Pasar Pedido a Albaran
            vCStock.cantidad = CSng(ComprobarCero(txtaux(3).Text))
        Else
            vCStock.cantidad = CSng(Data2.Recordset!cantidad)
        End If
        vCStock.Importe = CCur(Data2.Recordset!ImporteL)
    End If
    
    If ModificaLineas = 1 Then '1=Insertar Linea
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


Private Function ActualizarServidas() As Boolean
'Actualiza el campo "servidas" de la tabla "sliped"
Dim SQL As String

    On Error Resume Next
    
    SQL = "0"
    If txtaux(3).Text <> "" Then
        If InStr(1, txtaux(3).Text, ",") > 0 Then
            ' ---- [28/09/2009] (LAURA)
'            sql = TransformaComasPuntos(txtAux(3).Text)
            SQL = DBSet(txtaux(3).Text, "N")
            ' ----
        Else
            SQL = txtaux(3).Text
        End If
    End If
    SQL = "UPDATE sliped SET servidas= " & SQL
    SQL = SQL & ", bultosser=" & txtaux(9).Text
    SQL = SQL & " WHERE " & Replace(ObtenerWhereCP, NombreTabla, NomTablaLineas) & " AND numlinea=" & Data2.Recordset!numlinea
    conn.Execute SQL
    
    If Err.Number <> 0 Then
        ActualizarServidas = False
    Else
        ActualizarServidas = True
    End If
End Function


Private Sub PonerServidas(index As Integer)
Dim NumFila As Integer
Dim cadMen As String
Dim vStock As String
Dim SeSirve As Boolean
'    NumFila = DataGrid1.Row
    NumFila = Data2.Recordset.AbsolutePosition
    
    If index = 3 Then
        PonerFormatoDecimal txtaux(index), 1
        If txtaux(index).Text <> "" Then
            If (CCur(txtaux(index).Text) <> Data2.Recordset!servidas) Or txtaux(9).Text = "" Then
                '-- calcular nº bultos
                'Comprobar si el articulo se vende por cajas antes de entrar a la función
                cadMen = DevuelveDesdeBDNew(conAri, "sartic", "unicajas", "codartic", Me.Data2.Recordset!codArtic, "T")
            
                If cadMen <> "" Then
                    '- obtener el nº bultos: cantidad/unids.caja
                    txtaux(9).Text = CalcularNumBultos2(CCur(txtaux(3).Text), CInt(cadMen))
                End If
            End If
        End If
    End If
    
    ActualizarServidas
    CargaGrid2 DataGrid1, Data2, True
    SituarDataPosicion Data2, CLng(NumFila), ""
    
'    DataGrid1.Row = NumFila
    'Enero 2010
    'If SePuedeServir(vStock) Then
    SeSirve = SePuedeServir(vStock)
    If Not SeSirve Then
        cadMen = "No hay suficiente Stock para servir la cantidad solicitada."
        cadMen = cadMen & vbCrLf & "(Stock= " & vStock & ")" & vbCrLf
        cadMen = cadMen & "¿Continuar?"
        If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then SeSirve = True
    End If
    
    If SeSirve Then
        If CSng(txtaux(3).Text) > Data2.Recordset!cantidad Then
            cadMen = "La cantitad a servir debe ser menor o igual a al cantidad del pedido."
            cadMen = cadMen & vbCrLf
            MsgBox cadMen, vbInformation
            PonerFoco txtaux(3)
            
        Else
'            TxtAux_KeyDown 3, 40, 0
            If index = 3 Then
                PonerFoco txtaux(9)
            Else
                MoverSigRegistro
                If Screen.ActiveControl.Name <> "cmdAceptar" Then PonerFoco txtaux(3)
            End If
        End If
    Else
        
'        ' ---- [28/09/2009] (LAURA) : dejar pasar el foco aunque no haya stock en recepcion incompleta
'        If Index = 3 Then
'            PonerFoco txtAux(9)
'        Else
'            MoverSigRegistro
'            If Screen.ActiveControl.Name <> "cmdAceptar" Then PonerFoco txtAux(3)
'        End If
        PonerFoco txtaux(3)
        ' ----
    End If
    
End Sub


Private Function SePuedeServir(vStock As String) As Boolean
'Si se puede servir una determinada linea del pedido cuando se esta introduciendo
'la cantidad a servir para cada codalmac,codartic
'OUT -> vStock: stock existente
Dim SQL As String
Dim RS As ADODB.Recordset
Dim Dif As Long
Dim vCStock As CStock

    On Error GoTo EServir

    Set vCStock = New CStock
    vCStock.codArtic = Data2.Recordset!codArtic
    If Not vCStock.MueveStock Then
        SePuedeServir = True
        Set vCStock = Nothing
        Exit Function
    End If
    Set vCStock = Nothing

    
    SePuedeServir = False
    SQL = "SELECT sliped.codalmac, sliped.codartic, canstock , sum(servidas), canstock - SUM(servidas) as Dif "
    SQL = SQL & " FROM sliped INNER JOIN salmac ON sliped.codalmac=salmac.codalmac AND sliped.codartic=salmac.codartic "
    SQL = SQL & " WHERE " & Replace(ObtenerWhereCP, NombreTabla, NomTablaLineas) & " AND sliped.codAlmac = " & Data2.Recordset!codAlmac & " AND sliped.codartic=" & DBSet(Data2.Recordset!codArtic, "T")
    SQL = SQL & " GROUP by sliped.codalmac, sliped.codartic "

    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        Dif = RS!Dif
        SePuedeServir = (RS!Dif >= 0)
        vStock = CStr(RS!CanStock)
    End If
    RS.Close
    Set RS = Nothing

EServir:
    If Err.Number <> 0 Then SePuedeServir = False
End Function


Private Sub GenerarAlbaran(PasarTambienAFacturar As Boolean)
Dim numPed As Long 'Nº Pedido
Dim NumAlb As String 'Nº Albaran
Dim SQL As String
Dim ImprimeFactura As Boolean
Dim AlbaranGenerado As Boolean

    'Pedir: Operador de Albaran, Material Preparado por y forma de envio
    CadenaSQL = ""
    
    Set frmList = New frmListadoPed
    If PasarTambienAFacturar Then
        frmList.OpcionListado = 1000
    Else
        frmList.OpcionListado = 43
    End If
    frmList.NumCod = CodTipoMov
    frmList.Show vbModal
    
    Set frmList = Nothing
    
    If CadenaSQL = "" Then Exit Sub
    
    
    NumRegElim = Data1.Recordset.AbsolutePosition
    numPed = Data1.Recordset!NumPedcl
    
    'Si es factura el albaran NO se imprime, y se imprimira si lo ha marcado, la factura
    
    If PasarTambienAFacturar Then
        ImprimeFactura = ImprimeAlb
        ImprimeAlb = False 'El albaran NO se imprime generanod la factura
'        ImprimeEtiq = False
        ImprimeHojaExp = False
    End If
    
    'CadenaSQL, se obtiene desde frmList
    lblIndicador.Caption = "Gen. albaran"
    lblIndicador.Refresh
    
    

    
    AlbaranGenerado = PasarPedidoAAlbaran(CadenaSQL, NumAlb)
    'If PasarPedidoAAlbaran(CadenaSQL, NumAlb) Then
    If AlbaranGenerado Then
        'Esto estaba antes dentro de pasarpedido
        'ahora esta fuera del begintrans
        ComprobarNSeriesLineas (NumAlb)
        
'        'Comprobar si Hay Nº SERIE en compras, si hay Mostrar los Nº Serie y seleccionar
'        'sino, pedir los Nº de serie de aquellos articulos que lo requieran
'        ComprobarNSeriesLineas (NumAlb)
        Espera 0.4
        If Not PasarTambienAFacturar Then
            MsgBox "El Pedido de Venta Nº: " & Format(numPed, "0000000") & vbCrLf & vbCrLf & "ha generado el Albaran Nº: " & EsAMostrador2 & Format(NumAlb, "0000000"), vbInformation
        Else
            'Ahora genero la factura a partir del ALBARAN
            lblIndicador.Caption = "Gen FACTURA"
            DoEvents
            
            SQL = EsAMostrador2
            
            CadenaSQL = "scaalb.codtipom = '" & SQL & "' AND scaalb.numalbar = " & NumAlb
            Precio = "SELECT scaalb.*,sclien.nomclien FROM scaalb INNER JOIN sclien ON scaalb.codclien=sclien.codclien "
            Precio = Precio & " WHERE " & CadenaSQL
            TraspasoAlbaranesFacturas Precio, CadenaSQL, FechaAlb, CtaBancoPropi, Nothing, lblIndicador, ImprimeFactura, SQL, "", 1, True, False, False
        End If
            
        
        PonerModo 2
        If AlbCompleto Then
            'Se habra eliminado el pedido de (scaped, sliped)
            PosicionarDataTrasEliminar
        Else
            SQL = DevuelveDesdeBDNew(conAri, "scaped", "numpedcl", "numpedcl", Text1(0).Text, "N")
            If SQL = "" Then 'Ya no existe le pedido lo hemos eliminado
                PosicionarDataTrasEliminar
            Else
                PosicionarData
                CargaGrid DataGrid1, Data2, True, False
            End If
        End If
        Screen.MousePointer = vbDefault
        CargaTxtAuxServidas False, False
    
        'Imprimer albaran si se solicitó
        If ImprimeAlb Then ImprimirAlbaran 45, NumAlb
        
        'Imprimir Etiquetas si se solicito
        If ImprimeEtiq Then
            frmListado.NumCod = NumAlb
            
            AbrirListado 95
        End If
        
        'Imprimir Hoja Expedicion si se solicito
        If ImprimeHojaExp Then
            ImprimirHojaExpedicion 45, NumAlb, EsAMostrador2
        End If
        
'    Else 'Si no se ha pasado el Pedido a Albaran
        
    End If
End Sub


Private Function SePuedeServirPedido() As Boolean
'Si se puede servir el Pedido solicitado (parcial o completo) y pasar a albaran
Dim vCStock As CStock
Dim SQL As String
Dim b As Boolean
Dim RS As ADODB.Recordset

    On Error Resume Next

    'Verificar si hay stock para aquellas familias que no son instalacion
    Set vCStock = New CStock
    b = True
    
    If AlbCompleto Then
        SQL = "SELECT codalmac, codartic, SUM(cantidad) as cantidad from sliped "
    Else
        SQL = "SELECT codalmac, codartic, SUM(servidas) as servidas from sliped "
    End If
    SQL = SQL & " WHERE " & Replace(ObtenerWhereCP, NombreTabla, NomTablaLineas)
    SQL = SQL & " GROUP by codalmac, codartic"
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    'para cada linea del Pedido comprobar el stock si no es instalacion
    While (Not RS.EOF) And b
        If Not InicializarCStockAlbar(vCStock, "S", , RS) Then
            b = False
            Screen.MousePointer = vbDefault
            Set vCStock = Nothing
            RS.Close
            Set RS = Nothing
            Exit Function
        End If
        
        'Comprobar si se puede mover stock (hay stock, o si no hay pero no control de stock)
        If AlbCompleto Then
            If vCStock.MueveStock Then b = vCStock.MoverStock(False, False, True)
        Else
            If vCStock.MueveStock Then b = vCStock.MoverStock(False, False)
        End If
        RS.MoveNext
    Wend
    
    Set vCStock = Nothing
    RS.Close
    Set RS = Nothing
    SePuedeServirPedido = b
    
    If Err.Number <> 0 Then SePuedeServirPedido = False
End Function


Private Sub InicializarServidas()
'Pone el campo servidas a 0 en la tabla lineas de pedido (sliped)
Dim SQL As String

    SQL = "UPDATE " & NomTablaLineas & " SET servidas= 0, bultosser=0 "
    SQL = SQL & " WHERE " & Replace(ObtenerWhereCP, NombreTabla, NomTablaLineas)
    conn.Execute SQL
End Sub


Private Sub ComprobarNSeriesLineas(NumAlb As String)
'Al pasar de PEDIDO a ALBARAN
'control de Nº Series si hay algun articulo en las lineas de pedido que requiere Nº de serie
'Si NO se realiza control Nº series en compras pedirlos ahora
'Si se realiza control Nº Series en compras verificar que efectivamente estan introducidos
'y mostrarlos para seleccionarlos
Dim SQL As String
Dim RSLineas As ADODB.Recordset
Dim cadWhere As String
        
    On Error GoTo ECompNSerie
    
    cadWhere = " WHERE codtipom='" & EsAMostrador2 & "' and "
    cadWhere = cadWhere & " numalbar=" & NumAlb
    
    'Seleccionamos aquellas lineas de albaran que tienen Nº de Serie
    SQL = "SELECT slialb.codartic, sum(cantidad) as cantidad, slialb.numlinea "
    SQL = SQL & " FROM slialb INNER JOIN sartic on slialb.codartic=sartic.codartic "
    SQL = SQL & cadWhere & " And nseriesn = 1 "
    SQL = SQL & " GROUP BY codartic ORDER BY Codartic "

    Set RSLineas = New ADODB.Recordset
    RSLineas.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not RSLineas.EOF Then
        'Comprobar si NO Hay Nº SERIE en Compras y si no se realizo alli
        'Mostrar ahora ventana para pedir los Nº Serie de la cantidad introducida
        Me.cmdAux(1).Tag = NumAlb
        If Not vParamAplic.NumSeries Then
            PedirNSeries RSLineas
        Else 'Se realizo contro en COMPRAS, Mostramos los Nº y seleccionamos
            MostrarNSeries RSLineas
        End If
    End If
    RSLineas.Close
    Set RSLineas = Nothing
    
ECompNSerie:
    If Err.Number <> 0 Then MuestraError Err.Number, "Comprobar Nº Serie.", Err.Description
End Sub


Private Sub PedirNSeries(ByRef RS As ADODB.Recordset)
    On Error GoTo EPedirNSeries
    
    'Visualizar en pantalla el Grid, y rellenar los Nº Serie
    PedirNSeriesGnral RS, True

    Set frmNSerie = New frmRepCargarNSerie
    frmNSerie.DeVentas = True 'Se llama desde Alb. de Venta
    frmNSerie.Show vbModal
    Set frmNSerie = Nothing
        
EPedirNSeries:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub MostrarNSeries(ByRef RSLineas As ADODB.Recordset)
'Si los Nº de serie se introdujeron en ALBARAN COMPRAS se muestran
'los Nº de serie de los articulos comprados y se seleccionan tantos como cantidad de la linea
Dim SQL As String
Dim Campos As String
   
    SQL = MostrarNSeriesGnral(RSLineas, Campos)
    
    Set frmMen = New frmMensajes
    frmMen.cadWhere = SQL
    frmMen.cadWHERE2 = ""
    frmMen.OpcionMensaje = 4 'Nº Series Articulo
    frmMen.vCampos = Campos
    frmMen.Show vbModal
    Set frmMen = Nothing
End Sub


Private Sub InsertarCabecera()
Dim vTipoMov As CTiposMov 'Clase Tipo Movimiento
Dim SQL As String

    Set vTipoMov = New CTiposMov
    If vTipoMov.Leer(CodTipoMov) Then
        Text1(0).Text = vTipoMov.ConseguirContador(CodTipoMov)
        SQL = CadenaInsertarDesdeForm(Me)
        If SQL <> "" Then
            If InsertarPedido(SQL, vTipoMov) Then
'                            PosicionarData
                CadenaConsulta = "Select * from " & NombreTabla & " WHERE " & ObtenerWhereCP & Ordenacion
                PonerCadenaBusqueda
                PonerModo 2
                'Ponerse en Modo Insertar Lineas
                BotonMtoLineas 1, "Pedidos"
                BotonAnyadirLinea False
            End If
        End If
        Text1(0).Text = Format(Text1(0).Text, "0000000")
    End If
    Set vTipoMov = Nothing
    Me.SSTab1.Tab = 0
End Sub


Private Function InsertarNSerie(numSerie As String, codArtic As String, numlinea As String, nummante As String) As Boolean
'Inserta o Actualiza en la tabla sserie, si al pasar Pedido -> Albaran
'existen lineas con control de Nº Serie
'Dim CadValues As String, cadValuesU As String, CadValuesI As String
Dim devuelve As String
Dim TieneMan As Boolean
Dim Numalbar As String
Dim nSerie As CNumSerie
Dim b As Boolean

    On Error GoTo EInsertarNSerie
    
'    'Comprobar que el cliente tiene mantenimientos en esa direc/dpto
'    TieneMan = "0"
'    devuelve = DevuelveDesdeBDNew(conAri, "scaman", "nummante", "codclien", Text1(4).Text, "N", , "coddirec", Text1(12).Text, "N")
'    'El cliente tiene Mantenimientos
'    If devuelve <> "" Then TieneMan = "1"
'
    nummante = Trim(nummante)
    TieneMan = (nummante <> "")

    Set nSerie = New CNumSerie
    nSerie.Cliente = CInt(Text1(4).Text)
    nSerie.DirDpto = Text1(12).Text
    nSerie.conMante = TieneMan
    nSerie.tipoMov = "ALV"   'DE ALVARAN VENTA
    
    devuelve = DevuelveDesdeBDNew(conAri, "scaalb", "fechaalb", "codtipom", "ALV", "T", , "numalbar", Me.cmdAux(1).Tag, "N")
    If devuelve <> "" Then nSerie.FechaVta = devuelve
    
    nSerie.NumAlbaran = Me.cmdAux(1).Tag
    nSerie.NumLinAlb = numlinea
    nSerie.nummante = nummante

    'obtenemos los dias de garantia del articulo
    nSerie.ObtenFechaFinGarantia codArtic, Text1(1).Text
   
     'Comprobar si existe en la tabla sserie
     Numalbar = "numalbar" 'Nº albaran de Venta
     devuelve = DevuelveDesdeBDNew(conAri, "sserie", "numserie", "numserie", numSerie, "T", Numalbar, "codartic", codArtic, "T")
     If devuelve <> "" Then 'EXISTE en tabla sserie
        If Numalbar = "" Then b = nSerie.ActualizarNumSerie(True)
     Else
         nSerie.Articulo = codArtic
         nSerie.numSerie = numSerie
        b = nSerie.InsertarNumSerie
    End If
    InsertarNSerie = True
    Set nSerie = Nothing
         
EInsertarNSerie:
    If Err.Number <> 0 Then b = False
    InsertarNSerie = b
End Function

 
Private Sub PonerDatosCliente2(codClien As String, Optional nifClien As String)
Dim vCliente As CCliente
Dim Observaciones As String
    
    On Error GoTo EPonerDatos
    
    If codClien = "" Then
        LimpiarDatosCliente
        Exit Sub
    End If

    Set vCliente = New CCliente
    
    'si se ha modificado el cliente volver a cargar los datos
    If vCliente.Existe(codClien) Then
        If vCliente.LeerDatos(codClien) Then
            'si el cliente esta bloqueado salimos
            If vCliente.ClienteBloqueado Then
                LimpiarDatosCliente
                Set vCliente = Nothing
                Exit Sub
            End If
            
'            EsDeVarios = vCliente.EsClienteVarios(Text1(4).Text)
            EsDeVarios = vCliente.DeVarios
            BloquearDatosCliente (EsDeVarios)
        
            If Modo = 4 And EsDeVarios Then 'Modificar
                'si no se ha modificado el cliente no hacer nada
                If CLng(Text1(4).Text) = CLng(Data1.Recordset!codClien) Then
                    If Text1(5).Text = Data1.Recordset!NomClien Then
                        Set vCliente = Nothing
                        Exit Sub
                    End If
                End If
            End If
            
'            If (Not EsDeVarios) Or (EsDeVarios And modo = 3) Then
            Text1(4).Text = vCliente.Codigo
            FormateaCampo Text1(4)
            If (Modo = 3) Or (Modo = 4) Then 'Insertar o Modificar
                Text1(5).Text = vCliente.Nombre  'Nom clien
                Text1(8).Text = vCliente.Domicilio
                Text1(9).Text = vCliente.CPostal
                Text1(10).Text = vCliente.Poblacion
                Text1(11).Text = vCliente.Provincia
                Text1(6).Text = vCliente.NIF
                Text1(7).Text = vCliente.TfnoClien
            End If
            
            If Modo = 3 Then 'insertar
                Text1(14).Text = vCliente.ForPago
                Text2(14).Text = PonerNombreDeCod(Text1(14), conAri, "sforpa", "nomforpa")
                Text1(15).Text = Format(vCliente.DtoPPago, FormatoDescuento)
                Text1(16).Text = Format(vCliente.DtoGnral, FormatoDescuento)
                Text1(17).Text = vCliente.Agente
                Text2(17).Text = PonerNombreDeCod(Text1(17), conAri, "sagent", "nomagent")
                Me.cboFacturacion.ListIndex = vCliente.TipoFactu
            End If

            Observaciones = DBLet(vCliente.Observaciones)
            If Observaciones <> "" Then MsgBox Observaciones, vbInformation, "Observaciones del cliente"
                           
            'Comprobar si el cliente tiene cobros pendientes
            ComprobarCobrosCliente codClien, Text1(1).Text
            
            If vCliente.DeVarios Then
                PonerFoco Text1(5)
            Else
                PonerFoco Text1(17)
            End If
        
        End If
    Else
        LimpiarDatosCliente
    End If
    Set vCliente = Nothing

EPonerDatos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poner Datos Cliente", Err.Description
End Sub


Private Sub PonerDatosClienteVario(nifClien As String)
Dim vCliente As CCliente
Dim b As Boolean
   
    If nifClien = "" Then Exit Sub
    
    Set vCliente = New CCliente
    b = vCliente.LeerDatosCliVario(nifClien)
    Text1(5).Text = vCliente.Nombre  'Nom clien
    Text1(8).Text = vCliente.Domicilio
    Text1(9).Text = vCliente.CPostal
    Text1(10).Text = vCliente.Poblacion
    Text1(11).Text = vCliente.Provincia
'    Text1(6).Text = vCliente.NIF
    Text1(7).Text = DBLet(vCliente.TfnoClien, "T")
            
    If Not b Then PonerFoco Text1(6)
    Set vCliente = Nothing
End Sub


Private Sub BloquearDatosCliente(bol As Boolean)
Dim I As Byte

    'bloquear/desbloquear campos de datos segun sea de varios o no
    If Modo <> 5 Then
        Me.imgBuscar(1).visible = bol
        Me.imgBuscar(1).Enabled = bol
        Me.imgBuscar(6).Enabled = bol
        
        For I = 5 To 11 'si no es de varios no se pueden modificar los datos
            BloquearTxt Text1(I), Not bol
        Next I
    End If
End Sub


Private Function ActualizarClienteVarios(clien As String, NIF As String) As Boolean
Dim vCliente As CCliente

    On Error GoTo EActualizarCV

    ActualizarClienteVarios = False
    
    Set vCliente = New CCliente
    If EsClienteVarios(clien) Then
        vCliente.NIF = NIF
        vCliente.Nombre = Text1(5).Text
        vCliente.Domicilio = Text1(8).Text
        vCliente.CPostal = Text1(9).Text
        vCliente.Poblacion = Text1(10).Text
        vCliente.Provincia = Text1(11).Text
        vCliente.TfnoClien = Text1(7).Text
        vCliente.ActualizarClienteV (NIF)
    End If
    Set vCliente = Nothing
    
    ActualizarClienteVarios = True
    
EActualizarCV:
    If Err.Number <> 0 Then
        ActualizarClienteVarios = False
    Else
        ActualizarClienteVarios = True
    End If
End Function



Private Sub CalcularDatosFactura()
Dim I As Integer
Dim cadWhere As String, SQL As String
Dim vFactu As CFactura

    'Limpiar en el form los datos calculados de la factura
    'y volvemos a recalcular
    For I = 33 To 56
         Text3(I).Text = ""
    Next I
    
     'Comprobar que hay lineas de albaran para calcular totales
    cadWhere = ObtenerWhereCP
    SQL = "Select count(*) from " & NomTablaLineas & " Where " & Replace(cadWhere, NombreTabla, NomTablaLineas)
    If RegistrosAListar(SQL) = 0 Then Exit Sub
    
    Set vFactu = New CFactura
    vFactu.DtoPPago = CCur(ComprobarCero(Text1(15).Text))
    vFactu.DtoGnral = CCur(ComprobarCero(Text1(16).Text))
    vFactu.Cliente = Text1(4).Text
    If vFactu.CalcularDatosFactura(cadWhere, NombreTabla, NomTablaLineas, False) Then
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
        Text3(56).Text = vFactu.BaseImp
        Text3(55).Text = vFactu.TotalFac
        
        
        'Recargos de equivalencia
        Text3(49).Text = vFactu.PorceIVA1RE
        Text3(50).Text = vFactu.PorceIVA2RE
        Text3(51).Text = vFactu.PorceIVA3RE
        Text3(52).Text = vFactu.ImpIVA1RE
        Text3(53).Text = vFactu.ImpIVA2RE
        Text3(54).Text = vFactu.ImpIVA3RE
        
        
        
        FormatoDatosTotales
    Else
        MuestraError Err.Number, "Calculando Totales", Err.Description
    End If
    Set vFactu = Nothing
End Sub


Private Function FormatoDatosTotales()
Dim I As Byte

    For I = 33 To 36
        Text3(I).Text = QuitarCero(Text3(I).Text)
        Text3(I).Text = Format(Text3(I).Text, FormatoImporte)
    Next I
 
    For I = 49 To 54
        Text3(I).Text = QuitarCero(Text3(I).Text)
        Text3(I).Text = Format(Text3(I).Text, FormatoImporte)
    Next I
 
 
    'Desglose B.Imponible por IVA
    For I = 43 To 45
        If Text3(I).Text <> "" Then
             If CSng(Text3(I).Text) = 0 Then
                Text3(I).Text = QuitarCero(Text3(I).Text)
                Text3(I - 3).Text = QuitarCero(Text3(I - 3).Text)
                Text3(I - 6).Text = QuitarCero(Text3(I - 6).Text)
                Text3(I + 3).Text = QuitarCero(Text3(I).Text)
            Else
                Text3(I).Text = Format(Text3(I).Text, FormatoImporte)
                Text3(I - 3) = Format(Text3(I - 3).Text, FormatoDescuento)
                Text3(I + 3).Text = Format(Text3(I + 3).Text, FormatoImporte)
            End If
        End If
    Next I
    
    'TOTALES
    Text3(55).Text = Format(Text3(55).Text, FormatoImporte)
    Text3(56).Text = Format(Text3(56).Text, FormatoImporte)
End Function



Private Function PonerDptoEnCliente() As Boolean
Dim vClien As CCliente
Dim NomDpto As String

    Set vClien = New CCliente
    vClien.Codigo = Text1(4).Text
    
    'si existe el departamento para el cliente
    If Trim(Text1(12).Text) <> "" Then
        If vClien.DptoCliente(Text1(12).Text, NomDpto) Then
            Text2(12).Text = NomDpto
            Text1(31).Text = ""
            PonerDptoEnCliente = True
        Else
            PonerDptoEnCliente = False
        End If
    End If
    
    If Text1(31).Text = "" Then
        Text1(31).Text = vClien.Obtener_EMailConfirmacion(Text1(12).Text)
    End If
    
    Set vClien = Nothing
End Function



Private Sub ComprobarRefObligatoria()
Dim vClien As CCliente

    Set vClien = New CCliente
    vClien.Codigo = Text1(4).Text
    If vClien.TieneRefObligatoria(Text1(13).Text) Then
        If Text1(13).Text = "" Then PonerFoco Text1(13)
    End If
    Set vClien = Nothing
End Sub


'Para obtener los dtos por cantidad lo que hace es a partir de un
'subtring del articulo(poscion 3 a 6) va a sdesca con la suma de la cantidad
'si en sdesca y dentro de los desde /hasta cantidad encuentra un dto lo aplica


Private Sub DescuentosCantidad(Articulo As String)
Dim cad As String
Dim R As ADODB.Recordset
Dim NuevoDto As Currency
Dim Importe As Currency
Dim bAct As Boolean

    On Error GoTo EDescuentosCantidad
    
    If Not vParamAplic.DtoxCantidad Then Exit Sub ' ---- [14/09/2009] (LAURA)
    
    If MsgBox("¿Desea recalcular los descuentos por cantidad?", vbQuestion + vbYesNo) = vbYes Then    'masl 140909
    
        'Si no  tenemos portes, ni nos pasamos
        If vParamAplic.TipoPortes <> 1 Then Exit Sub
    
    
        Espera 0.2
        Set miRsAux = New ADODB.Recordset
        Set R = New ADODB.Recordset
    
        'variable articulo:
        'Si tiene valor es para no tener que recalcular todos los valores del albaran, solo los
        ' del substring() del articulo que acabamos de insertar/actualizar o eliminar
        ' Si no lleva nada recalcular los dtos para todas la lineas
        cad = " WHERE numpedcl = " & Text1(0).Text
        cad = "select substring(codartic,3,4) raiz,sum(cantidad) suma from " & NomTablaLineas & cad
        If Articulo <> "" Then cad = cad & " AND substring(codartic,3,4)= '" & Mid(Articulo, 3, 4) & "'"
        'Y origen PRECIO no es precio especial
        cad = cad & " AND origpre <> 'E'"
        cad = cad & " group by 1"
        miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
                cad = TransformaComasPuntos(CStr(miRsAux!Suma))
                cad = "select * from sdesca where desdecan <=" & cad & " and " & cad & " <= hastacan and envagran = '"
                cad = cad & miRsAux!raiz & "'"
                R.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                cad = ""
                If Not R.EOF Then cad = R!dtolinea
                R.Close
                
                If cad <> "" Then
                    'OK tiene nuevo descuento
                    NuevoDto = CCur(cad)
                    
                    'Cojo los articulos del albaran y le meto el dto
                    cad = " WHERE numpedcl = " & Text1(0).Text
                    cad = "select * from " & NomTablaLineas & cad
                    '                                 a partir de la 3era posicion
                    cad = cad & " AND codartic like '__" & miRsAux!raiz & "%'"
                    R.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                    While Not R.EOF
                        '-- comprobar si admite descuento
                        If R!origpre = "T" Then
                            cad = DevuelveDesdeBDNew(conAri, "sclien", "codtarif", "codclien", Text1(4).Text, "N")
                            cad = DevuelveDesdeBDNew(conAri, "slista", "dtopermi", "codartic", R!codArtic, "T", , "codlista", cad, "N")
                            bAct = (cad = "1")
                        ElseIf R!origpre = "A" Or R!origpre = "M" Then
                            bAct = True
                        Else
                            bAct = False
                        End If
                        
                        If bAct Then
                            cad = CalcularImporte(CStr(R!cantidad), CStr(R!precioar), CStr(NuevoDto), CStr(R!dtoline2), vParamAplic.TipoDtos)
                            Importe = CCur(cad)
                            cad = "UPDATE " & NomTablaLineas & " set dtoline1=" & TransformaComasPuntos(CStr(NuevoDto))
                            cad = cad & ", importel = " & TransformaComasPuntos(CStr(Importe))
                            cad = cad & " WHERE numpedcl = " & Text1(0).Text
                            cad = cad & " and numlinea = " & R!numlinea
                            conn.Execute cad
                        End If
                        'Siguiente
                        R.MoveNext
                    Wend
                    R.Close
                    
                End If
                'sig
                miRsAux.MoveNext
        Wend
        miRsAux.Close
    End If 'masl
    
EDescuentosCantidad:
    If Err.Number <> 0 Then MuestraError Err.Number, "DescuentosxCantidad"
    Set miRsAux = Nothing
    Set R = Nothing
End Sub




Private Function SumaKilosLineas(Optional ImporteL As Currency) As Currency
Dim C As String
    On Error GoTo ESumaKilosLineas
    SumaKilosLineas = 0
    Set miRsAux = New ADODB.Recordset
    C = " WHERE " & NomTablaLineas & ".numpedcl = " & Text1(0).Text
    C = C & " AND " & NomTablaLineas & ".codartic=sartic.codartic"
    C = C & " AND " & NomTablaLineas & ".codartic <> " & DBSet(vParamAplic.ArtPortesN, "T")
    C = "select sum(cantidad*pesoarti),sum(importel) from " & NomTablaLineas & ",sartic " & C
    
    
    'El enlzace
    
    miRsAux.Open C, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        SumaKilosLineas = DBLet(miRsAux.Fields(0), "N")
        ImporteL = DBLet(miRsAux.Fields(1), "N")
    End If
    miRsAux.Close
    
    
    'Fijo la zona y la ruta del cliente
    
    RutaCliente = -1
    ZonaCliente = -1
    C = "Select codzonas,codrutas from sclien where codclien = " & Val(Text1(4).Text)
    miRsAux.Open C, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        ZonaCliente = DBLet(miRsAux!codzonas, "N")
        RutaCliente = DBLet(miRsAux!codrutas, "N")
    End If
    miRsAux.Close
    
    
ESumaKilosLineas:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set miRsAux = Nothing
End Function



'Si devuelve cero nada
'si devuelve >0 marcara la linea de portes
Private Function HacerAccionesPortes() As Integer
Dim ImporteLineas As Currency
Dim KilosAhora As Currency
Dim C As String
Dim CodEnvio As Integer
Dim PrecioKilo As Currency
Dim DtoPorte As Currency
Dim DesdeKilo As Currency
Dim ImporteL_Portes As Currency

    HacerAccionesPortes = 0
    KilosAhora = SumaKilosLineas(ImporteLineas)
    
    
    ' Si no cambia los kilos me salgo
    '-----------------------------------------------
    If KilosAhora = KilosAnteriores Then Exit Function
    If Data2.Recordset.EOF Then Exit Function
    
    If MsgBox("Desea recalcular los portes?", vbQuestion + vbYesNo) = vbNo Then Exit Function
    
    
    Set miRsAux = New ADODB.Recordset
    
    
    If ZonaCliente > 0 Then
        'Ha encontrado la zona /ruta. Miro en sportes
        C = "select sporte.codenvio,nomenvio,PrecioKg,desdekgs from sporte,senvio where sporte.codenvio=senvio.codenvio "
        C = C & " AND codcentr = " & ZonaCliente
        'Los kilos  hastakgs
        C = C & " AND desdekgs <= " & TransformaComasPuntos(CStr(KilosAhora))
        C = C & " AND hastakgs >= " & TransformaComasPuntos(CStr(KilosAhora))
        C = C & " group by sporte.codenvio"
        miRsAux.Open C, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        NumRegElim = 0
        CodEnvio = 0
        If Not miRsAux.EOF Then
            'Por si acaso hay mas de uno
            CadenaDesdeOtroForm = ""
            While Not miRsAux.EOF
                CodEnvio = miRsAux!CodEnvio
                PrecioKilo = miRsAux!preciokg
                DesdeKilo = DBLet(miRsAux!DesdeKgs, "N")
                NumRegElim = NumRegElim + 1
                CadenaDesdeOtroForm = CadenaDesdeOtroForm & miRsAux!CodEnvio & "<" & miRsAux!nomEnvio & "<" & miRsAux!preciokg & "<" & DBLet(miRsAux!DesdeKgs, "N") & "|"
                miRsAux.MoveNext
            Wend
            
            
            If NumRegElim > 1 Then
                'Mostraremos un form para que seleccione la opcion correspondiente
                frmVarios.Opcion = 3
                frmVarios.Show vbModal
                If CadenaDesdeOtroForm <> "" Then
                    C = RecuperaValor(CadenaDesdeOtroForm, 1)
                    CodEnvio = Val(C)
                    
                    C = RecuperaValor(CadenaDesdeOtroForm, 3)
                    PrecioKilo = CCur(C)
                    
                    DesdeKilo = CCur(RecuperaValor(CadenaDesdeOtroForm, 4))
                End If
            Else
                    CadenaDesdeOtroForm = Replace(CadenaDesdeOtroForm, "<", "|")
                    CadenaDesdeOtroForm = RecuperaValor(CadenaDesdeOtroForm, 2)
            
            End If
            
        End If
        miRsAux.Close
        
        
        'Dto en portes
        DtoPorte = 0
        ImporteL_Portes = 0
        If RutaCliente = 1 Or RutaCliente = 3 Or RutaCliente = 4 Then DtoPorte = vParamAplic.AbonoKilos
        If RutaCliente = 1 Or RutaCliente = 2 Then PrecioKilo = 0
        If RutaCliente = 4 And ImporteLineas < vParamAplic.ImporteMinimo Then 'importe pedido menor que importe minimo todo a cero(preciokilo, dtokilo)
               PrecioKilo = 0
               DtoPorte = 0
               ImporteL_Portes = 0
        Else
            If RutaCliente = 4 Then ImporteL_Portes = PrecioKilo
        End If
        
        If DesdeKilo = 1 Then
            If RutaCliente <> 4 Then
                ImporteL_Portes = PrecioKilo
                KilosAhora = 1
            End If
        Else
            ImporteL_Portes = (PrecioKilo - DtoPorte) * KilosAhora
        End If
        If RutaCliente <> 1 And ImporteL_Portes < 0 Then ImporteL_Portes = 0 ' masl 090709
        
        'Ahora compruebo si tiene la linea de portes para aplicarle el importe
        C = " WHERE " & NomTablaLineas & ".numpedcl = " & Text1(0).Text
        C = "Select numlinea from " & NomTablaLineas & C & " and codartic ='" & vParamAplic.ArtPortesN & "'"
        miRsAux.Open C, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        NumRegElim = 0
        If Not miRsAux.EOF Then NumRegElim = DBLet(miRsAux!numlinea, "N")
        miRsAux.Close
        
        
        'SI ya existe la borro, para que siempre aparezca al final
        If NumRegElim > 0 Then
            C = " WHERE numpedcl = " & Text1(0).Text
            C = C & " AND numlinea = " & NumRegElim
            C = "DELETE FROM " & NomTablaLineas & C
            conn.Execute C
            Espera 0.1
            
        
        End If
        
       'If RutaCliente <> 1 And ImporteL_Portes < 0 Then ImporteL_Portes = 0 masl 090709
        
        
            'Si el precio es mayor k cero entonces SI pongo la linea
            C = " WHERE " & NomTablaLineas & ".numpedcl = " & Text1(0).Text
            C = "select codalmac,max(numlinea) from " & NomTablaLineas & C
            C = C & " GROUP BY codalmac"
            miRsAux.Open C, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If miRsAux.EOF Then
                MsgBox "NO deberia haberse producido", vbExclamation
                Exit Function
            End If
            NumRegElim = miRsAux.Fields(1) + 1
            HacerAccionesPortes = NumRegElim
    '            SQL = "INSERT INTO " & NomTablaLineas
    '            SQL = SQL & "(codtipom, numalbar,numlinea, codalmac, codartic, nomartic, ampliaci, cantidad, precioar, dtoline1, dtoline2, importel, origpre,codprovex) "
    '            SQL = SQL & "VALUES ('" & Text1(30).Text & "', " & Val(Text1(0).Text) & ", " & NumRegElim & ", "
            
            C = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", vParamAplic.ArtPortesN, "T")
            C = DevNombreSQL(C)
            C = miRsAux!codAlmac & ",'" & vParamAplic.ArtPortesN & "','" & C & "','"
            
            'Esto es propio de la tabla de lineas
            C = "INSERT INTO " & NomTablaLineas & "(numpedcl,numlinea,codalmac,codartic,nomartic,ampliaci,cantidad,servidas,precioar,dtoline1,dtoline2,importel,origpre)" & _
                " VALUES (" & Val(Text1(0).Text) & ", " & NumRegElim & ", " & C
            
            'ampliaci,cantidad,servidas,precioar,dtoline1,dtoline2,importel,origpre
            'Amplicacion
            C = C & CadenaDesdeOtroForm & "',"
            
          If RutaCliente <> 1 And RutaCliente <> 3 And RutaCliente <> 4 Then   'masl 090709
            'Cantidad, SERVIDAS,precioar dto1 dto2
            C = C & TransformaComasPuntos(CStr(KilosAhora)) & ",0," & TransformaComasPuntos(CStr(PrecioKilo))
            C = C & "," & TransformaComasPuntos(CStr(DtoPorte)) & ",0,"
            
                                                                        Else    'masl 090709
            C = C & TransformaComasPuntos(CStr(KilosAhora)) & ",0," & TransformaComasPuntos(CStr(DtoPorte * (-1)))
            C = C & ",0" & ",0,"
                                                                        
          End If
          
                        
            'importel
            C = C & TransformaComasPuntos(CStr(ImporteL_Portes))
            
            'origpre
            C = C & ",'M')"
        
            'Noviembre 2009.    Enero 2010.  SIEMPRE hay que meter la linea de portes
            'If ImporteL_Portes <> 0 Then conn.Execute C
            conn.Execute C
        
    End If
            
End Function



Private Sub AbrirForm_CentroCoste()
    Screen.MousePointer = vbHourglass
    cmdAux(0).Tag = "2"

    Set frmB = New frmBuscaGrid
    If vParamAplic.ContabilidadNueva Then
        frmB.vCampos = "Codigo|ccoste|codccost|T||20·Descripción|ccoste|nomccost|T||70·"
        frmB.vTabla = "ccoste"
    Else
        frmB.vCampos = "Codigo|cabccost|codccost|T||20·Descripción|cabccost|nomccost|T||70·"
        frmB.vTabla = "cabccost"
    End If
    frmB.vSQL = ""
    HaDevueltoDatos = False
    '###A mano
    frmB.vDevuelve = "0|1|"
    frmB.vTitulo = "Centros de coste"
    frmB.vselElem = 0
    frmB.vConexionGrid = conConta
    
    frmB.Show vbModal
    Set frmB = Nothing
    cmdAux(0).Tag = "-1"
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

Private Sub UpdateaNomDirec()
Dim N As Integer
Dim Ol As Integer
Dim C As String

    N = -1
    If Not IsNull(Data1.Recordset!CodDirec) Then N = Data1.Recordset!CodDirec
    
    Ol = -1
    If Text1(12).Text <> "" Then Ol = CInt(Text1(12).Text)
    
    If N <> Ol Then
        If Ol < 0 Then
            C = "UPDATE scaped set nomdirec=NULL "
        Else
            C = "UPDATE scaped SET nomdirec=" & DBSet(Text2(12).Text, "T")
        End If
        C = C & " WHERE numpedcl = " & Text1(0).Text
        ejecutar C, False
    End If
End Sub





'Nuevo. Cuando pulse MAS (y es el primer carcater abre el prismatico asociado)
Private Sub PulsarTeclaMas(InsertandoCabecera As Boolean, index As Integer)

    If InsertandoCabecera Then
        EsCabecera = True
        If imgBuscar(index).visible Then imgBuscar_Click index
        
    Else
        'Lineas
        EsCabecera = False
        If index = 11 Then index = 2
        cmdAux_Click index
        
        
    End If
        
End Sub





Private Sub PonerCampoActuacion()
Dim CADENA As String
            If Modo = 1 Then Exit Sub
            CADENA = ""
            If Text1(32).Text <> "" Then
                Text1(32).Text = UCase(Text1(32).Text)
                If Text1(4).Text = "" Or Text1(12).Text = "" Then
                    MsgBox "Falta cliente/obra", vbExclamation
                    Text1(32).Text = ""
                Else
                    CADENA = "codclien =" & Text1(4).Text & " AND coddirec= " & Text1(12).Text & " AND actuacion "
                
                    CADENA = DevuelveDesdeBDNew(conAri, "sactuaobra", "concat(fechaini,' ',observa)", CADENA, Text1(32).Text, "N")
                    If CADENA = "" Then
                        MsgBox "Ninguna actuacion con ese valor:" & Text1(32).Text, vbInformation
                        Text1(32).Text = ""
                    End If
                End If
                
            End If
            Text2(32).Text = CADENA

End Sub





Private Sub PonerForaGrid()
    'Dim RS As ADODB.Recordset
    'Dim SQL As String
    Dim Borrar As Boolean
    Dim J As Integer
    Dim Desde As Integer
    Dim Base As Integer
    Dim C As String
    
On Error GoTo Error1
  
        Borrar = True
  
        
        'Nuevo SAIL. codtipom numalbar numlinea
        'SQL = "select codcapit,codtipor, codtipor as codtraba,precoste,ampliaci from " & NomTablaLineas & Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas) & " and numlinea=" & Data2.Recordset!numlinea
        If Not Data2.Recordset.EOF Then
            Borrar = False
            
            Text2(16).Text = DBLet(Data2.Recordset!Ampliaci, "T")
        
            J = ModificaLineas
            
            
            C = DBLet(Data2.Recordset!codcapit, "T")
            If txtaux(13).Text <> C Then
                txtaux(13).Text = C
                PonerDatosNuevosLineaAlbaran False, 13
            End If
            
            
                 
            C = DBLet(Data2.Recordset!codtipor, "T")
            If txtaux(12).Text <> C Then
                txtaux(12).Text = C
                PonerDatosNuevosLineaAlbaran False, 12
            End If
  
            
            
            If vEmpresa.TieneAnalitica Then
                '- centro de coste
                ' ---- [19/10/2009] [LAURA]: añadir campo centro de coste familia
                Me.txtaux(11).Text = DBLet(Data2.Recordset!CodCCost, "T")
                Me.txtAux2(11).Text = PonerNombreCCoste(Me.txtaux(11))
            End If
        
            ModificaLineas = J

      Else
        'EOF
        For J = 11 To 13
            txtaux(J).Text = ""
            txtAux2(J).Text = ""
        Next
        Text2(16).Text = ""
        
      End If   'De EOF
        
    

    
    
    

    
Error1:
    If Err.Number <> 0 Then
        MsgBox Err.Number & ": " & Err.Description, vbExclamation
        Borrar = True
    End If
    
    If Borrar Then
        For J = 11 To 13
            txtaux(J).Text = ""
            txtAux2(J).Text = ""
            
        Next
        Text2(16).Text = ""
    End If

End Sub




Private Sub PonerDatosNuevosLineaAlbaran(Edicion As Boolean, index As Integer)
Dim devuelve As String

       devuelve = ""
            
                'Si es numerico
                'ORDEN TRABAJO=13
                
                If index = 13 Then
                   
                    If txtaux(index).Text <> "" Then
                        If Not EsNumerico(txtaux(index).Text) Then
                            txtaux(index).Text = ""
                            If Edicion Then PonerFoco txtaux(index)
                        End If
                    End If

                End If
                
                If txtaux(index).Text <> "" Then
                    If index = 13 Then
                        'codcapit nomcapit scapitulos
                        devuelve = DevuelveDesdeBD(conAri, "nomcapit", "scapitulos", "codcapit", txtaux(index).Text, "N")
                    ElseIf index = 12 Then
                        'stipor codtipor nomtipor
                        devuelve = DevuelveDesdeBD(conAri, "nomtipor", "stipor", "codtipor", txtaux(index).Text, "T")
                   ' Else
                   '     devuelve = DevuelveDesdeBD(conAri, "nomtraba", "straba", "codtraba", txtAux(Index).Text, "N")
                    End If
                    If devuelve = "" Then
                        MsgBox "No existe el registro para el campo: " & txtaux(index).Text & " en la tabla de " & txtaux(index).Tag, vbExclamation
                        txtaux(index).Text = ""
                        If Edicion Then PonerFoco txtaux(index)
                    End If
                End If
                
                txtAux2(index).Text = devuelve
                


End Sub
