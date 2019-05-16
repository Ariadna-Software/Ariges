VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmComEntPedidosSa 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pedidos Proveedor"
   ClientHeight    =   8805
   ClientLeft      =   45
   ClientTop       =   4335
   ClientWidth     =   14400
   Icon            =   "frmComEntPedidosSa.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8805
   ScaleWidth      =   14400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   910
      Left            =   120
      TabIndex        =   93
      Top             =   420
      Width           =   14055
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   5
         Left            =   10200
         MaxLength       =   40
         TabIndex        =   5
         Tag             =   "Nombre Proveedor|T|N|||scappr|nomprove||N|"
         Top             =   420
         Width           =   3645
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   4
         Left            =   9360
         MaxLength       =   30
         TabIndex        =   4
         Tag             =   "Cod. Proveedor|N|N|0|999999|scappr|codprove|000000|N|"
         Text            =   "Text1"
         Top             =   420
         Width           =   780
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   3
         Left            =   4680
         MaxLength       =   30
         TabIndex        =   3
         Tag             =   "Realizado Por|N|N|0|9999|scappr|codtraba|0000|N|"
         Text            =   "Text1"
         Top             =   420
         Width           =   780
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   3
         Left            =   5520
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   96
         Text            =   "Text2"
         Top             =   420
         Width           =   3645
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   1
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Fecha Pedido|F|N|||scappr|fecpedpr|dd/mm/yyyy|N|"
         Top             =   420
         Width           =   1185
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000013&
         Height          =   315
         Index           =   0
         Left            =   240
         MaxLength       =   7
         TabIndex        =   0
         Tag             =   "Nº Pedido|N|S|0||scappr|numpedpr|0000000|S|"
         Text            =   "Text1 7"
         Top             =   420
         Width           =   1125
      End
      Begin VB.CheckBox chkRestoPed 
         Caption         =   "Resto de Pedido"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2880
         TabIndex        =   2
         Tag             =   "Resto de Pedido|N|N|||scappr|restoped||N|"
         Top             =   360
         Width           =   1575
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   10200
         Picture         =   "frmComEntPedidosSa.frx":000C
         ToolTipText     =   "Buscar proveedor"
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Proveedor"
         Height          =   255
         Index           =   0
         Left            =   9360
         TabIndex        =   98
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Realizado Por"
         Height          =   255
         Index           =   21
         Left            =   4680
         TabIndex        =   97
         Top             =   240
         Width           =   1095
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   5760
         Picture         =   "frmComEntPedidosSa.frx":010E
         ToolTipText     =   "Buscar trabajador"
         Top             =   165
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Ped."
         Height          =   255
         Index           =   14
         Left            =   1440
         TabIndex        =   95
         Top             =   225
         Width           =   855
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   2400
         Picture         =   "frmComEntPedidosSa.frx":0210
         ToolTipText     =   "Buscar fecha"
         Top             =   180
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Pedido"
         Height          =   255
         Index           =   50
         Left            =   240
         TabIndex        =   94
         Top             =   220
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   475
      Index           =   0
      Left            =   120
      TabIndex        =   37
      Top             =   8175
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
      Left            =   13170
      TabIndex        =   20
      Top             =   8280
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   12000
      TabIndex        =   19
      Top             =   8280
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   9960
      Top             =   1440
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
      Width           =   14400
      _ExtentX        =   25400
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
            Object.ToolTipText     =   "Modificar descuento"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir Pedido"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Simular otro proveedor"
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Último"
            ImageIndex      =   9
         EndProperty
      EndProperty
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
         ForeColor       =   &H000000C0&
         Height          =   285
         Index           =   0
         Left            =   11640
         MaxLength       =   15
         TabIndex        =   136
         Text            =   "TOTAL"
         Top             =   100
         Width           =   885
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0FF&
         Height          =   285
         Index           =   50
         Left            =   12600
         MaxLength       =   15
         TabIndex        =   135
         Text            =   "Text1 7"
         Top             =   80
         Width           =   1530
      End
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   6480
         TabIndex        =   40
         Top             =   120
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc Data2 
      Height          =   330
      Left            =   8520
      Top             =   1440
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
      Height          =   6705
      Left            =   120
      TabIndex        =   41
      Top             =   1350
      Width           =   14085
      _ExtentX        =   24844
      _ExtentY        =   11827
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
      TabPicture(0)   =   "frmComEntPedidosSa.frx":029B
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(35)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(46)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(29)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(34)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(43)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "imgBuscar2(9)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "imgBuscar2(10)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "imgBuscar2(11)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Line2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Line3"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label1(51)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label1(52)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "imgBuscar2(0)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "imgBuscar2(1)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label1(53)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "imgBuscar2(2)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "imgAmpliaci"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "DataGrid1"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtAux(0)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtAux(1)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtAux(3)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txtAux(4)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txtAux(5)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "txtAux(6)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "txtAux(7)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "txtAux(2)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "cmdAux(0)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "cmdAux(1)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "FrameCliente"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "txtAux(8)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "cmdAux(2)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Text2(16)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "txtAux2(8)"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "txtAux(9)"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "txtAux(10)"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "txtAux(11)"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "txtDesc(9)"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "txtDesc(10)"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "txtDesc(11)"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "txtDesc(0)"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "txtAux(14)"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "txtAux(13)"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "txtAux(12)"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "txtAux(15)"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).ControlCount=   44
      TabCaption(1)   =   "Otros Datos"
      TabPicture(1)   =   "frmComEntPedidosSa.frx":02B7
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1(45)"
      Tab(1).Control(1)=   "Label1(28)"
      Tab(1).Control(2)=   "imgBuscar(11)"
      Tab(1).Control(3)=   "Label1(44)"
      Tab(1).Control(4)=   "imgFecha(1)"
      Tab(1).Control(5)=   "Label1(47)"
      Tab(1).Control(6)=   "Label1(48)"
      Tab(1).Control(7)=   "Label1(49)"
      Tab(1).Control(8)=   "imgBuscar(12)"
      Tab(1).Control(9)=   "Text1(17)"
      Tab(1).Control(10)=   "Text1(18)"
      Tab(1).Control(11)=   "Text1(19)"
      Tab(1).Control(12)=   "Text1(20)"
      Tab(1).Control(13)=   "Text1(21)"
      Tab(1).Control(14)=   "FrameDirMercancia"
      Tab(1).Control(15)=   "FrameDirFactura"
      Tab(1).Control(16)=   "FrameHco"
      Tab(1).Control(17)=   "Text1(27)"
      Tab(1).Control(18)=   "Text2(27)"
      Tab(1).Control(19)=   "Text1(28)"
      Tab(1).Control(20)=   "Text1(29)"
      Tab(1).Control(21)=   "Text1(30)"
      Tab(1).Control(22)=   "Text1(31)"
      Tab(1).Control(23)=   "Text2(4)"
      Tab(1).ControlCount=   24
      TabCaption(2)   =   "Totales"
      TabPicture(2)   =   "frmComEntPedidosSa.frx":02D3
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FrameFactura"
      Tab(2).ControlCount=   1
      Begin VB.TextBox txtAux 
         BackColor       =   &H80000013&
         Height          =   315
         Index           =   15
         Left            =   11520
         MaxLength       =   10
         TabIndex        =   164
         Tag             =   "pedido"
         Text            =   "cc"
         Top             =   4440
         Width           =   1095
      End
      Begin VB.TextBox txtAux 
         BackColor       =   &H80000013&
         Height          =   315
         Index           =   12
         Left            =   10560
         MaxLength       =   3
         TabIndex        =   51
         Tag             =   "codtipom"
         Text            =   "Codtipom"
         Top             =   3240
         Width           =   495
      End
      Begin VB.TextBox txtAux 
         BackColor       =   &H80000013&
         Height          =   315
         Index           =   13
         Left            =   11160
         MaxLength       =   20
         TabIndex        =   52
         Tag             =   "numeroalb"
         Text            =   "numeroalb"
         Top             =   3240
         Width           =   975
      End
      Begin VB.TextBox txtAux 
         BackColor       =   &H80000013&
         Height          =   315
         Index           =   14
         Left            =   12600
         MaxLength       =   20
         TabIndex        =   53
         Tag             =   "fechaalb"
         Text            =   "99/99/9999"
         Top             =   3240
         Width           =   1095
      End
      Begin VB.TextBox txtDesc 
         BackColor       =   &H80000018&
         Height          =   675
         Index           =   0
         Left            =   10560
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   161
         Text            =   "frmComEntPedidosSa.frx":02EF
         Top             =   3600
         Width           =   3405
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   4
         Left            =   -66480
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   159
         Text            =   "Text2"
         Top             =   2760
         Width           =   4350
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   31
         Left            =   -67320
         MaxLength       =   30
         TabIndex        =   25
         Tag             =   "Envio|N|S|0|999|scappr|codenvio|0000|N|"
         Text            =   "Text1"
         Top             =   2760
         Width           =   660
      End
      Begin VB.TextBox Text1 
         Height          =   280
         Index           =   30
         Left            =   -67920
         MaxLength       =   80
         TabIndex        =   27
         Tag             =   "T|T|S|||scappr|SReferencia||N|"
         Top             =   3480
         Width           =   6525
      End
      Begin VB.TextBox Text1 
         Height          =   280
         Index           =   29
         Left            =   -74760
         MaxLength       =   80
         TabIndex        =   26
         Tag             =   "O|T|S|||scappr|NReferencia||N|"
         Top             =   3480
         Width           =   6525
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   28
         Left            =   -68880
         MaxLength       =   10
         TabIndex        =   24
         Tag             =   "Fecha entraga|F|S|||scappr|fecentrega|dd/mm/yyyy|N|"
         Top             =   2760
         Width           =   1185
      End
      Begin VB.TextBox txtDesc 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   11
         Left            =   12120
         Locked          =   -1  'True
         TabIndex        =   155
         Text            =   "Text4"
         Top             =   3960
         Width           =   1575
      End
      Begin VB.TextBox txtDesc 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   10
         Left            =   11400
         Locked          =   -1  'True
         TabIndex        =   154
         Text            =   "Text4"
         Top             =   3240
         Width           =   2535
      End
      Begin VB.TextBox txtDesc 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   9
         Left            =   11400
         Locked          =   -1  'True
         TabIndex        =   153
         Text            =   "Text4"
         Top             =   2520
         Width           =   2535
      End
      Begin VB.TextBox txtAux 
         BackColor       =   &H80000013&
         Height          =   315
         Index           =   11
         Left            =   10560
         MaxLength       =   20
         TabIndex        =   55
         Tag             =   "actuacion"
         Top             =   3960
         Width           =   1575
      End
      Begin VB.TextBox txtAux 
         BackColor       =   &H80000013&
         Height          =   315
         Index           =   10
         Left            =   10560
         MaxLength       =   10
         TabIndex        =   54
         Tag             =   "Obra"
         Text            =   "cc"
         Top             =   3240
         Width           =   735
      End
      Begin VB.TextBox txtAux 
         BackColor       =   &H80000013&
         Height          =   315
         Index           =   9
         Left            =   10560
         MaxLength       =   10
         TabIndex        =   50
         Tag             =   "cliente"
         Text            =   "cc"
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox txtAux2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   8
         Left            =   11280
         Locked          =   -1  'True
         MaxLength       =   60
         TabIndex        =   148
         Text            =   "nom ccoste"
         Top             =   5160
         Visible         =   0   'False
         Width           =   2685
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   795
         Index           =   16
         Left            =   10560
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   57
         Top             =   5760
         Width           =   3405
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   27
         Left            =   -73080
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   145
         Text            =   "Text2"
         Top             =   2760
         Width           =   3510
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   27
         Left            =   -73680
         MaxLength       =   30
         TabIndex        =   23
         Tag             =   "Direc. recogida|N|S|0|999|scappr|coddirre|000|N|"
         Text            =   "Text1"
         Top             =   2760
         Width           =   540
      End
      Begin VB.CommandButton cmdAux 
         Appearance      =   0  'Flat
         Caption         =   "+"
         Height          =   240
         Index           =   2
         Left            =   11520
         TabIndex        =   143
         ToolTipText     =   "Buscar centro coste"
         Top             =   4920
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.TextBox txtAux 
         BackColor       =   &H80000013&
         Height          =   315
         Index           =   8
         Left            =   10560
         MaxLength       =   4
         TabIndex        =   56
         Tag             =   "centro coste"
         Text            =   "cc"
         Top             =   5160
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Frame FrameHco 
         Caption         =   "Datos  Eliminación"
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
         Height          =   2085
         Left            =   -66600
         TabIndex        =   137
         Top             =   4080
         Width           =   4935
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   24
            Left            =   720
            MaxLength       =   10
            TabIndex        =   33
            Top             =   230
            Width           =   1185
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   25
            Left            =   120
            MaxLength       =   30
            TabIndex        =   34
            Text            =   "Text1"
            Top             =   960
            Width           =   780
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   25
            Left            =   960
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   139
            Text            =   "Text2"
            Top             =   960
            Width           =   3285
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   26
            Left            =   135
            MaxLength       =   30
            TabIndex        =   35
            Text            =   "Text1"
            Top             =   1680
            Width           =   540
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   26
            Left            =   675
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   138
            Text            =   "Text2"
            Top             =   1680
            Width           =   2685
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha"
            Height          =   255
            Index           =   37
            Left            =   120
            TabIndex        =   142
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Trabajador"
            Height          =   255
            Index           =   38
            Left            =   75
            TabIndex        =   141
            Top             =   670
            Width           =   825
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   9
            Left            =   960
            Picture         =   "frmComEntPedidosSa.frx":02FF
            ToolTipText     =   "Buscar trabajador"
            Top             =   660
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Incidencia"
            Height          =   255
            Index           =   40
            Left            =   120
            TabIndex        =   140
            Top             =   1485
            Width           =   825
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   10
            Left            =   1080
            Picture         =   "frmComEntPedidosSa.frx":0401
            ToolTipText     =   "Buscar incidencia"
            Top             =   1440
            Width           =   240
         End
      End
      Begin VB.Frame FrameFactura 
         Height          =   3300
         Left            =   -73080
         TabIndex        =   103
         Top             =   1560
         Width           =   10575
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   33
            Left            =   240
            MaxLength       =   15
            TabIndex        =   120
            Text            =   "Text1 7"
            Top             =   555
            Width           =   1485
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   34
            Left            =   2160
            MaxLength       =   15
            TabIndex        =   119
            Text            =   "Text1 7"
            Top             =   555
            Width           =   1365
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   35
            Left            =   3960
            MaxLength       =   15
            TabIndex        =   118
            Text            =   "Text1 7"
            Top             =   555
            Width           =   1365
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   36
            Left            =   5760
            MaxLength       =   15
            TabIndex        =   117
            Text            =   "Text1 7"
            Top             =   555
            Width           =   1485
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   43
            Left            =   5760
            MaxLength       =   15
            TabIndex        =   116
            Text            =   "Text1 7"
            Top             =   1440
            Width           =   1485
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   37
            Left            =   4320
            MaxLength       =   4
            TabIndex        =   115
            Text            =   "Text1 7"
            Top             =   1440
            Width           =   525
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   40
            Left            =   5017
            MaxLength       =   5
            TabIndex        =   114
            Text            =   "Text1 7"
            Top             =   1440
            Width           =   570
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   46
            Left            =   7560
            MaxLength       =   15
            TabIndex        =   113
            Text            =   "Text1 7"
            Top             =   1440
            Width           =   1485
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   44
            Left            =   5760
            MaxLength       =   15
            TabIndex        =   112
            Text            =   "Text1 7"
            Top             =   1755
            Width           =   1485
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   38
            Left            =   4320
            MaxLength       =   4
            TabIndex        =   111
            Text            =   "Text1 7"
            Top             =   1755
            Width           =   525
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   41
            Left            =   5017
            MaxLength       =   5
            TabIndex        =   110
            Text            =   "Text1 7"
            Top             =   1755
            Width           =   570
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   47
            Left            =   7560
            MaxLength       =   15
            TabIndex        =   109
            Text            =   "Text1 7"
            Top             =   1755
            Width           =   1485
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   45
            Left            =   5760
            MaxLength       =   15
            TabIndex        =   108
            Text            =   "Text1 7"
            Top             =   2085
            Width           =   1485
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   39
            Left            =   4320
            MaxLength       =   4
            TabIndex        =   107
            Text            =   "Text1 7"
            Top             =   2085
            Width           =   525
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   42
            Left            =   5017
            MaxLength       =   5
            TabIndex        =   106
            Text            =   "Text1 7"
            Top             =   2085
            Width           =   570
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   48
            Left            =   7560
            MaxLength       =   15
            TabIndex        =   105
            Text            =   "Text1 7"
            Top             =   2085
            Width           =   1485
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0FF&
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
            Index           =   49
            Left            =   7200
            MaxLength       =   15
            TabIndex        =   104
            Text            =   "Text1 7"
            Top             =   2640
            Width           =   1845
         End
         Begin VB.Label Label1 
            Caption         =   "Base Imponible"
            Height          =   255
            Index           =   27
            Left            =   5760
            TabIndex        =   134
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Imp. Bruto"
            Height          =   255
            Index           =   24
            Left            =   240
            TabIndex        =   133
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Imp. Dto PP"
            Height          =   255
            Index           =   23
            Left            =   2160
            TabIndex        =   132
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Imp. Dto Gn"
            Height          =   255
            Index           =   22
            Left            =   3960
            TabIndex        =   131
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Base Imponible"
            Height          =   255
            Index           =   18
            Left            =   5760
            TabIndex        =   130
            Top             =   360
            Width           =   1215
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
            TabIndex        =   129
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
            TabIndex        =   128
            Top             =   480
            Width           =   135
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
            TabIndex        =   127
            Top             =   480
            Width           =   135
         End
         Begin VB.Label Label1 
            Caption         =   "Imp. IVA"
            Height          =   255
            Index           =   33
            Left            =   7560
            TabIndex        =   126
            Top             =   1230
            Width           =   1335
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
            Index           =   8
            Left            =   7320
            TabIndex        =   125
            Top             =   1320
            Width           =   135
         End
         Begin VB.Line Line1 
            X1              =   4320
            X2              =   7320
            Y1              =   1065
            Y2              =   1065
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
            TabIndex        =   124
            Top             =   2160
            Width           =   135
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
            ForeColor       =   &H000000C0&
            Height          =   285
            Index           =   39
            Left            =   5640
            TabIndex        =   123
            Top             =   2660
            Width           =   1530
         End
         Begin VB.Label Label1 
            Caption         =   "% IVA"
            Height          =   255
            Index           =   41
            Left            =   5040
            TabIndex        =   122
            Top             =   1230
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "Cod. IVA"
            Height          =   255
            Index           =   42
            Left            =   4320
            TabIndex        =   121
            Top             =   1200
            Width           =   735
         End
      End
      Begin VB.Frame FrameDirFactura 
         Caption         =   "Direc. Factura"
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
         Height          =   1900
         Left            =   -67320
         TabIndex        =   83
         Top             =   480
         Width           =   5175
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   2
            Left            =   1005
            MaxLength       =   30
            TabIndex        =   22
            Tag             =   "Direc. Factura|N|S|0|999|scappr|coddiref|000|N|"
            Text            =   "Text1"
            Top             =   360
            Width           =   540
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   2
            Left            =   1560
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   92
            Text            =   "Text2"
            Top             =   360
            Width           =   3510
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   24
            Left            =   1005
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   87
            Text            =   "Text1 Text1 Text1 Text1 Text22"
            Top             =   1425
            Width           =   2565
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   22
            Left            =   1005
            Locked          =   -1  'True
            MaxLength       =   6
            TabIndex        =   86
            Text            =   "Text15"
            Top             =   1065
            Width           =   630
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   23
            Left            =   1650
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   85
            Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwaq"
            Top             =   1065
            Width           =   3405
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   21
            Left            =   1005
            Locked          =   -1  'True
            MaxLength       =   35
            TabIndex        =   84
            Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwwwww aq"
            Top             =   720
            Width           =   4050
         End
         Begin VB.Image imgBuscar 
            Enabled         =   0   'False
            Height          =   240
            Index           =   5
            Left            =   720
            Picture         =   "frmComEntPedidosSa.frx":0503
            ToolTipText     =   "Buscar dirección"
            Top             =   360
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Provincia"
            Height          =   255
            Index           =   13
            Left            =   120
            TabIndex        =   91
            Top             =   1425
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Población"
            Height          =   255
            Index           =   12
            Left            =   120
            TabIndex        =   90
            Top             =   1035
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Domicilio"
            Height          =   255
            Index           =   11
            Left            =   120
            TabIndex        =   89
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Cod.Dir"
            Height          =   255
            Index           =   10
            Left            =   120
            TabIndex        =   88
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.Frame FrameDirMercancia 
         Caption         =   "Direc. Mercancia"
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
         Height          =   1665
         Left            =   -74760
         TabIndex        =   73
         Top             =   600
         Width           =   5175
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   17
            Left            =   1005
            Locked          =   -1  'True
            MaxLength       =   35
            TabIndex        =   78
            Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwwwww aq"
            Top             =   600
            Width           =   4050
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   19
            Left            =   1650
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   77
            Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwaq"
            Top             =   960
            Width           =   3405
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   18
            Left            =   1005
            Locked          =   -1  'True
            MaxLength       =   6
            TabIndex        =   76
            Text            =   "Text15"
            Top             =   960
            Width           =   630
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   20
            Left            =   1005
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   75
            Text            =   "Text1 Text1 Text1 Text1 Text22"
            Top             =   1320
            Width           =   2565
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   15
            Left            =   1005
            MaxLength       =   30
            TabIndex        =   21
            Tag             =   "Direc. Mercancia|N|S|0|999|scappr|coddirea|000|N|"
            Text            =   "Text1"
            Top             =   240
            Width           =   540
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   15
            Left            =   1560
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   74
            Text            =   "Text2"
            Top             =   240
            Width           =   3510
         End
         Begin VB.Label Label1 
            Caption         =   "Cod.Dir"
            Height          =   255
            Index           =   9
            Left            =   120
            TabIndex        =   82
            Top             =   240
            Width           =   600
         End
         Begin VB.Label Label1 
            Caption         =   "Domicilio"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   81
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Población"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   80
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Provincia"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   79
            Top             =   1320
            Width           =   735
         End
         Begin VB.Image imgBuscar 
            Enabled         =   0   'False
            Height          =   240
            Index           =   4
            Left            =   720
            Picture         =   "frmComEntPedidosSa.frx":0605
            ToolTipText     =   "Buscar dirección"
            Top             =   240
            Width           =   240
         End
      End
      Begin VB.Frame FrameCliente 
         Height          =   1770
         Left            =   240
         TabIndex        =   62
         Top             =   310
         Width           =   13695
         Begin VB.CheckBox chkObra 
            Caption         =   "Obra"
            Enabled         =   0   'False
            Height          =   255
            Left            =   4080
            TabIndex        =   12
            Tag             =   "Obra|N|N|||scappr|obra||N|"
            Top             =   1320
            Width           =   1215
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   1
            Left            =   8955
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   101
            Text            =   "Text2"
            Top             =   560
            Width           =   3525
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   23
            Left            =   8250
            MaxLength       =   30
            TabIndex        =   14
            Tag             =   "Solicitado Por|N|S|0|9999|scappr|codtrab1|0000|N|"
            Text            =   "Text1"
            Top             =   550
            Width           =   660
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   0
            Left            =   8955
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   99
            Text            =   "Text2"
            Top             =   190
            Width           =   3525
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   22
            Left            =   8250
            MaxLength       =   30
            TabIndex        =   13
            Tag             =   "Cliente|N|S|0|999999|scappr|codclien|000000|N|"
            Text            =   "Text1"
            Top             =   190
            Width           =   660
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   16
            Left            =   8280
            MaxLength       =   25
            TabIndex        =   16
            Tag             =   "Tipo Portes|T|S|||scappr|tipoporte||N|"
            Text            =   "Text1 wwwwwwwwwwwwwwwwwww"
            Top             =   1380
            Width           =   2445
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   11
            Left            =   1170
            MaxLength       =   30
            TabIndex        =   11
            Tag             =   "Provincia|T|N|||scappr|proprove||N|"
            Text            =   "Text1 Text1 Text1 Text1 Text22"
            Top             =   1300
            Width           =   2565
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   9
            Left            =   1170
            MaxLength       =   6
            TabIndex        =   9
            Tag             =   "CPostal|T|N|||scappr|codpobla||N|"
            Text            =   "Text15"
            Top             =   910
            Width           =   630
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   10
            Left            =   1820
            MaxLength       =   30
            TabIndex        =   10
            Tag             =   "Población|T|N|||scappr|pobprove||N|"
            Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwaq"
            Top             =   910
            Width           =   3525
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   7
            Left            =   3375
            MaxLength       =   20
            TabIndex        =   7
            Tag             =   "teléfono Proveedor|T|S|||scappr|telprove||N|"
            Text            =   "12345678911234567899"
            Top             =   190
            Width           =   1965
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   6
            Left            =   1170
            MaxLength       =   15
            TabIndex        =   6
            Tag             =   "NIF Proveedor|T|N|||scappr|nifprove||N|"
            Text            =   "123456789"
            Top             =   190
            Width           =   1230
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   12
            Left            =   8250
            MaxLength       =   30
            TabIndex        =   15
            Tag             =   "Forma de Pago|N|N|0|999|scappr|codforpa|000|N|"
            Text            =   "Text1"
            Top             =   910
            Width           =   660
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   12
            Left            =   8955
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   64
            Text            =   "Text2"
            Top             =   910
            Width           =   3525
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   13
            Left            =   11115
            MaxLength       =   7
            TabIndex        =   17
            Tag             =   "Descuento P.Pago|N|N|0|99.90|scaped|dtoppago|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   1390
            Width           =   525
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   14
            Left            =   12000
            MaxLength       =   7
            TabIndex        =   18
            Tag             =   "Descuento General|N|N|0|99.90|scaped|dtognral|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   1390
            Width           =   525
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   8
            Left            =   1170
            MaxLength       =   35
            TabIndex        =   8
            Tag             =   "Domicilio|T|N|||scappr|domprove||N|"
            Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwwwww aq"
            Top             =   550
            Width           =   4170
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   8
            Left            =   7965
            Picture         =   "frmComEntPedidosSa.frx":0707
            ToolTipText     =   "Buscar trabajador"
            Top             =   555
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Solicitado por"
            Height          =   255
            Index           =   6
            Left            =   6975
            TabIndex        =   102
            Top             =   555
            Width           =   1095
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   7
            Left            =   7965
            Picture         =   "frmComEntPedidosSa.frx":0809
            ToolTipText     =   "Buscar cliente"
            Top             =   210
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Para Cliente"
            Height          =   255
            Index           =   4
            Left            =   6975
            TabIndex        =   100
            Top             =   195
            Width           =   855
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   6
            Left            =   900
            Picture         =   "frmComEntPedidosSa.frx":090B
            ToolTipText     =   "Buscar proveedor varios"
            Top             =   210
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo Portes"
            Height          =   255
            Index           =   3
            Left            =   6960
            TabIndex        =   72
            Top             =   1380
            Width           =   855
         End
         Begin VB.Image imgBuscar 
            Enabled         =   0   'False
            Height          =   240
            Index           =   1
            Left            =   900
            Picture         =   "frmComEntPedidosSa.frx":0A0D
            ToolTipText     =   "Buscar población"
            Top             =   940
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Provincia"
            Height          =   255
            Index           =   17
            Left            =   120
            TabIndex        =   71
            Top             =   1300
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Población"
            Height          =   255
            Index           =   16
            Left            =   120
            TabIndex        =   70
            Top             =   910
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Teléfono"
            Height          =   255
            Index           =   19
            Left            =   2565
            TabIndex        =   69
            Top             =   195
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "NIF"
            Height          =   255
            Index           =   20
            Left            =   120
            TabIndex        =   68
            Top             =   190
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Forma Pago"
            Height          =   255
            Index           =   15
            Left            =   6975
            TabIndex        =   67
            Top             =   915
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Dto. P.P"
            Height          =   255
            Index           =   25
            Left            =   11040
            TabIndex        =   66
            Top             =   1215
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Dto. Gral"
            Height          =   255
            Index           =   26
            Left            =   11880
            TabIndex        =   65
            Top             =   1215
            Width           =   735
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   3
            Left            =   7965
            Picture         =   "frmComEntPedidosSa.frx":0B0F
            ToolTipText     =   "Buscar forma de pago"
            Top             =   930
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Domicilio"
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   63
            Top             =   550
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
         Top             =   3960
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
         Top             =   3960
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
         Top             =   3960
         Visible         =   0   'False
         Width           =   3045
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   7
         Left            =   9360
         MaxLength       =   12
         TabIndex        =   58
         Tag             =   "Importe"
         Text            =   "Importe"
         Top             =   3960
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   6
         Left            =   8880
         MaxLength       =   30
         TabIndex        =   49
         Tag             =   "Descuento 2"
         Text            =   "Dto2"
         Top             =   3960
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   5
         Left            =   8280
         MaxLength       =   5
         TabIndex        =   48
         Tag             =   "Descuento 1"
         Text            =   "Dto1"
         Top             =   3960
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   4
         Left            =   7320
         MaxLength       =   12
         TabIndex        =   47
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
         Left            =   6000
         MaxLength       =   16
         TabIndex        =   46
         Tag             =   "Cantidad"
         Text            =   "1,234,567,891.25"
         Top             =   3960
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
         Top             =   3900
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
         Top             =   3900
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   280
         Index           =   21
         Left            =   -74760
         MaxLength       =   80
         TabIndex        =   32
         Tag             =   "Observación 5|T|S|||scappr|observa5||N|"
         Top             =   5280
         Width           =   7485
      End
      Begin VB.TextBox Text1 
         Height          =   280
         Index           =   20
         Left            =   -74760
         MaxLength       =   80
         TabIndex        =   31
         Tag             =   "Observación 4|T|S|||scappr|observa4||N|"
         Top             =   5010
         Width           =   7485
      End
      Begin VB.TextBox Text1 
         Height          =   280
         Index           =   19
         Left            =   -74760
         MaxLength       =   80
         TabIndex        =   30
         Tag             =   "Observación 3|T|S|||scappr|observa3||N|"
         Top             =   4740
         Width           =   7485
      End
      Begin VB.TextBox Text1 
         Height          =   280
         Index           =   18
         Left            =   -74760
         MaxLength       =   80
         TabIndex        =   29
         Tag             =   "Observación 2|T|S|||scappr|observa2||N|"
         Top             =   4470
         Width           =   7485
      End
      Begin VB.TextBox Text1 
         Height          =   280
         Index           =   17
         Left            =   -74760
         MaxLength       =   80
         TabIndex        =   28
         Tag             =   "Observación 1|T|S|||scappr|observa1||N|"
         Top             =   4200
         Width           =   7485
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmComEntPedidosSa.frx":0C11
         Height          =   4305
         Left            =   240
         TabIndex        =   59
         Top             =   2280
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   7594
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
      Begin VB.Image imgAmpliaci 
         Height          =   240
         Left            =   11880
         Picture         =   "frmComEntPedidosSa.frx":0C26
         ToolTipText     =   "Buscar actuacion"
         Top             =   5520
         Width           =   240
      End
      Begin VB.Image imgBuscar2 
         Height          =   240
         Index           =   2
         Left            =   11160
         Picture         =   "frmComEntPedidosSa.frx":0D28
         ToolTipText     =   "Buscar actuacion"
         Top             =   4440
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Pedido"
         Height          =   255
         Index           =   53
         Left            =   10560
         TabIndex        =   165
         Top             =   4440
         Width           =   615
      End
      Begin VB.Image imgBuscar2 
         Height          =   240
         Index           =   1
         Left            =   13080
         Picture         =   "frmComEntPedidosSa.frx":0E2A
         ToolTipText     =   "Buscar cliente"
         Top             =   3000
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgBuscar2 
         Height          =   240
         Index           =   0
         Left            =   11160
         Picture         =   "frmComEntPedidosSa.frx":13B4
         ToolTipText     =   "Buscar actuacion"
         Top             =   3000
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha"
         Height          =   255
         Index           =   52
         Left            =   12600
         TabIndex        =   163
         Top             =   3000
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Albarán"
         Height          =   255
         Index           =   51
         Left            =   10560
         TabIndex        =   162
         Top             =   3000
         Width           =   975
      End
      Begin VB.Line Line3 
         X1              =   10560
         X2              =   14040
         Y1              =   4875
         Y2              =   4875
      End
      Begin VB.Line Line2 
         X1              =   10560
         X2              =   14040
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Image imgBuscar 
         Enabled         =   0   'False
         Height          =   240
         Index           =   12
         Left            =   -66720
         Picture         =   "frmComEntPedidosSa.frx":14B6
         ToolTipText     =   "Buscar dirección"
         Top             =   2520
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Envio"
         Height          =   255
         Index           =   49
         Left            =   -67320
         TabIndex        =   160
         Top             =   2520
         Width           =   960
      End
      Begin VB.Label Label1 
         Caption         =   "Su referencia"
         Height          =   255
         Index           =   48
         Left            =   -67920
         TabIndex        =   158
         Top             =   3240
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Nuestra referencia"
         Height          =   255
         Index           =   47
         Left            =   -74760
         TabIndex        =   157
         Top             =   3240
         Width           =   2175
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   1
         Left            =   -67920
         Picture         =   "frmComEntPedidosSa.frx":15B8
         ToolTipText     =   "Buscar fecha"
         Top             =   2520
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "F. Recogida"
         Height          =   255
         Index           =   44
         Left            =   -68880
         TabIndex        =   156
         Top             =   2565
         Width           =   1095
      End
      Begin VB.Image imgBuscar2 
         Height          =   240
         Index           =   11
         Left            =   11400
         Picture         =   "frmComEntPedidosSa.frx":1643
         ToolTipText     =   "Buscar proveedor varios"
         Top             =   3720
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgBuscar2 
         Height          =   240
         Index           =   10
         Left            =   10920
         Picture         =   "frmComEntPedidosSa.frx":1745
         ToolTipText     =   "Buscar proveedor varios"
         Top             =   3000
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgBuscar2 
         Height          =   240
         Index           =   9
         Left            =   11040
         Picture         =   "frmComEntPedidosSa.frx":1847
         ToolTipText     =   "Buscar proveedor varios"
         Top             =   2280
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Actuacion"
         Height          =   255
         Index           =   43
         Left            =   10560
         TabIndex        =   152
         Top             =   3720
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Obra"
         Height          =   255
         Index           =   34
         Left            =   10560
         TabIndex        =   151
         Top             =   3000
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente"
         Height          =   255
         Index           =   29
         Left            =   10560
         TabIndex        =   150
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Centro coste"
         Height          =   255
         Index           =   46
         Left            =   10560
         TabIndex        =   149
         Top             =   4920
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Ampliación Línea"
         Height          =   255
         Index           =   35
         Left            =   10560
         TabIndex        =   147
         Top             =   5520
         Width           =   1335
      End
      Begin VB.Image imgBuscar 
         Enabled         =   0   'False
         Height          =   240
         Index           =   11
         Left            =   -73920
         Picture         =   "frmComEntPedidosSa.frx":1949
         ToolTipText     =   "Buscar dirección"
         Top             =   2760
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Recogida"
         Height          =   255
         Index           =   28
         Left            =   -74760
         TabIndex        =   146
         Top             =   2760
         Width           =   720
      End
      Begin VB.Label Label1 
         Caption         =   "Observaciones"
         Height          =   255
         Index           =   45
         Left            =   -74760
         TabIndex        =   42
         Top             =   3960
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   13170
      TabIndex        =   36
      Top             =   8280
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
      Left            =   2760
      TabIndex        =   144
      Top             =   8400
      Width           =   6375
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
      Begin VB.Menu mnGeneraDtos 
         Caption         =   "Modificar &descuentos"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnImpPedido 
         Caption         =   "&Imprimir Pedido"
         Shortcut        =   ^I
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
Attribute VB_Name = "frmComEntPedidosSa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
'Public Event DatoSeleccionado(CadenaSeleccion As String)


Public MostrarDatos As String  'Para ver un dato enconcreto
Public EsHistorico As Boolean 'Si es true abrir el formulario con la tabla de
                              'de historico schppr, y solo en modo de consulta
                              
                              
Private WithEvents frmB As frmBuscaGrid 'Form para busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmF As frmCal 'Form Calendario Fecha
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmCP As frmCPostal 'Codigos Postales
Attribute frmCP.VB_VarHelpID = -1

Private WithEvents frmProv As frmComProveedores  'Form Mto Proveedores
Attribute frmProv.VB_VarHelpID = -1
Private WithEvents frmProveV As frmComProveV  'Form Mto Proveedores Varios
Attribute frmProveV.VB_VarHelpID = -1
Private WithEvents frmDir As frmComDirecciones
Attribute frmDir.VB_VarHelpID = -1
Private WithEvents frmFP As frmFacFormasPago 'Form Mto Formas de Pago
Attribute frmFP.VB_VarHelpID = -1
Private WithEvents frmT As frmAdmTrabajadores  'Form Mto Trabajadores
Attribute frmT.VB_VarHelpID = -1
Private WithEvents frmAlm As frmAlmAlPropios   'Form Almacenes Propios
Attribute frmAlm.VB_VarHelpID = -1

Private WithEvents FrmArt As frmAlmArticu2   'Form Articulos
Attribute FrmArt.VB_VarHelpID = -1
Private WithEvents FrmArtEul As frmAlmArticuEUL   'Form Articulos
Attribute FrmArtEul.VB_VarHelpID = -1
 
Private WithEvents frmCli As frmFacClientes 'form mantenimiento clientes
Attribute frmCli.VB_VarHelpID = -1
Private WithEvents frmInc As frmIncidencias  'form mantenimiento incidencias eliminacion
Attribute frmInc.VB_VarHelpID = -1

Private WithEvents frmNSerie As frmRepCargarNSerie  'Form Cargar nº Series
Attribute frmNSerie.VB_VarHelpID = -1
Private WithEvents frmNLote As frmAlmCargarNLote   'Form Cargar nº lote
Attribute frmNLote.VB_VarHelpID = -1
Private WithEvents frmList As frmListadoOfer 'Listados
Attribute frmList.VB_VarHelpID = -1
Private WithEvents frmFE As frmFacFormasEnvio
Attribute frmFE.VB_VarHelpID = -1
  
Private frmArt2 As frmAlmArticulos   'Form Articulos


Private WithEvents frmRecoge As frmComDirRecogida
Attribute frmRecoge.VB_VarHelpID = -1
Private WithEvents frmAc As frmObraActua
Attribute frmAc.VB_VarHelpID = -1

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
'-------------------------------------------------------------------------


Dim ModificaLineas As Byte
'1.- Añadir,  2.- Modificar,  3.- Borrar,  0.-Pasar control a Lineas

Dim TituloLinea As String 'Descripcion de la linea que estamos en Mantenimiento

Dim primeravez As Boolean
Dim txtAnterior  As String

'Codigo tipo de movimiento en función del valor en la tabla de parámetros: stipom
Dim CodTipoMov As String


Dim EsDeVarios As Boolean 'Si el Proveedor mostrado es de Varios o No

'SQL de la tabla principal del formulario
Private CadenaConsulta As String
Private CadenaSQL As String 'Para crear consulta de Generar Albaran a partir del Pedido

Private Ordenacion As String
Private NombreTabla As String  'Nombre de la tabla de Cabecera
Private NomTablaLineas As String 'Nombre de la Tabla de lineas
Private kCampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean

'Variable que indica el número del Boton  Anyadir en la Toolbar1
Dim btnAnyadir As Byte

'Variable que indica el número del Boton  PrimerRegistro en la Toolbar1
Dim btnPrimero As Byte


Private VieneDeBuscar As Boolean
'Para cuando devuelve dos poblaciones con el mismo codigo Postal. Si viene de pulsar prismatico
'de busqueda poner el valor de poblacion seleccionado y no volver a recuperar de la Base de Datos

Private HaCambiadoCP As Boolean
'Para saber si tras haber vuelto de prismaticos ha cambiado el valor del CPostal

Dim gridCargado As Boolean 'Saber si el grid esta cargado cuando se ejecuta DataGrid1_RowColChange

Dim AlbCompleto As Boolean 'Si se va a servir el Pedido Completo (slialb.cantidad=sliped.cantidad)
                            'o se va a servir una parte (slialb.cantidad=sliped.servidas)
Dim FormularioListAlbAbierto As Boolean
Dim PulsadoMas2 As Boolean

Private Sub chkObra_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

'================================================================================

Private Sub cmdAceptar_Click()
Dim vTipoMov As CTiposMov 'Clase Tipo Movimiento
Dim SQL As String
Dim PrimeraLin As Boolean 'Si se inserta la primera linea no esta creado el datagrid1 entonces llamar
                          ' a DataGrid, sino llamar solo a DataGrid2

    Screen.MousePointer = vbHourglass
    On Error GoTo Error1

    Select Case Modo
        Case 1  'BUSQUEDA
            HacerBusqueda
            
        Case 3 'INSERTAR Cabecera Pedido
            If DatosOk Then
                Set vTipoMov = New CTiposMov
                If vTipoMov.Leer(CodTipoMov) Then
                    Text1(0).Text = vTipoMov.ConseguirContador(CodTipoMov)
                    SQL = CadenaInsertarDesdeForm(Me)
                    If SQL <> "" Then
                        If InsertarPedido(SQL, vTipoMov) Then
                            CadenaConsulta = "Select * from " & NombreTabla & ObtenerWhereCP(True) & Ordenacion
                            PonerCadenaBusqueda
                            PonerModo 2
                            'Ponerse en Modo Insertar Lineas
                            BotonMtoLineas 1, "Pedidos"
                            BotonAnyadirLinea
                        End If
                    End If
                    FormateaCampo Text1(0)
                End If
                Set vTipoMov = Nothing
            End If
            Me.SSTab1.Tab = 0
            
        Case 4  'MODIFICAR Cabecera Pedido
            If DatosOk Then
                If ModificaDesdeFormulario(Me, 1) Then
                    'Actualizar los datos del Proveedor si es de varios
                    ActualizarProveVarios Text1(4).Text, Text1(6).Text
                    TerminaBloquear
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
                    BotonAnyadirLinea
                End If
            ElseIf ModificaLineas = 2 Then 'MODIFICAR lineas
                If ModificarLinea Then
                     NumRegElim = Data2.Recordset!numlinea
                    TerminaBloquear
                    CargaTxtAux False, False
                    CargaGrid2 DataGrid1, Data2
                    
                    ModificaLineas = 0
                    PosicionarData2
                    CamposObractua2
                    PonerBotonCabecera True
                    BloquearTxt Text2(16), True
                End If
                Me.DataGrid1.Enabled = True
            End If
            CalcularDatosFactura
            
        Case 6 'Pasar Pedido a Albaran
            If BLOQUEADesdeFormulario(Me) Then GenerarAlbaran
            TerminaBloquear
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
            PonerFoco txtAux(index)
            
        Case 1 'Busqueda de Cod. Artic
            If vParamAplic.NumeroInstalacion = 4 Then
                'EULER  As
                Set FrmArtEul = New frmAlmArticuEUL
                'frmArt.DatosADevolverBusqueda3 = "@1@" 'Poner en modo busqueda
                FrmArtEul.FechaDoc = CDate(Text1(1).Text)
                FrmArtEul.Codprove = CLng(Text1(4).Text)
                FrmArtEul.Show vbModal
                Set FrmArtEul = Nothing
            
            Else
                'SALIL
                Set FrmArt = New frmAlmArticu2
                'frmArt.DatosADevolverBusqueda3 = "@1@" 'Poner en modo busqueda
                FrmArt.DesdeTPV = False
                FrmArt.Show vbModal
                Set FrmArt = Nothing
            End If
            PonerFoco txtAux(index)
            
        Case 2 'COD. CENTRO DE COSTE
            If vEmpresa.TieneAnalitica Then
                'centro de coste
                AbrirForm_CentroCoste
                PonerFoco txtAux(8)
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
            If ModificaLineas = 1 Then 'INSERTAR
                ModificaLineas = 0
                DataGrid1.AllowAddNew = False
                If Not Data2.Recordset.EOF Then Data2.Recordset.MoveFirst
              
            Else
                ModificaLineas = 0
            End If
            PonerBotonCabecera True
            Me.DataGrid1.Enabled = True
            Me.lblF.Caption = ""
        Case 6 'Insertar servidas en Generar Albaran (Pedido --> Albaran)
            If MsgBox("Desea cancelar la introducción de unidades del pedido?", vbQuestion + vbYesNo) = vbYes Then
                TerminaBloquear
                InicializarServidas
                PonerModo 2
                CargaTxtAuxServidas False, False
                CargaGrid DataGrid1, Data2, True, False
                
            Else
                PonerFoco Me.txtAux(3)
            End If
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

    'Poner el nombre del trabajador que esta conectado
    Text1(3).Text = PonerTrabajadorConectado(NomTraba)
    Text2(3).Text = NomTraba

    Text1(1).Text = Format(Now, "dd/mm/yyyy") 'Fecha Oferta
    PonerFoco Text1(1)
End Sub


Private Sub BotonAnyadirLinea()
    'Si no estaba modificando lineas salimos
    ' Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 2 Then Exit Sub
       
    ModificaLineas = 1 'Ponemos Modo Añadir Linea
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerBotonCabecera False
    lblIndicador.Caption = "INSERTAR"
    
    AnyadirLinea DataGrid1, Data2
    CargaTxtAux True, True
    
    'Poner el Almacen por defecto del Trabajador
    txtAux(0).Text = DevuelveDesdeBDNew(conAri, "straba", "codalmac", "codtraba", Text1(3).Text, "N")
    If txtAux(0).Text <> "" Then txtAux(0).Text = Format(txtAux(0).Text, "000")
    
    'Campo Ampliacion Linea
    Text2(16).Text = ""
    BloquearTxt Text2(16), False
    
    ' ---- [20/10/2009] [LAURA]: añadir campo centro de coste
    'si contab. analitica por trabajador traer su centro de coste
    If vEmpresa.TieneAnalitica And vParamAplic.ModoAnalitica = 0 Then
        txtAux(8).Text = DevuelveDesdeBDNew(conAri, "straba", "codccost", "codtraba", Text1(3).Text, "N")
        Me.txtAux2(8).Text = PonerNombreCCoste(Me.txtAux(8))
    End If
    
    ' ----
    
    
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
Dim SQL As String
On Error GoTo EModificar

    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 4
    
    PonerFoco Text1(1)
            
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
'Prepara el Form para Modificar una linea de Pedido (tabla: sliped)
Dim vWhere As String
On Error GoTo EModificarLinea

    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Then Exit Sub '1= Insertar
    
    If Data2.Recordset.EOF Then Exit Sub
    vWhere = ObtenerWhereCP(False) & " and numlinea=" & Data2.Recordset!numlinea
    vWhere = Replace(vWhere, NombreTabla, NomTablaLineas)
    If Not BloqueaRegistro(NomTablaLineas, vWhere) Then Exit Sub
    
    CargaTxtAux True, False
    CamposObractua2
    ModificaLineas = 2 'Modificar
    'Añadiremos el boton de aceptar y demas objetos para insertar
    Me.lblIndicador.Caption = "MODIFICAR"
    PonerBotonCabecera False
    BloquearTxt Text2(16), False 'Campo Ampliacion Linea
    BloquearTxt txtAux(2), True 'campo nombre articulo
    PonerFoco txtAux(0)
    Me.DataGrid1.Enabled = False
    
EModificarLinea:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub BotonEliminar()
'Eliminar Registro de la Cabecera: Tabla de Pedidos (scaped)
' y los registros correspondientes de las tablas de lineas (sliped)
Dim cad As String
Dim vTipoMov As CTiposMov
Dim NumPedElim As Long 'Numero del Pedido que se ha Eliminado

    On Error GoTo EEliminar

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
    cad = "Cabecera de Pedidos Compras." & vbCrLf
    cad = cad & "--------------------------------------" & vbCrLf & vbCrLf
    cad = cad & "Va a eliminar el Pedido:            "
    cad = cad & vbCrLf & "Nº:  " & Format(Text1(0).Text, "0000000")
    cad = cad & vbCrLf & "Proveedor:  " & Format(Text1(4).Text, "000000") & " - " & Text1(5).Text
    cad = cad & vbCrLf & vbCrLf & " ¿Desea Eliminarlo? "
       
    'Borramos
    If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then
        Screen.MousePointer = vbHourglass
        
        NumRegElim = Data1.Recordset.AbsolutePosition
        NumPedElim = Data1.Recordset.Fields(0).Value
        
        CadenaSQL = ""
        Set frmList = New frmListadoOfer
        frmList.OpcionListado = 81
        frmList.Show vbModal
        Set frmList = Nothing
    
        If CadenaSQL = "" Then Exit Sub
        cad = ""
        cad = DBSet(RecuperaValor(CadenaSQL, 1), "F") & " as fechelim,"
        cad = cad & RecuperaValor(CadenaSQL, 2) & " as trabelim,"
        cad = cad & DBSet(RecuperaValor(CadenaSQL, 3), "T") & " as codincid"
        CadenaSQL = cad
        
        
        If Not Eliminar() Then Exit Sub
        PosicionarDataTrasEliminar
        
        'Devolvemos contador, si no estamos actualizando
        Set vTipoMov = New CTiposMov
        vTipoMov.DevolverContador CodTipoMov, NumPedElim
        Set vTipoMov = Nothing
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
    
    If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        NumRegElim = Data2.Recordset.AbsolutePosition
        SQL = "Delete from " & NomTablaLineas & Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas)
        
        SQL = SQL & " and numlinea=" & Data2.Recordset!numlinea
        conn.Execute SQL
        
        ModificaLineas = 0
        CargaGrid2 DataGrid1, Data2
        SituarDataTrasEliminar Data2, NumRegElim
        CalcularDatosFactura
'        CancelaADODC
    End If
    PonerFocoBtn Me.cmdRegresar
    
EEliminarLinea:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Lineas Mantenimientos", Err.Description
End Sub


Private Sub BotonGenerarAlbaran()
    'Pasar una Pedido a Albaran
Dim Resp As Byte

    'Comprobar que hay un Pedido seleccionado
    If Text1(0).Text = "" Then Exit Sub
        
    'Comprobar que hay lineas
    Resp = 1
    If Not (Data2.Recordset Is Nothing) Then
        If Not Data2.Recordset.EOF Then
            If Not IsNull(Data2.Recordset!numlinea) Then Resp = 0
        End If
    End If
    If Resp = 1 Then
        MsgBox "Pedido sin lineas", vbExclamation
        Exit Sub
    End If
    'Preguntar si se Recibe el pedido completo o no
    Resp = MsgBox("¿Recibir el pedido completo?", vbYesNoCancel + vbQuestion)
    If Resp = vbCancel Then Exit Sub
    
    If Resp = vbYes Then 'RECIBIR EL PEDIDO COMPLETO
        AlbCompleto = True
        Screen.MousePointer = vbHourglass

        GenerarAlbaran
        TerminaBloquear
        
    ElseIf Resp = vbNo Then 'RECIBIR PEDIDO INCOMPLETO
        AlbCompleto = False
        Me.SSTab1.Tab = 0
        TerminaBloquear
        'Si no se va a servir completo Mostrar lineas para que se indiquen las Servidas
        MsgBox "Introduzca la cantidad  a recibir para cada línea.", vbInformation
        Modo = 6
        gridCargado = False
        Me.cmdAceptar.visible = True
        Me.cmdCancelar.visible = True
        PonerModoOpcionesMenu Modo
        CargaGrid DataGrid1, Data2, True, True
        CargaTxtAuxServidas True, True
        primeravez = True
    Else
        TerminaBloquear
    End If

End Sub





Private Sub cmdRegresar_Click()
'Este es el boton Cabecera
Dim cad As String

    'Quitar lineas y volver a la cabecera
    If Modo = 5 Then  'modo 5: Mantenimientos Lineas
        PonerModo 2
        'BloquearTabs False
        Me.lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
        If DataGrid1.Row >= 0 Then
            DeseleccionaGrid DataGrid1
            DataGrid1.Bookmark = 1
        End If
        cmdRegresar.Caption = "Regresar"
    Else 'Se llama desde algún Prismatico de otro Form al Mantenimiento de Trabajadores
        If Data1.Recordset.EOF Then
            MsgBox "Ningún registro devuelto.", vbExclamation
            Exit Sub
        End If
        'DAVID. Pongo a piñon el numero de pedido. YA NO SE UTILIZA
        'cad = Data1.Recordset.Fields(0)
        'RaiseEvent DatoSeleccionado(cad)
        Unload Me
    End If
End Sub


Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error GoTo Error1
    
    If Modo = 6 And gridCargado Then '6: Pasar Pedido a Albaran no Completo (Introducir las servidas)
        CargaTxtAuxServidas True, True
        txtAux(3).Text = Format(Data2.Recordset!recibida, FormatoImporte)
    End If
    
    
    If Not Data2.Recordset.EOF Then
        If Not DGrid_CambiarFila(DataGrid1) Then Exit Sub
    End If
    
   If Not Data2.Recordset.EOF And ModificaLineas <> 1 Then
    
    
        If Not Data2.Recordset.EOF And ModificaLineas <> 1 Then '1: Insertar
            'Poner descripcion de ampliacion lineas
            Text2(16).Text = DevuelveDesdeBDNew(conAri, NomTablaLineas, "ampliaci", "numpedpr", Text1(0).Text, "N", , "numlinea", Data2.Recordset!numlinea, "N")
                        
        Else
            Text2(16).Text = ""
        End If
        
        
        
        '- centro de coste
        ' ---- [20/10/2009] [LAURA]: añadir campo centro de coste familia
        If vEmpresa.TieneAnalitica Then
            Me.txtAux(8).Text = DBLet(Data2.Recordset!CodCCost, "T")
            Me.txtAux2(8).Text = PonerNombreCCoste(Me.txtAux(8))
        Else
            txtAux2(8).Text = ""
        End If
        
        CamposObractua2
        
    End If
    Exit Sub
    
Error1:
    MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub PonerForaGrid()
    
    
    
        If Not Data2.Recordset.EOF Then  '1: Insertar
            'Poner descripcion de ampliacion lineas
            Text2(16).Text = DevuelveDesdeBDNew(conAri, NomTablaLineas, "ampliaci", "numpedpr", Text1(0).Text, "N", , "numlinea", Data2.Recordset!numlinea, "N")
            If vEmpresa.TieneAnalitica Then
                Me.txtAux(8).Text = DBLet(Data2.Recordset!CodCCost, "T")
                Me.txtAux2(8).Text = PonerNombreCCoste(Me.txtAux(8))
            End If
        Else
            Text2(16).Text = ""
            If vEmpresa.TieneAnalitica Then Me.txtAux(8).Text = "":                Me.txtAux2(8).Text = ""
            
        End If
        
        
        
        '- centro de coste
        
        
        CamposObractua2
        
    
End Sub



Private Sub Form_activate()
    If MostrarDatos <> "" Then
        MostrarDatos = ""
        PonerCampos
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub Form_Load()
Dim SelectInicial As String

    'Icono del formulario
    Me.Icon = frmPpal.Icon
    
    ' ICONITOS DE LA BARRA
    btnAnyadir = 5
    btnPrimero = 19
    With Me.Toolbar1
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 1   'Botón Buscar
        .Buttons(2).Image = 2   'Botón Todos
        .Buttons(5).Image = 3   'Insertar Nuevo
        .Buttons(6).Image = 4   'Modificar
        .Buttons(7).Image = 5   'Borrar
        .Buttons(10).Image = 10 'Mto Lineas Ofertas
        .Buttons(11).Image = 26 'Generar Albaran
        
        'OCtubre 2011
        .Buttons(12).Image = 43 'Modificar descuentos
        
        .Buttons(14).Image = 16 'Imprimir Pedido
        .Buttons(16).Image = 45  'simular
        .Buttons(17).Image = 15  'Salir
        .Buttons(btnPrimero).Image = 6  'Primero
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
        .Buttons(btnPrimero + 3).Image = 9 'Último
    End With
    Me.SSTab1.Tab = 0
      
    LimpiarCampos   'Limpia los campos TextBox
    cmdAux(2).Tag = -1
          
    '## A mano
     Me.FrameHco.visible = EsHistorico
    
    
    'Si no lleva datosdvolverbusqueda
    
    If Not EsHistorico Then
        NombreTabla = "scappr"
        NomTablaLineas = "slippr" 'Tabla lineas de Pedido
        Me.Caption = "Pedidos Proveedores"
        Ordenacion = " ORDER BY numpedpr "

    Else
        NombreTabla = "schppr"
        NomTablaLineas = "slhppr"
        CargarTagsHco Me, "scappr", NombreTabla
        'Estos campos solo estan en la tabla del histórico
        Text1(24).Tag = "Fecha Eliminación|F|N|||" & NombreTabla & "|fechelim|dd/mm/yyyy|N|"
        Text1(25).Tag = "Trabajador Eliminación|N|N|0|9999|" & NombreTabla & "|trabelim|0000|N|"
        Text1(26).Tag = "Incidencia elim.|T|N|||" & NombreTabla & "|codincid||N|"
        Me.Caption = "Histórico Pedidos Proveedores"
        Ordenacion = " ORDER BY numpedpr,fecpedpr "
    End If
    
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = conn
    If MostrarDatos = "" Then
        CodTipoMov = "-1"
    Else
        CodTipoMov = MostrarDatos
    End If
    Data1.RecordSource = "Select * from " & NombreTabla & "  WHERE numpedpr= " & CodTipoMov
    Data1.Refresh
    
    Me.Tag = "" 'Para que no carge los datos
 
    If MostrarDatos = "" Then
        PonerModo 0
    Else
        PonerModo 2
    End If
    
    
    CodTipoMov = "PEC"
    VieneDeBuscar = False
    
    'CC
    Label1(46).visible = vEmpresa.TieneAnalitica And vParamAplic.ModoAnalitica = 0
    txtAux(8).visible = vEmpresa.TieneAnalitica And vParamAplic.ModoAnalitica = 0
    txtAux2(8).visible = vEmpresa.TieneAnalitica And vParamAplic.ModoAnalitica = 0
    
    
    Euler_O_Sail 'Pondra visible s unos campos u otros
    
    
    'Cargar el DataGrid de lineas de Revisiones inicialmente a nada DATA2
    primeravez = True

    'Poner los grid sin apuntar a nada
    If MostrarDatos = "" Then LimpiarDataGrids
End Sub


Private Sub LimpiarCampos()
On Error Resume Next

    limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador.Caption = ""
    'Aqui va el especifico de cada form es
    '### a mano
    Me.chkRestoPed.Value = 0
    Me.chkObra.Value = 0
    Text3(0).Text = "TOTAL"
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    conn.Execute "DELETE FROM tmpnseries WHERE codusu=" & vUsu.Codigo
    'DatosADevolverBusqueda2 = "
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub frmAc_DatoSeleccionado(CadenaSeleccion As String)
    CadenaSQL = CadenaSeleccion
End Sub

Private Sub frmAlm_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Almacenes Propios
    txtAux(0).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000") 'Cod Almacen
End Sub

Private Sub frmArt_DatoSeleccionado(CadenaSeleccion As String)
'Mantenimiento de Articulos
    txtAux(1).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod Artic
    txtAux(2).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Artic
End Sub


Private Sub FrmArtEul_DatoSeleccionado(CadenaSeleccion As String)
    'Mantenimiento de Articulos
    txtAnterior = txtAux(1).Text
    txtAux(1).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod Artic
    txtAux(2).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Artic
    
    If txtAnterior <> txtAux(1).Text Then
        txtAux(4).Text = ""
        txtAux(5).Text = ""
        txtAux(6).Text = ""
        txtAux(7).Text = ""
    End If

    
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
Dim CadB As String
Dim Aux As String
      
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        
        If Val(cmdAux(2).Tag) > 0 Then
            'Llama desde boton busqueda centros de coste
            ' ---- [20/10/2009] [LAURA]: añadir campo centro de coste familia
            Me.txtAux(8).Text = RecuperaValor(CadenaDevuelta, 1)
            Me.txtAux2(8).Text = PonerNombreCCoste(Me.txtAux(8))
        Else
            CadB = ""
            Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
            CadB = Aux
            If EsHistorico Then
                Aux = ValorDevueltoFormGrid(Text1(1), CadenaDevuelta, 1)
                CadB = CadB & " and " & Aux
            End If
            
            CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
            PonerCadenaBusqueda
            Text1(0).Text = Format(RecuperaValor(CadenaDevuelta, 1), "0000000")
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub frmCli_DatoSeleccionado(CadenaSeleccion As String)

    CadenaSQL = CadenaSeleccion

   ' Text1(22).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod cliente
   ' FormateaCampo Text1(22)
   ' Text2(0).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom clien
End Sub

Private Sub frmCP_DatoSeleccionado(CadenaSeleccion As String)
'Formulario Mantenimiento C. Postales
Dim Indice As Byte
Dim devuelve As String

    Indice = 9
    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'CPostal
    'poblacion
    Text1(Indice + 1).Text = ObtenerPoblacion(Text1(Indice).Text, devuelve)
    'provincia
    Text1(Indice + 2).Text = devuelve
End Sub


Private Sub frmDir_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Direcciones
Dim Indice As Byte
    Indice = CByte(Me.imgBuscar(0).Tag)
    Text1(Indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000") 'Cod Direccion
    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Direc

    CargarDatosDirec Text1(Indice).Text, Indice
End Sub

Private Sub frmF_Selec(vFecha As Date) 'Calendario Fechas
'Dim Indice As Byte
'    Indice = CByte(Me.imgFecha(0).Tag) + 1
'    Text1(Indice).Text = Format(vFecha, "dd/mm/yyyy")
    CadenaSQL = Format(vFecha, "dd/mm/yyyy")
End Sub


Private Sub frmFE_DatoSeleccionado(CadenaSeleccion As String)
    CadenaSQL = CadenaSeleccion
End Sub

Private Sub frmFP_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Formas de Pago
Dim Indice As Byte

    Indice = 12
    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod Forma Pago
    FormateaCampo Text1(Indice)
    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Forma Pago
End Sub


Private Sub frmInc_DatoSeleccionado(CadenaSeleccion As String)
'Mantenimiento de incidencias
    Text1(26).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod incidencia
    Text2(26).Text = RecuperaValor(CadenaSeleccion, 2) 'nom incidencia
End Sub

Private Sub frmList_DatoSeleccionado(CadenaSeleccion As String)
'Aqui devuelve los valores que se introducen en el Listado
'para pasar de Pedido a Albaran, o para pasar al historico
    
    CadenaSQL = CadenaSeleccion
End Sub



Private Sub frmNSerie_CargarNumSeries()
'Insertar un registro en la tabla "sserie" por cada uno de los
'Nº de Serie introducidos en la Tabla Temporal
Dim RStmp As ADODB.Recordset
Dim RSalb As ADODB.Recordset
Dim SQL As String
Dim I As Byte
Dim B As Boolean
    
    On Error GoTo EInsertar

    
    SQL = "SELECT slialp.codartic, numlinea, cantidad "
    SQL = SQL & " FROM slialp INNER JOIN sartic on slialp.codartic=sartic.codartic "
    SQL = SQL & " WHERE numalbar=" & DBSet(Me.cmdAux(1).Tag, "T") & " and fechaalb=" & DBSet(Me.cmdAux(0).Tag, "F") & " and "
    SQL = SQL & "slialp.codprove=" & Text1(4).Text
    SQL = SQL & " And nseriesn = 1 "
    SQL = SQL & " ORDER BY codartic, numlinea "

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
        
        B = True
        For I = 1 To RSalb!cantidad
            If Not RStmp.EOF Then
                InsertarNSerie RStmp!numSerie, RStmp!codArtic, RSalb!numlinea
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


Private Sub frmProv_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Proveedores
Dim Indice As Byte

    Indice = 4
    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod Prove
    FormateaCampo Text1(Indice)
End Sub

Private Sub frmProveV_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento Proveedores varios
Dim Indice As Byte

    Indice = 6
    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'nif Prove
    Text1(Indice - 1).Text = RecuperaValor(CadenaSeleccion, 2) 'nom Prove
    PonerDatosProveVario Text1(Indice).Text
End Sub

Private Sub frmRecoge_DatoSeleccionado(CadenaSeleccion As String)
    Text1(27).Text = RecuperaValor(CadenaSeleccion, 1) 'coddirre
    Text2(27).Text = RecuperaValor(CadenaSeleccion, 2) 'nomdirre
    FormateaCampo Text1(27)
End Sub

Private Sub frmT_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Trabajadores
Dim Indice As Byte

    Indice = Val(Me.imgBuscar(0).Tag)
    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod Trabajador
    FormateaCampo Text1(Indice)
    If Indice = 23 Then Indice = 1
    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Trabajador
End Sub


Private Sub imgAmpliaci_Click()
Dim B As Boolean
    
    
    If Modo < 2 Then Exit Sub
    
    CadenaDesdeOtroForm = ""
    If Not Data2.Recordset.EOF Then CadenaDesdeOtroForm = DBLet(Data2.Recordset!Ampliaci, "T")
            
    B = False
    If Modo = 5 Then
        If ModificaLineas = 1 Or ModificaLineas = 2 Then B = True
    End If
    frmFacClienteObser.Modificar = B
    frmFacClienteObser.Text1 = CadenaDesdeOtroForm
    frmFacClienteObser.Show vbModal
    If Modo = 5 Then
        If RecuperaValor(CadenaDesdeOtroForm, 1) = "1" Then
            Text2(16).Text = Mid(CadenaDesdeOtroForm, 3)
        End If
    End If
    
End Sub

Private Sub imgBuscar_Click(index As Integer)
Dim Indice As Byte

    If Modo = 2 Or Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass
    Indice = 0
    Select Case index
        Case 0 'Cod. Proveedor
            Indice = 4
            Set frmProv = New frmComProveedores
            frmProv.DatosADevolverBusqueda = "0"
            frmProv.Show vbModal
            Set frmProv = Nothing
            
        Case 1 'Cod. Postal
            Set frmCP = New frmCPostal
            frmCP.DatosADevolverBusqueda = "0"
            frmCP.Show vbModal
            Set frmCP = Nothing
            Indice = 9
            VieneDeBuscar = True
            
        Case 2, 8 'Realizada Por Trabajador
            If index = 2 Then
                Indice = 3
            Else
                Indice = 23
            End If
            Me.imgBuscar(0).Tag = Indice
            Set frmT = New frmAdmTrabajadores
            frmT.DatosADevolverBusqueda = "0"
            frmT.Show vbModal
            Set frmT = Nothing
            
        Case 3 'Forma de Pago
            Indice = 12
            Set frmFP = New frmFacFormasPago
            frmFP.DatosADevolverBusqueda = "0"
            frmFP.Show vbModal
            Set frmFP = Nothing
            
        Case 4, 5 'Direccion
            If index = 4 Then Indice = 15
            If index = 5 Then Indice = 2
            Me.imgBuscar(0).Tag = Indice
            Set frmDir = New frmComDirecciones
            frmDir.DatosADevolverBusqueda = "0"
            frmDir.Show vbModal
            Set frmDir = Nothing
            
        Case 6 'NIF de Proveedores VARIOS
            Indice = 6
            Set frmProveV = New frmComProveV
            frmProveV.DatosADevolverBusqueda = "0"
            frmProveV.Show vbModal
            Set frmProveV = Nothing
            
        Case 7 'Cliente
            Indice = 22
            CadenaSQL = ""
            Set frmCli = New frmFacClientes
            frmCli.DatosADevolverBusqueda = "0"
            frmCli.Show vbModal
            Set frmCli = Nothing
            If CadenaSQL <> "" Then
                Text1(22).Text = RecuperaValor(CadenaSQL, 1) 'Cod cliente
                FormateaCampo Text1(22)
                Text2(0).Text = RecuperaValor(CadenaSQL, 2) 'Nom clien
                CadenaSQL = ""
            End If
            
        Case 10 'Incidencias
            Indice = 26
            Set frmInc = New frmIncidencias
            frmInc.DatosADevolverBusqueda = "0"
            frmInc.Show vbModal
            Set frmInc = Nothing
            
        Case 11
            If Text1(4).Text = "" Then
                MsgBox "Ponga primero el proveedor", vbExclamation
                PonerFoco Text1(4)
            Else
                Indice = 27
                Set frmRecoge = New frmComDirRecogida
                frmRecoge.Codprove = CLng(Text1(4).Text)
                frmRecoge.nomprove = Text1(5).Text
                If Text1(Indice).Text <> "" Then
                    frmRecoge.VerDatoDpto = Text1(Indice).Text
                Else
                    frmRecoge.VerDatoDpto = -1
                End If
                frmRecoge.Show vbModal
                Set frmRecoge = Nothing
            
            End If
        Case 12
            Set frmFE = New frmFacFormasEnvio
            frmFE.DatosADevolverBusqueda = "0|1|"
            CadenaSQL = ""
            frmFE.Show vbModal
            Set frmFE = Nothing
            If CadenaSQL <> "" Then
                Text1(31).Text = RecuperaValor(CadenaSQL, 1)
                Text2(4).Text = RecuperaValor(CadenaSQL, 2)
                CadenaSQL = ""
            End If
    End Select
    If Indice > 0 Then PonerFoco Text1(Indice)
    Screen.MousePointer = vbDefault
End Sub


Private Sub imgBuscar2_Click(index As Integer)
    If Modo <> 5 Then Exit Sub
    
    
    If index = 1 Then
         
        'Fecha albaran para EULER
        If ModificaLineas = 0 Then Exit Sub
        imgFecha_Click 1000
    
    ElseIf index = 0 Or index = 2 Then
        'EULER
        If ModificaLineas = 0 Then Exit Sub
        If vParamAplic.NumeroInstalacion = 4 Then LanzarBuscarAlbaranEuler index = 2 'Abrimos
        
        
    ElseIf index = 9 Then
            CadenaSQL = ""
            Set frmCli = New frmFacClientes
            frmCli.DatosADevolverBusqueda = "0"
            frmCli.Show vbModal
            Set frmCli = Nothing
            If CadenaSQL <> "" Then
                txtAux(9).Text = RecuperaValor(CadenaSQL, 1) 'Cod cliente
                Me.txtDesc(9).Text = RecuperaValor(CadenaSQL, 2) 'Nom clien
                CadenaSQL = ""
                If vParamAplic.NumeroInstalacion = 4 Then
                    LanzarBuscarAlbaranEuler False  'Abrimos
                Else
                    PonerFoco txtAux(10)
                End If
            End If
            
    Else
        'Obra actuacion. Llamaraemos al mismo
        If Me.txtAux(9).Text = "" Then
            MsgBox "Indique el cliente", vbExclamation
            PonerFoco txtAux(9)
            
        Else
            CadenaSQL = ""
            Set frmAc = New frmObraActua
            frmAc.DatosADevolverBusqueda = txtAux(9).Text & "|" & txtAux(10).Text & "|"
            frmAc.Show vbModal
            Set frmAc = Nothing
            If CadenaSQL <> "" Then
                txtAux(11).Text = RecuperaValor(CadenaSQL, 3)
                txtDesc(11).Text = RecuperaValor(CadenaSQL, 4) & "  " & RecuperaValor(CadenaSQL, 5)
                
                If txtAux(10).Text = "" Then
                    txtAux(10).Text = RecuperaValor(CadenaSQL, 2)
                    PonerClieObraActuacion 10, False
                End If
                CadenaSQL = ""
            End If
        End If
    End If
    
    
End Sub

Private Sub imgFecha_Click(index As Integer) 'Abre calendario Fechas
Dim Indice As Byte

   If Modo = 2 Or Modo = 0 Then Exit Sub
   Screen.MousePointer = vbHourglass
   Set frmF = New frmCal
   frmF.Fecha = Now
   
   If index = 1000 Then
        Indice = 128
   ElseIf index = 0 Then
        Indice = 1 'Index + 1
   Else
        Indice = 28
   End If

   If Indice < 100 Then
        PonerFormatoFecha Text1(Indice)
        If Text1(Indice).Text <> "" Then frmF.Fecha = CDate(Text1(Indice).Text)
   Else
        PonerFormatoFecha txtAux(14)
        If txtAux(14).Text <> "" Then frmF.Fecha = CDate(txtAux(14))
   End If
   Screen.MousePointer = vbDefault
   CadenaSQL = ""
   frmF.Show vbModal
   Set frmF = Nothing
   If CadenaSQL <> "" Then
        If Indice = 128 Then
            txtAux(14).Text = CadenaSQL
            PonerFoco txtAux(14)
        Else
            Text1(Indice).Text = CadenaSQL
            PonerFoco Text1(Indice)
        End If
        CadenaSQL = ""
   End If
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


Private Sub mnGenAlbaran_Click()
    'bloqueamos el pedido y lo pasamos a Albaran
    If BLOQUEADesdeFormulario(Me) Then BotonGenerarAlbaran
End Sub


Private Sub mnGeneraDtos_Click()
Dim B As Boolean
    If Text1(0).Text = "" Then Exit Sub 'por si las moscas
    If Data2.Recordset Is Nothing Then Exit Sub
    If Data2.Recordset.EOF Then Exit Sub
    
    'Modifica los descuentos de este albaran y recalcula los importes de las lineas(y por ende el total)
     If BLOQUEADesdeFormulario(Me) Then
        
        B = False
        
        CadenaDesdeOtroForm = Text1(5).Text & "(" & Text1(4).Text & ")|" & Text1(0).Text & " de " & Text1(1).Text & "|"
        'en el load pone a "" la variable
        frmVarios.Opcion = 9
        frmVarios.Show vbModal
        If CadenaDesdeOtroForm <> "" Then
            conn.BeginTrans
            B = ActualizarDtos
            If B Then
                conn.CommitTrans
            Else
                conn.RollbackTrans
            End If
        End If
        'Termina bloquear
        TerminaBloquear
        If B Then PonerCampos
     End If
End Sub

Private Sub mnImpPedido_Click()
'Imprime un Pedido
       frmListadoOfer.NumCod = Text1(0).Text    'Nº de Pedido
       frmListadoOfer.codClien = Text1(4).Text 'Cod.Proveedor
       If EsHistorico Then
            AbrirListadoOfer (56) '59: Informe de Pedidos Compras (Historico)
            frmListadoOfer.FecEntre = Text1(1).Text
       Else
            AbrirListadoOfer (55) '55: Informe de Pedidos Compras
       End If
End Sub

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
         BotonAnyadirLinea
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


'Los metodos del text tendran que estar
'Los descomentamos cuando esten puestos ya los controles
Private Sub Text1_GotFocus(index As Integer)
Dim cadkey As Integer

    cadkey = ObtenerCadKey(kCampo, index)
    kCampo = index
    If index = 9 Then HaCambiadoCP = False 'CPostal
    ConseguirFoco Text1(index), Modo
End Sub


Private Sub Text1_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
Dim Ind As Integer
Dim B As Boolean
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
    
    If KeyCode = 43 Or KeyCode = 107 Or KeyCode = 187 Then
        B = False
        If Text1(index).Text = "" Then
            B = True
        Else
            If Text1(index).SelLength = Len(Text1(index).Text) Then B = True
        End If
        If B Then
                Ind = -1
                Select Case index
                Case 2
                    Ind = 5
                Case 3
                    Ind = 2
                Case 4
                    Ind = 0
                Case 6
                    Ind = 6
                Case 9
                    Ind = 1
                Case 12
                    Ind = 3
                Case 15
                    Ind = 4
                Case 22, 23
                    Ind = index - 15
                Case 27
                    Ind = 11
                End Select
                If Ind >= 0 Then
                    PulsadoMas2 = True
                    PulsarTeclaMas True, Ind
                End If
            End If
        End If
End Sub


Private Sub Text1_KeyPress(index As Integer, KeyAscii As Integer)
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
Private Sub Text1_LostFocus(index As Integer)
Dim devuelve As String
Dim I As Byte
        
        
    'Han pulsado el mas
    If PulsadoMas2 Then
        'Para que cuando pulse el mas abra el form
        PulsadoMas2 = False
        Text1(index).Text = ""
        Exit Sub
    End If
    
    If Not PerderFocoGnral(Text1(index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
       
    'Si queremos hacer algo ..
    Select Case index
        Case 1, 28 'Fecha Oferta, Fecha Entrega
            PonerFormatoFecha Text1(index)
        
        Case 3, 23, 31 'Cod Trabajador
            I = index
            If index = 23 Then I = 1
            If index = 31 Then I = 4
            If PonerFormatoEntero(Text1(index)) Then
                If index = 31 Then
                    Text2(I).Text = PonerNombreDeCod(Text1(index), conAri, "senvio", "nomenvio", "codenvio", "el envio")
                Else
                    Text2(I).Text = PonerNombreDeCod(Text1(index), conAri, "straba", "nomtraba", "codtraba", "el Trabajador")
                End If
            Else
                Text2(I).Text = ""
            End If
            
        Case 4 'Cod. Prove
            If PonerFormatoEntero(Text1(index)) Then
                If Modo = 1 Then 'Busqueda
                    Text1(5).Text = PonerNombreDeCod(Text1(index), conAri, "sprove", "nomprove")
                Else ' cargar datos de Tabla sprove
                    PonerDatosProveedor (Text1(index).Text)
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
            PonerDatosProveVario (Text1(index).Text)
             
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
            
        Case 12 'Forma de Pago
            If PonerFormatoEntero(Text1(index)) Then
                Text2(index).Text = PonerNombreDeCod(Text1(index), conAri, "sforpa", "nomforpa")
            Else
                Text2(index).Text = ""
            End If
            
        Case 13, 14 'Descuentos
            If PonerFormatoDecimal(Text1(index), 4) Then   'Tipo 4: Decimal(4,2)
                 If Modo = 4 Then CalcularDatosFactura
            End If
            
        Case 15, 2 'Cod. Direccion
            If PonerFormatoEntero(Text1(index)) Then
                Me.imgBuscar(0).Tag = index
                If Not CargarDatosDirec(Text1(index).Text, CByte(index)) Then
                    PonerFoco Text1(index)
                End If
            Else
                LimpiarDatosDirec CByte(index)
            End If
            
        Case 22 'cod.cliente
            If PonerFormatoEntero(Text1(index)) Then
                Text2(0).Text = PonerNombreDeCod(Text1(index), conAri, "sclien", "nomclien")
            Else
                Text2(0).Text = ""
            End If
            
        Case 21
            If Me.ActiveControl.Name = "SSTab1" Then PonerFocoBtn Me.cmdAceptar
            
        Case 26 'cod Incidencia de eliminacion
            If EsHistorico Then
                Text2(index).Text = PonerNombreDeCod(Text1(index), conAri, "sincid", "nomincid")
                If Not (Text2(index).Text = "" And Text1(index).Text <> "") Then
                    PonerFocoBtn Me.cmdAceptar
                Else
                    PonerFoco Text1(index)
                End If
            End If
            
        Case 27
            devuelve = ""
            If Text1(4).Text <> "" Then
                If Text1(index).Text <> "" Then
                    If PonerFormatoEntero(Text1(index)) Then
                        devuelve = PonerNombreDeCod(Text1(index), conAri, "sdirRecog", "nomdirre", "codprove=" & Text1(4).Text & " AND coddirre")
                        If devuelve = "" Then PonerFoco Text1(index)
                    End If
                End If
            Else
                If Modo > 2 Then
                    MsgBox "Debe poner proveedor", vbExclamation
                    PonerFoco Text1(4)  'que ponga el proveedor
                End If
            End If
            Text2(27).Text = devuelve
            If devuelve = "" And Text1(index).Text <> "" Then Text1(index).Text = ""
    End Select
End Sub


Private Sub HacerBusqueda()
Dim CadB As String

    CadB = ObtenerBusqueda(Me, False)
    If chkVistaPrevia = 1 Then
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
Dim devuelve As String
    'Llamamos a al form
    '##A mano
    cad = ""
'    If EsCabecera Then
        cad = cad & ParaGrid(Text1(0), 15, "Nº Pedido")
        cad = cad & ParaGrid(Text1(1), 20, "Fecha Ped.")
        cad = cad & ParaGrid(Text1(4), 15, "Proveedor")
        cad = cad & ParaGrid(Text1(5), 50, "Nombre Prov.")
        tabla = NombreTabla
        Titulo = "Pedidos Compras"
        If EsHistorico Then
            Titulo = "Histórico de Pedidos"
            devuelve = "0|1|"
        Else
            Titulo = "Pedidos"
            devuelve = "0|"
        End If
'    End If
    
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
'        If Not EsCabecera Then frmB.Label1.FontSize = 11
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
            PonerFoco Text1(0)
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
'Carga las Pestañas con las tablas de lineas del Trabajador seleccionado para mostrar
On Error GoTo EPonerLineas

    Screen.MousePointer = vbHourglass

    'Datos de la tabla slippr
    CargaGrid DataGrid1, Data2, True
    CamposObractua2

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
    
    'Realizado por
    Text2(3).Text = PonerNombreDeCod(Text1(3), conAri, "straba", "nomtraba")
    Text2(12).Text = PonerNombreDeCod(Text1(12), conAri, "sforpa", "nomforpa")
    'Cliente para
    Text2(0).Text = PonerNombreDeCod(Text1(22), conAri, "sclien", "nomclien")
    'Solicitado por
    Text2(1).Text = PonerNombreDeCod(Text1(23), conAri, "straba", "nomtraba", "codtraba")
    
    'Direccion de recogida
    If Text1(27).Text <> "" Then
        Text2(27).Text = PonerNombreDeCod(Text1(27), conAri, "sdirRecog", "nomdirre", "codprove=" & Text1(4).Text & " AND coddirre")
    Else
        Text2(27).Text = ""
    End If
    'Envio
    Text2(4).Text = PonerNombreDeCod(Text1(31), conAri, "senvio", "nomenvio", "codenvio", "el envio")
    
    
    'Poner las direcciones
    CargarDatosDirec Text1(15).Text, 15
    CargarDatosDirec Text1(2).Text, 2
    
    
    PonerCamposLineas 'Pone los datos de las tablas de lineas de Pedidos
    
    If EsHistorico Then
        'poner datos de eliminacion
        Text2(25).Text = PonerNombreDeCod(Text1(25), conAri, "straba", "nomtraba", "codtraba")
        Text2(26).Text = PonerNombreDeCod(Text1(26), conAri, "sincid", "nomincid", "codincid")
    End If
    
    PonerForaGrid
    
    CalcularDatosFactura 'rellenar campos pestaña de totales
    
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
Dim B As Boolean
On Error GoTo EPonerModo

    'Actualiza Iconos Insertar,Modificar,Eliminar
    ActualizarToolbarGnral Me.Toolbar1, Modo, Kmodo, btnAnyadir
    lblF.Caption = ""
    
    Modo = Kmodo
    PonerIndicador lblIndicador, Modo
    
    If Modo = 6 Then Me.lblIndicador.Caption = "Insertar Cant. Servidas"
    
    'Modo 2. Hay datos y estamos visualizandolos
    '=========================================
    B = (Modo = 2)
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos

        cmdRegresar.visible = False
 
        
    'Poner Flechas de desplazamiento visibles
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero, B, NumReg
        
        
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1
    BloquearText1 Me, Modo
    'Campo Numero de Albaran siempre bloqueado, excepto si estamos en modo de busqueda
    B = (Modo <> 1)
    BloquearTxt Text1(0), B, True
       
    'datos cliente siempre bloqueados hasta que sea de varios
    If Modo = 3 Then
        EsDeVarios = False
        BloquearDatosProve (EsDeVarios)
    End If
       
       
    '-----  Datos Totales de Factura siempre bloqueado
    For I = 33 To 50
        BloquearTxt Text3(I), True
    Next I
    'Campo B.Imp y Imp. IVA siempre en azul
    Text3(36).BackColor = &HFFFFC0
    Text3(46).BackColor = &HFFFFC0
    Text3(47).BackColor = &HFFFFC0
    Text3(48).BackColor = &HFFFFC0
    Text3(49).BackColor = &HC0C0FF    'Tatal factura
    Text3(50).BackColor = &HC0C0FF    'Tatal factura
    '---------------------------------------------------
       
       
    'Si no es modo lineas Boquear los TxtAux
    For I = 0 To 7
        BloquearTxt txtAux(I), (Modo <> 5)
    Next I
    If Modo <> 5 Then
        'Los foragrid
        For I = 8 To txtAux.Count - 1
            BloquearTxt txtAux(I), (Modo <> 5)
        Next I
    End If
    BloquearTxt Text2(16), (Modo <> 5)
    
    
    B = (Modo <> 0 And Modo <> 2 And Modo <> 5)
    cmdCancelar.visible = B
    cmdAceptar.visible = B
    chkObra.Enabled = B
    
    For I = 0 To Me.imgFecha.Count - 1
        Me.imgFecha(I).Enabled = B
    Next I
    
    'El ultimo index NO es secuencial
    For I = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(I).Enabled = B
    Next I
    Me.imgBuscar(1).visible = False
    
           
    'Modo Linea de Ofertas. Poner el campo ampliacion linea
    'Me.Label1(35).visible = (Modo = 5)
    'Me.Text2(16).visible = (Modo = 5)
    BloquearTxt Text2(16), True
    
    ' ---- [20/10/2009] [LAURA] : añadir del centro de coste
    If txtAux2(8).visible Then BloquearTxt txtAux2(8), True
    
    
       
    Me.chkVistaPrevia.Enabled = (Modo <= 2)
       
    'Poner el tamaño de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos
       
    PonerModoOpcionesMenu (Modo) 'Activar opciones de menu según modo
    PonerOpcionesMenu 'Activar opciones de menu según nivel de permisos del usuario
EPonerModo:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
    PonerLongCamposGnral Me, Modo, 1
End Sub


Private Function DatosOk() As Boolean
'Comprueba si los datos de la cabecera son correctos antes de Insertar o Modificar el
'Pedido
Dim B As Boolean
On Error GoTo EDatosOK

    DatosOk = False
    B = CompForm(Me, 1) 'Comprobar formato datos ok
    If Not B Then Exit Function
            
            
    If B Then
        'El trabajador debe existir
        CadenaSQL = ""
        If Text2(3).Text = "" Then CadenaSQL = CadenaSQL & vbCrLf & "   - Trabajador pedido"
        'Recogida
        If Text1(27).Text <> "" Then
            If Text2(27).Text = "" Then CadenaSQL = CadenaSQL & vbCrLf & "   - Direccion recogida de mercancia"
        End If
        'Solicitado por
        If Text1(23).Text <> "" Then
            If Text2(1).Text = "" Then CadenaSQL = CadenaSQL & vbCrLf & "   - Trabajador que solicita pedido"
        End If


        If CadenaSQL <> "" Then
            CadenaSQL = "Error en campos: " & vbCrLf & CadenaSQL
            MsgBox CadenaSQL, vbExclamation
            B = False
        End If
    End If
    CadenaSQL = ""
    DatosOk = B
    
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Function DatosOkLinea() As Boolean
'Comprueba si los datos de una linea son correctos antes de Insertar o Modificar
'una linea del Pedido
Dim B As Boolean
'Dim devuelve As String
Dim I As Byte
Dim vArtic As CArticulo
Dim Aux As String
Dim TipoDto As Byte
Dim cad As String

    On Error GoTo EDatosOkLinea

    DatosOkLinea = False
    
    'Febrero 2010   Si han apretado Alt+A NO recalcula
    '----------------------------------------------------------------------------------
    'txtAux(8).Text = CalcularImporte(txtAux(3).Text, txtAux(4).Text, txtAux(6).Text, txtAux(7).Text, vParamAplic.TipoDtos)
    TipoDto = DevuelveDesdeBDNew(conAri, "sprove", "tipodtos", "codprove", Text1(4).Text, "N")
    Aux = RecalculoImporteLineas(txtAux(3), txtAux(4), txtAux(5), txtAux(6), TipoDto)
    Aux = Format(Aux, FormatoImporte)
    If Aux <> txtAux(7).Text Then txtAux(7).Text = Aux
    

    'ANALITICA
    'Si va por trabajador--> Si esta "" ponemos la del trabajador
    If vEmpresa.TieneAnalitica And vParamAplic.ModoAnalitica = 0 Then
        
        If txtAux(8).Text = "" Then
            txtAux(8).Text = DevuelveDesdeBDNew(conAri, "straba", "codccost", "codtraba", Text1(3).Text, "N")
            Me.txtAux2(8).Text = PonerNombreCCoste(Me.txtAux(8))
        End If
    End If
    
    
    B = True
    'Comprobar que los campos NOT NULL tienen valor
    For I = 0 To 8
        If txtAux(I).Text = "" Then
            If I = 8 And vEmpresa.TieneAnalitica = False Then
                'no hace nada pq puede ser nulo
            Else
                MsgBox "El campo " & txtAux(I).Tag & " no puede ser nulo", vbExclamation
                B = False
                PonerFoco txtAux(I)
                Exit Function
            End If
        End If
    Next I
        
    
        
    'obra actuacion
    Aux = ""
    TipoDto = 0 'para saber si ha puesto alguna de ellas
    For I = 9 To 11
       If txtAux(I).Text = "" Xor Me.txtDesc(I).Text = "" Then Aux = Aux & vbCrLf & txtAux(I).Tag
       If txtAux(I).Text <> "" Then TipoDto = 1
       
    Next
    If Aux <> "" Then Aux = "Error en: " & vbCrLf & Aux
        
    
    'Si indica alguno, debe indicarlos todos
    If Aux = "" Then
        If TipoDto = 1 Then
            'Ha puesto alguno de los campos(no deberia haber pasado)
            If txtAux(9).Text = "" Or txtAux(10).Text = "" Or txtAux(11).Text = "" Then
                 'Si es euler NO controlo este error
                If vParamAplic.NumeroInstalacion = 4 Then
                    Aux = ""
                Else
                    Aux = "Faltan campos en la obra actuacion"
                End If
            End If
        End If
    End If
    
    If Aux <> "" Then
        MsgBox Aux, vbExclamation
        PonerFoco txtAux(9)
        Exit Function
    End If
    
    
    'Numerero de albaran
    If vParamAplic.NumeroInstalacion = 4 Then
        Aux = ""
        cad = "" 'para saber si ha puesto alguna de ellas
        For I = 12 To 14
            If txtAux(I).Text <> "" Then cad = cad & "1"
             
        Next
        
        If cad <> "" Then
            If Len(cad) <> 3 Then
                MsgBox "Falta identificar el albaran correctamente", vbExclamation
                Exit Function
            Else
                'LEN 3, vemaos si existe
                Aux = "NO EXISTE"
                cad = txtDesc(0).Text
                If cad = "" Then cad = Aux
                If cad = Aux Then
                    cad = "No existe el albaran indicado. ¿Continuar de igual modo?"
                    If MsgBox(cad, vbQuestion + vbYesNo) = vbNo Then Exit Function
                End If
            End If
        End If
        
        
        If txtAux(15).Text <> "" Then
            Aux = DevuelveDesdeBD(conAri, "codclien", "scaped", "numpedcl", txtAux(15).Text, "N")
            If Val(Aux) <> Val(txtAux(9).Text) Then
                MsgBox "El pedido no pertenece al cliente indicado", vbExclamation
                Exit Function
            End If
        End If
    End If
    
    
    
    'Comprobar que existe el articulo en el almacen seleccionado
    Set vArtic = New CArticulo
    vArtic.Codigo = txtAux(1).Text
    If Not vArtic.ExisteEnAlmacen(txtAux(0).Text) Then
        B = False
        PonerFoco txtAux(1)
    End If
    Set vArtic = Nothing
    
'    devuelve = DevuelveDesdeBDNew(conAri, "salmac", "codartic", "codartic", txtAux(1).Text, "T", , "codalmac", txtAux(0).Text, "N")
'    If devuelve = "" Then
'        MsgBox "No existen unidades del Artículo: " & txtAux(1).Text & "  en el Almacen: " & txtAux(0).Text, vbExclamation
'        b = False
'        PonerFoco txtAux(1)
'    End If
    
    DatosOkLinea = B
    Exit Function
    
EDatosOkLinea:
    MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Sub Text2_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
    If index = 16 And KeyCode = 40 Then 'campo Ampliacion linea y Flecha hacia abajo
        PonerFocoBtn Me.cmdAceptar
    Else
        KEYdown KeyCode
    End If
End Sub

Private Sub Text2_KeyPress(index As Integer, KeyAscii As Integer)
    If index = 16 And KeyAscii = 13 Then 'campo Ampliación linea y ENTER
        PonerFocoBtn Me.cmdAceptar
    End If
End Sub

Private Sub Text2_LostFocus(index As Integer)
    If index = 16 And (Text2(index).Locked = False) Then Text2(index).Text = UCase(Text2(index).Text)
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.index
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
        Case 10  'Lineas
            mnLineas_Click
        Case 11 'Generar Albaran
            mnGenAlbaran_Click
        Case 12
            mnGeneraDtos_Click
        Case 14 'Imprimir Pedido
             mnImpPedido_Click
        Case 16
            SimularOtroProveedor
        Case 17   'Salir
            mnSalir_Click
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
Dim numlinea As String, vWhere As String
Dim cantidad As Currency
Dim J As Integer
Dim TipoDto  As Byte

On Error GoTo EInsertarLinea

    InsertarLinea = False
    SQL = ""

    If DatosOkLinea() Then 'Lineas de Pedidos
        'Conseguir el siguiente numero de linea
        vWhere = Replace(ObtenerWhereCP(False), NombreTabla, NomTablaLineas)
        numlinea = SugerirCodigoSiguienteStr(NomTablaLineas, "numlinea", vWhere)
        
        cantidad = ImporteFormateado(txtAux(3).Text)
        
        SQL = "INSERT INTO " & NomTablaLineas
        SQL = SQL & "(numpedpr,numlinea, codalmac, codartic, nomartic, ampliaci, cantidad, recibida, precioar, dtoline1, dtoline2, importel,codccost,codclien,coddirec,actuacion,codtipomV,numalbarV,fechaalbV,numpedV) "
        SQL = SQL & "VALUES (" & Val(Text1(0).Text) & ", " & numlinea & ", " & Val(txtAux(0).Text) & ","
        SQL = SQL & DBSet(txtAux(1).Text, "T") & ", " & DBSet(txtAux(2).Text, "T") & ", " & DBSet(Text2(16).Text, "T") & ", "
        SQL = SQL & DBSet(txtAux(3).Text, "N") & ", 0,"
        SQL = SQL & DBSet(txtAux(4).Text, "S") & "," & DBSet(txtAux(5).Text, "N") & ", "   'Mayo 2009   La "N" es ahora una "S"
        SQL = SQL & DBSet(txtAux(6).Text, "N") & ", " 'Dto 2
        SQL = SQL & DBSet(txtAux(7).Text, "N") & "," 'importe
        SQL = SQL & DBSet(txtAux(8).Text, "T", "S") & "," 'centro coste
        
        'Sept 2012. Cliente obra actuacion
        SQL = SQL & DBSet(txtAux(9).Text, "N", "S") & ", "  'cliente
        SQL = SQL & DBSet(txtAux(10).Text, "T", "S") & "," 'obra 'LE pongo TXT
        SQL = SQL & DBSet(txtAux(11).Text, "T", "S") 'actuac
        
         If vParamAplic.NumeroInstalacion = 4 Then
            SQL = SQL & "," & DBSet(txtAux(12).Text, "T", "S")
            SQL = SQL & "," & DBSet(txtAux(13).Text, "N", "S")
            SQL = SQL & "," & DBSet(txtAux(14).Text, "F", "S")
            SQL = SQL & "," & DBSet(txtAux(15).Text, "N", "S")
        Else
            SQL = SQL & ",NULL,NULL,NULL,NULL"
        End If
        
        
        SQL = SQL & ")"
    End If
    
    If SQL <> "" Then
        conn.Execute SQL
        InsertarLinea = True
    End If
    
    
    'Si el articulo es de conjuntos, preguntara si quiere insertar la lineas de los conjuntos
    If InsertarLinea = True Then
        SQL = DevuelveDesdeBD(conAri, "conjunto", "sartic", "codartic", txtAux(1).Text, "T")
        If SQL = "1" Then
        
            'SI!!!!!!, es de conjuntos
            If MsgBox("Articulo con componentes. Desea insertar las lineas?", vbQuestion + vbYesNo) <> vbYes Then Exit Function
            
            
            
            SQL = DevuelveDesdeBDNew(conAri, "sprove", "tipodtos", "codprove", Text1(4).Text, "N")
            TipoDto = CByte(SQL)
            
            
            SQL = "Select sarti1.*,nomartic from sarti1,sartic where sarti1.codarti1=sartic.codartic and sarti1.codartic=" & DBSet(txtAux(1).Text, "T")
            Set miRsAux = New ADODB.Recordset
            'miRsAux.Open SQL, conn, adOpenKeyset, adLockPessimistic, adCmdText
            miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            SQL = ""
            While Not miRsAux.EOF
                'Limpiamos todo menos el almacen y el CC si lo tuviera
                For J = 1 To 7
                    txtAux(J).Text = ""
                Next
            
                txtAux(1).Text = miRsAux!codarti1
                txtAux(2).Text = miRsAux!NomArtic
                'Cantidad es la cantidad de la linea ppal * la del escandallo
                txtAux(3).Text = cantidad * miRsAux!cantidad
            
                ObtenerPrecioCompra
            
                
                txtAux(7).Text = CalcularImporteSng(txtAux(3).Text, txtAux(4).Text, txtAux(5).Text, txtAux(6).Text, TipoDto)
            
            
            
                numlinea = Val(numlinea) + 1
                SQL = "INSERT INTO " & NomTablaLineas
                SQL = SQL & "(numpedpr,numlinea, codalmac, codartic, nomartic, ampliaci, cantidad, recibida, precioar, dtoline1, dtoline2, importel,codccost) "
                SQL = SQL & "VALUES (" & Val(Text1(0).Text) & ", " & numlinea & ", " & Val(txtAux(0).Text) & ","
                SQL = SQL & DBSet(txtAux(1).Text, "T") & ", " & DBSet(txtAux(2).Text, "T") & ", " & DBSet(Text2(16).Text, "T") & ", "
                SQL = SQL & DBSet(txtAux(3).Text, "N") & ", 0,"
                SQL = SQL & DBSet(txtAux(4).Text, "S") & "," & DBSet(txtAux(5).Text, "N") & ", "   'Mayo 2009   La "N" es ahora una "S"
                SQL = SQL & DBSet(txtAux(6).Text, "N") & ", " 'Dto 2
                SQL = SQL & DBSet(txtAux(7).Text, "N") & "," 'importe
                SQL = SQL & DBSet(txtAux(8).Text, "T", "S") 'centro coste
                SQL = SQL & ")"
            
            
                If Not ejecutar(SQL, True) Then MsgBox "Error insertando articulo componente: " & miRsAux!codArtic & " " & miRsAux!NomArtic, vbExclamation
            
            
            
            
                miRsAux.MoveNext
            Wend
            
        End If
        
    End If 'insertar =OK
    
    
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
        SQL = "UPDATE " & NomTablaLineas & " Set codalmac = " & txtAux(0).Text & ", codartic=" & DBSet(txtAux(1).Text, "T") & ", "
        SQL = SQL & "nomartic=" & DBSet(txtAux(2).Text, "T") & ", ampliaci=" & DBSet(Text2(16).Text, "T") & ", "
        SQL = SQL & "cantidad= " & DBSet(txtAux(3).Text, "N") & ", "
        SQL = SQL & "precioar= " & DBSet(txtAux(4).Text, "S") & ", "   'MAYO 2009.  La "N" es ahora una "S"
        SQL = SQL & "dtoline1= " & DBSet(txtAux(5).Text, "N") & ", dtoline2= " & DBSet(txtAux(6).Text, "N") & ", "
        SQL = SQL & "importel= " & DBSet(txtAux(7).Text, "N") & ", "
        SQL = SQL & "codccost= " & DBSet(txtAux(8).Text, "T", "S") & ","

        'Sept 2012
        'codclien , CodDirec,actuacion
        SQL = SQL & "codclien= " & DBSet(txtAux(9).Text, "N", "S") & ","
        SQL = SQL & "CodDirec= " & DBSet(txtAux(10).Text, "T") & ", "  'le he puesto TIPO Texto
        SQL = SQL & "actuacion= " & DBSet(txtAux(11).Text, "T", "S")
                
        'Agosto 2015. Euler
        If vParamAplic.NumeroInstalacion = 4 Then
            'codtipomV numalbarV fechaalbV
            SQL = SQL & "," & "codtipomv=" & DBSet(txtAux(12).Text, "T", "S")
            SQL = SQL & "," & "numalbarV=" & DBSet(txtAux(13).Text, "N", "S")
            SQL = SQL & "," & "fechaalbV=" & DBSet(txtAux(14).Text, "F", "S")
            SQL = SQL & "," & "numpedV=" & DBSet(txtAux(15).Text, "N", "S")
        End If
                    
                
        
        SQL = SQL & Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas) & " AND numlinea=" & Data2.Recordset!numlinea
    End If
    
    If SQL <> "" Then
        conn.Execute SQL
        ModificarLinea = True
    End If
    Exit Function
EModificarLinea:
    MuestraError Err.Number, "Modificar Lineas Pedido" & vbCrLf & Err.Description
End Function


Private Sub PonerBotonCabecera(B As Boolean)
'Pone el boton de Regresar a la Cabecera si pasamos a MAntenimiento de Lineas
'o Pone los botones de Aceptar y cancelar en Insert,update o delete lineas
On Error Resume Next

    Me.cmdAceptar.visible = Not B
    Me.cmdCancelar.visible = Not B
    Me.cmdRegresar.visible = B
    Me.cmdRegresar.Caption = "Cabecera"
    
    If cmdRegresar.visible Then
        cmdRegresar.Cancel = True
    Else
        cmdCancelar.Cancel = True
    End If
    
    If B Then
        Me.lblIndicador.Caption = "Líneas " & TituloLinea
        PonerFocoBtn Me.cmdRegresar
    End If
    
    'Habilitar las opciones correctas del menu según Modo
    PonerModoOpcionesMenu (Modo)
    PonerOpcionesMenu 'Habilitar las opciones correctas del menu según Nivel de Acceso
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub CargaGrid(ByRef vDataGrid As DataGrid, ByRef vData As Adodc, enlaza As Boolean, Optional conServidas As Boolean)
'IN: enlaza= si carga el grid con valores de la tabla o lo muestra vacio si no enlaza
'    conServidas=si enlaza, se muestra la columna de servidas solo cuando se va a generar el Albaran no completo
Dim B As Boolean
Dim SQL As String

On Error GoTo ECargaGrid

    B = DataGrid1.Enabled
    
    SQL = MontaSQLCarga(enlaza, conServidas)
    CargaGridGnral vDataGrid, vData, SQL, primeravez

    If conServidas Then
        vDataGrid.ClearFields
        vDataGrid.ReBind
        vDataGrid.Refresh
    End If
    
    CargaGrid2 vDataGrid, vData, conServidas
    vDataGrid.ScrollBars = dbgAutomatic
    
    B = (Modo = 5) And (ModificaLineas = 1 Or ModificaLineas = 2) '5:Modo Mto Lineas (Insertando o Modificando linea)
    vDataGrid.Enabled = Not B
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
    I = 1
    Select Case vDataGrid.Name
        Case "DataGrid1" 'Cod. Almacen
                I = I + 1
                vDataGrid.Columns(I).Caption = "Alm."
                If conServidas Then
                    vDataGrid.Columns(I).Width = 400
                Else
                    vDataGrid.Columns(I).Width = 450
                End If
                vDataGrid.Columns(I).NumberFormat = "000"
                
                I = I + 1
                vDataGrid.Columns(I).Caption = "Articulo"
                If conServidas Then
                    vDataGrid.Columns(I).Width = 1550
                Else
                    vDataGrid.Columns(I).Width = 1650
                End If
                
                I = I + 1
                vDataGrid.Columns(I).Caption = "Desc. Artículo"
                If conServidas Then
                    If vEmpresa.TieneAnalitica Then
                        vDataGrid.Columns(I).Width = 3000
                    Else
                        vDataGrid.Columns(I).Width = 3100
                    End If
                Else
                    If vEmpresa.TieneAnalitica Then
                        vDataGrid.Columns(I).Width = 3300
                    Else
                        vDataGrid.Columns(I).Width = 3500
                    End If
                End If
                
                I = I + 1
                'vDataGrid.Columns(i).Caption = "Ampl. Línea"
                'vDataGrid.Columns(i).Width = 7980
                vDataGrid.Columns(I).visible = False
                
                I = I + 1
                vDataGrid.Columns(I).Caption = "Cantidad"
                vDataGrid.Columns(I).Width = 900
                vDataGrid.Columns(I).Alignment = dbgRight
                vDataGrid.Columns(I).NumberFormat = FormatoImporte
                
                I = I + 1
                If conServidas Then
                    'Cargar el grid con la columna de cantidad servida
                    vDataGrid.Columns(I).Caption = "Recibidas"
                    vDataGrid.Columns(I).Width = 800
                    vDataGrid.Columns(I).Alignment = dbgRight
                    vDataGrid.Columns(I).NumberFormat = FormatoImporte
                    I = I + 1
                End If
                vDataGrid.Columns(I).Caption = "Precio"
                
                vDataGrid.Columns(I).Width = 1000
                
                vDataGrid.Columns(I).Alignment = dbgRight
                vDataGrid.Columns(I).NumberFormat = FormatoPrecio2   'Mayo 2009
                
                
                I = I + 1
                vDataGrid.Columns(I).Caption = "Dto.1"
                If conServidas Then
                    vDataGrid.Columns(I).Width = 500
                Else
                    vDataGrid.Columns(I).Width = 550
                End If
                vDataGrid.Columns(I).Alignment = dbgRight
                vDataGrid.Columns(I).NumberFormat = FormatoDescuento
                
                I = I + 1
                vDataGrid.Columns(I).Caption = "Dto.2"
                If conServidas Then
                    vDataGrid.Columns(I).Width = 500
                Else
                    vDataGrid.Columns(I).Width = 550
                End If
                vDataGrid.Columns(I).Alignment = dbgRight
                vDataGrid.Columns(I).NumberFormat = FormatoDescuento
            
                I = I + 1
                vDataGrid.Columns(I).Caption = "Importe"
                If conServidas Then
                    vDataGrid.Columns(I).Width = 900
                Else
                    vDataGrid.Columns(I).Width = 1000
                End If
                vDataGrid.Columns(I).Alignment = dbgRight
                vDataGrid.Columns(I).NumberFormat = FormatoImporte
                
                
                '---- [19/10/2009] [LAURA] : se añade el centro de coste
                I = I + 1
'                If vEmpresa.TieneAnalitica Then
'                    vDataGrid.Columns(i).Caption = "CCoste"
'                    If conServidas Then
'                        vDataGrid.Columns(i).Width = 650
'                    Else
'                        vDataGrid.Columns(i).Width = 700
'                    End If
'                Else
                    vDataGrid.Columns(I).visible = False   'lo hemos sacado fuera
                'End If
                
                'cliente,obra actuacion
                I = I + 1
                vDataGrid.Columns(I).visible = False
                vDataGrid.Columns(I + 1).visible = False
                vDataGrid.Columns(I + 2).visible = False
                
                 'Julio 2015
                'codtipomV numalbarV  fechaalbV numpedv
                vDataGrid.Columns(I + 3).visible = False
                vDataGrid.Columns(I + 4).visible = False
                vDataGrid.Columns(I + 5).visible = False
                vDataGrid.Columns(I + 6).visible = False
                
                

    End Select

    For I = 0 To vDataGrid.Columns.Count - 1
        vDataGrid.Columns(I).Locked = True
        vDataGrid.Columns(I).AllowSizing = False
    Next I
    vDataGrid.HoldFields
    
    Exit Sub
ECargaGrid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub


'Esta funcion sustituye a LlamaLineas
Private Sub CargaTxtAux(visible As Boolean, limpiar As Boolean)
'IN: visible: si es true ponerlos visibles en la posición adecuada
'    limpiar: si es true vaciar los txtAux
Dim alto As Single
Dim I As Byte

    If Not visible Then
        'Fijamos el alto (ponerlo en la parte inferior del form)
        For I = 0 To 7  '
            txtAux(I).Top = 290
            txtAux(I).visible = visible
        Next I
        cmdAux(0).visible = visible
        cmdAux(1).visible = visible
        cmdAux(2).visible = visible
        
        If Not Me.Data2.Recordset.EOF Then
            
            For I = 9 To 11
                Me.txtAux(I).Text = DBLet(Me.Data2.Recordset.Fields(I + 3), "T")
                PonerClieObraActuacion CInt(I), True
                BloquearTxt txtAux(I), True
            Next I
            If vParamAplic.NumeroInstalacion = 4 Then
                For I = 12 To 15
                    Me.txtAux(I).Text = DBLet(Me.Data2.Recordset.Fields(I + 3), IIf(I = 14, "F", "T"))
                    BloquearTxt txtAux(I), True
                Next I
                If txtAux(15).Text <> "" Then txtAux(15).Text = Format(txtAux(15).Text, "00000")
                PonerDatosAlbaranFacturaEuler
            End If
            
            BloquearTxt txtAux(8), True
        Else
            'EOF
            
            For I = 0 To txtAux.Count - 1
                If I >= 9 And I <= 11 Then Me.txtDesc(I).Text = ""
                txtAux(I).Text = ""
                BloquearTxt txtAux(I), True
            Next I
            txtDesc(0).Text = ""
            txtAux2(8).Text = ""
        End If
         
    Else
        If limpiar Then 'Vaciar los textBox (Vamos a Insertar)
            DeseleccionaGrid DataGrid1
            For I = 0 To txtAux.Count - 1
                If I >= 9 And I <= 11 Then Me.txtDesc(I).Text = ""
                txtAux(I).Text = ""
                BloquearTxt txtAux(I), False
            Next I
            txtDesc(0).Text = ""
        Else 'Vamos a modificar
            For I = 0 To txtAux.Count - 1
                If I < 3 Then
                    txtAux(I).Text = DataGrid1.Columns(I + 2).Text
                Else
                    txtAux(I).Text = DataGrid1.Columns(I + 3).Text
                End If
                txtAux(I).Locked = False
            Next I
            
            'Desbloqueamos el foragrid
            '
            For I = 8 To txtAux.Count - 1
                'txtAux(i).Text = ""
                BloquearTxt txtAux(I), False
            Next I
            
        End If
        
        'El campo Importe es calculado y lo bloqueamos.
        BloquearTxt txtAux(7), True
        
        
        ' ---- [20/10/2009] [LAURA] : añadir centro de coste
        'BloquearTxt txtAux(8), Not (vEmpresa.TieneAnalitica And vParamAplic.ModoAnalitica = 2)
        BloquearTxt txtAux(8), Not (vEmpresa.TieneAnalitica)
        Me.cmdAux(2).Enabled = (vEmpresa.TieneAnalitica)
        Me.cmdAux(2).visible = (vEmpresa.TieneAnalitica)
        ' ----
        
        

        'Fijamos altura(Height) y posición Top
        '-------------------------------
        alto = ObtenerAlto(DataGrid1, 20)
        
        For I = 0 To 7
            txtAux(I).Top = alto
            txtAux(I).Height = DataGrid1.RowHeight
        Next I
        cmdAux(0).Top = alto
        cmdAux(1).Top = alto
        'cmdAux(2).Top = alto
        cmdAux(0).Height = DataGrid1.RowHeight
        cmdAux(1).Height = DataGrid1.RowHeight
        cmdAux(2).Height = DataGrid1.RowHeight
        
        
        'Fijamos anchura y posicion Left
        '--------------------------------
        'Cod. Almac
        txtAux(0).Left = DataGrid1.Left + 330
        txtAux(0).Width = DataGrid1.Columns(2).Width - 160
        cmdAux(0).Left = txtAux(0).Left + txtAux(0).Width - 40
        'Cod Artic
        txtAux(1).Left = cmdAux(0).Left + cmdAux(0).Width + 20
        txtAux(1).Width = DataGrid1.Columns(3).Width - 160
        cmdAux(1).Left = txtAux(1).Left + txtAux(1).Width - 50
        'Nom Artic
        txtAux(2).Left = cmdAux(1).Left + cmdAux(1).Width
        txtAux(2).Width = DataGrid1.Columns(4).Width - 10
        'Cantidad
        txtAux(3).Left = txtAux(2).Left + txtAux(2).Width + 10
        txtAux(3).Width = DataGrid1.Columns(6).Width - 10
        'Precio, Dto1, Dto2, Precio
        For I = 4 To 7
            txtAux(I).Left = txtAux(I - 1).Left + txtAux(I - 1).Width + 10
            txtAux(I).Width = DataGrid1.Columns(I + 3).Width - 10
        Next I
        
        'cmdAux(2).Left = txtAux(i - 1).Left + txtAux(i - 1).Width - cmdAux(2).Width
        
                
        
        
        
        'Los ponemos Visibles o No
        '--------------------------
        For I = 0 To 7
            txtAux(I).visible = visible
        Next I
        txtAux(8).visible = visible And vEmpresa.TieneAnalitica
        cmdAux(0).visible = visible
        cmdAux(1).visible = visible

        
        
    End If
    
        
        
    Me.imgBuscar2(1).visible = visible And vParamAplic.NumeroInstalacion = 4
    Me.imgBuscar2(0).visible = visible And vParamAplic.NumeroInstalacion = 4
    Me.imgBuscar2(2).visible = visible And vParamAplic.NumeroInstalacion = 4
    'Anlitica
    'Me.imgBuscar2(9).visible = visible And vEmpresa.TieneAnalitica
    
    Me.imgBuscar2(9).visible = visible
    Me.imgBuscar2(10).visible = visible And vParamAplic.NumeroInstalacion <> 4
    Me.imgBuscar2(11).visible = visible And vParamAplic.NumeroInstalacion <> 4
        
        
End Sub


Private Sub CargaTxtAuxServidas(visible As Boolean, limpiar As Boolean)
'IN: visible: si es true ponerlos visibles en la posición adecuada
'    limpiar: si es true vaciar los txtAux
'Carga el TxtAux(3) con el campo RECIBIDAS de la tabla slippr
Dim alto As Single
Dim I As Byte

    I = 3
    If Not visible Then
        'Fijamos el alto (ponerlo en la parte inferior del form)
        txtAux(I).Top = 290
        txtAux(I).visible = visible
        txtAux(I).BackColor = vbWhite
    Else
        If limpiar Then 'Vaciar los textBox (Vamos a Insertar)
            DeseleccionaGrid DataGrid1
            txtAux(I).Text = ""
            BloquearTxt txtAux(I), False
            txtAux(I).BackColor = &H80000013
        End If
      
        'Fijamos altura(Height) y posición Top
        '-------------------------------
        If DataGrid1.Row < 0 Then
            alto = DataGrid1.Top + 230
        Else
            alto = DataGrid1.Top + DataGrid1.RowTop(DataGrid1.Row) + 20
        End If
        
        txtAux(I).Top = alto
        txtAux(I).Height = DataGrid1.RowHeight
        
        'Fijamos anchura y posicion Left
        '--------------------------------
        'Cantidad servida
        alto = DataGrid1.Left + 330 + DataGrid1.Columns(2).Width + DataGrid1.Columns(3).Width
        alto = alto + DataGrid1.Columns(4).Width + DataGrid1.Columns(6).Width
        txtAux(I).Left = alto
        txtAux(I).Width = DataGrid1.Columns(7).Width - 15
        
        'Los ponemos Visibles o No
        '--------------------------
        txtAux(I).visible = visible
        PonerFoco txtAux(I)
    End If
End Sub


Private Sub txtAux_GotFocus(index As Integer)
Dim cadkey As Integer
    If Modo <> 5 Then Exit Sub
    txtAnterior = txtAux(index).Text
    cadkey = ObtenerCadKey(kCampo, index)
    kCampo = index
    ConseguirFocoLin txtAux(index), cadkey
    If index = 3 Or index = 4 Then
        If Modo <> 6 Then
            If index = 3 Then
                lblF.Caption = "F2- Ver articulo"
            Else
                lblF.Caption = "F2- Ver precio          F3- Precio proveedor"
            End If
        End If
    Else
        lblF.Caption = ""
    End If
End Sub

Private Sub txtAux_KeyPress(index As Integer, KeyAscii As Integer)
    If Modo <> 6 Then
        KEYpress KeyAscii
    Else 'Pasar el Pedido a Albaran
        If KeyAscii = 13 Then 'ENTER
            PonerServidas True
'            ConseguirFoco txtAux(3), Modo
        End If
    End If
End Sub




Private Sub TxtAux_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    If Modo <> 6 Then 'Pasar de Pedido a Albaran
        ' ---- [02/11/2009] [LAURA] : al pulsar F2 para abrir articulos en la solapa Documentos|Pedidos
        If KeyCode = 113 Then
            If index = 3 Then AbrirForm_Articulos
            If index = 4 And txtAux(1).Text <> "" Then
                frmListadoPrecios.Opcion = 0
                frmListadoPrecios.CadenaPasoDatos = txtAux(1).Text & "|" & Text1(4).Text & "|"
                frmListadoPrecios.Show vbModal
            End If
        
        
        Else
            If KeyCode = 114 Then
                If index = 4 And txtAux(1).Text <> "" Then
                    frmListadoPrecios.Opcion = 1  'Presicos de slispr
                    frmListadoPrecios.CadenaPasoDatos = txtAux(1).Text & "|" & Text1(4).Text & "|"
                    frmListadoPrecios.Show vbModal
                End If
            Else
                If KeyCode = 43 Or KeyCode = 107 Or KeyCode = 187 Then
                      If index < 2 Or index = 8 Then  'Para los que tienen busqueda
                          If Modo = 5 And ModificaLineas = 1 Then
                              If txtAux(index).Text = "" Then
                                  PulsadoMas2 = True
                                  KeyCode = 0
                      
                                  PulsarTeclaMas False, index
                              End If
                          End If
                      End If
                  End If
            End If
        End If
    Else 'Modo lineas
        Select Case KeyCode
            Case 38 'Desplazamieto Fecha Hacia Arriba
                If DataGrid1.Row > 0 Then
                    DataGrid1.Row = DataGrid1.Row - 1
                    CargaTxtAuxServidas True, True
                Else
                    If Data2.Recordset.BOF Then
                        PonerFoco txtAux(3)
                    Else
                        gridCargado = False
                        Data2.Recordset.MovePrevious
                        gridCargado = True
                        If Data2.Recordset.BOF Then Data2.Recordset.MoveFirst
                         If DataGrid1.Row > 0 Then
                            DataGrid1.Row = DataGrid1.Row - 1
                            CargaTxtAuxServidas True, True

                        End If
                    End If
                End If
                txtAux(3).Text = Format(Data2.Recordset!recibida, FormatoImporte)
                ConseguirFoco txtAux(3), Modo
                
            Case 40 'Desplazamiento Flecha Hacia Abajo
'                If DataGrid1.Row < Data2.Recordset.RecordCount - 1 Then
'                    DataGrid1.Row = DataGrid1.Row + 1
'                    CargaTxtAuxServidas True, True
'                Else
'                    PonerFocoBtn Me.cmdAceptar
'                End If
'                txtAux(3).Text = Format(Data2.Recordset!recibida, FormatoImporte)
'                ConseguirFoco txtAux(3), Modo
                
                PonerServidas True
        End Select
    End If
End Sub



Private Sub txtAux_LostFocus(index As Integer)
Dim devuelve As String
'Dim vPrecio As CPreciosCom
Dim TipoDto As Byte
Dim B As Boolean


    If PulsadoMas2 Then
        'Para que cuando pulse el mas abra el form
        PulsadoMas2 = False
        txtAux(index).Text = ""
        Exit Sub
    End If

    If Not PerderFocoGnralLineas(txtAux(index), ModificaLineas) Then Exit Sub
    
    Select Case index
        Case 0 'Cod ALMACEN
            'Comprobar que existe el almacen
            devuelve = PonerAlmacen(txtAux(index).Text)
            txtAux(index).Text = devuelve
            If devuelve = "" Then PonerFoco txtAux(index)

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
            
            If ModificaLineas = 2 Then
                 If txtAux(1).Text <> Data2.Recordset!codArtic Then
                       
                    txtAux(4).Text = "": txtAux(5).Text = "": txtAux(6).Text = "": txtAux(7).Text = ""
                        
                       
                End If
            End If
            If PonerArticulo(txtAux(1), txtAux(2), txtAux(0).Text, CodTipoMov, ModificaLineas, , , devuelve) Then
                
                '---- [20/10/2009] [LAURA] : añadir centro de coste
                If Not vEmpresa.TieneAnalitica Then
                    txtAux(8).Text = ""
                ElseIf vParamAplic.ModoAnalitica = 1 Then 'por familia
                    txtAux(8).Text = devuelve
                    Me.txtAux2(8).Text = PonerNombreCCoste(Me.txtAux(8))
                End If
                
                'Ha camabiado
                If txtAnterior <> txtAux(1).Text Then
                    txtAux(4).Text = ""
                    txtAux(5).Text = ""
                    txtAux(6).Text = ""
                    txtAux(7).Text = ""
                End If
                
                '----
                B = (Me.ActiveControl.Name = "txtAux")
                If B Then B = (Me.ActiveControl.index = 0)
                
                If Not B Then
'                    If txtAux(2).Locked Then PonerFoco txtAux(3)
                Else
                    PonerFoco txtAux(0)
                End If
            Else
                PonerFoco txtAux(index)
            End If
            
            
'            If PonerArticulo(txtAux(1), txtAux(2), txtAux(0).Text, CodTipoMov) Then
'                If txtAux(2).Locked Then PonerFoco txtAux(3)
'                'Si es articulo de varios podemos modificar la descripción del articulo, sino bloqueamos.
''                If Not EsArticuloVarios(txtAux(Index).Text) Then
''                    BloquearTxt txtAux(2), True
''                Else
''                    BloquearTxt txtAux(2), False
''                    PonerFoco txtAux(2)
''                End If
'            Else
'                PonerFoco txtAux(Index)
'            End If
            
        Case 2 'Desc. Articulo
            If txtAux(index).Locked = False Then txtAux(index).Text = UCase(txtAux(index).Text)
            
        Case 3 'CANTIDAD
            If PonerFormatoDecimal(txtAux(index), 1) Then  'Tipo 1: Decimal(12,2)
                'Comprobar si hay suficiente stock
                If (Modo = 5) And (ModificaLineas = 1 Or (ModificaLineas = 2 And txtAux(4).Text = "")) Then 'Modo Insertar en Mto Lineas
                    'Obtener el precio correspondiente y los descuentos
                    ObtenerPrecioCompra
                    
'                    Set vPrecio = New CPreciosCom
'                    If vPrecio.Leer(txtAux(1).Text, Text1(4).Text) Then
'                        If vPrecio.ComprobarCantidad(CInt(txtAux(3).Text)) Then
'                            txtAux(4).Text = vPrecio.ObtenerPrecio(Text1(1).Text)
'                            PonerFormatoDecimal txtAux(4), 2
'                            txtAux(5).Text = vPrecio.Descuento1
'                            PonerFormatoDecimal txtAux(5), 4
'                            txtAux(6).Text = vPrecio.Descuento2
'                            PonerFormatoDecimal txtAux(6), 4
'                        Else
'                            PonerFoco txtAux(Index)
'                        End If
'                    End If
'                    Set vPrecio = Nothing
                End If
            End If
            
        Case 4 'Precio
            PonerFormatoDecimal_Single txtAux(index), 9 'Tipo 9: Decimal(10,5)parametros
        Case 5, 6 'Descuentos
            PonerFormatoDecimal txtAux(index), 4 'Tipo 4: Decimal(4,2)
        Case 7 'Importe Linea
            PonerFormatoDecimal txtAux(index), 1 'Tipo 3: Decimal(12,2)
            
        Case 8 'COD. CENTRO DE COSTE
            ' ---- [20/10/2009] [LAURA]: añadir centro de coste a la linea
            If txtAux(index).Text = "" Then
                 txtAux2(index).Text = ""
            ElseIf vEmpresa.TieneAnalitica Then
                'centro de coste
                ' ---- [20/10/2009] [LAURA]: añadir campo centro de coste familia
                Me.txtAux2(index).Text = PonerNombreCCoste(Me.txtAux(index))
            End If
            
        Case 9, 10, 11
            PonerClieObraActuacion index, False
            If index = 9 Then
                If vParamAplic.NumeroInstalacion = 4 Then LanzarBuscarAlbaranEuler False  'Abrimos
            End If
            
            'EULER
        Case 12
            txtAux(index).Text = UCase(txtAux(index).Text)
            
        Case 13
            'NUmero
            If Not PonerFormatoEntero(txtAux(index)) Then txtAux(index).Text = ""
            
        Case 14
            'Fecha
            If txtAux(index).Text <> "" Then PonerFormatoFecha txtAux(index)
        
        Case 15
                
            If Not PonerFormatoEntero(txtAux(index)) Then
                txtAux(index).Text = ""
            Else
            
                devuelve = DevuelveDesdeBD(conAri, "codclien", "scaped", "numpedcl", txtAux(index).Text, "N")
                If devuelve = "" Then
                    MsgBox "No existe el pedido: " & txtAux(index).Text, vbExclamation
                    txtAux(index).Text = ""
                Else
                    If txtAux(9).Text = "" Then
                        txtAux(9).Text = devuelve
                    Else
                        If Val(txtAux(9).Text) <> Val(devuelve) Then
                            MsgBox "El pedido no es de el cliente seleccionado:  " & Val(devuelve) & " // " & Val(txtAux(10).Text), vbExclamation
                            txtAux(index).Text = ""
                        Else
                            txtAux(12).Text = "": txtAux(13).Text = "": txtAux(14).Text = "": txtDesc(0).Text = ""
                        End If
                    End If
                End If
                If txtAux(index).Text = "" Then PonerFoco txtAux(index)
            End If
    
            
    End Select
    
    If index >= 12 And index <= 14 Then
        'Buscamos el albaran-factura
        PonerDatosAlbaranFacturaEuler
        
    End If
    
    
    If Modo = 5 Then
         If (index = 3 Or index = 4 Or index = 5 Or index = 6) Then 'Cant., Precio, Dto1, Dto2
'            If Trim(TxtAux(3).Text) = "" Or Trim(TxtAux(4).Text) = "" Then Exit Sub
'            If Trim(TxtAux(6).Text) = "" Or Trim(TxtAux(7).Text) = "" Then Exit Sub
            If txtAux(1).Text = "" Then Exit Sub
            TipoDto = DevuelveDesdeBDNew(conAri, "sprove", "tipodtos", "codprove", Text1(4).Text, "N")
            txtAux(7).Text = CalcularImporteSng(txtAux(3).Text, txtAux(4).Text, txtAux(5).Text, txtAux(6).Text, TipoDto)
            PonerFormatoDecimal txtAux(7), 1
        End If
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
            txtAux(4).Text = vPrecio.ObtenerPrecio(Text1(1).Text)    'FALTARA QUE DEVUELVE 5 decimales
'            PonerFormatoDecimal txtAux(4), 2
            txtAux(5).Text = vPrecio.Descuento1
'            PonerFormatoDecimal txtAux(5), 4
            txtAux(6).Text = vPrecio.Descuento2
'            PonerFormatoDecimal txtAux(6), 4
        Else
            PonerFoco txtAux(3)
            Exit Sub
        End If
    Else
        'Obtener el ult. precio de compra de ese articulo (sartic)
        cad = DevuelveDesdeBDNew(conAri, "sartic", "preciouc", "codartic", txtAux(1).Text, "T")
        If cad <> "" Then txtAux(4).Text = cad
        
        'Septiembre 2010   'Descuentos
        vPrecio.CodigoArtic = txtAux(1).Text
        vPrecio.CodigoProve = Text1(4).Text
        cad = vPrecio.ObtenerDescuentos2(Text1(1).Text, Aux2)
        If cad = "" Then cad = "0"
        If Aux2 = "" Then Aux2 = "0"
        txtAux(5).Text = cad
        txtAux(6).Text = Aux2
    
    End If
    PonerFormatoDecimal_Single txtAux(4), 9   '10,5
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
        FormularioListAlbAbierto = False
        PonerModo 5
        PonerBotonCabecera True
        DataGrid1_RowColChange 1, 1
End Sub


Private Function Eliminar() As Boolean
Dim B As Boolean
Dim vWhere As String
On Error GoTo FinEliminar

        conn.BeginTrans
         vWhere = ObtenerWhereCP(False)

'        If opt = 1 Then 'ELIMINAR
'            b = EliminarPedido(Data1.Recordset!numpedpr)
'        Else 'Pasar al HISTORICO
            B = ActualizarElTraspaso("", vWhere, CodTipoMov, CadenaSQL)
'        End If
        
FinEliminar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar Pedido"
        B = False
    End If
    If Not B Then
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
         vWhere = ObtenerWhereCP(False)
         vWhere = Replace(vWhere, NombreTabla & ".", "")
         If SituarData(Data1, vWhere, Indicador) Then
             PonerModo 2
             lblIndicador.Caption = Indicador
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


Private Function ObtenerWhereCP(conW As Boolean) As String
'Obtiene la where de la Clave Primaria de la tabla de Cabecera: scaped
Dim SQL As String
On Error Resume Next
    SQL = ""
    If conW Then SQL = " WHERE "
    SQL = SQL & NombreTabla & ".numpedpr= " & Val(Text1(0).Text)
    If EsHistorico Then SQL = SQL & " AND " & NomTablaLineas & ".fecpedpr=" & DBSet(Text1(1).Text, "F")
    ObtenerWhereCP = SQL
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
    
    SQL = "SELECT numpedpr, numlinea, codalmac, codartic, nomartic, ampliaci, cantidad, "
    If conServidas Then SQL = SQL & "recibida, "
'    SQL = SQL & "precioar, origpre, dtoline1, dtoline2,importel "
    SQL = SQL & "precioar, dtoline1, dtoline2,importel,codccost "
    
    'Neuvo Sept 2012
    SQL = SQL & ",codclien,coddirec,actuacion"
    
      'Julio 2015
    SQL = SQL & ",codtipomV,numalbarV, fechaalbV , numpedV"
    
    
    SQL = SQL & " FROM " & NomTablaLineas
    If enlaza Then
        SQL = SQL & " " & Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas)
        If EsHistorico Then SQL = SQL & " and fecpedpr='" & Format(Text1(1).Text, FormatoFecha) & "'"
    Else
        SQL = SQL & " WHERE numpedpr = -1"
    End If
    SQL = SQL & " Order by numpedpr, numlinea"
    MontaSQLCarga = SQL
End Function


Private Sub PonerModoOpcionesMenu(Modo As Byte)
'Activas unas Opciones de Menu y Toolbar según el modo en que estemos
Dim B As Boolean

        B = (Modo = 2) Or (Modo = 5 And ModificaLineas = 0)
        'Insertar
        Toolbar1.Buttons(5).Enabled = (B Or Modo = 0) And Not EsHistorico
        Me.mnNuevo.Enabled = (B Or Modo = 0) And Not EsHistorico
        'Modificar
        Toolbar1.Buttons(6).Enabled = B And Not EsHistorico
        Me.mnModificar.Enabled = B And Not EsHistorico
        'eliminar
        Toolbar1.Buttons(7).Enabled = B And Not EsHistorico
        Me.mnEliminar.Enabled = B And Not EsHistorico
            
        B = (Modo = 2) And Not EsHistorico
        'Mantenimiento lineas
        Toolbar1.Buttons(10).Enabled = (Modo = 2)
        Me.mnLineas.Enabled = (Modo = 2)
        'Generar Albaran desde Pedido
        Toolbar1.Buttons(11).Enabled = B
        Me.mnGenAlbaran.Enabled = B
        
        'Octubre 2011
        'Modifica descuentos
        Toolbar1.Buttons(12).Enabled = B
        Me.mnGeneraDtos.Enabled = B
        
        
        B = (Modo >= 3) Or Modo = 1
        'Buscar
        Toolbar1.Buttons(1).Enabled = Not B
        Me.mnBuscar.Enabled = Not B
        'Ver Todos
        Toolbar1.Buttons(2).Enabled = Not B
        Me.mnVerTodos.Enabled = Not B
End Sub


Private Function CargarDatosDirec(CodDirec As String, Indice As Byte) As Boolean
'Direcciones Propias
Dim RS As ADODB.Recordset
Dim devuelve As String
Dim B As Boolean
On Error GoTo ECargarProve

    B = False
    If CodDirec <> "" Then
        devuelve = "Select nomdirec, domdirec, codpobla, pobdirec, prodirec "
        devuelve = devuelve & " FROM sdirpr Where coddirec=" & Val(CodDirec)
        
        Set RS = New ADODB.Recordset
        RS.Open devuelve, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        If Not RS.EOF Then
            Text1(Indice).Text = Format(CodDirec, "000")
            Text2(Indice).Text = RS.Fields!nomdirec 'Nom Direccion
            If Indice = 2 Then
                Indice = 21
            Else
                Indice = 17
            End If
            Text2(Indice).Text = RS.Fields!domdirec 'Domicilio
            Text2(Indice + 1).Text = RS.Fields!codpobla
            Text2(Indice + 2).Text = RS.Fields!pobdirec
            Text2(Indice + 3).Text = RS.Fields!prodirec
            B = True
        Else
            MsgBox "No existe la dirección: " & Text1(Indice).Text, vbInformation
            LimpiarDatosDirec (Indice)
        End If
        RS.Close
        Set RS = Nothing
    Else
        LimpiarDatosDirec (Indice)
        B = True
    End If
    
    CargarDatosDirec = B
    
ECargarProve:
    If Err.Number <> 0 Then CargarDatosDirec = False
End Function


Private Sub LimpiarDatosDirec(Indice As Byte)
    Text2(Indice).Text = ""
    If Indice = 2 Then
        Indice = 21
    Else
        Indice = 17
    End If
    Text2(Indice).Text = "" 'Domicilio
    Text2(Indice + 1).Text = "" 'cpostal
    Text2(Indice + 2).Text = "" 'poblacion
    Text2(Indice + 3).Text = "" 'provincia
End Sub


Private Function InsertarPedido(vSQL As String, vTipoMov As CTiposMov) As Boolean
'Insertar la Cabecera de un Pedido, tabla: scaped
Dim MenError As String
Dim bol As Boolean, Existe As Boolean
Dim cambiaSQL As Boolean
Dim devuelve As String
On Error GoTo EInsertarOferta
    
    bol = True
    
    cambiaSQL = False
    'Comprobar si mientras tanto se incremento el contador de Pedidos
    'para ello vemos si existe un Pedido con ese contador y si existe lo incrementamos
    Do
        devuelve = DevuelveDesdeBDNew(conAri, NombreTabla, "numpedpr", "numpedpr", Text1(0).Text, "N")
        If devuelve <> "" Then
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
    MenError = "Error al insertar en la tabla Cabecera de Pedidos (" & NombreTabla & ")."
    conn.Execute vSQL, , adCmdText
    
    
    'Actualizar los datos del proveedor si es de varios
    If EsDeVarios Then
        'Si es cliente de varios actualizar datos cliente en tabla:sclvar
        MenError = "Modificando datos proveedor varios."
        bol = ActualizarProveVarios(Text1(4).Text, Text1(6).Text)
    End If
    
    MenError = "Error al actualizar el contador del Pedido."
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


Private Sub LimpiarDatosProve()
'Limpia los campos del Form con datos del Proveedor
Dim I As Byte

    For I = 4 To 14
        Text1(I).Text = ""
    Next I
End Sub
    





Private Function PasarPedidoAAlbaran(NumAlb As String, FechaAlb As String) As Boolean
'OUT -> numalb: Devuelve el Nº de albaran asignado al pedido
Dim bol As Boolean
Dim MenError As String
Dim SQL As String
Dim RS As ADODB.Recordset
Dim vWhere As String
Dim cProve As CProveedor

    On Error GoTo EGenPedido

    bol = False
            
    
    'Aqui empieza transaccion
    conn.BeginTrans
    
    'Insertar en tablas de Albaranes Proveedor el Pedido  (scaalp, slialp)
    bol = InsertarAlbaran(MenError, NumAlb)
    
    
    
    
    
    
    'Para cada linea del pedido:
    ' Actualizar precio medio ponderado del articulo
    ' Actualizar precio y fecha ultima compra del articulo
    
    
    '17 Febrero 2011
    'LO QUITAMOS DE AQUI
'    If bol Then
'        MenError = "Actualizando Stocks"
'        bol = InsertarMovStock2(NumAlb, FechaAlb)
'    End If

    If bol Then
        'Actualizar la ult.fecha de compra del Proveedor
        MenError = "Actualizando ultima fecha compra en Proveedor."
        Set cProve = New CProveedor
        bol = cProve.ActualizaFechaUltCompra(Text1(4).Text, FechaAlb)
        Set cProve = Nothing
        
'        If bol Then
'            'Actualizar ult. fecha de compra y el precio ult compra de los articulos del Albaran
'            MenError = "Actualizando ultima fecha compra en Artículos."
'            SQL = "numalbar=" & DBSet(NumAlb, "T") & " and fechaalb=" & DBSet(FechaAlb, "F") & " and slialp.codprove=" & Text1(4).Text
'            bol = ActualizarUltFechaCom(SQL)
'        End If
    End If
    
    
    If bol Then
        If AlbCompleto Then  'Si se inserta Albaran
            'Borrar el Pedido de las tablas de Pedidos (scaped, sliped)
            MenError = "Eliminando cabecera y lineas del Pedido."
            bol = EliminarPedido(CLng(Text1(0).Text))
        Else
            'Actualizar la cantidad=cantidad-recibida y recibida= 0 en slippr
            bol = ActualizarPedido()
            'Marcar Resto de pedido: restoped=1
            If bol Then bol = ActualizarCabPedido
        End If
    End If
    
    
    
    If bol Then
        'si se ha generado correctamente el ALBARAN ver si hay alguna línea que tiene
        'el artículo con control de nº de lote y pedir los nº de lotes.
        ComprobarNumLotesLineas NumAlb, FechaAlb
        
    End If
    
    
    
    
    If bol Then
        'Se ha generado correctamente el ALBARAN y vemos si tiene Nº Series
'        FechaAlb = RecuperaValor(CadenaSQL, 3)
        'Comprobar si Hay Nº SERIE en compras y Mostrar
        'ventana para pedir los Nº Serie de la cantidad introducida si lo requiere algun articulo
        ComprobarNSeriesLineas NumAlb, FechaAlb
        
        
        If Not AlbCompleto Then
            'Eliminar las filas del pedido que se servieron completas (slippr)
            MenError = "Eliminando lineas pedidido servidas completas."
            SQL = "DELETE FROM " & NomTablaLineas & Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas) & " AND cantidad=0"
            conn.Execute SQL
            
            'Comprobar que si no quedan lineas en el pedido se elimine la cabecera del pedido
            MenError = "Eliminando cabecera del pedido."
            SQL = "select codalmac, codartic FROM " & NomTablaLineas & " WHERE numpedpr=" & Text1(0).Text
            Set RS = New ADODB.Recordset
            RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If RS.EOF Then 'No hay lineas de pedido --> Eliminar la cabecera
                SQL = "DELETE FROM " & NombreTabla & " WHERE numpedcl=" & Text1(0).Text
                conn.Execute SQL
            End If
            RS.Close
            Set RS = Nothing
        End If
        bol = True
    End If
    
    
EGenPedido:
    If Err.Number <> 0 Then
'        MenError = "Pasando Pedido a Albaran." & vbCrLf & "----------------------------" & vbCrLf & MenError
'        MuestraError Err.Number, MenError, Err.Description
        bol = False
    End If
    If bol Then
        conn.CommitTrans
        PasarPedidoAAlbaran = True
    Else
        conn.RollbackTrans
        PasarPedidoAAlbaran = False
        MenError = "Pasando Pedido a Albaran." & vbCrLf & "----------------------------" & vbCrLf & MenError
        MuestraError Err.Number, MenError, Err.Description
    End If
End Function


Private Function InsertarAlbaran(MenError As String, NumAlb As String) As Boolean
'Devuelve el mensaje de error si se produce
Dim bol As Boolean
Dim vSQL As String
Dim FechaAlb As String
Dim TrabAlb As String

    On Error GoTo EInsertarAlbaran
    
    bol = False
    InsertarAlbaran = bol
    
    NumAlb = RecuperaValor(CadenaSQL, 2)
    FechaAlb = RecuperaValor(CadenaSQL, 3)
    TrabAlb = RecuperaValor(CadenaSQL, 1)
    
    vSQL = "INSERT INTO scaalp (numalbar, fechaalb, codprove, nomprove, domprove, codpobla, pobprove, proprove, nifprove, telprove, codforpa, codtraba, codtrab1, dtoppago, dtognral, observa1, observa2, observa3, observa4, observa5, "
    vSQL = vSQL & " numpedpr, fecpedpr,codenvio,NReferencia,SReferencia,fecentrega)"
    vSQL = vSQL & " SELECT " & DBSet(NumAlb, "T") & " as numalbar, " & DBSet(FechaAlb, "F") & " as fechaalb, "
    vSQL = vSQL & "codprove, nomprove, domprove, codpobla, pobprove, proprove, nifprove, telprove, codforpa, "
    vSQL = vSQL & TrabAlb & " as codtraba,codtraba as codtrab1, dtoppago, dtognral, observa1, observa2, observa3, observa4, observa5 "
    vSQL = vSQL & " ,numpedpr, fecpedpr,codenvio,NReferencia,SReferencia,fecentrega"
    vSQL = vSQL & " FROM " & NombreTabla & " WHERE numpedpr=" & Text1(0).Text

    'Insertar Cabecera
    MenError = "Error al insertar en la tabla Cabecera de Albaranes Proveedor (scaalp)."
    conn.Execute vSQL, , adCmdText
    
    'Insertar Lineas Albaran desde Pedido
    MenError = "Error al insertar en la tabla Lineas de Albaran (slialp)."
    If Not InsertarLineasAlbaran(NumAlb, FechaAlb) Then Exit Function
    
    bol = True
    
EInsertarAlbaran:
        If Err.Number <> 0 Then
            bol = False
            MenError = MenError & vbCrLf & Err.Description
        End If
        If bol Then
            InsertarAlbaran = True
        Else
            InsertarAlbaran = False
        End If
End Function


Private Function InsertarLineasAlbaran(NumAlb As String, FechaAlb As String) As Boolean
'Inserta en la tabla de lineas de albaran (slialb)
'IN -> TipoM, numAlb
Dim SQL2 As String
Dim RS As ADODB.Recordset
Dim ImpLinea As String
Dim TipoDto As Byte
'Dim InsertDirecto As Boolean
Dim cantidad As Currency
Dim ImpReciclado As Single
Dim numlinea As Integer
Dim ErrorFechaInventario As String
On Error GoTo EInsertarLinAlb


    
        'NO insert directo.
        'Es o bien pq no es completio o pq tiene tasa reciclado
        SQL2 = "select * from " & NomTablaLineas
        SQL2 = SQL2 & Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas)
        SQL2 = SQL2 & " ORDER BY numlinea"
        Set RS = New ADODB.Recordset
        RS.Open SQL2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        numlinea = 1
        While Not RS.EOF 'Para cada linea de pedido insertar una de albaran si recibidas >0
            SQL2 = ""
            If AlbCompleto Then
                'Va la linea entera
                SQL2 = "INSERT INTO slialp (numalbar, fechaalb, codprove, numlinea,codartic, codalmac, nomartic, ampliaci, cantidad, precioar, dtoline1, dtoline2, importel,codccost,codclien,coddirec,actuacion,codtipomV,numalbarV, fechaalbV,numpedV ) "
                SQL2 = SQL2 & " VALUES(" & DBSet(NumAlb, "T") & ", " & DBSet(FechaAlb, "F") & ", " & Val(Text1(4).Text) & ", " & numlinea & ", "
                SQL2 = SQL2 & DBSet(RS!codArtic, "T") & "," & RS!codAlmac & ", " & DBSet(RS!NomArtic, "T") & ", " & DBSet(RS!Ampliaci, "T") & ", "
                SQL2 = SQL2 & DBSet(RS!cantidad, "N") & ", " & DBSet(RS!precioar, "S") & ", " & DBSet(RS!dtoline1, "N") & ", " & DBSet(RS!dtoline2, "N") & ", "
                SQL2 = SQL2 & DBSet(RS!ImporteL, "N") & ","
                SQL2 = SQL2 & DBSet(RS!CodCCost, "T", "S") & ","
                'Sept 2012   client obra actuacion
                SQL2 = SQL2 & DBSet(RS!codClien, "N", "S") & "," & DBSet(RS!CodDirec, "N", "S") & "," & DBSet(RS!actuacion, "T", "S") & ","
                
                'Agosto 2015
                If vParamAplic.NumeroInstalacion = 4 Then
                    SQL2 = SQL2 & DBSet(RS!codtipomv, "T", "S") & "," & DBSet(RS!numalbarV, "T", "S") & "," & DBSet(RS!fechaalbV, "F", "S") & "," & DBSet(RS!numpedv, "T", "S")
                Else
                    SQL2 = SQL2 & "NULL,NULL,NULL,NULL"
                End If
                SQL2 = SQL2 & ")"
                
                
                
                cantidad = RS!cantidad
                ImpLinea = RS!ImporteL
            Else
                If RS!recibida > 0 Then
                    TipoDto = DevuelveDesdeBDNew(conAri, "sprove", "tipodtos", "codprove", Text1(4).Text, "N")
                    ImpLinea = CalcularImporte(RS!recibida, RS!precioar, RS!dtoline1, RS!dtoline2, TipoDto)
                    SQL2 = "INSERT INTO slialp (numalbar, fechaalb, codprove, numlinea,codartic, codalmac, nomartic, ampliaci, cantidad, precioar, dtoline1, dtoline2, importel,codccost,codclien,coddirec,actuacion,codtipomV,numalbarV, fechaalbV,numpedV ) "
                    SQL2 = SQL2 & " VALUES(" & DBSet(NumAlb, "T") & ", " & DBSet(FechaAlb, "F") & ", " & Val(Text1(4).Text) & ", " & numlinea & ", "
                    SQL2 = SQL2 & DBSet(RS!codArtic, "T") & "," & RS!codAlmac & ", " & DBSet(RS!NomArtic, "T") & ", " & DBSet(RS!Ampliaci, "T") & ", "
                    SQL2 = SQL2 & DBSet(RS!recibida, "N") & ", " & DBSet(RS!precioar, "S") & ", " & DBSet(RS!dtoline1, "N") & ", " & DBSet(RS!dtoline2, "N") & ", "
                    SQL2 = SQL2 & DBSet(ImpLinea, "N") & ","
                    SQL2 = SQL2 & DBSet(RS!CodCCost, "T", "S") & ","
                    'Sept 2012   client obra actuacion
                    SQL2 = SQL2 & DBSet(RS!codClien, "N", "S") & "," & DBSet(RS!CodDirec, "N", "S") & "," & DBSet(RS!actuacion, "T", "S") & ","
                    'Agosto 2015
                    If vParamAplic.NumeroInstalacion = 4 Then
                        SQL2 = SQL2 & DBSet(RS!codtipomv, "T", "S") & "," & DBSet(RS!numalbarV, "T", "S") & "," & DBSet(RS!fechaalbV, "F", "S") & "," & DBSet(RS!numpedv, "T", "S")
                    Else
                        SQL2 = SQL2 & "NULL,NULL,NULL,NULL"
                    End If
                    SQL2 = SQL2 & ")"
    
                    cantidad = RS!recibida
                End If
            End If
            
            
            
            If SQL2 <> "" Then
                
                conn.Execute SQL2, , adCmdText
                
                
                'AQui habria que hacer lo del stock
                InsertarMovStock3 NumAlb, FechaAlb, numlinea, cantidad, CCur(ImpLinea), RS!codAlmac, RS!codArtic
                
                
                
                
                numlinea = numlinea + 1
                'TASA RECILCADO
                If vParamAplic.ArtReciclado <> "" Then
                    If ArticuloConTasaReciclado(CStr(RS!codArtic), ImpReciclado) Then
                        ImpLinea = Round2(cantidad * ImpReciclado, 2)
                        SQL2 = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", vParamAplic.ArtReciclado, "T")
                        'OCTUBRE 2011
                        'Error. Ponia rs!codartic en lugar de artrecicla: SQL2 = numlinea & ", " & DBSet(RS!codArtic, "T") .....
                        SQL2 = numlinea & ", " & DBSet(vParamAplic.ArtReciclado, "T") & "," & RS!codAlmac & ", " & DBSet(SQL2, "T") & ", " & DBSet("", "T") & ", "
                        SQL2 = " VALUES(" & DBSet(NumAlb, "T") & ", " & DBSet(FechaAlb, "F") & ", " & Val(Text1(4).Text) & ", " & SQL2
                        'SQL2 = SQL2 & DBSet(Cantidad, "N") & ", " & DBSet(ImpReciclado, "S") & ", " & DBSet(RS!dtoline1, "N") & ", " & DBSet(RS!dtoline2, "N") & ", "
                        SQL2 = SQL2 & DBSet(cantidad, "N") & ", " & DBSet(ImpReciclado, "S") & ",0,0,"
                        SQL2 = SQL2 & DBSet(ImpLinea, "N") & ","
                        SQL2 = SQL2 & DBSet(RS!CodCCost, "T", "S") & ")"
                        SQL2 = "INSERT INTO slialp (numalbar, fechaalb, codprove, numlinea,codartic, codalmac, nomartic, ampliaci, " & _
                            "cantidad, precioar, dtoline1, dtoline2, importel,codccost) " & SQL2
                        
                        conn.Execute SQL2
                        numlinea = numlinea + 1
            
                    End If
                End If
                
            End If
            RS.MoveNext
        Wend
        RS.Close
        Set RS = Nothing
    'End If
    
EInsertarLinAlb:
    If Err.Number <> 0 Then
         'Hay error , almacenamos y salimos
        InsertarLineasAlbaran = False
        MuestraError Err.Number, "Insertar lineas albaran.", Err.Description
    Else
        InsertarLineasAlbaran = True
    End If
End Function



Private Function EliminarPedido(numPed As Long) As Boolean
'Eliminar las lineas y la Cabecera de un Pedido. Tablas: scaped, sliped
Dim SQL As String
On Error GoTo EEliminarPed

     SQL = " WHERE  numpedpr=" & numPed

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
Dim TipoDto As Byte

    On Error GoTo EActPedido

    SQL = "select numlinea, codalmac, codartic, cantidad, recibida, precioar, dtoline1, dtoline2 from " & NomTablaLineas
    SQL = SQL & Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas)
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not RS.EOF 'Para cada linea
        TipoDto = DevuelveDesdeBDNew(conAri, "sprove", "tipodtos", "codprove", Text1(4).Text, "N")
        ImpLinea = CalcularImporte(RS!cantidad - RS!recibida, RS!precioar, RS!dtoline1, RS!dtoline2, TipoDto)
        SQL = "UPDATE " & NomTablaLineas & " SET cantidad=cantidad-recibida, recibida=0, importel=" & DBSet(ImpLinea, "N")
'        SQL = SQL & " WHERE codalmac=" & RS!codAlmac & " AND codartic='" & RS!codArtic & "'"
        SQL = SQL & Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas)
        SQL = SQL & " AND numlinea=" & RS!numlinea
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

    SQL = "UPDATE " & NombreTabla & " SET restoped=1 " & ObtenerWhereCP(True)
    conn.Execute SQL
    If Err.Number <> 0 Then
        ActualizarCabPedido = False
    Else
        ActualizarCabPedido = True
    End If
End Function


Private Function InsertarMovStock3(NumAlb As String, FechaAlb As String, NLin As Integer, cantidad As Currency, Importe As Currency, codAlmac As Integer, codArtic As String) As Boolean
Dim vCStock As CStock
Dim B As Boolean
Dim SQL As String
Dim cart As CArticulo


    'No lleva error, que salte en la rutina ppal
    On Error Resume Next

    InsertarMovStock3 = False
    
    Set vCStock = New CStock
    B = True
   
    vCStock.FechaMov = FechaAlb
    vCStock.tipoMov = "E"
    vCStock.DetaMov = "ALC"
    vCStock.Trabajador = CLng(Text1(4).Text) 'En codigope ponemos el Proveedor
    vCStock.codArtic = codArtic
    vCStock.codAlmac = CInt(codAlmac)
    
    
    vCStock.cantidad = CSng(cantidad)
    vCStock.Importe = CCur(Importe)
    
    
    vCStock.LineaDocu = NLin
    vCStock.Documento = NumAlb
    If vCStock.cantidad <> 0 Then
        '==== Laura 22/09/2006
        '-- antes de actualizar el stock calculamos el precio medio ponderado del articulo
        Set cart = New CArticulo
        If cart.LeerDatos(vCStock.codArtic) Then
            '17 Junio 2009
            'Si la cantidad es negativa no actualiza ni precio medio ponderado NI fecha ult compra
            If vCStock.cantidad >= 0 Then
            
                'Laura 19/12/2006: Calcular precio_med_pond con el precio con los descuentos,e.d. importe/cantidad
                'If Not cArt.ActualizarPrecioMedPond(CCur(vCStock.Cantidad), CCur(RS!precioar)) Then b = False
                If Not cart.ActualizarPrecioMedPond(CCur(vCStock.cantidad), Round2(CCur(vCStock.Importe) / CCur(vCStock.cantidad), 4)) Then B = False
                
                '--actualizar fecha y precio ultima compra del articulo
                'Laura 19/12/2006: actualizar precio_ult_compra con el precio con los descuentos,e.d. importe/cantidad
                'If Not cArt.ActualizarUltFechaCompra(vCStock.Fechamov, CStr(RS!precioar)) Then b = False
                If Not cart.ActualizarUltFechaCompra(vCStock.FechaMov, Round2(CCur(vCStock.Importe) / CCur(vCStock.cantidad), 4)) Then B = False


                

            End If 'De cantidad >=0
        End If
        Set cart = Nothing
        '====
    
    
        'en actualizar stock comprobamos si el articulo tiene control de stock
        B = vCStock.ActualizarStock
    
    Else
        B = True  'Si no inserta pq la cantidad es cero n pasa nada
    End If
    InsertarMovStock3 = B
    
End Function

Private Sub ImprimirAlbaran(Numalbar As String, FechaAlb As String, Codprove As Long)
Dim cadNomRPT As String
Dim SQL As String
Dim numP As Byte
Dim param As String

    
    
    'Albaran socio
    If Not PonerParamRPT2(27, param, numP, cadNomRPT, pImprimeDirecto, pPdfRpt, pRptvMultiInforme) Then Exit Sub
    
    
    



    
    
    SQL = CadenaDesdeHasta(CStr(FechaAlb), CStr(FechaAlb), "{scaalp.fechaalb}", "F")
    SQL = SQL & " AND  {scaalp.codprove} = " & Codprove
    SQL = SQL & " AND  {scaalp.numalbar} = """ & DevNombreSQL(Numalbar) & """"
    



    
     With frmImprimir
        .FormulaSeleccion = SQL
        .OtrosParametros = param
        .NumeroParametros = numP
        .SeleccionaRPTCodigo = pRptvMultiInforme
        .SoloImprimir = False
        .EnvioEMail = False
        .Opcion = 2000 + 10   '2000 mas la opcion de entrada
        .NombrePDF = ""
        '.NombrePDF = cadPDFrpt
        .NombreRPT = cadNomRPT
        .ConSubInforme = True
        .Show vbModal
    End With
End Sub


Private Function ActualizarServidas() As Boolean
Dim SQL As String
On Error Resume Next

    SQL = "UPDATE " & NomTablaLineas & " SET recibida= " & DBSet(txtAux(3).Text, "N")
    SQL = SQL & Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas) & " AND numlinea=" & Data2.Recordset!numlinea
    conn.Execute SQL
    
    If Err.Number <> 0 Then
        ActualizarServidas = False
    Else
        ActualizarServidas = True
    End If
End Function


Private Sub PonerServidas(HaciaAlante As Boolean)
Dim NumFila As Integer
Dim cadMen As String

'    NumFila = DataGrid1.Row
    NumFila = Data2.Recordset.AbsolutePosition
    If PonerFormatoDecimal(txtAux(3), 1) Then  'Tipo 1: Decimal(12,2)
        If CCur(txtAux(3).Text) > Data2.Recordset!cantidad Then
            cadMen = "La cantidad a Recibir no puede ser superior a la del Pedido."
            MsgBox cadMen, vbExclamation
            PonerFoco txtAux(3)
            Exit Sub
        End If
    End If
    ActualizarServidas
    CargaGrid2 DataGrid1, Data2, True
'    DataGrid1.Row = NumFila
    SituarDataPosicion Data2, CLng(NumFila), ""
    If HaciaAlante Then MoverSigRegistro
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
    txtAux(3).Text = Format(Data2.Recordset!recibida, FormatoImporte)
    PonerFoco txtAux(3)
    ConseguirFocoLin txtAux(3)
    txtAux(3).Refresh
EMover:
    If Err.Number <> 0 Then MuestraError Err.Description, "Mover registro.", Err.Description
End Sub





Private Sub GenerarAlbaran()
Dim numPed As Long 'Nº Pedido
Dim NumAlb As String 'Nº Albaran
Dim FechaAlb As String 'Fecha del Albaran
Dim SQL As String
Dim RS As ADODB.Recordset
Dim B As Boolean
Dim ImprimeAlb As Long   'Si queremos imprimir guardare el codprove
Dim ArticuloEsEscandallo As String



    NumRegElim = Data1.Recordset.AbsolutePosition
    numPed = Data1.Recordset!numpedpr
    
    'pedir por pantalla: el operador, Nº albaran y fecha albaran
    Set frmList = New frmListadoOfer
    
    frmList.codClien = Text1(4).Text  'Julio18
    
    frmList.OpcionListado = 57
    CadenaSQL = ""
    frmList.Show vbModal
    Set frmList = Nothing
    
    If CadenaSQL = "" Then Exit Sub
    FechaAlb = RecuperaValor(CadenaSQL, 3)
    SQL = RecuperaValor(CadenaSQL, 4)
    ImprimeAlb = -1
    If SQL = "1" Then ImprimeAlb = CLng(Text1(4).Text)
    
    
    'Mostraremos un msg si algunos de los articulos tienen fecha inventario posterior
    SQL = "SELECT  codalmac,salmac.codartic,nomartic,fechainv FROM salmac,sartic where salmac.codartic=sartic.codartic and artvario=0 and "
    SQL = SQL & " fechainv > " & DBSet(FechaAlb, "F")
    SQL = SQL & " and (codalmac,salmac.codartic) in ("
    SQL = SQL & " select codalmac,codartic from slippr WHERE numpedpr=" & numPed
    'seleccionar solo de las que se vayan a recibir
    If Not AlbCompleto Then SQL = SQL & " and slippr.recibida>0 "
    SQL = SQL & ")"
    B = ObtenerRSprecios(RS, SQL)
    SQL = ""
    If Not B Then
        MsgBox "Error obteniendo datos cruzados con inventarios", vbExclamation
    Else
        If Not RS.EOF Then
            
            While Not RS.EOF
                SQL = SQL & "   -" & RS!codArtic & "  " & RS!NomArtic & "   inventariado el " & RS!FechaINV & vbCrLf
                RS.MoveNext
            Wend
            
            
            If SQL <> "" Then
                SQL = "Las siguientes referencias tiene fecha inventario posterior al del albaran:" & vbCrLf & vbCrLf & SQL
                SQL = SQL & vbCrLf & "¿Continuar?"
                If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then SQL = ""
            End If
        End If
    End If
    RS.Close
    Set RS = Nothing
    If SQL <> "" Then Exit Sub
    
    
    'Antes de pasar el pedido al albaran nos guardamos los articulos cuyo precio_compra
    'se han modificado para preguntar despues si se quiere actualizar precios_venta
    'hay q guardarlo antes de pasar pedido a albaran ya q aqui se actualiza el precio_ult_compra
    '-- Laura 19/12/2006: calcular precio_med_pond con el precio aplicados los descuentos, ed. importe/cantidad
    ' Iremos cambiando el numero de decimales poc a poco ANTES era un 4
    SQL = "SELECT slippr.codartic,sartic.nomartic,round(slippr.importel/slippr.cantidad," & PrecioDecimales & ")"
    SQL = SQL & " as precioar,sartic.preciouc,sum(cantidad) "
    SQL = SQL & " FROM slippr INNER JOIN sartic ON slippr.codartic=sartic.codartic "
    'SQL = SQL & " WHERE numpedpr=" & numPed & " and (slippr.precioar<>sartic.preciouc)"
    SQL = SQL & " WHERE numpedpr=" & numPed & " and (round(slippr.importel/slippr.cantidad,4)<>sartic.preciouc)"
    'seleccionar solo de las que se vayan a recibir
    If Not AlbCompleto Then SQL = SQL & " and slippr.recibida>0 "
    SQL = SQL & " group by slippr.codartic,slippr.precioar,sartic.preciouc "
    SQL = SQL & " Having Sum(Cantidad) > 0"
    B = ObtenerRSprecios(RS, SQL)
    
    
    
    If PasarPedidoAAlbaran(NumAlb, FechaAlb) Then
        'Imprime los pedidos de cliente vinculados con los articulos del albaran de proveedor generado
        If Not ComprobarPedidosClientesDesdeAlbProveedor(NumAlb, CDate(FechaAlb), Text1(4).Text) Then MsgBox "Se ha generado correctamente el Albaran: " & NumAlb, vbInformation
                
'        FechaAlb = RecuperaValor(CadenaSQL, 3)
'        'Comprobar si Hay Nº SERIE en compras y Mostrar
'        'ventana para pedir los Nº Serie de la cantidad introducida si lo requiere algun articulo
'        ComprobarNSeriesLineas NumAlb, FechaAlb

        PonerModo 2
        
        
        'comprobar si hay lineas de artículos cuyo precio_ultima_compra
        'se ha modificado y preguntar si que quieren actualizar los precio_venta
        '--------------------------------------------------------
        If B Then
            ArticuloEsEscandallo = ""
            While Not RS.EOF
            
                'Primero compruebo si es escandallo de otro select count(*) from ariges3.sarti1 where codarti1='0020080939'
                SQL = DevuelveDesdeBD(conAri, "count(*)", "sarti1", "codarti1", RS!codArtic, "T")
                If SQL <> "" Then
                    If Val(SQL) > 0 Then
                        ArticuloEsEscandallo = ArticuloEsEscandallo & RS!codArtic & "|"
                    End If
                End If
                SQL = "Se ha modificado el precio última compra del artículo:" & vbCrLf
                SQL = SQL & RS!codArtic & ":  " & RS!NomArtic & vbCrLf
                SQL = SQL & vbCrLf & "¿Desea actualizar los precios de venta?"
                If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
                    'Comprobar que el artículo tiene margen comercial
                    If ArticuloTieneMargen(RS!codArtic) Then
                        'Aplicar margen comercial a los precios
                        'Modificar precios de venta en articulo y tarifas
                        frmComActPrecios.parCodArtic = RS!codArtic
                        frmComActPrecios.parNomArtic = RS!NomArtic
                        frmComActPrecios.Show vbModal
                    End If
                End If
                RS.MoveNext
            Wend
            RS.Close
            Set RS = Nothing
            
            
            If ArticuloEsEscandallo <> "" Then
                frmListado4.Opcion = 1
                frmListado4.vCadena = ArticuloEsEscandallo
                frmListado4.Show vbModal
            End If
            
            
            
        End If
        
       
        
        
        If AlbCompleto Then
            'Se habra eliminado el pedido de (scaped, sliped)
            PosicionarDataTrasEliminar
        Else
            SQL = DevuelveDesdeBDNew(conAri, "scappr", "numpedpr", "numpedpr", Text1(0).Text, "N")
            If SQL = "" Then 'Ya no existe le pedido lo hemos eliminado
                PosicionarDataTrasEliminar
            Else
                PosicionarData
                PonerCampos
                CargaGrid DataGrid1, Data2, True, False
            End If
            CargaTxtAuxServidas False, False
        
            'Eliminar las filas del pedido que se servieron completas (slippr)
'            SQL = "DELETE FROM " & NomTablaLineas & Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas) & " AND cantidad=0"
'            Conn.Execute SQL
            
            'Comprobar que si no quedan lineas en el pedido se elimine la cabecera del pedido
'            SQL = "select codalmac, codartic FROM " & NomTablaLineas & " WHERE numpedpr=" & numPed
'            Set RS = New ADODB.Recordset
'            RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'            If RS.EOF Then 'No hay lineas de pedido --> Eliminar la cabecera
'                SQL = "DELETE FROM " & NombreTabla & " WHERE numpedcl=" & numPed
'                Conn.Execute SQL
'                PosicionarDataTrasEliminar
'            Else 'Quedan lineas en el pedido --> Actualizar las lineas
'                PosicionarData
'                PonerCampos
'                CargaGrid DataGrid1, Data2, True, False
'            End If
'            RS.Close
'            Set RS = Nothing
'            CargaTxtAuxServidas False, False
        End If
       
        
'        Imprimer albaran si se solicitó
        If ImprimeAlb >= 0 Then ImprimirAlbaran NumAlb, FechaAlb, ImprimeAlb
        Screen.MousePointer = vbDefault
    Else 'Si no se ha pasado el Pedido a Albaran
        
    End If
End Sub


Private Sub InicializarServidas()
'Pone el campo servidas a 0 en la tabla lineas de pedido (sliped)
Dim SQL As String
    On Error Resume Next
    SQL = "UPDATE " & NomTablaLineas & " SET recibida= 0 "
    SQL = SQL & Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas)
    conn.Execute SQL
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub ComprobarNumLotesLineas(NumAlb As String, FechaAlb As String)
'Al pasar de PEDIDO a ALBARAN
'control de Nº Lotes si hay algun articulo en las lineas de pedido que
'requiere Nº de lote en compras pedirlos
Dim SQL As String
Dim RSLineas As ADODB.Recordset
Dim cadWhere As String

    On Error GoTo ErrLotes

    cadWhere = " WHERE numalbar=" & DBSet(NumAlb, "T") & " AND "
    cadWhere = cadWhere & " fechaalb=" & DBSet(FechaAlb, "F") & " AND "
    cadWhere = cadWhere & " slialp.codprove=" & Text1(4).Text

    'seleccionamos aquellas lineas del albaran insertado que tengan control de lote
    SQL = "SELECT slialp.* "
    SQL = SQL & " FROM (slialp INNER JOIN sartic ON slialp.codartic=sartic.codartic) "
    SQL = SQL & " LEFT OUTER JOIN scateg ON sartic.codcateg=scateg.codcateg "
    SQL = SQL & cadWhere
    SQL = SQL & " AND scateg.ctrlotes = 1"


    Set RSLineas = New ADODB.Recordset
    RSLineas.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    If Not RSLineas.EOF Then
        'Comprobar si NO Hay Nº SERIE en Compras y si no se realizo alli
        'Mostrar ahora ventana para pedir los Nº Serie de la cantidad introducida
'        Me.cmdAux(1).Tag = NumAlb
'        Me.cmdAux(0).Tag = FechaAlb
        PedirNLotes RSLineas
    
'        Set frmNLote = New frmAlmCargarNLote
'        frmNLote.parSQL = SQL
'        frmNLote.Show vbModal
'        Set frmNLote = Nothing

    End If
    
    RSLineas.Close
    Set RSLineas = Nothing
    Exit Sub

ErrLotes:
    MuestraError Err.Number, "Pedir Nº de lote.", Err.Description
End Sub




Private Sub ComprobarNSeriesLineas(NumAlb As String, FechaAlb As String)
'Al pasar de PEDIDO a ALBARAN
'control de Nº Series si hay algun articulo en las lineas de pedido que requiere Nº de serie
'y hay control de Nº de serie en compras pedirlos
Dim SQL As String
Dim RSLineas As ADODB.Recordset
Dim cadWhere As String
        
    If vParamAplic.NumSeries Then 'So control de Nº Series en COMPRAS
        cadWhere = " WHERE numalbar=" & DBSet(NumAlb, "T") & " AND "
        cadWhere = cadWhere & " fechaalb=" & DBSet(FechaAlb, "F") & " AND "
        cadWhere = cadWhere & " slialp.codprove=" & Text1(4).Text
        
        'Seleccionamos aquellas lineas de albaran que tienen Nº de Serie
        SQL = "SELECT slialp.codartic, sum(cantidad) as cantidad, slialp.numlinea "
        SQL = SQL & " FROM slialp INNER JOIN sartic on slialp.codartic=sartic.codartic "
        SQL = SQL & cadWhere & " And nseriesn = 1 "
        SQL = SQL & " GROUP BY codartic ORDER BY Codartic "
    
        Set RSLineas = New ADODB.Recordset
        RSLineas.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        If Not RSLineas.EOF Then
            'Comprobar si NO Hay Nº SERIE en Compras y si no se realizo alli
            'Mostrar ahora ventana para pedir los Nº Serie de la cantidad introducida
            Me.cmdAux(1).Tag = NumAlb
            Me.cmdAux(0).Tag = FechaAlb
            PedirNSeries RSLineas
        End If
        RSLineas.Close
        Set RSLineas = Nothing
    End If
End Sub


Private Sub PedirNSeries(ByRef RS As ADODB.Recordset)
On Error GoTo EPedirNSeries
        
        'Visualizar en pantalla el Grid, y rellenar los Nº Serie
        PedirNSeriesGnral RS, True

        Set frmNSerie = New frmRepCargarNSerie
        frmNSerie.DeVentas = False 'Se llama desde Alb. de Venta
        frmNSerie.Show vbModal
        Set frmNSerie = Nothing
        
EPedirNSeries:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub



Private Sub PedirNLotes(ByRef RS As ADODB.Recordset)
Dim cadSel As String

    On Error GoTo EPedirNLotes
        
    cadSel = "numalbar=" & DBSet(RS!Numalbar, "T") & " AND fechaalb=" & DBSet(RS!FechaAlb, "F") & " AND codprove=" & DBSet(RS!Codprove, "N")
    
    'Visualizar en pantalla el Grid, y rellenar los Nº Serie
    If Not PedirNLotesGnral(RS, True) Then
'             Visualizar en pantalla el Grid, y rellenar los Nº Serie
        MsgBox "No se han podido mostrar todos los Artículos con Nº de Lote.", vbInformation
    End If

        Set frmNLote = New frmAlmCargarNLote
        frmNLote.Desde2 = "" 'Desde proveedores
        frmNLote.parSelSQL = cadSel
        frmNLote.Show vbModal
        Set frmNLote = Nothing
        
        
     'Eliminar de la tabla temporal tmpnlotes los lotes introducidos
    DescargarDatosTMPNumLotes "tmpnlotes", cadSel
        
EPedirNLotes:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub



Private Function InsertarNSerie(numSerie As String, codArtic As String, numlinea As String) As Boolean
'Inserta o Actualiza en la tabla sserie, si al pasar Pedido -> Albaran
'existen lineas con control de Nº Serie
'Dim CadValues As String, cadValuesU As String
Dim devuelve As String
Dim Numalbar As String
Dim nSerie As CNumSerie
Dim B As Boolean

    On Error GoTo EInsertarNS

    Set nSerie = New CNumSerie
    nSerie.numSerie = numSerie
    nSerie.Articulo = codArtic
    nSerie.Proveedor = CInt(Text1(4).Text)
    nSerie.NumAlbProve = Me.cmdAux(1).Tag
    nSerie.fechacom = Me.cmdAux(0).Tag
    nSerie.NumLinAlbPr = numlinea
    'calculamos la fecha de fin garantia para el articulo comprado
    nSerie.ObtenFechaFinGarantia codArtic, Me.cmdAux(0).Tag
    
    'Comprobar si existe en la tabla sserie
    Numalbar = "numalbpr" 'Nº albaran de Compra
    devuelve = DevuelveDesdeBDNew(conAri, "sserie", "numserie", "numserie", numSerie, "T", Numalbar, "codartic", codArtic, "T")
    If devuelve <> "" Then 'EXISTE en tabla sserie
        If Numalbar = "" Then
            B = nSerie.ActualizarNumSerie(False)
        End If
    Else
        B = nSerie.InsertarNumSerie
    End If
    Set nSerie = Nothing
    
EInsertarNS:
    If Err.Number <> 0 Then B = False
    If Not B Then
        InsertarNSerie = False
    Else
        InsertarNSerie = True
    End If
End Function



Private Sub PonerDatosProveedor(Codprove As String, Optional nifProve As String)
'lee de la tabla de proveedores y pone los valores
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
            If Modo = 3 Then
                Observaciones = ""
            Else
                'Modificando y no ha puesto nada
                Observaciones = Trim(Text1(17) & Text1(18) & Text1(19) & Text1(20) & Text1(21))
            End If
            If Observaciones = "" Then vProve.PonerObservaciones Text1(17), Text1(18), Text1(19), Text1(20), Text1(21)


            Observaciones = DBLet(vProve.Observaciones)
            If Observaciones <> "" Then
                MsgBox Observaciones, vbInformation, "Observaciones del proveedor"
            End If
            
            If Modo = 3 Then
                'Insertando
                If Not EsDeVarios Then
                    PonerFocoChk Me.chkObra
                Else
                    PonerFoco Text1(5)
                End If
            End If
        End If
    Else
        LimpiarDatosProve
        PonerFoco Text1(4)
    End If
    Set vProve = Nothing

EPonerDatos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poner Datos Proveedor", Err.Description
End Sub


Private Sub PonerDatosProveVario(nifProve As String)
'Poner el los campos Text el valor del cliente
Dim vProve As CProveedor
Dim B As Boolean
   
    If nifProve = "" Then Exit Sub
   
    Set vProve = New CProveedor
    B = vProve.LeerDatosProveVario(nifProve)
    
    If B Then
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
Dim I As Byte

    'bloquear/desbloquear campos de datos segun sea proveedor de varios o no
    If Modo <> 5 Then
        Me.imgBuscar(6).visible = bol 'NIF
        Me.imgBuscar(6).Enabled = bol 'NIF
        Me.imgBuscar(1).Enabled = bol 'poblacion
        
        For I = 5 To 11 'si no es de varios no se pueden modificar los datos
            BloquearTxt Text1(I), Not bol
        Next I
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
        'Actualiza la tabla de proveedores varios con los datos que tenemos
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
Dim I As Byte
Dim cadWhere As String
Dim vFactu As CFacturaCom

    'Limpiar en el form los datos calculados de la factura
    'y volvemos a recalcular
    For I = 33 To 50
         Text3(I).Text = ""
    Next I
    
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
Dim I As Byte

    For I = 33 To 36
        If I = 34 Or I = 35 Then Text3(I).Text = QuitarCero(Text3(I).Text)
        Text3(I).Text = Format(Text3(I).Text, FormatoImporte)
    Next I
    
    'Desglose B.Imponible por IVA
    For I = 43 To 45
        If Text3(I).Text <> "" Then
             If CSng(Text3(I).Text) = 0 And Text3(I - 6).Text = "" Then
                Text3(I).Text = QuitarCero(Text3(I).Text)
                Text3(I - 3).Text = QuitarCero(Text3(I - 3).Text)
                Text3(I - 6).Text = QuitarCero(Text3(I - 6).Text)
                Text3(I + 3).Text = QuitarCero(Text3(I + 3).Text)
            Else
                Text3(I).Text = Format(Text3(I).Text, FormatoImporte)
                Text3(I - 3) = Format(Text3(I - 3).Text, FormatoDescuento)
    '            Text3(i - 6) = Format(Text3(i - 6).Text, "000")
                Text3(I + 3).Text = Format(Text3(I + 3).Text, FormatoImporte)
            End If
        End If
    Next I
    
    'Los Totales
    For I = 49 To 50
'        Text3(i).Text = QuitarCero(Text3(i).Text)
        Text3(I).Text = Format(Text3(I).Text, FormatoImporte)
    Next I
End Sub




Private Function ActualizarUltFechaCom(cadW As String) As Boolean
''Actualiza la ultima fecha de compra y el ult. precio de compra
''en el articulo, poniendo los valores del albaran de compra
'Dim SQL As String
'Dim RS As ADODB.Recordset
'
'    On Error GoTo EActualizaFecha
'
'    SQL = "select distinct numalbar,fechaalb,slialp.codartic,max(slialp.precioar) as precioar , sartic.ultfecco "
'    SQL = SQL & " from slialp INNER JOIN sartic ON slialp.codartic=sartic.codartic "
''    SQL = SQL & " where numalbar='K2500088' and fechaalb='2005-10-06' and slialp.codprove=21"
'    SQL = SQL & " WHERE " & cadW
'    SQL = SQL & " and (fechaalb>ultfecco or isnull(ultfecco))"
'    SQL = SQL & " group by numalbar,fechaalb,slialp.codartic "
'    SQL = SQL & " order by codartic "
'
'    Set RS = New ADODB.Recordset
'    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'    While Not RS.EOF
'        SQL = "UPDATE sartic SET ultfecco=" & DBSet(RS!FechaAlb, "F") & ", preciouc=" & DBSet(RS!precioar, "N")
'        SQL = SQL & " WHERE codartic=" & DBSet(RS!codArtic, "T")
'        Conn.Execute SQL
'        RS.MoveNext
'    Wend
'    RS.Close
'    Set RS = Nothing
'
'EActualizaFecha:
'    If Err.Number <> 0 Then
'        ActualizarUltFechaCom = False
'    Else
'        ActualizarUltFechaCom = True
'    End If
End Function



Private Function ObtenerRSprecios(ByRef RS As ADODB.Recordset, cadSQL As String) As Boolean
    On Error GoTo ErrRS
    Set RS = New ADODB.Recordset
    RS.Open cadSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    ObtenerRSprecios = True
    Exit Function
    
ErrRS:
    ObtenerRSprecios = False
    If Not RS Is Nothing Then Set RS = Nothing
    MuestraError Err.Number, "Cargando RS precios ultima compra.", Err.Description
End Function




Private Sub AbrirForm_CentroCoste()


    Screen.MousePointer = vbHourglass
    cmdAux(2).Tag = "2"

    Set frmB = New frmBuscaGrid
    frmB.vCampos = "Codigo|cabccost|codccost|T||20·Descripción|cabccost|nomccost|T||70·"
    frmB.vTabla = "cabccost"
    frmB.vSQL = ""
    HaDevueltoDatos = False
    '###A mano
    frmB.vDevuelve = "0|1|"
    frmB.vTitulo = "Centros de coste"
    frmB.vselElem = 0
    frmB.vConexionGrid = conConta
    
    frmB.Show vbModal
    Set frmB = Nothing
    cmdAux(2).Tag = "-1"
End Sub



' ---- [02/11/2009] [LAURA] : al pulsar F2 para abrir articulos en la solapa Documentos|Pedidos
Private Sub AbrirForm_Articulos()
    If Trim(txtAux(1).Text) = "" Then Exit Sub
    
    Set frmArt2 = New frmAlmArticulos
    frmArt2.DatosADevolverBusqueda = "::" & Trim(txtAux(1).Text)  'DevNombreSQL(Data2.Recordset!codarti1)
    frmArt2.parNumTAb = 6
    frmArt2.Show vbModal
    Set frmArt2 = Nothing
End Sub
' -----


'Nuevo. Cuando pulse MAS (y es el primer carcater abre el prismatico asociado)
Private Sub PulsarTeclaMas(InsertandoCabecera As Boolean, index As Integer)

    If InsertandoCabecera Then
        If imgBuscar(index).visible Then imgBuscar_Click index
        
    Else
        'Lineas
        If index = 8 Then index = 2
        cmdAux_Click index
        
        
    End If
        
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



'Private Function InsertarMovStock3(NumAlb As String, FechaAlb As String) As Boolean
'Dim vCStock As CStock
'Dim b As Boolean
'Dim RS As ADODB.Recordset
'Dim SQL As String
'Dim cart As CArticulo
'
'    On Error Resume Next
'
'    InsertarMovStock2 = False
'
'    Set vCStock = New CStock
'    b = True
'
'    SQL = Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas)
'    SQL = "select * from " & NomTablaLineas & SQL
'    Set RS = New ADODB.Recordset
'    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'    vCStock.Fechamov = FechaAlb
'
'    'para cada linea del Pedido Insertar en smoval y Actualizar Stock en salmac
'    While (Not RS.EOF) And b
'        If InicializarCStockAlbar(vCStock, "E", CStr(RS!numlinea), RS) Then
'            vCStock.Documento = NumAlb
'            If vCStock.Cantidad <> 0 Then
'                '==== Laura 22/09/2006
'                '-- antes de actualizar el stock calculamos el precio medio ponderado del articulo
'                Set cart = New CArticulo
'                If cart.LeerDatos(vCStock.codArtic) Then
'                    '17 Junio 2009
'                    'Si la cantidad es negativa no actualiza ni precio medio ponderado NI fecha ult compra
'                    If vCStock.Cantidad >= 0 Then
'
'                        'Laura 19/12/2006: Calcular precio_med_pond con el precio con los descuentos,e.d. importe/cantidad
'                        'If Not cArt.ActualizarPrecioMedPond(CCur(vCStock.Cantidad), CCur(RS!precioar)) Then b = False
'                        If Not cart.ActualizarPrecioMedPond(CCur(vCStock.Cantidad), Round2(CCur(vCStock.Importe) / CCur(vCStock.Cantidad), 4)) Then b = False
'
'                        '--actualizar fecha y precio ultima compra del articulo
'                        'Laura 19/12/2006: actualizar precio_ult_compra con el precio con los descuentos,e.d. importe/cantidad
'                        'If Not cArt.ActualizarUltFechaCompra(vCStock.Fechamov, CStr(RS!precioar)) Then b = False
'                        If Not cart.ActualizarUltFechaCompra(vCStock.Fechamov, Round2(CCur(vCStock.Importe) / CCur(vCStock.Cantidad), 4)) Then b = False
'
'                    End If 'De cantidad >=0
'                End If
'                Set cart = Nothing
'                '====
'
'
'                'en actualizar stock comprobamos si el articulo tiene control de stock
'                b = vCStock.ActualizarStock
'            End If
'        Else
'            b = False
'        End If
'        RS.MoveNext
'    Wend
'    Set vCStock = Nothing
'    RS.Close
'    Set RS = Nothing
'
'    InsertarMovStock2 = b
'
'End Function





Private Function ActualizarDtos() As Boolean
Dim SQL As String
Dim TipoDto As Byte
Dim Dto1 As Currency
Dim Dto2 As Currency

On Error GoTo eActualizarDtos

         TipoDto = DevuelveDesdeBDNew(conAri, "sprove", "tipodtos", "codprove", Text1(4).Text, "N")
         SQL = "SELECT numlinea,cantidad, precioar FROM " & NomTablaLineas
         SQL = SQL & " " & Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas) & " ORDER BY numlinea"
         Set miRsAux = New ADODB.Recordset
         miRsAux.Open SQL, conn, adOpenKeyset, adLockOptimistic, adCmdText
         
         While Not miRsAux.EOF
         
            SQL = CalcularImporteSng(CStr(miRsAux!cantidad), CStr(miRsAux!precioar), RecuperaValor(CadenaDesdeOtroForm, 1), RecuperaValor(CadenaDesdeOtroForm, 2), TipoDto)
            'Ya tengo el importe
            SQL = "UPDATE " & NomTablaLineas & " SET importel = " & DBSet(SQL, "N")
            SQL = SQL & ", dtoline1=" & DBSet(RecuperaValor(CadenaDesdeOtroForm, 1), "N")
            SQL = SQL & ", dtoline2=" & DBSet(RecuperaValor(CadenaDesdeOtroForm, 2), "N")
            SQL = SQL & " " & Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas)
            SQL = SQL & " AND numlinea = " & miRsAux!numlinea
            conn.Execute SQL
            
            miRsAux.MoveNext
            
        Wend
        miRsAux.Close
        ActualizarDtos = True 'Ok
eActualizarDtos:
    If Err.Number <> 0 Then MuestraError Err.Number
    Set miRsAux = Nothing
End Function



Private Sub PonerClieObraActuacion(Cual As Integer, DesdePonerCampos As Boolean)
Dim D As String
Dim Msg As String

    D = txtAux(Cual).Text
    If D = "" Then
        Me.txtDesc(Cual).Text = ""
        Exit Sub
    End If
    D = ""
    Msg = ""
    Select Case Cual
    Case 9
        'Codclien y direc. Es numericio
        If Not IsNumeric(txtAux(Cual).Text) Then
            Msg = "Campo numerico"
        Else
            
            D = DevuelveDesdeBD(conAri, "nomclien", "sclien", "codclien", txtAux(Cual).Text)
            If D = "" Then Msg = "No existe cliente"
            
        End If
        Me.txtDesc(Cual).Text = D
        
        If Msg <> "" Then
            
            
            If Not DesdePonerCampos Then
                MsgBox Msg, vbExclamation
                PonerFoco txtAux(Cual)
            End If
            txtAux(Cual).Text = ""
        
            
        End If
    Case 10
        If txtAux(9).Text = "" Then
            If Not DesdePonerCampos Then
                MsgBox "Ponga el cliente", vbExclamation
                txtAux(Cual).Text = ""
                PonerFoco txtAux(9)
                Exit Sub
            End If
        End If
        D = ""
        If Not IsNumeric(txtAux(Cual).Text) Then
            Msg = "Campo numerico"
        Else
            D = "codclien = " & Val(txtAux(9).Text) & " and coddirec "
            D = DevuelveDesdeBD(conAri, "nomdirec", "sdirec", D, txtAux(10).Text)
            If D = "" Then Msg = "No existe la obra para el cliente"
                
        End If
        
        If Msg <> "" Then
            MsgBox Msg, vbExclamation
            txtAux(Cual).Text = ""
            PonerFoco txtAux(Cual)
        End If
        
        Me.txtDesc(Cual).Text = D
        
    Case 11
        'Actuacion
        If txtAux(9).Text = "" Or txtAux(10).Text = "" Then
            If Not DesdePonerCampos Then
                MsgBox "Ponga el cliente/obra", vbExclamation
                txtAux(Cual).Text = ""
                If txtAux(9).Text = "" Then
                    PonerFoco txtAux(9)
                Else
                    PonerFoco txtAux(10)
                End If
                Exit Sub
            End If
            D = ""
        End If
        
        
        D = "codclien =" & txtAux(9).Text & " AND coddirec= " & txtAux(10).Text & " AND actuacion "
                
        D = DevuelveDesdeBDNew(conAri, "sactuaobra", "concat(fechaini,' ',if(observa is null,'',observa))", D, txtAux(11).Text, "T")
        If D = "" Then
            If Not DesdePonerCampos Then
                MsgBox "No existe la obra para el cliente", vbExclamation
                txtAux(Cual).Text = ""
                PonerFoco txtAux(Cual)
            End If
        End If
        Me.txtDesc(Cual).Text = D
        
        
    End Select
End Sub


Private Sub CamposObractua2()
Dim B As Boolean
    
    B = False
    If Not Me.Data2.Recordset Is Nothing Then
        If Not Data2.Recordset.EOF Then B = True
    End If
    If B Then
        Me.txtAux(9).Text = DBLet(Data2.Recordset!codClien, "T")
        Me.txtAux(10).Text = DBLet(Data2.Recordset!CodDirec, "T")
        Me.txtAux(11).Text = DBLet(Data2.Recordset!actuacion, "T")
    Else
        Me.txtAux(9).Text = ""
        Me.txtAux(10).Text = ""
        Me.txtAux(11).Text = ""
    End If
    PonerClieObraActuacion 9, True
    PonerClieObraActuacion 10, True
    PonerClieObraActuacion 11, True
    
    
    If vParamAplic.NumeroInstalacion = 4 Then
        If B Then
            Me.txtAux(12).Text = DBLet(Me.Data2.Recordset.Fields(15), "T")
            Me.txtAux(13).Text = DBLet(Me.Data2.Recordset.Fields(16), "T")
            Me.txtAux(14).Text = DBLet(Me.Data2.Recordset.Fields(17), "F")
            Me.txtAux(15).Text = DBLet(Me.Data2.Recordset.Fields(18), "n")
            If txtAux(15).Text <> "" Then txtAux(15).Text = Format(txtAux(15).Text, "00000")
            PonerDatosAlbaranFacturaEuler
        Else
            txtAux(12).Text = "": txtAux(13).Text = "": txtAux(14).Text = "": txtAux(15).Text = ""
            txtDesc(0).Text = ""
        End If
    End If
    
    
    
    If vEmpresa.TieneAnalitica Then Me.txtAux2(8).Text = PonerNombreCCoste(Me.txtAux(8))
    
End Sub



Private Sub SimularOtroProveedor()
    If Modo <> 2 Then Exit Sub
    If Data1.Recordset.EOF Then Exit Sub
    If Data2.Recordset.EOF Then Exit Sub
    
    
    
    CadenaDesdeOtroForm = ""
    frmListado3.OtrosDatos = Text1(1).Text & "|" & Text1(13).Text & "|" & Text1(14).Text & "|" & Text1(0).Text & "|"
    frmListado3.Opcion = 57
    frmListado3.Show vbModal
    
    If CadenaDesdeOtroForm <> "" Then
        CadenaConsulta = "select * from " & NombreTabla & " WHERE numpedpr IN (" & CadenaDesdeOtroForm & ") ORDER BY numpedpr"
        PonerCadenaBusqueda
        CadenaDesdeOtroForm = ""
    End If
End Sub




Private Sub Euler_O_Sail()
Dim Euler As Boolean
    
    'En euler saldran codtipon numalbar fechaalb   y para sail saldran codcapit y numlotes
    Euler = vParamAplic.NumeroInstalacion = 4
        
    'SAIL
    Me.txtAux(10).visible = Not Euler
    Me.txtDesc(10).visible = Not Euler
    Me.txtAux(11).visible = Not Euler
    Me.txtAux(11).visible = Not Euler
    Me.txtDesc(11).visible = Not Euler
    Label1(34).visible = Not Euler
    Label1(43).visible = Not Euler

    Me.imgBuscar2(10).visible = Not Euler
    Me.imgBuscar2(11).visible = Not Euler
    
    
    'EULER
    Line2.visible = Euler
    Line3.visible = Euler
    Me.txtAux(12).visible = Euler
    Me.txtAux(13).visible = Euler
    Me.txtAux(14).visible = Euler
    Me.txtAux(15).visible = Euler
    Me.txtDesc(0).visible = Euler
    Label1(51).visible = Euler
    Label1(52).visible = Euler
    Label1(53).visible = Euler
    
End Sub

Private Sub PonerDatosAlbaranFacturaEuler()
Dim cad As String

    txtDesc(0).Text = ""
    If txtAux(12).Text <> "" And Me.txtAux(13).Text <> "" And Me.txtAux(14).Text <> "" Then
        'Buscamos en albaranes
        cad = "codtipom=" & DBSet(txtAux(12).Text, "T") & " AND fechaalb =" & DBSet(txtAux(14).Text, "F")
        cad = cad & " AND numalbar"
        cad = DevuelveDesdeBD(conAri, "concat(codclien,' ',nomclien)", "scaalb", cad, txtAux(13).Text)
        
        If cad = "" Then
            cad = "scafac.codtipom=scafac1.codtipom and scafac.numfactu=scafac1.numfactu and "
            cad = cad & " scafac.fecfactu=scafac1.fecfactu AND scafac1.codtipoa=" & DBSet(txtAux(12).Text, "T")
            cad = cad & " AND fechaalb =" & DBSet(txtAux(14).Text, "F") & " AND numalbar"
            cad = DevuelveDesdeBD(conAri, "concat(scafac.codtipom,right(concat('00000',scafac.numfactu),10),' de ',DATE_FORMAT(scafac.fecfactu, '%d/%m/%Y'),'|',codclien,' ',nomclien,'|')", "scafac,scafac1", cad, txtAux(13).Text)
            
            If cad = "" Then
                cad = "NO EXISTE"
            Else
                cad = "ALBARAN FACTURADO.     Fra:" & RecuperaValor(cad, 1) & vbCrLf & RecuperaValor(cad, 2)
            End If
        Else
            cad = "ALBARAN: " & vbCrLf & cad
        End If
        txtDesc(0).Text = cad
    End If
        
End Sub



Private Sub LanzarBuscarAlbaranEuler(EsPedido As Boolean)

    If txtAux(9).Text = "" Then Exit Sub
    If FormularioListAlbAbierto Then Exit Sub
    FormularioListAlbAbierto = True
    frmListado5.OpcionListado = 14
    frmListado5.OtrosDatos = IIf(EsPedido, "1", "0") & Val(txtAux(9).Text)
    frmListado5.Show vbModal
    FormularioListAlbAbierto = False
    'CadenaDesdeOtroForm ="" esta puesto en el form
    If CadenaDesdeOtroForm <> "" Then
        EsPedido = Mid(CadenaDesdeOtroForm, 1, 1) = "1"
        CadenaDesdeOtroForm = Mid(CadenaDesdeOtroForm, 2)
        If Not EsPedido Then
            txtAux(12).Text = RecuperaValor(CadenaDesdeOtroForm, 1)
            txtAux(13).Text = RecuperaValor(CadenaDesdeOtroForm, 2)
            txtAux(14).Text = RecuperaValor(CadenaDesdeOtroForm, 3)
             txtAux_LostFocus 12
        Else
            txtAux(15).Text = RecuperaValor(CadenaDesdeOtroForm, 2)
            txtAux_LostFocus 15
        End If
    End If
End Sub
