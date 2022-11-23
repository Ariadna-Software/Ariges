VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFacHcoFacturas2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Histórico de Facturas Clientes"
   ClientHeight    =   7965
   ClientLeft      =   45
   ClientTop       =   4335
   ClientWidth     =   14580
   Icon            =   "frmFacHcoFacturas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmFacHcoFacturas.frx":000C
   ScaleHeight     =   7965
   ScaleWidth      =   14580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAux2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   9
      Left            =   3480
      Locked          =   -1  'True
      MaxLength       =   60
      TabIndex        =   149
      Top             =   7515
      Visible         =   0   'False
      Width           =   6045
   End
   Begin VB.Frame Frame2 
      Height          =   710
      Left            =   120
      TabIndex        =   130
      Top             =   400
      Width           =   14175
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   2625
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   315
         Width           =   2505
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   5
         Left            =   8850
         MaxLength       =   40
         TabIndex        =   6
         Tag             =   "Nombre Cliente|T|N|||scafac|nomclien||N|"
         Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwwwwwwwwww aq"
         Top             =   240
         Width           =   4350
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   4
         Left            =   7920
         MaxLength       =   6
         TabIndex        =   5
         Tag             =   "Cod. Cliente|N|N|0|999999|scafac|codclien|000000|N|"
         Text            =   "Text1"
         Top             =   240
         Width           =   780
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   1
         Left            =   2640
         TabIndex        =   2
         Tag             =   "Tipo Factura|T|N|||scafac|codtipom||S|"
         Text            =   "Text3"
         Top             =   315
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   2
         Left            =   1350
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Fecha Factura|F|N|||scafac|fecfactu|dd/mm/yyyy|S|"
         Top             =   315
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
         Tag             =   "Nº Factura|N|N|||scafac|numfactu|0000000|S|"
         Text            =   "Text1 7"
         Top             =   315
         Width           =   980
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Contabilizado"
         Height          =   375
         Left            =   5320
         TabIndex        =   4
         Tag             =   "Contabilizado|N|N|0|1|scafac|intconta||N|"
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblSerie 
         Caption         =   "Label2"
         Height          =   255
         Left            =   3840
         TabIndex        =   255
         Top             =   360
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente"
         Height          =   255
         Index           =   0
         Left            =   7200
         TabIndex        =   134
         Top             =   300
         Width           =   495
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   7695
         ToolTipText     =   "Buscar cliente"
         Top             =   270
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Fact."
         Height          =   255
         Index           =   29
         Left            =   1350
         TabIndex        =   133
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Factura"
         Height          =   255
         Index           =   28
         Left            =   240
         TabIndex        =   132
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Fact."
         Height          =   255
         Index           =   27
         Left            =   2640
         TabIndex        =   131
         Top             =   120
         Width           =   795
      End
   End
   Begin MSAdodcLib.Adodc Data2 
      Height          =   330
      Left            =   6000
      Top             =   1080
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
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   12480
      Top             =   1080
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
      Height          =   5880
      Left            =   120
      TabIndex        =   42
      Tag             =   "Fecha Oferta|F|N|||scapre|fecentre|dd/mm/yyyy|N|"
      Top             =   1080
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   10372
      _Version        =   393216
      Style           =   1
      Tabs            =   4
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
      TabPicture(0)   =   "frmFacHcoFacturas.frx":0A0E
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameCliente"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "FrameFactura"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Albaranes"
      TabPicture(1)   =   "frmFacHcoFacturas.frx":0A2A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "imgBuscar(7)"
      Tab(1).Control(1)=   "Label1(9)"
      Tab(1).Control(2)=   "Label1(23)"
      Tab(1).Control(3)=   "Label1(24)"
      Tab(1).Control(4)=   "Label1(21)"
      Tab(1).Control(5)=   "Label1(2)"
      Tab(1).Control(6)=   "Label1(6)"
      Tab(1).Control(7)=   "Label1(18)"
      Tab(1).Control(8)=   "Label1(22)"
      Tab(1).Control(9)=   "Label1(40)"
      Tab(1).Control(10)=   "imgBuscar(8)"
      Tab(1).Control(11)=   "imgBuscar(9)"
      Tab(1).Control(12)=   "imgBuscar(6)"
      Tab(1).Control(13)=   "Label1(47)"
      Tab(1).Control(14)=   "imgBuscar(10)"
      Tab(1).Control(15)=   "Label1(48)"
      Tab(1).Control(16)=   "Label1(49)"
      Tab(1).Control(17)=   "Label1(57)"
      Tab(1).Control(18)=   "Label1(61)"
      Tab(1).Control(19)=   "Label1(62)"
      Tab(1).Control(20)=   "FrameTelefonia"
      Tab(1).Control(21)=   "Text3(15)"
      Tab(1).Control(22)=   "chkEnvio"
      Tab(1).Control(23)=   "FrameObserva"
      Tab(1).Control(24)=   "DataGrid2"
      Tab(1).Control(25)=   "DataGrid1"
      Tab(1).Control(26)=   "txtAux(8)"
      Tab(1).Control(27)=   "txtAux(7)"
      Tab(1).Control(28)=   "txtAux(6)"
      Tab(1).Control(29)=   "txtAux(4)"
      Tab(1).Control(30)=   "Text3(2)"
      Tab(1).Control(31)=   "Text2(2)"
      Tab(1).Control(32)=   "Text3(1)"
      Tab(1).Control(33)=   "Text2(1)"
      Tab(1).Control(34)=   "Text3(0)"
      Tab(1).Control(35)=   "Text2(0)"
      Tab(1).Control(36)=   "Text3(8)"
      Tab(1).Control(37)=   "Text3(6)"
      Tab(1).Control(38)=   "Text3(7)"
      Tab(1).Control(39)=   "Text3(5)"
      Tab(1).Control(40)=   "Text3(4)"
      Tab(1).Control(41)=   "Text3(3)"
      Tab(1).Control(42)=   "Text2(3)"
      Tab(1).Control(43)=   "cmdObserva3"
      Tab(1).Control(44)=   "txtAux(0)"
      Tab(1).Control(45)=   "txtAux(1)"
      Tab(1).Control(46)=   "txtAux(2)"
      Tab(1).Control(47)=   "txtAux(3)"
      Tab(1).Control(48)=   "txtAux(5)"
      Tab(1).Control(49)=   "txtAux3(0)"
      Tab(1).Control(50)=   "txtAux3(1)"
      Tab(1).Control(51)=   "txtAux3(2)"
      Tab(1).Control(52)=   "Text3(14)"
      Tab(1).Control(53)=   "txtAux(9)"
      Tab(1).Control(54)=   "txtAux(10)"
      Tab(1).Control(55)=   "cmdaux"
      Tab(1).Control(56)=   "txtAux(11)"
      Tab(1).Control(57)=   "Text3(17)"
      Tab(1).Control(58)=   "Text2(18)"
      Tab(1).Control(59)=   "Text3(18)"
      Tab(1).Control(60)=   "chkPedxCli"
      Tab(1).Control(61)=   "FrameEuler"
      Tab(1).Control(62)=   "cmdInsertaAlbaranCostes"
      Tab(1).Control(63)=   "FrameCampos"
      Tab(1).Control(64)=   "Text3(28)"
      Tab(1).Control(65)=   "Text3(29)"
      Tab(1).ControlCount=   66
      TabCaption(2)   =   "Costes"
      TabPicture(2)   =   "frmFacHcoFacturas.frx":0A46
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lwCostes"
      Tab(2).Control(1)=   "cmdLineasCostes(2)"
      Tab(2).Control(2)=   "cmdLineasCostes(0)"
      Tab(2).Control(3)=   "cmdLineasCostes(1)"
      Tab(2).Control(4)=   "txtCostes(0)"
      Tab(2).Control(5)=   "txtCostes(1)"
      Tab(2).Control(6)=   "txtCostes(2)"
      Tab(2).Control(7)=   "txtCostes(3)"
      Tab(2).Control(8)=   "txtCostes(4)"
      Tab(2).Control(9)=   "txtCostes(5)"
      Tab(2).Control(10)=   "txtCostes(6)"
      Tab(2).Control(11)=   "txtCostes(7)"
      Tab(2).ControlCount=   12
      TabCaption(3)   =   "Impresion lineas"
      TabPicture(3)   =   "frmFacHcoFacturas.frx":0A62
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lwEulerLineas"
      Tab(3).Control(1)=   "cmdLineasImpresion(0)"
      Tab(3).Control(2)=   "cmdLineasImpresion(1)"
      Tab(3).Control(3)=   "cmdLineasImpresion(2)"
      Tab(3).Control(4)=   "cmdLineasImpresion(3)"
      Tab(3).ControlCount=   5
      Begin VB.TextBox Text3 
         Height          =   315
         Index           =   29
         Left            =   -64320
         TabIndex        =   259
         Tag             =   "Rped|T|S|||scafac1|refpedidol|||"
         Text            =   "Text1 7"
         Top             =   2640
         Width           =   2685
      End
      Begin VB.TextBox Text3 
         Height          =   315
         Index           =   28
         Left            =   -68880
         TabIndex        =   258
         Tag             =   "rf|T|S|||scafac1|refcontrato||N|"
         Text            =   "Text1 7"
         Top             =   2640
         Width           =   2085
      End
      Begin VB.Frame FrameCampos 
         Caption         =   "Campos / Fitosanitarios"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   2520
         Left            =   -75000
         TabIndex        =   158
         Top             =   3300
         Visible         =   0   'False
         Width           =   13815
         Begin VB.Frame FrameCamposMani 
            Caption         =   "Frame3"
            Height          =   2055
            Left            =   120
            TabIndex        =   191
            Top             =   360
            Width           =   5055
            Begin VB.CommandButton cmdCarnetManipulador 
               Caption         =   "+"
               Height          =   315
               Left            =   3240
               TabIndex        =   257
               Top             =   0
               Width           =   255
            End
            Begin VB.TextBox Text2 
               BackColor       =   &H80000018&
               Height          =   315
               Index           =   4
               Left            =   1080
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   198
               Text            =   "Bajo tiene el texts de scafac1"
               Top             =   1200
               Width           =   2805
            End
            Begin VB.TextBox Text3 
               Height          =   315
               Index           =   21
               Left            =   1080
               MaxLength       =   20
               TabIndex        =   196
               Tag             =   "ManipuladorFecCaducidad|F|S|||scafac1|ManipuladorFecCaducidad||N|"
               Text            =   "Text1 Text1 Text1 Te"
               Top             =   810
               Width           =   1125
            End
            Begin VB.TextBox Text3 
               Height          =   315
               Index           =   20
               Left            =   1080
               MaxLength       =   50
               TabIndex        =   194
               Tag             =   "ManipuladorNombre|T|S|||scafac1|ManipuladorNombre||N|"
               Text            =   "Text1 Text1 Text1 Te"
               Top             =   390
               Width           =   3885
            End
            Begin VB.TextBox Text3 
               Height          =   315
               Index           =   19
               Left            =   1080
               MaxLength       =   20
               TabIndex        =   192
               Tag             =   "ManipuladorNumCarnet|T|S|||scaalb|scafac1||N|"
               Text            =   "Text1 Text1 Text1 Te"
               Top             =   0
               Width           =   1725
            End
            Begin VB.TextBox Text3 
               Height          =   315
               Index           =   22
               Left            =   1200
               MaxLength       =   20
               TabIndex        =   200
               Tag             =   "TipoCarnet|N|S|||scafac1|TipoCarnet||N|"
               Text            =   "Text1 Text1 Text1 Te"
               Top             =   1200
               Width           =   285
            End
            Begin VB.Label Label1 
               Caption         =   "Tipo"
               Height          =   195
               Index           =   56
               Left            =   0
               TabIndex        =   199
               Top             =   1200
               Width           =   690
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "F.Caducidad"
               Height          =   195
               Index           =   55
               Left            =   0
               TabIndex        =   197
               Top             =   840
               Width           =   900
            End
            Begin VB.Label Label1 
               Caption         =   "Nombre"
               Height          =   195
               Index           =   54
               Left            =   0
               TabIndex        =   195
               Top             =   480
               Width           =   690
            End
            Begin VB.Label Label1 
               Caption         =   "Nº Carnet"
               Height          =   195
               Index           =   53
               Left            =   0
               TabIndex        =   193
               Top             =   0
               Width           =   690
            End
         End
         Begin VB.CommandButton cmdMtoCampos 
            Height          =   375
            Index           =   1
            Left            =   5280
            Picture         =   "frmFacHcoFacturas.frx":0A7E
            Style           =   1  'Graphical
            TabIndex        =   160
            ToolTipText     =   "Eliminar campo"
            Top             =   840
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.CommandButton cmdMtoCampos 
            Height          =   375
            Index           =   0
            Left            =   5280
            Picture         =   "frmFacHcoFacturas.frx":1480
            Style           =   1  'Graphical
            TabIndex        =   159
            ToolTipText     =   "Añadir campo"
            Top             =   360
            Visible         =   0   'False
            Width           =   375
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   1935
            Left            =   5760
            TabIndex        =   161
            Top             =   360
            Width           =   7695
            _ExtentX        =   13573
            _ExtentY        =   3413
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   5
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Campo"
               Object.Width           =   1323
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Partida"
               Object.Width           =   4701
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Variedad"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Socio"
               Object.Width           =   1834
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Nombre"
               Object.Width           =   5186
            EndProperty
         End
      End
      Begin VB.CommandButton cmdInsertaAlbaranCostes 
         Height          =   375
         Left            =   -71040
         Picture         =   "frmFacHcoFacturas.frx":7CD2
         Style           =   1  'Graphical
         TabIndex        =   256
         ToolTipText     =   "Añadir albaran de costes"
         Top             =   1080
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdLineasImpresion 
         Height          =   375
         Index           =   3
         Left            =   -73080
         Style           =   1  'Graphical
         TabIndex        =   235
         ToolTipText     =   "Imprimir factura lineas"
         Top             =   480
         Width           =   375
      End
      Begin VB.CommandButton cmdLineasImpresion 
         Height          =   375
         Index           =   2
         Left            =   -73920
         Style           =   1  'Graphical
         TabIndex        =   234
         ToolTipText     =   "Eliminar"
         Top             =   480
         Width           =   375
      End
      Begin VB.CommandButton cmdLineasImpresion 
         Height          =   375
         Index           =   1
         Left            =   -74400
         Style           =   1  'Graphical
         TabIndex        =   233
         ToolTipText     =   "Modificar"
         Top             =   480
         Width           =   375
      End
      Begin VB.CommandButton cmdLineasImpresion 
         Height          =   375
         Index           =   0
         Left            =   -74880
         Style           =   1  'Graphical
         TabIndex        =   232
         ToolTipText     =   "Insertar "
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox txtCostes 
         Height          =   285
         Index           =   7
         Left            =   -69360
         TabIndex        =   218
         Text            =   "Text4"
         Top             =   4320
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtCostes 
         Height          =   285
         Index           =   6
         Left            =   -69960
         TabIndex        =   217
         Text            =   "Text4"
         Top             =   4320
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtCostes 
         Height          =   285
         Index           =   5
         Left            =   -70680
         TabIndex        =   216
         Text            =   "Text4"
         Top             =   4320
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtCostes 
         Height          =   285
         Index           =   4
         Left            =   -71520
         TabIndex        =   215
         Text            =   "Text4"
         Top             =   4320
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtCostes 
         Height          =   285
         Index           =   3
         Left            =   -72240
         TabIndex        =   214
         Text            =   "Text4"
         Top             =   4320
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtCostes 
         Height          =   285
         Index           =   2
         Left            =   -72960
         TabIndex        =   213
         Text            =   "Text4"
         Top             =   4320
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtCostes 
         Height          =   285
         Index           =   1
         Left            =   -73800
         TabIndex        =   212
         Text            =   "Text4"
         Top             =   4320
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtCostes 
         Height          =   285
         Index           =   0
         Left            =   -74760
         TabIndex        =   211
         Text            =   "Text4"
         Top             =   4320
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdLineasCostes 
         Height          =   375
         Index           =   1
         Left            =   -74280
         Style           =   1  'Graphical
         TabIndex        =   210
         ToolTipText     =   "Modificar articulo"
         Top             =   480
         Width           =   375
      End
      Begin VB.CommandButton cmdLineasCostes 
         Height          =   375
         Index           =   0
         Left            =   -74760
         Style           =   1  'Graphical
         TabIndex        =   209
         ToolTipText     =   "Insertar articulo"
         Top             =   480
         Width           =   375
      End
      Begin VB.CommandButton cmdLineasCostes 
         Height          =   375
         Index           =   2
         Left            =   -73800
         Style           =   1  'Graphical
         TabIndex        =   208
         ToolTipText     =   "eliminar articulo"
         Top             =   480
         Width           =   375
      End
      Begin VB.Frame FrameEuler 
         Height          =   2655
         Left            =   -74880
         TabIndex        =   170
         Top             =   3300
         Visible         =   0   'False
         Width           =   13935
         Begin VB.Frame FrameTaxco 
            Height          =   2385
            Left            =   360
            TabIndex        =   236
            Top             =   120
            Visible         =   0   'False
            Width           =   13065
            Begin VB.TextBox txtTaxco 
               Height          =   315
               Index           =   5
               Left            =   1080
               TabIndex        =   243
               Text            =   "Text5"
               Top             =   1920
               Width           =   1815
            End
            Begin VB.TextBox txtTaxco 
               Height          =   315
               Index           =   6
               Left            =   5400
               TabIndex        =   244
               Text            =   "Text5"
               Top             =   1920
               Width           =   1815
            End
            Begin VB.TextBox txtTaxco 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   7
               Left            =   4440
               TabIndex        =   238
               Text            =   "Text5"
               Top             =   240
               Width           =   1335
            End
            Begin VB.TextBox txtTaxco 
               Height          =   315
               Index           =   1
               Left            =   1080
               TabIndex        =   239
               Text            =   "Text5"
               Top             =   840
               Width           =   3135
            End
            Begin VB.TextBox txtTaxco 
               Height          =   315
               Index           =   2
               Left            =   5400
               TabIndex        =   240
               Text            =   "Text5"
               Top             =   840
               Width           =   3375
            End
            Begin VB.TextBox txtTaxco 
               Height          =   315
               Index           =   3
               Left            =   1080
               TabIndex        =   241
               Text            =   "Text5"
               Top             =   1380
               Width           =   3135
            End
            Begin VB.TextBox txtTaxco 
               Height          =   315
               Index           =   4
               Left            =   5400
               TabIndex        =   242
               Text            =   "Text5"
               Top             =   1380
               Width           =   1935
            End
            Begin VB.TextBox txtTaxco 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   0
               Left            =   1080
               TabIndex        =   237
               Text            =   "Text1"
               Top             =   240
               Width           =   1695
            End
            Begin VB.Line Line2 
               X1              =   120
               X2              =   8760
               Y1              =   690
               Y2              =   690
            End
            Begin VB.Label Label3 
               Caption         =   "Licencia"
               Height          =   195
               Index           =   22
               Left            =   120
               TabIndex        =   252
               Top             =   1980
               Width           =   945
            End
            Begin VB.Label Label3 
               Caption         =   "Taximetro"
               Height          =   195
               Index           =   24
               Left            =   4440
               TabIndex        =   251
               Top             =   1980
               Width           =   690
            End
            Begin VB.Label Label3 
               Caption         =   "Kms"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   28
               Left            =   3840
               TabIndex        =   250
               Top             =   240
               Width           =   945
            End
            Begin VB.Label Label3 
               Caption         =   "Matrícula"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   13
               Left            =   120
               TabIndex        =   249
               Top             =   240
               Width           =   945
            End
            Begin VB.Label Label3 
               Caption         =   "Bastidor"
               Height          =   195
               Index           =   15
               Left            =   120
               TabIndex        =   248
               Top             =   900
               Width           =   945
            End
            Begin VB.Label Label3 
               Caption         =   "Motor"
               Height          =   195
               Index           =   17
               Left            =   4440
               TabIndex        =   247
               Top             =   900
               Width           =   405
            End
            Begin VB.Label Label3 
               Caption         =   "Marca/Modelo"
               Height          =   195
               Index           =   18
               Left            =   120
               TabIndex        =   246
               Top             =   1440
               Width           =   945
            End
            Begin VB.Label Label3 
               Caption         =   "Neumáticos"
               Height          =   195
               Index           =   20
               Left            =   4440
               TabIndex        =   245
               Top             =   1440
               Width           =   840
            End
         End
         Begin VB.Frame FrameReparEuler 
            Height          =   2415
            Left            =   480
            TabIndex        =   203
            Top             =   120
            Width           =   13440
            Begin VB.TextBox TextmatriculaTaxco 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Left            =   3600
               TabIndex        =   254
               Text            =   "Text4"
               Top             =   1920
               Visible         =   0   'False
               Width           =   1935
            End
            Begin VB.TextBox txtEuler 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1515
               Index           =   8
               Left            =   3600
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   206
               Text            =   "frmFacHcoFacturas.frx":825C
               Top             =   240
               Width           =   7575
            End
            Begin VB.CommandButton cmdReparEuler 
               Height          =   375
               Index           =   0
               Left            =   3000
               Picture         =   "frmFacHcoFacturas.frx":8262
               Style           =   1  'Graphical
               TabIndex        =   205
               ToolTipText     =   "Ver datos reparacion"
               Top             =   240
               Width           =   375
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Matrícula"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   60
               Left            =   840
               TabIndex        =   253
               Top             =   1920
               Visible         =   0   'False
               Width           =   2160
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Datos reparación"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   58
               Left            =   720
               TabIndex        =   204
               Top             =   270
               Width           =   2160
            End
         End
         Begin VB.TextBox txtEuler 
            Height          =   2235
            Index           =   7
            Left            =   9240
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   188
            Text            =   "frmFacHcoFacturas.frx":EAB4
            Top             =   240
            Width           =   4575
         End
         Begin VB.Frame FrameALE 
            Height          =   2415
            Left            =   8000
            TabIndex        =   171
            Top             =   120
            Visible         =   0   'False
            Width           =   8175
            Begin VB.TextBox txtEuler 
               Height          =   1995
               Index           =   6
               Left            =   1080
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   189
               Text            =   "frmFacHcoFacturas.frx":EABA
               Top             =   240
               Width           =   6975
            End
            Begin VB.Label Label3 
               Caption         =   "Notas operario"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   795
               Index           =   1
               Left            =   120
               TabIndex        =   190
               Top             =   240
               Width           =   1095
            End
         End
         Begin VB.TextBox txtEuler 
            Height          =   315
            Index           =   4
            Left            =   1200
            TabIndex        =   186
            Text            =   "Text1"
            Top             =   1920
            Width           =   2175
         End
         Begin VB.TextBox txtEuler 
            Height          =   315
            Index           =   5
            Left            =   3480
            TabIndex        =   185
            Text            =   "Text1"
            Top             =   1920
            Width           =   4815
         End
         Begin VB.TextBox txtEuler 
            Height          =   315
            Index           =   2
            Left            =   1200
            TabIndex        =   181
            Text            =   "Text1"
            Top             =   1320
            Width           =   2175
         End
         Begin VB.TextBox txtEuler 
            Height          =   315
            Index           =   3
            Left            =   3480
            TabIndex        =   180
            Text            =   "Text1"
            Top             =   1320
            Width           =   4815
         End
         Begin VB.TextBox txtEuler 
            Height          =   315
            Index           =   1
            Left            =   5160
            TabIndex        =   176
            Text            =   "Text1"
            Top             =   600
            Width           =   1695
         End
         Begin VB.TextBox txtEuler 
            Height          =   315
            Index           =   0
            Left            =   1920
            TabIndex        =   175
            Text            =   "Text4"
            Top             =   600
            Width           =   2415
         End
         Begin VB.OptionButton optEuler 
            Caption         =   "Pagados"
            Height          =   195
            Index           =   1
            Left            =   2280
            TabIndex        =   173
            Top             =   240
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton optEuler 
            Caption         =   "Debidos"
            Height          =   195
            Index           =   0
            Left            =   1200
            TabIndex        =   172
            Top             =   240
            Width           =   975
         End
         Begin VB.Image imgBuscarEULER 
            Height          =   240
            Left            =   0
            ToolTipText     =   "Ver datos extendido"
            Top             =   120
            Width           =   240
         End
         Begin VB.Label Label3 
            Caption         =   "Motor"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   600
            TabIndex        =   187
            Top             =   1920
            Width           =   2655
         End
         Begin VB.Label Label3 
            Caption         =   "Bombas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   16
            Left            =   480
            TabIndex        =   184
            Top             =   1320
            Width           =   2655
         End
         Begin VB.Label Label3 
            Caption         =   "Modelo"
            Height          =   195
            Index           =   14
            Left            =   5160
            TabIndex        =   183
            Top             =   1080
            Width           =   705
         End
         Begin VB.Label Label3 
            Caption         =   "Marca"
            Height          =   195
            Index           =   12
            Left            =   2160
            TabIndex        =   182
            Top             =   1080
            Width           =   705
         End
         Begin VB.Label Label3 
            Caption         =   "Ref."
            Height          =   195
            Index           =   2
            Left            =   1200
            TabIndex        =   179
            Top             =   600
            Width           =   825
         End
         Begin VB.Label Label3 
            Caption         =   "Pedido"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   3
            Left            =   480
            TabIndex        =   178
            Top             =   600
            Width           =   2655
         End
         Begin VB.Label Label3 
            Caption         =   "Fecha"
            Height          =   195
            Index           =   4
            Left            =   4560
            TabIndex        =   177
            Top             =   600
            Width           =   945
         End
         Begin VB.Label Label3 
            Caption         =   "Portes"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   21
            Left            =   480
            TabIndex        =   174
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.CheckBox chkPedxCli 
         Height          =   375
         Left            =   -63000
         TabIndex        =   201
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   18
         Left            =   -68880
         MaxLength       =   30
         TabIndex        =   36
         Tag             =   "Dir envio|N|S|0|99999|scafac1|coddiren|0000|N|"
         Text            =   "Text1"
         Top             =   2160
         Width           =   660
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   18
         Left            =   -68160
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   154
         Text            =   "Text2"
         Top             =   2160
         Width           =   3525
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   17
         Left            =   -64320
         TabIndex        =   152
         Tag             =   "Fecha|F|S|||scafac1|fecenvio|dd/mm/yyyy||"
         Top             =   2160
         Width           =   1185
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   11
         Left            =   -70920
         MaxLength       =   9
         TabIndex        =   150
         Tag             =   "Nº Bultos|N|N|0||slifac|numbultos|#,###,##0|N|"
         Text            =   "numbultos"
         Top             =   4320
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdaux 
         Caption         =   "+"
         Height          =   320
         Left            =   -65520
         TabIndex        =   124
         Top             =   4320
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   10
         Left            =   -65280
         MaxLength       =   15
         TabIndex        =   146
         Tag             =   "Nº Lote|T|S|||slifac|numlote||N|"
         Text            =   "NLote"
         Top             =   4320
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   9
         Left            =   -66120
         MaxLength       =   30
         TabIndex        =   123
         Tag             =   "Cod. Proveedor|N|N|||slifac|codprovex|0||"
         Text            =   "prove"
         Top             =   4320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   14
         Left            =   -62520
         MaxLength       =   7
         TabIndex        =   135
         Tag             =   "Nº Venta|N|S|||scafac1|numventa|0000000|N|"
         Text            =   "Text1 7"
         Top             =   1440
         Width           =   1185
      End
      Begin VB.TextBox txtAux3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   2
         Left            =   -72960
         MaxLength       =   30
         TabIndex        =   129
         Tag             =   "Fecha Albaran|F|N|||scafac1|fechaalb|dd/mm/yyyy|N|"
         Text            =   "fecalbar"
         Top             =   2160
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtAux3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   1
         Left            =   -73920
         MaxLength       =   15
         TabIndex        =   128
         Tag             =   "Nº Albaran|N|N|||scafac1|numalbar|0000000|N|"
         Text            =   "numalbar"
         Top             =   2160
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtAux3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   0
         Left            =   -74640
         MaxLength       =   7
         TabIndex        =   127
         Tag             =   "Tipo Albaran|T|N|||scafac1|codtipoa||N|"
         Text            =   "codtipoa"
         Top             =   2160
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   5
         Left            =   -69360
         MaxLength       =   5
         TabIndex        =   118
         Text            =   "origp"
         Top             =   4320
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   3
         Left            =   -71760
         MaxLength       =   12
         TabIndex        =   116
         Tag             =   "Cantidad|N|N|0||slifac|cantidad|#,###,###,##0.00|N|"
         Text            =   "cantidad"
         Top             =   4320
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   2
         Left            =   -72840
         TabIndex        =   115
         Tag             =   "Nombre Art.|T|N|||slifac|nomartic||N|"
         Text            =   "nomartic"
         Top             =   4320
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   1
         Left            =   -73680
         MaxLength       =   12
         TabIndex        =   114
         Tag             =   "Art.|T|N|||slifac|codartic||N|"
         Text            =   "codartic"
         Top             =   4320
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   0
         Left            =   -74640
         MaxLength       =   12
         TabIndex        =   113
         Tag             =   "Almacen|N|N|0|999|slifac|codalmac|000|N|"
         Text            =   "almacen"
         Top             =   4320
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton cmdObserva3 
         Height          =   375
         Left            =   -71040
         Style           =   1  'Graphical
         TabIndex        =   112
         Top             =   520
         Width           =   375
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   3
         Left            =   -68160
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   105
         Text            =   "Text2"
         Top             =   1740
         Width           =   3525
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   3
         Left            =   -68880
         MaxLength       =   30
         TabIndex        =   35
         Tag             =   "Cod. Envío|N|N|0|999|scafac1|codenvio|000|N|"
         Text            =   "Text1"
         Top             =   1740
         Width           =   660
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   4
         Left            =   -64320
         MaxLength       =   7
         TabIndex        =   99
         Tag             =   "Nº Oferta|N|S|||scafac1|numofert|0000000|N|"
         Text            =   "Text1 7"
         Top             =   1440
         Width           =   885
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   5
         Left            =   -62520
         MaxLength       =   10
         TabIndex        =   98
         Tag             =   "Fecha Oferta|F|S|||scafac1|fecofert|dd/mm/yyyy|N|"
         Top             =   1440
         Width           =   1185
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   7
         Left            =   -63240
         MaxLength       =   10
         TabIndex        =   97
         Tag             =   "Fecha Pedido|F|S|||scafac1|fecpedcl|dd/mm/yyyy|N|"
         Top             =   720
         Width           =   1185
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   6
         Left            =   -64320
         MaxLength       =   7
         TabIndex        =   96
         Tag             =   "Nº Pedido|N|S|||scafac1|numpedcl|0000000|N|"
         Text            =   "Text1 7"
         Top             =   720
         Width           =   885
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   8
         Left            =   -61920
         MaxLength       =   10
         TabIndex        =   95
         Tag             =   "Semana Entrega|N|S|||scafac1|sementre||N|"
         Top             =   720
         Width           =   705
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   0
         Left            =   -68160
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   94
         Text            =   "Text2"
         Top             =   480
         Width           =   3525
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   0
         Left            =   -68880
         MaxLength       =   30
         TabIndex        =   32
         Tag             =   "Trabajador Albaran|N|N|0|9999|scafac1|codtraba|0000|N|"
         Text            =   "Text1"
         Top             =   480
         Width           =   660
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   1
         Left            =   -68160
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   93
         Text            =   "Text2"
         Top             =   900
         Width           =   3525
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   1
         Left            =   -68880
         MaxLength       =   30
         TabIndex        =   33
         Tag             =   "Trabajador pedido|N|S|0|9999|scafac1|codtrab1|0000|N|"
         Text            =   "Text1"
         Top             =   900
         Width           =   660
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   2
         Left            =   -68160
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   92
         Text            =   "Text2"
         Top             =   1320
         Width           =   3525
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   2
         Left            =   -68880
         MaxLength       =   30
         TabIndex        =   34
         Tag             =   "Preparador materia|N|N|0|9999|scafac1|codtrab2|0000|N|"
         Text            =   "Text1"
         Top             =   1320
         Width           =   660
      End
      Begin VB.Frame FrameFactura 
         Height          =   2100
         Left            =   960
         TabIndex        =   65
         Top             =   3360
         Width           =   12255
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   17
            Left            =   4560
            MaxLength       =   5
            TabIndex        =   27
            Tag             =   "Descuento General|N|N|0|99.90|scafac|dtognral|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   345
            Width           =   525
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   16
            Left            =   2400
            MaxLength       =   5
            TabIndex        =   25
            Tag             =   "Descuento P.Pago|N|N|0|99.90|scafac|dtoppago|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   345
            Width           =   525
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   44
            Left            =   6720
            MaxLength       =   5
            TabIndex        =   142
            Tag             =   "% IVA 1|N|S|0|99.90|scafac|porciva3re|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   1605
            Width           =   525
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   43
            Left            =   7320
            MaxLength       =   15
            TabIndex        =   141
            Tag             =   "Importe IVA 1|N|S|||scafac|imporiv3re|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   1605
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   42
            Left            =   6720
            MaxLength       =   5
            TabIndex        =   140
            Tag             =   "% IVA 1|N|S|0|99.90|scafac|porciva2re|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   1275
            Width           =   525
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   41
            Left            =   7320
            MaxLength       =   15
            TabIndex        =   139
            Tag             =   "Importe IVA 1|N|S|||scafac|imporiv2re|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   1275
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   40
            Left            =   6720
            MaxLength       =   5
            TabIndex        =   138
            Tag             =   "% IVA 1|N|S|0|99.90|scafac|porciva1re|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   960
            Width           =   525
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   39
            Left            =   7320
            MaxLength       =   15
            TabIndex        =   137
            Tag             =   "Importe IVA 1|N|S|||scafac|imporiv1re|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   960
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFC0&
            Height          =   285
            Index           =   38
            Left            =   9720
            MaxLength       =   15
            TabIndex        =   87
            Tag             =   "Total Factura|N|N|||scafac|totalfac|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   1560
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   37
            Left            =   5160
            MaxLength       =   15
            TabIndex        =   82
            Tag             =   "Importe IVA 3|N|S|||scafac|imporiv3|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   1605
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   31
            Left            =   4560
            MaxLength       =   5
            TabIndex        =   81
            Tag             =   "% IVA 3|N|S|0|99.90|scafac|porciva3|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   1605
            Width           =   525
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   28
            Left            =   2400
            MaxLength       =   4
            TabIndex        =   80
            Tag             =   "Cod. IVA 3|N|S|0|9999|scafac|codigiv3|0000|N|"
            Text            =   "Text1 7"
            Top             =   1605
            Width           =   525
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   34
            Left            =   3000
            MaxLength       =   15
            TabIndex        =   79
            Tag             =   "Base Imponible 3|N|S|||scafac|baseimp3|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   1605
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   36
            Left            =   5160
            MaxLength       =   15
            TabIndex        =   78
            Tag             =   "Importe IVA 2|N|S|||scafac|imporiv2|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   1275
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   30
            Left            =   4560
            MaxLength       =   5
            TabIndex        =   77
            Tag             =   "% IVA 2|N|S|0|99.90|scafac|porciva2|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   1275
            Width           =   525
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   27
            Left            =   2400
            MaxLength       =   4
            TabIndex        =   76
            Tag             =   "Cod. IVA 2|N|S|0|9999|scafac|codigiv2|0000|N|"
            Text            =   "Text1 7"
            Top             =   1275
            Width           =   525
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   33
            Left            =   3000
            MaxLength       =   15
            TabIndex        =   75
            Tag             =   "Base Imponible 2 |N|S|||scafac|baseimp2|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   1275
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   35
            Left            =   5160
            MaxLength       =   15
            TabIndex        =   74
            Tag             =   "Importe IVA 1|N|N|||scafac|imporiv1|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   960
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   29
            Left            =   4560
            MaxLength       =   5
            TabIndex        =   73
            Tag             =   "% IVA 1|N|S|0|99.90|scafac|porciva1|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   960
            Width           =   525
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   26
            Left            =   2400
            MaxLength       =   4
            TabIndex        =   72
            Tag             =   "Cod. IVA 1|N|S|0|9999|scafac|codigiv1|0000|N|"
            Text            =   "Text1 7"
            Top             =   960
            Width           =   525
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   32
            Left            =   3000
            MaxLength       =   15
            TabIndex        =   71
            Tag             =   "Base Imponible 1|N|N|||scafac|baseimp1|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   960
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   25
            Left            =   7320
            MaxLength       =   15
            TabIndex        =   66
            Text            =   "Text1 7"
            Top             =   360
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   24
            Left            =   5160
            MaxLength       =   15
            TabIndex        =   28
            Tag             =   "Imp. Dto Gn|N|N|||scafac|impdtogr|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   360
            Width           =   1245
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   23
            Left            =   3000
            MaxLength       =   15
            TabIndex        =   26
            Tag             =   "Imp. Dto PP|N|N|||scafac|impdtopp|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   360
            Width           =   1125
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   22
            Left            =   240
            MaxLength       =   15
            TabIndex        =   24
            Tag             =   "Imp.Bruto|N|N|||scafac|brutofac|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   360
            Width           =   1245
         End
         Begin VB.Label Label1 
            Caption         =   "Dto. Gral"
            Height          =   255
            Index           =   26
            Left            =   4440
            TabIndex        =   169
            Top             =   120
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Dto. P.P"
            Height          =   255
            Index           =   25
            Left            =   2400
            TabIndex        =   168
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Importe RE"
            Height          =   195
            Index           =   44
            Left            =   7560
            TabIndex        =   145
            Top             =   720
            Width           =   825
         End
         Begin VB.Label Label1 
            Caption         =   "% RE"
            Height          =   255
            Index           =   43
            Left            =   6720
            TabIndex        =   144
            Top             =   720
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "Importe IVA"
            Height          =   195
            Index           =   37
            Left            =   5520
            TabIndex        =   143
            Top             =   720
            Width           =   825
         End
         Begin VB.Line Line1 
            X1              =   2280
            X2              =   2280
            Y1              =   960
            Y2              =   1920
         End
         Begin VB.Label Label1 
            Caption         =   "Desglose IVA"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   42
            Left            =   960
            TabIndex        =   126
            Top             =   1320
            Width           =   1155
         End
         Begin VB.Label Label1 
            Caption         =   "% IVA"
            Height          =   255
            Index           =   41
            Left            =   4560
            TabIndex        =   125
            Top             =   720
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "TOTAL FACTURA"
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
            Left            =   9720
            TabIndex        =   90
            Top             =   1320
            Width           =   1530
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
            Index           =   38
            Left            =   9360
            TabIndex        =   89
            Top             =   1560
            Width           =   135
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
            TabIndex        =   88
            Top             =   2160
            Width           =   135
         End
         Begin VB.Label Label1 
            Caption         =   "Base impo. IVA"
            Height          =   255
            Index           =   33
            Left            =   3120
            TabIndex        =   86
            Top             =   720
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
            Left            =   6840
            TabIndex        =   85
            Top             =   360
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
            Left            =   4320
            TabIndex        =   84
            Top             =   360
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
            Left            =   1680
            TabIndex        =   83
            Top             =   360
            Width           =   135
         End
         Begin VB.Label Label1 
            Caption         =   "Base Imponible"
            Height          =   255
            Index           =   14
            Left            =   7440
            TabIndex        =   70
            Top             =   120
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Imp. Dto Gn"
            Height          =   255
            Index           =   12
            Left            =   5520
            TabIndex        =   69
            Top             =   120
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Imp. Dto PP"
            Height          =   255
            Index           =   11
            Left            =   3120
            TabIndex        =   68
            Top             =   120
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Imp. Bruto"
            Height          =   255
            Index           =   10
            Left            =   240
            TabIndex        =   67
            Top             =   120
            Width           =   855
         End
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   4
         Left            =   -70200
         MaxLength       =   12
         TabIndex        =   117
         Tag             =   "Precio|N|N|0|999999.0000|slifac|precioar|###,##0.0000|N|"
         Text            =   "Precio"
         Top             =   4320
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   6
         Left            =   -68640
         MaxLength       =   5
         TabIndex        =   119
         Tag             =   "Dto 1|N|N|0|99.90|slifac|dtoline1|#0.00|N|"
         Text            =   "Dto1"
         Top             =   4320
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   7
         Left            =   -67920
         MaxLength       =   30
         TabIndex        =   120
         Tag             =   "Dto 2|N|N|0|99.90|slifac|dtolinea|#0.00|N|"
         Text            =   "Dto2"
         Top             =   4320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   8
         Left            =   -67320
         MaxLength       =   12
         TabIndex        =   122
         Tag             =   "Importe|N|N|0||slifac|importel|#,###,###,##0.00|N|"
         Text            =   "Importe"
         Top             =   4320
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Frame FrameCliente 
         Caption         =   "Datos Cliente"
         ForeColor       =   &H00972E0B&
         Height          =   2295
         Left            =   960
         TabIndex        =   44
         Top             =   600
         Width           =   12255
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   46
            Left            =   7845
            MaxLength       =   4
            TabIndex        =   19
            Tag             =   "IBAN|T|S|||scafac|iban|||"
            Text            =   "Text1 7"
            Top             =   1740
            Width           =   525
         End
         Begin VB.CheckBox Check2 
            Caption         =   "FacturaE"
            Height          =   375
            Left            =   3720
            TabIndex        =   29
            Tag             =   "En Factura E|N|N|0|1|scafac|EnFacturaE||N|"
            Top             =   1320
            Width           =   1095
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   3
            Left            =   11280
            MaxLength       =   10
            TabIndex        =   151
            Tag             =   "Aportacion|N|S|||scafac|portes|#,##0.00|N|"
            Text            =   "Portes"
            Top             =   1740
            Visible         =   0   'False
            Width           =   525
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   45
            Left            =   5040
            MaxLength       =   10
            TabIndex        =   14
            Tag             =   "Aportacion|N|S|||scafac|aportacion|#,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   1740
            Visible         =   0   'False
            Width           =   1125
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   21
            Left            =   9930
            MaxLength       =   10
            TabIndex        =   23
            Tag             =   "Cuenta Bancaria|T|S|||scafac|cuentaba|0000000000|N|"
            Text            =   "9999999999"
            Top             =   1740
            Width           =   1125
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   20
            Left            =   9570
            MaxLength       =   2
            TabIndex        =   22
            Tag             =   "Digito Control|T|S|||scafac|digcontr|00|N|"
            Text            =   "Text1 7"
            Top             =   1740
            Width           =   315
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   19
            Left            =   9000
            MaxLength       =   4
            TabIndex        =   21
            Tag             =   "Sucursal|N|S|0|9999|scafac|codsucur|0000|N|"
            Text            =   "Text1 7"
            Top             =   1740
            Width           =   525
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   18
            Left            =   8400
            MaxLength       =   4
            TabIndex        =   20
            Tag             =   "Banco|N|S|0|9999|scafac|codbanco|0000|N|"
            Text            =   "Text1 7"
            Top             =   1740
            Width           =   525
         End
         Begin VB.TextBox Text3 
            Height          =   315
            Index           =   16
            Left            =   1125
            MaxLength       =   20
            TabIndex        =   13
            Tag             =   "Refere. Cliente|T|S|||scafac1|referenc|||"
            Text            =   "Text1 Text1 Text1 Te"
            Top             =   1740
            Width           =   1725
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   13
            Left            =   8430
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   16
            Tag             =   "Direccion/Dpto.|T|S|||scafac|nomdirec||N|"
            Text            =   "Text1"
            Top             =   285
            Width           =   3405
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   12
            Left            =   7845
            MaxLength       =   3
            TabIndex        =   15
            Tag             =   "Direccion/Dpto.|N|S|0|999|scafac|coddirec|000|N|"
            Text            =   "Text1"
            Top             =   285
            Width           =   540
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   11
            Left            =   1125
            MaxLength       =   30
            TabIndex        =   12
            Tag             =   "Provincia|T|N|||scafac|proclien||N|"
            Text            =   "Text1 Text1 Text1 Text1 Text22"
            Top             =   1350
            Width           =   2445
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   9
            Left            =   1125
            MaxLength       =   6
            TabIndex        =   10
            Tag             =   "CPostal|T|N|||scafac|codpobla||N|"
            Text            =   "Text15"
            Top             =   990
            Width           =   630
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   10
            Left            =   1755
            MaxLength       =   30
            TabIndex        =   11
            Tag             =   "Población|T|N|||scafac|pobclien||N|"
            Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwaq"
            Top             =   990
            Width           =   3405
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   7
            Left            =   3195
            MaxLength       =   20
            TabIndex        =   8
            Tag             =   "teléfono Cliente|T|S|||scafac|telclien||N|"
            Text            =   "12345678911234567899"
            Top             =   285
            Width           =   1965
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   6
            Left            =   1125
            MaxLength       =   15
            TabIndex        =   7
            Tag             =   "NIF Cliente|T|N|||scafac|nifclien||N|"
            Text            =   "123456789"
            Top             =   285
            Width           =   1110
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   14
            Left            =   7845
            MaxLength       =   4
            TabIndex        =   17
            Tag             =   "Cod. Agente|N|N|0|9999|scafac|codagent|0000|N|"
            Text            =   "Text1"
            Top             =   645
            Width           =   540
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   14
            Left            =   8430
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   48
            Text            =   "Text2"
            Top             =   645
            Width           =   3405
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   15
            Left            =   7845
            MaxLength       =   3
            TabIndex        =   18
            Tag             =   "Forma de Pago|N|N|0|999|scafac|codforpa|000|N|"
            Text            =   "Text1"
            Top             =   990
            Width           =   540
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   15
            Left            =   8430
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   46
            Text            =   "Text2"
            Top             =   990
            Width           =   3405
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   8
            Left            =   1125
            MaxLength       =   35
            TabIndex        =   9
            Tag             =   "Domicilio|T|N|||scafac|domclien||N|"
            Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwwwww aq"
            Top             =   645
            Width           =   4030
         End
         Begin VB.Label Label1 
            Caption         =   "IBAN"
            Height          =   255
            Index           =   52
            Left            =   6720
            TabIndex        =   167
            Top             =   1800
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Aportación"
            Height          =   255
            Index           =   45
            Left            =   5040
            TabIndex        =   147
            Top             =   1485
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   4
            Left            =   7560
            ToolTipText     =   "Buscar agente"
            Top             =   645
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Cuenta"
            Height          =   255
            Index           =   8
            Left            =   10320
            TabIndex        =   64
            Top             =   1560
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "DC"
            Height          =   255
            Index           =   5
            Left            =   9840
            TabIndex        =   63
            Top             =   1560
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label1 
            Caption         =   "Sucursal"
            Height          =   255
            Index           =   4
            Left            =   9120
            TabIndex        =   62
            Top             =   1560
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Banco"
            Height          =   255
            Index           =   3
            Left            =   8340
            TabIndex        =   61
            Top             =   1560
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "Ref. Cliente"
            Height          =   255
            Index           =   13
            Left            =   120
            TabIndex        =   55
            Top             =   1740
            Width           =   855
         End
         Begin VB.Image imgBuscar 
            Enabled         =   0   'False
            Height          =   240
            Index           =   2
            Left            =   855
            ToolTipText     =   "Buscar población"
            Top             =   1005
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Direc./Dpto"
            Height          =   255
            Index           =   1
            Left            =   6660
            TabIndex        =   54
            Top             =   285
            Width           =   855
         End
         Begin VB.Image imgBuscar 
            Enabled         =   0   'False
            Height          =   240
            Index           =   3
            Left            =   7560
            ToolTipText     =   "Buscar direc./dpto"
            Top             =   285
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Provincia"
            Height          =   255
            Index           =   17
            Left            =   120
            TabIndex        =   53
            Top             =   1350
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Población"
            Height          =   255
            Index           =   16
            Left            =   120
            TabIndex        =   52
            Top             =   990
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Teléfono"
            Height          =   255
            Index           =   19
            Left            =   2445
            TabIndex        =   51
            Top             =   285
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "NIF"
            Height          =   255
            Index           =   20
            Left            =   120
            TabIndex        =   50
            Top             =   285
            Width           =   615
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   1
            Left            =   855
            ToolTipText     =   "Buscar cliente varios"
            Top             =   300
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Agente"
            Height          =   255
            Index           =   34
            Left            =   6660
            TabIndex        =   49
            Top             =   645
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Forma Pago"
            Height          =   255
            Index           =   15
            Left            =   6660
            TabIndex        =   47
            Top             =   990
            Width           =   855
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   5
            Left            =   7560
            ToolTipText     =   "Buscar forma de pago"
            Top             =   990
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Domicilio"
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   45
            Top             =   645
            Width           =   735
         End
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmFacHcoFacturas.frx":EAC0
         Height          =   2265
         Left            =   -74760
         TabIndex        =   60
         Top             =   3300
         Width           =   13815
         _ExtentX        =   24368
         _ExtentY        =   3995
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
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "frmFacHcoFacturas.frx":EAD5
         Height          =   2550
         Left            =   -74760
         TabIndex        =   91
         Top             =   525
         Width           =   3600
         _ExtentX        =   6350
         _ExtentY        =   4498
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
         Caption         =   "Albaranes de la Factura"
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
      Begin VB.Frame FrameObserva 
         Caption         =   "Observaciones"
         ForeColor       =   &H00972E0B&
         Height          =   2295
         Left            =   -74880
         TabIndex        =   106
         Tag             =   "Observación 4|T|S|||scafac1|observa4||N|"
         Top             =   3300
         Width           =   13695
         Begin VB.Frame FrameRecepMercan 
            Caption         =   "Recepción mercancia"
            Height          =   1815
            Left            =   7920
            TabIndex        =   219
            Top             =   120
            Visible         =   0   'False
            Width           =   5655
            Begin VB.TextBox Text3 
               Height          =   300
               Index           =   27
               Left            =   4320
               MaxLength       =   80
               TabIndex        =   224
               Tag             =   "Geo-Long|N|S|||scafac1|longitud|#0.00000|N|"
               Text            =   "Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Te"
               Top             =   1320
               Width           =   1185
            End
            Begin VB.TextBox Text3 
               Height          =   300
               Index           =   26
               Left            =   2760
               MaxLength       =   80
               TabIndex        =   223
               Tag             =   "Geo-Latitud|N|S|||scafac1|latitud|#0.00000|N|"
               Text            =   "Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Te"
               Top             =   1320
               Width           =   1185
            End
            Begin VB.TextBox Text3 
               Height          =   300
               Index           =   25
               Left            =   240
               MaxLength       =   80
               TabIndex        =   222
               Tag             =   "T|T|S|||scafac1|dnient||N|"
               Text            =   "Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Te"
               Top             =   1320
               Width           =   2025
            End
            Begin VB.TextBox Text3 
               Height          =   300
               Index           =   24
               Left            =   2160
               MaxLength       =   80
               TabIndex        =   221
               Tag             =   "C|T|S|||scafac1|perrecep||N|"
               Text            =   "Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Te"
               Top             =   600
               Width           =   3345
            End
            Begin VB.TextBox Text3 
               Height          =   300
               Index           =   23
               Left            =   240
               MaxLength       =   80
               TabIndex        =   220
               Tag             =   "Observación 1|FH|S|||scafac1|fechaent|dd/mm/yyy hh:nn:ss||"
               Top             =   600
               Width           =   1665
            End
            Begin VB.Image imgFirmaRecep 
               Height          =   480
               Left            =   5160
               Picture         =   "frmFacHcoFacturas.frx":EAEA
               ToolTipText     =   "Firma de la recepción del albaran"
               Top             =   0
               Visible         =   0   'False
               Width           =   480
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Recepción mercancia"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   59
               Left            =   120
               TabIndex        =   230
               Top             =   0
               Width           =   1845
            End
            Begin VB.Image imgGeolocalizacion 
               Height          =   255
               Left            =   3840
               Picture         =   "frmFacHcoFacturas.frx":EDF4
               Stretch         =   -1  'True
               Tag             =   "-1"
               ToolTipText     =   "Abrir web"
               Top             =   1080
               Width           =   255
            End
            Begin VB.Label Label1 
               Caption         =   "Fecha"
               Height          =   255
               Index           =   78
               Left            =   240
               TabIndex        =   229
               Top             =   360
               Width           =   735
            End
            Begin VB.Label Label1 
               Caption         =   "Persona recepcion"
               Height          =   195
               Index           =   80
               Left            =   2160
               TabIndex        =   228
               Top             =   360
               Width           =   1335
            End
            Begin VB.Label Label1 
               Caption         =   "DNI"
               Height          =   195
               Index           =   81
               Left            =   240
               TabIndex        =   227
               Top             =   1080
               Width           =   1335
            End
            Begin VB.Label Label1 
               Caption         =   "Latitud"
               Height          =   195
               Index           =   82
               Left            =   2760
               TabIndex        =   226
               Top             =   1080
               Width           =   480
            End
            Begin VB.Label Label1 
               Caption         =   "Longitud"
               Height          =   195
               Index           =   83
               Left            =   4320
               TabIndex        =   225
               Top             =   1080
               Width           =   615
            End
            Begin VB.Line Line3 
               X1              =   0
               X2              =   0
               Y1              =   240
               Y2              =   1680
            End
         End
         Begin VB.TextBox Text3 
            Height          =   300
            Index           =   13
            Left            =   2520
            MaxLength       =   80
            TabIndex        =   111
            Tag             =   "Observación 5|T|S|||scafac1|observa5||N|"
            Text            =   "Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Te"
            Top             =   1560
            Width           =   8940
         End
         Begin VB.TextBox Text3 
            Height          =   300
            Index           =   12
            Left            =   2520
            MaxLength       =   80
            TabIndex        =   110
            Tag             =   "Observación 4|T|S|||scafac1|observa4||N|"
            Text            =   "Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Te"
            Top             =   1230
            Width           =   8940
         End
         Begin VB.TextBox Text3 
            Height          =   300
            Index           =   11
            Left            =   2520
            MaxLength       =   80
            TabIndex        =   109
            Tag             =   "Observación 3|T|S|||scafac1|observa3||N|"
            Text            =   "Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Te"
            Top             =   900
            Width           =   8940
         End
         Begin VB.TextBox Text3 
            Height          =   300
            Index           =   10
            Left            =   2520
            MaxLength       =   80
            TabIndex        =   108
            Tag             =   "Observación 2|T|S|||scafac1|observa2||N|"
            Text            =   "Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Te"
            Top             =   570
            Width           =   8940
         End
         Begin VB.TextBox Text3 
            Height          =   300
            Index           =   9
            Left            =   2520
            MaxLength       =   80
            TabIndex        =   107
            Tag             =   "Observación 1|T|S|||scafac1|observa1||N|"
            Text            =   "Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Te"
            Top             =   240
            Width           =   8940
         End
      End
      Begin VB.CheckBox chkEnvio 
         Height          =   375
         Left            =   -63000
         TabIndex        =   156
         Top             =   2160
         Width           =   375
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   15
         Left            =   -64320
         MaxLength       =   7
         TabIndex        =   136
         Tag             =   "Nº Terminal|N|S|||scafac1|numtermi||N|"
         Text            =   "Text1 7"
         Top             =   1440
         Width           =   885
      End
      Begin VB.Frame FrameTelefonia 
         Height          =   2535
         Left            =   -74880
         TabIndex        =   162
         Top             =   3300
         Visible         =   0   'False
         Width           =   13935
         Begin MSComctlLib.ListView ListView2 
            Height          =   2055
            Left            =   1080
            TabIndex        =   163
            Top             =   240
            Width           =   5655
            _ExtentX        =   9975
            _ExtentY        =   3625
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Concepto"
               Object.Width           =   6950
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   1
               Text            =   "Importe"
               Object.Width           =   1940
            EndProperty
         End
         Begin MSComctlLib.ListView ListView3 
            Height          =   2055
            Left            =   7800
            TabIndex        =   165
            Top             =   240
            Width           =   5895
            _ExtentX        =   10398
            _ExtentY        =   3625
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Tipo"
               Object.Width           =   4234
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Numero"
               Object.Width           =   1940
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Fecha"
               Object.Width           =   1235
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Text            =   "Importe"
               Object.Width           =   1411
            EndProperty
         End
         Begin VB.Label Label1 
            Caption         =   "Detalles"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008080&
            Height          =   195
            Index           =   51
            Left            =   6960
            TabIndex        =   166
            Top             =   240
            Width           =   705
         End
         Begin VB.Label Label1 
            Caption         =   "Conceptos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   255
            Index           =   50
            Left            =   120
            TabIndex        =   164
            Top             =   240
            Width           =   1455
         End
      End
      Begin MSComctlLib.ListView lwCostes 
         Height          =   3615
         Left            =   -74760
         TabIndex        =   207
         Top             =   1080
         Width           =   13695
         _ExtentX        =   24156
         _ExtentY        =   6376
         SortKey         =   8
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   10
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Tipo"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   5115
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Documento"
            Object.Width           =   2116
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Fecha"
            Object.Width           =   2205
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Descripción"
            Object.Width           =   6350
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Cantidad"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Precio"
            Object.Width           =   2010
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Importe"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "ORDEN"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "codartic"
            Object.Width           =   0
         EndProperty
      End
      Begin MSComctlLib.ListView lwEulerLineas 
         Height          =   3495
         Left            =   -74880
         TabIndex        =   231
         Top             =   1320
         Width           =   13695
         _ExtentX        =   24156
         _ExtentY        =   6165
         SortKey         =   5
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Articulo"
            Object.Width           =   3704
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   11642
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Cantidad"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Precio"
            Object.Width           =   2010
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Dto"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Importe"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "ORDEN"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "linea"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "descripcionReal"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "Referencia pedido"
         Height          =   195
         Index           =   62
         Left            =   -66000
         TabIndex        =   261
         Top             =   2640
         Width           =   1530
      End
      Begin VB.Label Label1 
         Caption         =   "Referencia contrato"
         Height          =   195
         Index           =   61
         Left            =   -70560
         TabIndex        =   260
         Top             =   2640
         Width           =   1530
      End
      Begin VB.Label Label1 
         Caption         =   "Pedido por cliente"
         Height          =   255
         Index           =   57
         Left            =   -62640
         TabIndex        =   202
         Top             =   1875
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Documento archivado"
         Height          =   255
         Index           =   49
         Left            =   -62640
         TabIndex        =   157
         Top             =   2240
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Direccion envio"
         Height          =   195
         Index           =   48
         Left            =   -70560
         TabIndex        =   155
         Top             =   2220
         Width           =   1140
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   10
         Left            =   -69120
         ToolTipText     =   "Buscar codigo direccion envio"
         Top             =   2220
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha envio"
         Height          =   255
         Index           =   47
         Left            =   -64320
         TabIndex        =   153
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   6
         Left            =   -69120
         ToolTipText     =   "Buscar trabajador"
         Top             =   480
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   9
         Left            =   -69120
         ToolTipText     =   "Buscar forma de envio"
         Top             =   1785
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   8
         Left            =   -69120
         ToolTipText     =   "Buscar trabajador"
         Top             =   1350
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Oferta"
         Height          =   255
         Index           =   40
         Left            =   -64320
         TabIndex        =   104
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Oferta"
         Height          =   255
         Index           =   22
         Left            =   -62400
         TabIndex        =   103
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Pedido"
         Height          =   255
         Index           =   18
         Left            =   -63240
         TabIndex        =   102
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Pedido"
         Height          =   255
         Index           =   6
         Left            =   -64320
         TabIndex        =   101
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Sem. Entrega"
         Height          =   255
         Index           =   2
         Left            =   -62040
         TabIndex        =   100
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Trabajador Albaran"
         Height          =   255
         Index           =   21
         Left            =   -70560
         TabIndex        =   59
         Top             =   525
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo  Envío"
         Height          =   195
         Index           =   24
         Left            =   -70560
         TabIndex        =   58
         Top             =   1839
         Width           =   1140
      End
      Begin VB.Label Label1 
         Caption         =   "Prepar. Material"
         Height          =   255
         Index           =   23
         Left            =   -70560
         TabIndex        =   57
         Top             =   1401
         Width           =   1185
      End
      Begin VB.Label Label1 
         Caption         =   "Trabajador Pedido"
         Height          =   255
         Index           =   9
         Left            =   -70560
         TabIndex        =   56
         Top             =   963
         Width           =   1425
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   7
         Left            =   -69120
         ToolTipText     =   "Buscar trabajador"
         Top             =   915
         Width           =   240
      End
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   360
      Index           =   16
      Left            =   3480
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   121
      Top             =   7080
      Visible         =   0   'False
      Width           =   7485
   End
   Begin VB.Frame Frame1 
      Height          =   475
      Index           =   0
      Left            =   120
      TabIndex        =   38
      Top             =   6975
      Width           =   2175
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   240
         Left            =   240
         TabIndex        =   39
         Top             =   180
         Width           =   1755
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   13050
      TabIndex        =   31
      Top             =   7200
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   11760
      TabIndex        =   30
      Top             =   7200
      Width           =   1035
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   40
      Top             =   0
      Width           =   14580
      _ExtentX        =   25718
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   26
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
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "1"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   "0"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Lineas Factura"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir Factura"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir albarán"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Asignar numero LOTE"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Campos"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Comision linea "
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cambiar fecha/reestablecer albaran"
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Valoracion"
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Facturas/albaranes firmados"
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Último"
            ImageIndex      =   9
         EndProperty
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   12000
         TabIndex        =   41
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   13050
      TabIndex        =   37
      Top             =   7200
      Visible         =   0   'False
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data3 
      Height          =   330
      Left            =   3000
      Top             =   1080
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
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   11
      Left            =   3240
      ToolTipText     =   "Buscar codigo direccion envio"
      Top             =   7080
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Proveedor"
      Height          =   255
      Index           =   46
      Left            =   2400
      TabIndex        =   148
      Top             =   7515
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Ampliación "
      Height          =   195
      Index           =   35
      Left            =   2400
      TabIndex        =   43
      Top             =   7080
      Visible         =   0   'False
      Width           =   810
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
      Begin VB.Menu mnModificar 
         Caption         =   "&Modificar"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnEliminar 
         Caption         =   "&Eliminar"
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
      Begin VB.Menu mnImprimir 
         Caption         =   "&Imprimir"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnImprimirAlbaran 
         Caption         =   "Imprimir &albarán"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnModLotes 
         Caption         =   "Cambiar &numeros de lote"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnEditarCampos 
         Caption         =   "Asignación campos"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnTipoPreciosLinea 
         Caption         =   "Tipo de precios lineas"
         Shortcut        =   ^T
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
Attribute VB_Name = "frmFacHcoFacturas2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'========== VBLES PUBLICAS ====================
Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

'Si se llama de la busqueda en el frmAlmMovimArticulos se accede
'a las tablas del Albaran o de Facturas de movimiento seleccionado (solo consulta)
Public hcoCodMovim As String 'cod. movim
Public hcoCodTipoM As String 'Codigo detalle de Movimiento(ALC)
Public hcoFechaMov As String 'fecha del movimiento

Public DesdeFichaCliente As Boolean

Private frmNLote As frmAlmCargarNLote

'========== VBLES PRIVADAS ====================
Private WithEvents frmB As frmBuscaGrid 'Form para busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmCP As frmCPostal 'Codigos Postales
Attribute frmCP.VB_VarHelpID = -1

Private WithEvents frmC As frmBasico2 'Form Mto Clientes
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmCV As frmBasico2 'frmFacClientesV  'Form Mto Clientes Varios
Attribute frmCV.VB_VarHelpID = -1
Private WithEvents frmFP As frmBasico2 'frmFacFormasPago 'Form Mto Formas de Pago
Attribute frmFP.VB_VarHelpID = -1
Private WithEvents frmT As frmBasico2 'frmAdmTrabajadores  'Form Mto Trabajadores
Attribute frmT.VB_VarHelpID = -1
Private WithEvents frmA As frmBasico2 '%=%=frmFacAgentesCom   'Form Mto Agentes
Attribute frmA.VB_VarHelpID = -1
Private WithEvents frmFE As frmFacFormasEnvio  'Form Formas de Envio
Attribute frmFE.VB_VarHelpID = -1
Private WithEvents frmP As frmBasico2 '%=%=frmComProveedores
Attribute frmP.VB_VarHelpID = -1

Private WithEvents frmAlb As frmBasico2 'Albaranes EULER para poder insertar
Attribute frmAlb.VB_VarHelpID = -1

Private WithEvents frmCliV As frmFacClientesV
Attribute frmCliV.VB_VarHelpID = -1

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
'-------------------------------------------------------------------------
'-------------------------------------------------------------------------

Dim ModificaLineas As Byte
'1.- Añadir,  2.- Modificar,  3.- Borrar,  0.-Pasar control a Lineas

Dim CodTipoMov As String
'Codigo tipo de movimiento en función del valor en la tabla de parámetros: stipom

Dim TituloLinea As String 'Descripcion de la linea que estamos en Mantenimiento

Dim PrimeraVez As Boolean

Dim EsCabecera2 As Byte    '0-Cabcera     1-Dpto
'Para saber en MandaBusquedaPrevia si busca en la tabla scapla o en la tabla sdirec


Dim EsDeVarios As Boolean
'Si el cliente mostrado es de Varios o No

'SQL de la tabla principal del formulario
Private CadenaConsulta As String

Private Ordenacion As String
Private NombreTabla As String  'Nombre de la tabla o de la
Private NomTablaLineas As String 'Nombre de la Tabla de lineas
Private kCampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean

Dim btnPrimero As Byte
'Variable que indica el número del Boton  PrimerRegistro en la Toolbar1


Private VieneDeBuscar As Boolean
'Para cuando devuelve dos poblaciones con el mismo codigo Postal. Si viene de pulsar prismatico
'de busqueda poner el valor de poblacion seleccionado y no volver a recuperar de la Base de Datos
Private HaCambiadoCP As Boolean
'Para saber si tras haber vuelto de prismaticos ha cambiado el valor del CPostal

Private UnaVez As Boolean
Private BuscaChekc As String

Private LetrasFraTelefonia As String
Private SolapaCamposFito As Boolean
Dim CarpetaImagenesEULER  As String



'Para cuando pueda modificar algo
Private TipoModificacionFactura As Byte
'       0.- Se puede TODO. NO esta contabilizada
'       1.- Se puede "parcial", lo que hacia antes
'       2.- Solo se puede trabajador, zona,referencias.....
'       3.- NADA. Fuera ejercicios





Private Sub Check1_Click()
    If Modo = 1 Then CheckCadenaBusqueda Check1, BuscaChekc
End Sub

Private Sub Check1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Check2_Click()
     If Modo = 1 Then CheckCadenaBusqueda Check2, BuscaChekc
End Sub

Private Sub Check2_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkEnvio_Click()
    If Modo = 1 Then CheckCadenaBusqueda chkEnvio, BuscaChekc
End Sub

Private Sub chkEnvio_KeyPress(KeyAscii As Integer)
     KEYpress KeyAscii
End Sub


Private Sub chkPedxCli_Click()
    If Modo = 1 Then CheckCadenaBusqueda chkPedxCli, BuscaChekc
End Sub

Private Sub chkPedxCli_KeyPress(KeyAscii As Integer)
     KEYpress KeyAscii
End Sub


Private Sub cmdAceptar_Click()
Dim i As Integer
Dim Cambios As String

    Screen.MousePointer = vbHourglass
    On Error GoTo Error1

    Select Case Modo
        Case 1  'BUSQUEDA
            
            HacerBusqueda
            
        Case 4  'MODIFICAR
            If DatosOk Then
                
               If ModificarFactura Then
               
                    FijarCadenaModificaUsuarioNormal Cambios
                    Set LOG = New cLOG
                    LOG.Insertar 8, vUsu, "Factura modificada: " & Text1(1).Text & Text1(0).Text & " " & Text1(2).Text & vbCrLf & Cambios
                    Set LOG = Nothing
               
               
               
               
                    Espera 0.2
                    TerminaBloquear
                    PosicionarData
                    FormatoDatosTotales
                    i = Data3.Recordset.AbsolutePosition
                    PonerCamposLineas
                    SituarDataPosicion Data3, CLng(i), ""
                End If
            End If
            
         Case 5 'InsertarModificar LINEAS
            If ModificaLineas = 2 Then 'MODIFICAR lineas
                If ModificarLinea Then
                
                        'INSERTA LOG
                        '-------------------------------------------------
                        Set LOG = New cLOG
                        BuscaChekc = "     Alb : " & Data2.Recordset!Numalbar & "   Linea: " & Data2.Recordset!numlinea
                        BuscaChekc = "Modificar linea: " & Text1(1).Text & Text1(0).Text & " " & Text1(2).Text & BuscaChekc
                        LOG.Insertar 8, vUsu, BuscaChekc
                        Set LOG = Nothing
                        BuscaChekc = ""
                
                    TerminaBloquear
                    CargaGrid DataGrid1, Data2, True
                    ModificaLineas = 0
                    PonerBotonCabecera True
                    BloquearTxt Text2(16), True
            
                    LLamaLineas Modo, 0, "DataGrid1"
                    PosicionarData
                Else
                    TerminaBloquear
                End If
                Me.DataGrid1.Enabled = True
            End If
    End Select
    Screen.MousePointer = vbDefault

Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub cmdAux_Click()
'    Set frmP = New frmComProveedores
'    frmP.DatosADevolverBusqueda = "0|1|"
'    frmP.Show vbModal
    Set frmP = New frmBasico2
    AyudaProveedores frmP, txtAux(9).Text
    Set frmP = Nothing
    PonerFoco txtAux(9)

End Sub

Private Sub cmdCancelar_Click()
    Select Case Modo
        Case 0, 1 'Busqueda, Insertar
            LimpiarCampos
            'Poner los grid sin apuntar a nada
            LimpiarDataGrids
            PonerModo 0
            LLamaLineas Modo, 0, "DataGrid2"
            PonerFoco Text1(0)
            
        Case 4  'Modificar
            TerminaBloquear
            PonerModo 2
            PonerCampos
            PonerFoco Text1(0)
            
        Case 5 'Lineas Detalle
            TerminaBloquear
            BloquearTxt Text2(16), True
            If ModificaLineas = 1 Then 'INSERTAR
                ModificaLineas = 0
                DataGrid1.AllowAddNew = False
                If Not Data2.Recordset.EOF Then Data2.Recordset.MoveFirst
            End If
            DataGrid2.Enabled = True
            ModificaLineas = 0
            LLamaLineas Modo, 0, "DataGrid1"
            PonerBotonCabecera True
            Me.DataGrid1.Enabled = True
        Case 6
            'Campos asociados a la factura
            '
            Me.cmdMtoCampos(0).visible = False
            Me.cmdMtoCampos(1).visible = False
    End Select
End Sub


Private Sub BotonBuscar()
Dim anc As Single

    'Buscar
    If Modo <> 1 Then
        BuscaChekc = ""
        LimpiarCampos
        'Poner los grid sin apuntar a nada
        LimpiarDataGrids
        PonerModo 1
        
        'poner los txtaux para buscar por lineas de albaran
        anc = DataGrid2.Top
        If DataGrid2.Row < 0 Then
            anc = anc + 440
        Else
            anc = anc + DataGrid2.RowTop(DataGrid2.Row) + 20
        End If
        LLamaLineas Modo, anc, "DataGrid2"
        
        
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
Dim cadB As String

    If vParamAplic.NumeroInstalacion = vbHerbelca Then
        'SOLO PARA HERBELCA
        If (vUsu.Codigo Mod 1000) > 0 Then
    
            cadB = " scafac.codtipom "
            If Val(vUsu.AlmacenPorDefecto2) = vParamAplic.AlmacenB Then
                cadB = cadB & " = "
            Else
                cadB = cadB & " <> "
            End If
            cadB = cadB & " 'FAZ'"
        Else
            cadB = " 1=1"
        End If
        If vUsu.CodigoAgente > 0 Then
             cadB = cadB & " AND (codagent = " & vUsu.CodigoAgente
            If vUsu.ClientesEnQueAgenteEsVisitador <> "" Then cadB = cadB & " OR codclien in (" & vUsu.ClientesEnQueAgenteEsVisitador & ")"
            cadB = cadB & ")"
        End If
    Else
        cadB = " 1=1"
        If vParamAplic.NumeroInstalacion = vbFenollar Then
            If Not HaMostradoCanal2_El_B Then cadB = "scafac.codtipom<>'FAZ'"
        End If
    End If

    If chkVistaPrevia.Value = 1 Then
        EsCabecera2 = 0
        MandaBusquedaPrevia cadB
    Else
        lblIndicador.Caption = "Preparando bus."
        lblIndicador.Refresh
        LimpiarCampos
        LimpiarDataGrids
        DoEvents
        
        CadenaConsulta = "Select scafac.* "
        CadenaConsulta = CadenaConsulta & "from " & NombreTabla & " whERE " & cadB
'        CadenaConsulta = CadenaConsulta & " WHERE scafac.codtipom='" & CodTipoMov & "'"
        lblIndicador.Caption = "Obteniendo reg."
        lblIndicador.Refresh
        PonerCadenaBusqueda
    End If
End Sub


Private Sub Desplazamiento(Index As Integer)
'Botones de Desplazamiento de la Toolbar
'Para desplazarse por los registros de control Data
    If Data1.Recordset.EOF Then Exit Sub
    DesplazamientoData Data1, Index
    PonerCampos
End Sub


Private Sub BotonModificar()
Dim DeVarios As Boolean
Dim EnTesoreria  As String

    
    
    
    If vParamAplic.ArticuloRecargoFinanciero <> "" Then
        'Tiene recargo financiero
        'Veremos si la factura tienen la linea con recargo financiero
        EnTesoreria = ObtenerWhereCP(False)
        EnTesoreria = EnTesoreria & " and codartic = '" & vParamAplic.ArticuloRecargoFinanciero & "' AND 1"
        
        EnTesoreria = DevuelveDesdeBD(conAri, "count(*)", "slifac", EnTesoreria, "1")
        
        If Val(EnTesoreria) > 0 Then
            'Tienen linea recargo financiero
            'Veremos si el cliente la tienen ahora
            EnTesoreria = DevuelveDesdeBD(conAri, "RecargoFinanciero", "sclien", "codclien", CStr(Data1.Recordset!codClien))
            If EnTesoreria = "" Then EnTesoreria = "0"
            If Val(EnTesoreria) = 0 Then
                EnTesoreria = "La factura tiene recargo financiero y el cliente ahora no lo tiene. Continuar?"
                If MsgBox(EnTesoreria, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
            End If
        End If
        EnTesoreria = ""
    End If
    
    
    'solo se puede modificar la factura si no esta contabilizada
    If FactContabilizada(EnTesoreria) Then
        TerminaBloquear
        Exit Sub
    End If
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 4
    
    'PonerFocoChk Me.Check1
        
    'Inserto en slog
    
    
    If EnTesoreria <> "" Then
        Set LOG = New cLOG
        EnTesoreria = "Tesoreria: " & vbCrLf & EnTesoreria
        EnTesoreria = Text1(1).Text & Text1(0).Text & " " & Text1(2).Text & vbCrLf & EnTesoreria
        EnTesoreria = "Pulsa mod factura: " & EnTesoreria
        LOG.Insertar 8, vUsu, EnTesoreria
        Set LOG = Nothing
        Espera 0.3
    End If
    
    
    
    If Text1(16).Text = "" Then Text1(16).Text = Format(Data1.Recordset!impdtopp, FormatoDescuento)
    If Text1(17).Text = "" Then Text1(17).Text = Format(Data1.Recordset!impdtogr, FormatoDescuento)

    
    
    'Si es Cliente de Varios no se pueden modificar sus datos
    DeVarios = EsClienteVarios(Text1(4).Text)
    BloquearDatosCliente (DeVarios)
End Sub


Private Sub BotonModificarLinea()
'Modificar una linea
Dim vWhere As String
Dim anc As Single
Dim J As Byte
Dim EstaEnTesoreria As String
    On Error GoTo EModificarLinea


     'solo se puede modificar la factura si no esta contabilizada
    If FactContabilizada(EstaEnTesoreria) Then
        TerminaBloquear
        Exit Sub
    End If

    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Then  '1= Insertar
        TerminaBloquear
        Exit Sub
    End If
    
    If Data2.Recordset.EOF Then
        TerminaBloquear
        Exit Sub
    End If
    
    
    'Comprobar recargo financiero
    '
    If vParamAplic.ArticuloRecargoFinanciero <> "" Then
        'Tiene recargo financiero
        'Veremos si la factura tienen la linea con recargo financiero
        vWhere = ObtenerWhereCP(False)
        vWhere = vWhere & " and codartic = '" & vParamAplic.ArticuloRecargoFinanciero & "' AND 1"
        
        vWhere = DevuelveDesdeBD(conAri, "count(*)", "slifac", vWhere, "1")
        
        If Val(vWhere) > 0 Then
            'Tienen linea recargo financiero
            'Veremos si el cliente la tienen ahora
            vWhere = DevuelveDesdeBD(conAri, "RecargoFinanciero", "sclien", "codclien", CStr(Data1.Recordset!codClien))
            If vWhere = "" Then vWhere = "0"
            If Val(vWhere) = 0 Then
                vWhere = "La factura tiene recargo financiero y el cliente ahora no lo tiene. Continuar?"
                If MsgBox(vWhere, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
            End If
        End If
    End If
    
    
    
    vWhere = ObtenerWhereCP(False)
    vWhere = vWhere & " AND codtipoa='" & Data3.Recordset.Fields!Codtipoa & "' AND numalbar=" & Data3.Recordset.Fields!Numalbar
    vWhere = vWhere & " and numlinea=" & Data2.Recordset!numlinea
    If Not BloqueaRegistro(NomTablaLineas, vWhere) Then
        TerminaBloquear
        Exit Sub
    End If


    


    'INSERTA LOG
    '-------------------------------------------------
    Set LOG = New cLOG
    If EstaEnTesoreria <> "" Then EstaEnTesoreria = "Tesoreria: " & EstaEnTesoreria
    EstaEnTesoreria = "     Alb : " & Data2.Recordset!Numalbar & "   Linea: " & Data2.Recordset!numlinea & vbCrLf & EstaEnTesoreria
    EstaEnTesoreria = "Pulsa mod linea: " & Text1(1).Text & Text1(0).Text & " " & Text1(2).Text & EstaEnTesoreria
    LOG.Insertar 8, vUsu, EstaEnTesoreria
    Set LOG = Nothing


    If Text1(16).Text = "" Then Text1(16).Text = Format(Data1.Recordset!impdtopp, FormatoDescuento)
    If Text1(17).Text = "" Then Text1(17).Text = Format(Data1.Recordset!impdtogr, FormatoDescuento)
    
    
    If Me.Text1(1).Text = "FAG" Then
        'Es para cmabiar el consumo
        CadenaDesdeOtroForm = ""
        vWhere = ObtenerWhereCP(False) & " AND numlinea"
        anc = DevuelveDesdeBD(conAri, "cantidad", "slifac", vWhere, "20")
        
        vWhere = Replace(Data3.Recordset!observa1, "  ", "@") & "|"
        
        vWhere = vWhere & Replace(Data3.Recordset!observa2, " ", "@") & "|" & Val(anc) & "|"
        vWhere = Replace(vWhere, "@", "     ")
        'Añadimos el select
        vWhere = vWhere & ObtenerWhereCP(False) & "|"
        frmListado5.OpcionListado = 8
        frmListado5.OtrosDatos = vWhere
        frmListado5.Show vbModal
        
        If CadenaDesdeOtroForm <> "" Then
            CalcularDatosFactura
            TerminaBloquear
            ModificarFactura
            PosicionarData
            FormatoDatosTotales
            J = Data3.Recordset.AbsolutePosition
            PonerCamposLineas
            SituarDataPosicion Data3, CLng(J), ""
            
        End If
        Exit Sub
    End If

    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        
        If DataGrid1.Bookmark < DataGrid1.FirstRow Then
            J = 0
        Else
            J = DataGrid1.Bookmark - DataGrid1.FirstRow
        End If
        DataGrid1.Scroll 0, J
        DataGrid1.Refresh
    End If
    
'    anc = ObtenerAlto(Me.DataGrid1)
    anc = DataGrid1.Top
    If DataGrid1.Row < 0 Then
        anc = anc + 210
    Else
        anc = anc + DataGrid1.RowTop(DataGrid1.Row) + 10
    End If

    For J = 0 To 2
        txtAux(J).Text = DataGrid1.Columns(J + 5).Text
    Next J
    Text2(16).Text = DataGrid1.Columns(J + 5).Text
    
    'cantidad
    J = 4
    txtAux(J - 1).Text = DataGrid1.Columns(J + 5).Text
    'num bultos
    J = 5
    txtAux(11).Text = DataGrid1.Columns(J + 5).Text
    
    J = 4
    For J = J + 1 To 9
        txtAux(J - 1).Text = DataGrid1.Columns(J + 6).Text
    Next J
    
    If vParamAplic.NumeroInstalacion = 2 Then
        
        
    Else
        'Para todas las demas..
        txtAux(9).Text = DataGrid1.Columns(16).Text
        If vEmpresa.TieneAnalitica Then
            Me.txtAux(9).Text = DBLet(Data2.Recordset!CodCCost, "T")
            Me.txtAux2(9).Text = PonerNombreCCoste(Me.txtAux(9))
        Else
        
            txtAux2(9).Text = DataGrid1.Columns(17).Text
        End If
    End If
    'num lote
    txtAux(10).Text = DataGrid1.Columns(19).Text
    
    ModificaLineas = 2 'Modificar
    LLamaLineas ModificaLineas, anc, "DataGrid1"
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    Me.lblIndicador.Caption = "MODIFICAR"
    PonerBotonCabecera False
    BloquearTxt Text2(16), False 'Campo Ampliacion Linea
    PonerFoco txtAux(4)
    Me.DataGrid1.Enabled = False
    
    
    'Si es de varios desbloqueo el nomartic por si se han equivocado

    vWhere = DevuelveDesdeBD(conAri, "artvario", "sartic", "codartic", txtAux(1).Text, "T")
    txtAux(2).visible = False
    If vWhere = "1" Then
        txtAux(2).Height = txtAux(4).Height
        txtAux(2).Top = txtAux(4).Top
        txtAux(2).visible = True
    End If
    BloquearTxt txtAux(2), vWhere <> "1"

EModificarLinea:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub LLamaLineas(xModo As Byte, Optional alto As Single, Optional grid As String)
Dim jj As Integer
Dim B As Boolean

    Select Case grid
        Case "DataGrid1"
            DeseleccionaGrid Me.DataGrid1
            'PonerModo xModo + 1
    
            B = (xModo = 1 Or xModo = 2) 'Insertar o Modificar Lineas
            If xModo = 2 And TipoModificacionFactura >= 2 Then B = False
            
            For jj = 0 To txtAux.Count - 1
                If jj = 4 Or (jj >= 6 And jj <= 10) Then
                    txtAux(jj).Height = DataGrid1.RowHeight
                    txtAux(jj).Top = alto
                    txtAux(jj).visible = B
                End If
            Next jj
            
            
           
            
            cmdAux.Top = alto
            cmdAux.visible = B
            txtAux(2).visible = False  'Por si acso
            
            If vParamAplic.NumeroInstalacion = 2 Then
                txtAux(9).visible = False  'Por si acso
                cmdAux.visible = False
            End If
        Case "DataGrid2"
            DeseleccionaGrid Me.DataGrid2
            B = (xModo = 1)
             For jj = 0 To txtAux3.Count - 1
                txtAux3(jj).Height = DataGrid2.RowHeight
                txtAux3(jj).Top = alto
                txtAux3(jj).visible = B
            Next jj
    End Select
End Sub


Private Sub BotonEliminar()
'Eliminar Registro de la Cabecera: Tabla de Facturas (scafac)
' y los registros correspondientes de las tablas cab. albaranes (scafac1)
' y las lineas de la factura (slifac)
Dim Cad As String
Dim EstaEnTesoreria As String
Dim EliminarElApunte As String
'Dim vTipoMov As CTiposMov

    On Error GoTo EEliminar

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
    'solo se puede modificar la factura si no esta contabilizada
    If FactContabilizada(EstaEnTesoreria) Then Exit Sub
    
    Cad = "E L I M I N A R" & vbCrLf
    Cad = Cad & String(40, "=") & vbCrLf & String(40, "=") & vbCrLf & vbCrLf & vbCrLf
    Cad = Cad & "Va a eliminar la Factura:            "
    Cad = Cad & vbCrLf & "Tipo:  " & Text1(1).Text
    Cad = Cad & vbCrLf & "Nº Fact.:  " & Format(Text1(0).Text, "0000000")
    Cad = Cad & vbCrLf & "Fecha:  " & Format(Text1(2).Text, "dd/mm/yyyy") & vbCrLf
    Cad = Cad & vbCrLf & String(40, "=") & vbCrLf & String(40, "=") & vbCrLf
    Cad = Cad & vbCrLf & vbCrLf & " ¿Desea continuar con el borre de la factura? "

    'Borramos
    If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
'        On Error GoTo EEliminar
        Screen.MousePointer = vbHourglass
        NumRegElim = Data1.Recordset.AbsolutePosition
'        NumPedElim = Data1.Recordset.Fields(1).Value
        CodTipoMov = Text1(1).Text
        
        If Not Eliminar() Then
            Screen.MousePointer = vbDefault
            Exit Sub
        Else
        
            
            Set LOG = New cLOG
            LOG.Insertar 8, vUsu, "Factura eliminada: " & Text1(1).Text & Text1(0).Text & " " & Text1(2).Text & vbCrLf & EstaEnTesoreria
            Set LOG = Nothing
        
            If SituarDataTrasEliminar(Data1, NumRegElim) Then
                PonerCampos
            Else
                LimpiarCampos
                'Poner los grid sin apuntar a nada
                LimpiarDataGrids
                PonerModo 0
            End If
        End If
'        'Devolvemos contador, si no estamos actualizando
'        Set vTipoMov = New CTiposMov
'        vTipoMov.DevolverContador CodTipoMov, NumPedElim
'        Set vTipoMov = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
    
EEliminar:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Eliminar Albaran", Err.Description
End Sub


Private Sub cmdCarnetManipulador_Click()
    
    If Modo <> 2 Then Exit Sub
    If Data1.Recordset!codtipom <> "FAV" Then MsgBox "No debe indicar carnet en tipo factura <> FAV", vbExclamation
    
    BuscaChekc = DevuelveDesdeBD(conAri, "clivario", "sclien", "codclien", Text1(4).Text)
    
    CadenaDesdeOtroForm = ""
    If BuscaChekc = "1" Then
        'Es cliente VARIOS, con lo cual lanzamos el mto de clientes varios
         BuscaChekc = ""
         Set frmCliV = New frmFacClientesV
         frmCliV.vNif = Text1(6).Text
         frmCliV.DatosADevolverBusqueda = "0|"
         frmCliV.Show vbModal
         Set frmCliV = Nothing
         
         If BuscaChekc = "" Then Exit Sub
         
         If BuscaChekc <> Text1(6).Text Then
                
            If vUsu.Nivel > 0 Then
                MsgBox "No es el NIF de la venta", vbExclamation
                Exit Sub
            Else
                If MsgBox("No es el NIF de la venta." & vbCrLf & "¿Continuar?", vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
            End If
         
         End If
                
         CadenaDesdeOtroForm = " concat(coalesce(ManipuladorNumCarnet,''),'|',date_format(coalesce(fcaducidad,now()),'%d/%m/%Y'),'|'"
         CadenaDesdeOtroForm = CadenaDesdeOtroForm & ",coalesce(nomclien,''),'|',if(coalesce(ManipuladortipoCarnet,1)=1,'Basico','CUALIFICADO'),'|') as devuelve"
         BuscaChekc = DevuelveDesdeBD(conAri, CadenaDesdeOtroForm, "sclvar", "nifclien", BuscaChekc, "T")
         CadenaDesdeOtroForm = ""
         If RecuperaValor(BuscaChekc, 1) = "" Then
              MsgBox "No tienen Nº de carnet", vbExclamation
         Else
             CadenaDesdeOtroForm = BuscaChekc
          End If
            
        
        
    Else
        'Adasda
        frmFitoCarnet.Cliente = Val(Text1(4).Text)
        frmFitoCarnet.Show vbModal
    End If
    
    If CadenaDesdeOtroForm <> "" Then
    
        'Va a actualizar el carnet de manipulador
        BuscaChekc = "Vaa actualizar los datos del carnet de manipulador: " & vbCrLf
        BuscaChekc = BuscaChekc & "Nº carnet: " & RecuperaValor(CadenaDesdeOtroForm, 1) & vbCrLf
        BuscaChekc = BuscaChekc & "Caducidad: " & RecuperaValor(CadenaDesdeOtroForm, 2) & vbCrLf
        BuscaChekc = BuscaChekc & "Nombre: " & RecuperaValor(CadenaDesdeOtroForm, 3) & vbCrLf
        BuscaChekc = BuscaChekc & "Tipo. " & RecuperaValor(CadenaDesdeOtroForm, 4) & vbCrLf & "¿Continuar?"
        If MsgBox(BuscaChekc, vbQuestion + vbYesNoCancel) = vbYes Then
    
            If CDate(RecuperaValor(CadenaDesdeOtroForm, 2)) < CDate(Text1(2).Text) Then
                If MsgBox("Carnet caducado.  ¿Desea continuar?", vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
            End If
            
            
            BuscaChekc = "UPDATE scafac1 SET ManipuladorNumCarnet =" & DBSet(RecuperaValor(CadenaDesdeOtroForm, 1), "T")
            BuscaChekc = BuscaChekc & ", ManipuladorFecCaducidad = " & DBSet(RecuperaValor(CadenaDesdeOtroForm, 2), "F")
            BuscaChekc = BuscaChekc & ", ManipuladorNombre = " & DBSet(RecuperaValor(CadenaDesdeOtroForm, 3), "T")
            BuscaChekc = BuscaChekc & ", TipoCarnet = " & IIf(UCase(RecuperaValor(CadenaDesdeOtroForm, 4)) = "CUALIFICADO", 2, 1)
            BuscaChekc = BuscaChekc & " WHERE codtipom = " & DBSet(Data1.Recordset!codtipom, "T") & " AND numfactu = " & DBSet(Data1.Recordset!Numfactu, "N")
            BuscaChekc = BuscaChekc & " AND fecfactu = " & DBSet(Data1.Recordset!FecFactu, "F") & " AND codtipoa = " & DBSet(Data3.Recordset!Codtipoa, "T")
            BuscaChekc = BuscaChekc & " AND numalbar = " & DBSet(Data3.Recordset!Numalbar, "N")
            If ejecutar(BuscaChekc, False) Then
                Me.Text3(19).Text = RecuperaValor(CadenaDesdeOtroForm, 1)
                Me.Text3(20).Text = RecuperaValor(CadenaDesdeOtroForm, 3)
                Me.Text3(21).Text = RecuperaValor(CadenaDesdeOtroForm, 2)
                Text3(22).Text = IIf(UCase(Text2(0).Text) = "CUALIFICADO", 2, 1)
                Text2(4).Text = RecuperaValor(CadenaDesdeOtroForm, 4)
                
            End If
        End If
    End If
    
    BuscaChekc = ""
    
End Sub

Private Sub cmdInsertaAlbaranCostes_Click()
    If Modo <> 2 Then Exit Sub
    
    'PResentaremos ualbaranes que cumplan losmrequisitos.
    BuscaChekc = ""
    
    'Mismo cliente.   Sin lineas de albaran.
    
    Set frmAlb = New frmBasico2
    AyudaFacturasAnyadeAlbaranCostes frmAlb, Text1(4).Text & "|" & Text1(5).Text & "|", False
    Set frmAlb = Nothing
    
    
    
    If BuscaChekc <> "" Then
            'Vamos con las comprobaciones,
            ' 1 Si tiene costes los sum
            ' 2 Si tiene horas, las sumo
            ' 3 Msg si quiere insertar
            Screen.MousePointer = vbHourglass
            If HacerAnyadirAlbaranCosteEuler Then
                
                
                    Espera 0.2
                    TerminaBloquear
                    PosicionarData
                    FormatoDatosTotales
                    NumRegElim = Data3.Recordset.AbsolutePosition
                    PonerCamposLineas
                    SituarDataPosicion Data3, NumRegElim, ""

                
                
            End If
            Screen.MousePointer = vbDefault
        
            
    
    End If
    BuscaChekc = ""
    NumRegElim = 0
End Sub

Private Sub cmdLineasCostes_Click(Index As Integer)
Dim Tipo As Byte
Dim Aux As String

    If Modo <> 2 Then Exit Sub
    If Data1.Recordset Is Nothing Then Exit Sub
    'If Data1.Recordset!codtipom = "FAV" Then Exit Sub
    If Data1.Recordset.EOF Then Exit Sub
    
    If Index > 0 Then
        If lwCostes.ListItems.Count = 0 Then Exit Sub
        If lwCostes.SelectedItem Is Nothing Then Exit Sub
        
        If Trim(lwCostes.SelectedItem.Text) = "" Then
            MsgBox "Esta linea es de totales", vbExclamation
            Exit Sub
        End If
        
    End If
    
    If Index = 2 Then
        'ELIMINAR
        BuscaChekc = lwCostes.ColumnHeaders(1).Text & ": " & lwCostes.SelectedItem.Text & vbCrLf
        For NumRegElim = 2 To lwCostes.ColumnHeaders.Count
            If Trim(lwCostes.SelectedItem.SubItems(NumRegElim - 1)) <> "" Then
                BuscaChekc = BuscaChekc & lwCostes.ColumnHeaders(NumRegElim).Text & ":   " & lwCostes.SelectedItem.SubItems(NumRegElim - 1) & vbCrLf
            End If
        Next
        
        BuscaChekc = "Eliminar el coste: " & vbCrLf & BuscaChekc & "¿Continuar?"
        If MsgBox(BuscaChekc, vbQuestion + vbYesNoCancel) = vbYes Then
            BuscaChekc = "DELETE FROM slifac_eu WHERE " & CadenaWhereCostes
            conn.Execute BuscaChekc
        
            PonerCamposFicha True
        End If
        
    Else
        BuscaChekc = ""
        If Index = 1 Then
            If lwCostes.SelectedItem.Text = "HOR" Then
                'HORAS TRABAJADAS
                BuscaChekc = "0"
                Aux = "HORAS TRABAJADAS"
            Else
                If lwCostes.SelectedItem.Text = "VEH" Then
                    BuscaChekc = "3"
                    Aux = "Gastos vehiculo"
                Else
                    If lwCostes.SelectedItem.Text = "ALB" Or lwCostes.SelectedItem.Text = "MAT" Then
                        BuscaChekc = "1"
                    Else
                        'Proveedor
                        BuscaChekc = "2"
                    End If
                    Aux = lwCostes.SelectedItem.SubItems(4)
                End If
            End If
            BuscaChekc = BuscaChekc & lwCostes.SelectedItem.SubItems(3) & "|"
            BuscaChekc = BuscaChekc & lwCostes.SelectedItem.ListSubItems(7).Tag & "|"
            BuscaChekc = BuscaChekc & Aux & "|"
            BuscaChekc = BuscaChekc & lwCostes.SelectedItem.SubItems(5) & "|"
            BuscaChekc = BuscaChekc & lwCostes.SelectedItem.SubItems(6) & "|"
            BuscaChekc = BuscaChekc & lwCostes.SelectedItem.SubItems(7) & "|"
            
            
        
        End If
        CadenaDesdeOtroForm = ""
        frmListado3.Opcion = 70
        frmListado3.OtrosDatos = BuscaChekc
        frmListado3.Show vbModal
        If CadenaDesdeOtroForm <> "" Then
            Aux = Mid(CadenaDesdeOtroForm, 1, 1)
            Tipo = CByte(Val(Aux))
            CadenaDesdeOtroForm = Mid(CadenaDesdeOtroForm, 2)
            If Index = 0 Then
                
                Aux = ObtenerWhereCP(False)
                Aux = Aux & " AND codtipoa='" & Data3.Recordset.Fields!Codtipoa & "' "
                Aux = Aux & " AND numalbar=" & Data3.Recordset.Fields!Numalbar & " AND 1"
                Aux = DevuelveDesdeBD(conAri, "Max(numlinea)", "slifac_eu", Aux, "1")
                BuscaChekc = Val(Aux) + 1
                
                'slifac_eu(codtipom,numfactu,fecfactu,codtipoa,numalbar,numlinea,fechamov,codalmac,codartic,nomartic,cantidad,precioar,Tipo)
                Aux = "('" & Data2.Recordset!codtipom & "'," & Data1.Recordset!Numfactu & "," & DBSet(Data1.Recordset!FecFactu, "F") & ",'"
                Aux = Aux & Data3.Recordset!Codtipoa & "'," & Data3.Recordset!Numalbar & "," & BuscaChekc & ","
                If Tipo = 0 Then
                    Aux = Aux & DBSet(Data3.Recordset!FecFactu, "F")
                Else
                    Aux = Aux & DBSet(RecuperaValor(CadenaDesdeOtroForm, 1), "F")
                End If
                Aux = Aux & ",1," & DBSet(RecuperaValor(CadenaDesdeOtroForm, 2), "T", "N") & ","
                Aux = Aux & DBSet(RecuperaValor(CadenaDesdeOtroForm, 3), "T") & ","
                Aux = Aux & DBSet(RecuperaValor(CadenaDesdeOtroForm, 4), "T") & ","
                Aux = Aux & DBSet(RecuperaValor(CadenaDesdeOtroForm, 5), "T") & "," & IIf(Tipo = 2, 4, IIf(Tipo = 3, 5, Tipo)) & ")"
                
                BuscaChekc = "INSERT INTO slifac_eu(codtipom,numfactu,fecfactu,codtipoa,numalbar,numlinea,fechamov,codalmac,codartic,nomartic,cantidad,precioar,Tipo) VALUES " & Aux
                
            Else
                
                BuscaChekc = "UPDATE slifac_eu SET  cantidad = " & DBSet(RecuperaValor(CadenaDesdeOtroForm, 4), "N")
                BuscaChekc = BuscaChekc & ", precioar =" & DBSet(RecuperaValor(CadenaDesdeOtroForm, 5), "N")
                BuscaChekc = BuscaChekc & ", nomartic =" & DBSet(RecuperaValor(CadenaDesdeOtroForm, 3), "T")
                If Tipo > 0 Then
                    BuscaChekc = BuscaChekc & ", codartic =" & DBSet(RecuperaValor(CadenaDesdeOtroForm, 2), "T")
                Else
                    BuscaChekc = BuscaChekc & ", fechamov =" & DBSet(RecuperaValor(CadenaDesdeOtroForm, 1), "F")
                End If
                BuscaChekc = BuscaChekc & " WHERE " & CadenaWhereCostes
            End If
            conn.Execute BuscaChekc
        
            PonerCamposFicha True
        End If
        
    End If
    BuscaChekc = ""
End Sub


Private Function CadenaWhereCostes() As String
    
        CadenaWhereCostes = "codtipom ='" & RecuperaValor(lwCostes.SelectedItem.Tag, 1)
        CadenaWhereCostes = CadenaWhereCostes & "' AND numfactu =" & RecuperaValor(lwCostes.SelectedItem.Tag, 2)
        CadenaWhereCostes = CadenaWhereCostes & " AND fecfactu =" & DBSet(RecuperaValor(lwCostes.SelectedItem.Tag, 3), "F")
        CadenaWhereCostes = CadenaWhereCostes & " AND codtipoa ='" & RecuperaValor(lwCostes.SelectedItem.Tag, 4)
        CadenaWhereCostes = CadenaWhereCostes & "' AND numalbar =" & RecuperaValor(lwCostes.SelectedItem.Tag, 5)
        CadenaWhereCostes = CadenaWhereCostes & " AND numlinea =" & RecuperaValor(lwCostes.SelectedItem.Tag, 6)
        CadenaWhereCostes = CadenaWhereCostes & " AND tipo =" & RecuperaValor(lwCostes.SelectedItem.Tag, 7)
        
      
End Function

Private Sub cmdLineasImpresion_Click(Index As Integer)
    If Modo <> 2 Then Exit Sub
    If Data1.Recordset.EOF Then Exit Sub
    
    
    If Index > 0 Then
        If lwEulerLineas.ListItems.Count = 0 Then
            MsgBox "Ningun dato", vbExclamation
            Exit Sub
        End If
        If Index < 3 Then
            'Modificar eliminar.
            'el seleccionado
            If Me.lwEulerLineas.SelectedItem Is Nothing Then
                MsgBox "Seleccione una linea", vbExclamation
                Exit Sub
            End If
        End If
    Else
        If Data2.Recordset.EOF Then Exit Sub
    End If
    CadenaDesdeOtroForm = ""
    
    If Index < 2 Then
        'nuevo modificar
        If Index = 1 Then
            CadenaDesdeOtroForm = Mid(lwEulerLineas.SelectedItem.Key, 2, 3)
            frmListado5.OtrosDatos = Data1.Recordset!codtipom & "|" & Data1.Recordset!Numfactu & "|" & Data1.Recordset!FecFactu & "|" & lwEulerLineas.SelectedItem.Tag
        Else
            CadenaDesdeOtroForm = ""  '"" = nuevo   id= linea
            frmListado5.OtrosDatos = Data1.Recordset!codtipom & "|" & Data1.Recordset!Numfactu & "|" & Data1.Recordset!FecFactu & "|" & Data3.Recordset!Codtipoa & "|" & Data3.Recordset!Numalbar & "|"
        End If
        frmListado5.OpcionListado = 27
        frmListado5.Show vbModal
        
    
    Else
        If Index = 2 Then
            'Eliminar
            BuscaChekc = "Va a eliminar linea impresion" & vbCrLf & "Articulo : " & Me.lwEulerLineas.SelectedItem.Text & vbCrLf
            BuscaChekc = BuscaChekc & "Descripcion : " & Me.lwEulerLineas.SelectedItem.SubItems(1) & vbCrLf
            BuscaChekc = BuscaChekc & "Importe : " & Me.lwEulerLineas.SelectedItem.SubItems(4) & vbCrLf
            If MsgBox(BuscaChekc, vbQuestion + vbYesNoCancel) = vbYes Then
                BuscaChekc = " WHERE codtipom='" & Data1.Recordset!codtipom & "' AND numfactu = " & Data1.Recordset!Numfactu
                BuscaChekc = BuscaChekc & " AND fecfactu = " & DBSet(Data1.Recordset!FecFactu, "F")
                BuscaChekc = BuscaChekc & " AND codtipoa = '" & RecuperaValor(lwEulerLineas.SelectedItem.Tag, 1) & "' AND numalbar = " & RecuperaValor(lwEulerLineas.SelectedItem.Tag, 2)
                BuscaChekc = "DELETE FROM slifac_eu2 " & BuscaChekc & " AND numlinea= " & Mid(Me.lwEulerLineas.SelectedItem.Key, 2, 3)
                If ejecutar(BuscaChekc, False) Then CadenaDesdeOtroForm = "OK"
            End If
            BuscaChekc = ""
        Else
            'imprimir
            If lwEulerLineas.Tag <> "" Then
                MsgBox lwEulerLineas.Tag, vbExclamation
            Else
                BotonImprimir 89
            End If
        End If
    End If
    
    If CadenaDesdeOtroForm <> "" Then PonerCamposFicha True
    
    
    
End Sub

Private Sub cmdMtoCampos_Click(Index As Integer)
    If Index = 0 Then
        'Añadir mas campos
            CadenaDesdeOtroForm = ""
            frmADVvarios.Opcion = 0
            frmADVvarios.vCampos = Text1(4).Text
            frmADVvarios.Show vbModal
            If CadenaDesdeOtroForm <> "" Then
                
                
                    
                MultiInsercionCampos
                
                'Cargamos GRID
                
            End If
            CargaDatosCampos
                
    Else
        BuscaChekc = ""
        If Me.ListView1.ListItems.Count > 0 Then
            If Not Me.ListView1.SelectedItem Is Nothing Then
                BuscaChekc = "Va a eliminar el campo: "
                BuscaChekc = BuscaChekc & vbCrLf & "Codigo : " & Me.ListView1.SelectedItem.Text
                BuscaChekc = BuscaChekc & vbCrLf & "Partida : " & Me.ListView1.SelectedItem.SubItems(1)
                BuscaChekc = BuscaChekc & vbCrLf & "Variedad : " & Me.ListView1.SelectedItem.SubItems(2)
                BuscaChekc = BuscaChekc & vbCrLf & vbCrLf & "¿Continuar?"
                If MsgBox(BuscaChekc, vbQuestion + vbYesNo) = vbYes Then
                    'El tag tiene codcampo
                    BuscaChekc = "DELETE FROM slifaccampos WHERE  codtipom = " & DBSet(Data1.Recordset!codtipom, "T")
                    BuscaChekc = BuscaChekc & " AND numfactu = " & Data1.Recordset!Numfactu
                    BuscaChekc = BuscaChekc & " AND fecfactu = " & DBSet(Data1.Recordset!FecFactu, "F")
                    'De momento dejamos borrar sin ver el albaran
                    'BuscaChekc = BuscaChekc & " AND codtipoa = " & DBSet(data3.Recordset!codtipoa, "T")
                    'BuscaChekc = BuscaChekc & " AND numalbar = " & DBSet(data3.Recordset!NumAlbar, "N")
                    BuscaChekc = BuscaChekc & " AND codcampo  = " & CStr(Val(Me.ListView1.SelectedItem.Text))
                    conn.Execute BuscaChekc
                    
                    Me.ListView1.ListItems.Remove Me.ListView1.SelectedItem.Index
    
                End If
            End If
        End If
    End If
End Sub

Private Sub cmdObserva3_Click()
    If Modo <> 2 And Modo <> 4 And Modo <> 1 Then Exit Sub
    If cmdObserva3.Tag = "" Then cmdObserva3.Tag = "0"
    cmdObserva3.Tag = cmdObserva3.Tag + 1
    
    'Campos, pero SI no hay parametros..
    
    If cmdObserva3.Tag = 2 Then
        If Not SolapaCamposFito Then
            If vParamAplic.TieneTelefonia2 > 0 Then
                cmdObserva3.Tag = 3
            Else
                If vParamAplic.NumeroInstalacion = vbEuler Then
                    cmdObserva3.Tag = 4
                Else
                    cmdObserva3.Tag = 0
                End If
            End If
        End If
    ElseIf cmdObserva3.Tag = 3 Then
         If Not vParamAplic.TieneTelefonia2 > 0 Then cmdObserva3.Tag = 0
         
    ElseIf cmdObserva3.Tag = 4 Then
        If Not InstalacionEsEulerTaxco Then cmdObserva3.Tag = 0

    End If
    If cmdObserva3.Tag >= 5 Then cmdObserva3.Tag = 0
    
    
    
    VisualizarPorTipoAlbaran False
    
    
End Sub



Private Sub VisualizarPorTipoAlbaran(DesdeInicioBusqueda As Boolean)
Dim Codtipoa As String

    Me.DataGrid1.visible = cmdObserva3.Tag = 0
    Me.FrameObserva.visible = cmdObserva3.Tag = 1
    Me.FrameCampos.visible = cmdObserva3.Tag = 2
    Me.FrameTelefonia.visible = cmdObserva3.Tag = 3
    Me.FrameEuler.visible = cmdObserva3.Tag = 4
    
    If Not InstalacionEsEulerTaxco Then
        FrameALE.visible = False
        FrameTAXCO.visible = False
    Else
        If Modo = 2 Or DesdeInicioBusqueda Then
            Codtipoa = "ALO"
            If Not Data3.Recordset.EOF Then Codtipoa = Data3.Recordset!Codtipoa
                
            
            If vParamAplic.NumeroInstalacion = vbEuler Then
                FrameALE.visible = Codtipoa = "ALE" 'Or Data3.Recordset!codtipoa = "ALO"
            ElseIf vParamAplic.NumeroInstalacion = vbTaxco Then
                FrameALE.visible = Codtipoa = "ALE" 'Or Data3.Recordset!codtipoa = "ALO"
                FrameTAXCO.visible = Codtipoa = "ALO"
            End If
            FrameReparEuler.visible = Codtipoa = "ALR"
            
            If FrameEuler.visible Then FrameEuler.Enabled = FrameReparEuler.visible
            
            
        End If
    End If
    If cmdObserva3.Tag = 4 Then
        
        
    End If
    Select Case (Me.cmdObserva3.Tag)
    Case 1
        If vParamAplic.Ariagro <> "" Then
            Me.cmdObserva3.ToolTipText = "Ver campos asociados  "
            Me.cmdObserva3.Picture = frmPpal.imgListComun.ListImages(48).Picture
            
        Else
            If vParamAplic.TieneTelefonia2 > 0 Then
                Me.cmdObserva3.ToolTipText = "Ver datos telefono "
                Me.cmdObserva3.Picture = frmPpal.imgListComun.ListImages(49).Picture
            
            Else
                If InstalacionEsEulerTaxco Then
                    Me.cmdObserva3.ToolTipText = "Orden trabajo / trabajo exterior"
                    Me.cmdObserva3.Picture = frmPpal.imgListComun.ListImages(26).Picture
                Else
                    Me.cmdObserva3.Picture = frmPpal.imgListComun.ListImages(18).Picture
                    Me.cmdObserva3.ToolTipText = "  lineas de factura"
                End If
            End If
            
        End If
        

        BloqueaText3
        

    Case 2
        If Not vParamAplic.TieneTelefonia2 > 0 Then
            Me.cmdObserva3.Picture = frmPpal.imgListComun.ListImages(18).Picture
    '        CargarICO Me.cmdObserva, "message.ico"
            Me.cmdObserva3.ToolTipText = "lineas de factura"
        Else
            Me.cmdObserva3.ToolTipText = "Ver datos telefono "
            Me.cmdObserva3.Picture = frmPpal.imgListComun.ListImages(49).Picture
        End If
        
        
    Case 3
        
        Me.cmdObserva3.Picture = frmPpal.imgListComun.ListImages(18).Picture
'        CargarICO Me.cmdObserva, "message.ico"
        Me.cmdObserva3.ToolTipText = "lineas de factura"
               
    Case 4
        Me.cmdObserva3.Picture = frmPpal.imgListComun.ListImages(18).Picture
        Me.cmdObserva3.ToolTipText = "lineas de factura"
    Case Else 'el cero
        
        Me.cmdObserva3.Picture = frmPpal.imgListComun.ListImages(41).Picture
'        CargarICO Me.cmdObserva, "message.ico"
        Me.cmdObserva3.ToolTipText = "ver observaciones albaran"
        
        
        
    End Select
    SSTab1_Click 0


End Sub



Private Sub BloqueaText3()
Dim i As Byte
Dim B As Boolean
    
'       tipoModificacionFactura As Byte
'       0.- Se puede TODO. NO esta contabilizada
'       1.- Se puede "parcial", lo que hacia antes
'       2.- Solo se puede trabajador, zona,referencias.....
'       3.- NADA. Fuera ejercicios



    'bloquear los Text3 que son las lineas de scafac1
    ' Trab trab prpa material  envio
    B = Modo <> 4 And Modo <> 1
    For i = 0 To 3
        BloquearTxt Text3(i), B
    Next i
    
    'Datos direccion envio
    If vParamAplic.DireccionesEnvio Then BloquearTxt Text3(18), B
    
    Me.chkEnvio.Enabled = Not B
    If Not B Then
        If Modo <> 1 Then B = vUsu.Nivel > 0
    End If
    chkPedxCli.Enabled = Not B
    
    'Observaciones
    For i = 9 To 13
        BloquearTxt Text3(i), (Modo <> 4 And Modo <> 1)
    Next i

    
    
    
    
    If InstalacionEsEulerTaxco Then
        For i = 23 To 27
            BloquearTxt Text3(i), (Modo <> 4 And Modo <> 1)
        Next i
    End If
    'Referencia cliente    referencias facelec
    BloquearTxt Text3(16), B
    BloquearTxt Text3(28), B
    BloquearTxt Text3(29), B
    
    
    'ofer fecofe ped  fecped smanaentreg
    B = Modo <> 1
    For i = 4 To 8
        BloquearTxt Text3(i), B
    Next i
    
    'datos venta TPV
    BloquearTxt Text3(14), B
    BloquearTxt Text3(15), B
    'BloquearTxt Text3(17), B
 
    For i = 17 To 22
        'eL 18 ya lo hace arriba
        'If i <> 18 Then BloquearTxt Text3(i), B
        BloquearTxt Text3(i), B
    Next i

    
 
End Sub


Private Sub cmdRegresar_Click()
'Este es el boton Cabecera
Dim Cad As String

    'Quitar lineas y volver a la cabecera
    If Modo >= 5 Then  'modo 5: Mantenimientos Lineas
        If Modo = 6 Then
            Me.cmdMtoCampos(0).visible = False
            Me.cmdMtoCampos(1).visible = False
        End If
        PonerModo 2
        DataGrid2.Enabled = True
        If Not Data1.Recordset.EOF Then _
            Me.lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount


        

    Else 'Se llama desde algún Prismatico de otro Form al Mantenimiento de Trabajadores
        If Data1.Recordset.EOF Then
            MsgBox "Ningún registro devuelto.", vbExclamation
            Exit Sub
        End If
        Cad = Data1.Recordset.Fields(0) & "|"
        Cad = Cad & Data1.Recordset.Fields(1) & "|"
        RaiseEvent DatoSeleccionado(Cad)
        Unload Me
    End If
End Sub


Private Sub cmdReparEuler_Click(Index As Integer)
    If Modo <> 2 Then Exit Sub
    CadenaDesdeOtroForm = ObtenerWhereCP(True)
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & " AND codtipoa='" & Data3.Recordset.Fields!Codtipoa & "' "
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & " AND numalbar=" & Data3.Recordset.Fields!Numalbar
    frmFacEulerDatosRep.Buscar = False
    frmFacEulerDatosRep.Show vbModal
    
End Sub

Private Sub Combo1_Click()
    If Modo = 1 Then
        If vParamAplic.NumeroInstalacion = vbTaxco Then
            cmdObserva3.Tag = 4
            FrameTAXCO.Enabled = True
            Me.SSTab1.Tab = 1
            VisualizarPorTipoAlbaran True
            Me.FrameEuler.Enabled = True
        End If
    End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub DataGrid1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Ayuda de Etiqueta de precio de salida de la Función de Precios
On Error Resume Next

    If Data2.Recordset.EOF Then Exit Sub
    If (Modo = 2) Or (Modo = 5 And ModificaLineas = 0) Then
        Me.DataGrid1.ToolTipText = ""
        If X > 7790 And X < 8170 Then
            Select Case DataGrid1.Columns(11).Value
                Case "P": Me.DataGrid1.ToolTipText = "P: Promoción"
                Case "E": Me.DataGrid1.ToolTipText = "E: Precio Especial"
                Case "T": Me.DataGrid1.ToolTipText = "T: Tarifa Artículo"
                Case "A": Me.DataGrid1.ToolTipText = "A: Precio Artículo"
                Case "M": Me.DataGrid1.ToolTipText = "M: Manual"
'                Case Else
'                    Me.DataGrid1.ToolTipText = ""
            End Select
'        Else
'            Me.DataGrid1.ToolTipText = ""
        End If
    End If
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error GoTo Error1

    If Not Data2.Recordset.EOF Then
        If ModificaLineas <> 1 Then
            Text2(16).Text = DBLet(Data2.Recordset.Fields!Ampliaci, "T")
            If vEmpresa.TieneAnalitica Then
                '- centro de coste
                ' ---- [19/10/2009] [LAURA]: añadir campo centro de coste familia
                Me.txtAux(9).Text = DBLet(Data2.Recordset!CodCCost, "T")
                Me.txtAux2(9).Text = PonerNombreCCoste(Me.txtAux(9))
            Else
                txtAux2(9).Text = DBLet(Data2.Recordset.Fields!nomprove, "T")
            End If
        End If
    Else
        Text2(16).Text = ""
        txtAux2(9).Text = ""
    End If
    
    Exit Sub

Error1:
    MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub DataGrid2_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim i As Byte

    If Not Data3.Recordset.EOF Then
        'Trabajador Albaran
        Text3(0).Text = Data3.Recordset.Fields!CodTraba
        Text3_LostFocus (0)
        'Trabajador pedido
        Text3(1).Text = DBLet(Data3.Recordset.Fields!CodTrab1, "T")
        Text3_LostFocus (1)
        'Trab. Prepara Material
        Text3(2).Text = Data3.Recordset.Fields!codtrab2
        Text3_LostFocus (2)
        Text3(3).Text = Data3.Recordset.Fields!CodEnvio
        Text3_LostFocus (3)
        
        'oferta
        Text3(4).Text = DBLet(Data3.Recordset.Fields!NumOfert, "N")
        If Text3(4).Text <> "0" Then
            FormateaCampo Text3(4)
        Else
            Text3(4).Text = ""
        End If
        Text3(5).Text = DBLet(Data3.Recordset.Fields!fecofert, "F")
        'pedido
        Text3(6).Text = DBLet(Data3.Recordset.Fields!NumPedcl, "N")
        If Text3(6).Text <> "0" Then
            FormateaCampo Text3(6)
        Else
            Text3(6).Text = ""
        End If
        Text3(7).Text = DBLet(Data3.Recordset.Fields!fecpedcl, "F")
        If Text3(7).Text <> "" Then FormateaCampo Text3(7)
        Text3(8).Text = DBLet(Data3.Recordset.Fields!sementre, "N")
        If Text3(8).Text = "0" Then Text3(8).Text = ""
        'venta
        Text3(15).Text = DBLet(Data3.Recordset.Fields!NumTermi, "N")
        Text3(14).Text = DBLet(Data3.Recordset.Fields!NumVenta, "N")
        FormateaCampo Text3(14)
'        If Text3(14).Text = "0" Then Text3(14).Text = ""
'        If Text3(15).Text = "0" Then Text3(15).Text = ""
        
        'Observaciones
        Text3(9).Text = DBLet(Data3.Recordset.Fields!observa1, "T")
        Text3(10).Text = DBLet(Data3.Recordset.Fields!observa2, "T")
        Text3(11).Text = DBLet(Data3.Recordset.Fields!observa3, "T")
        Text3(12).Text = DBLet(Data3.Recordset.Fields!observa4, "T")
        Text3(13).Text = DBLet(Data3.Recordset.Fields!observa5, "T")
        
        
        Text3(16).Text = DBLet(Data3.Recordset.Fields!referenc, "T")
        Text3(17).Text = DBLet(Data3.Recordset.Fields!FecEnvio, "F")
        
        
        If vParamAplic.DireccionesEnvio Then
            Text3(18).Text = DBLet(Data3.Recordset.Fields!coddiren, "F")
            If Text3(18).Text = "0" Then Text3(18).Text = ""
            Text3_LostFocus 18
        End If
        
        chkEnvio.Value = DBLet(Data3.Recordset!docarchiv, "N")
        chkPedxCli.Value = DBLet(Data3.Recordset!PideCliente, "N")
        
        'EULER
        If InstalacionEsEulerTaxco Then
            VisualizarPorTipoAlbaran False
            'Recepcion mercancia
            For i = 23 To 27
                Text3(i).Text = DBLet(Data3.Recordset.Fields(i + 7), "T")
                
                If i = 23 And Text3(i).Text <> "" Then Text3(i).Text = Format(Data3.Recordset.Fields(i + 7), "dd/mm/yyyy hh:nn:ss")
                If i = 26 And Text3(i).Text <> "" Then Text3(i).Text = Format(Data3.Recordset.Fields(i + 7), "#0.00000")
                If i = 27 And Text3(i).Text <> "" Then Text3(i).Text = Format(Data3.Recordset.Fields(i + 7), "#0.00000")
            Next
            
            PonerImagenFirma
            
        End If
        
        
        'Si lleva fitosanitarios
        Text2(4).Text = ""
        If SolapaCamposFito Then
            'ManipuladorNumCarnet,ManipuladorFecCaducidad,ManipuladorNombre,TipoCarnet
            Text3(19).Text = DBLet(Data3.Recordset!ManipuladorNumCarnet, "T")
            Text3(20).Text = DBLet(Data3.Recordset!ManipuladorNombre, "T")
            Text3(21).Text = ""
            Text3(22).Text = ""
            
            If DBLet(DBLet(Data3.Recordset!ManipuladorFecCaducidad, "T")) <> "" Then Text3(21).Text = Format(Data3.Recordset!ManipuladorFecCaducidad, "dd/mm/yyyy")
            If Val(DBLet(Data3.Recordset!TipoCarnet, "N")) > 0 Then
                Text3(22).Text = Data3.Recordset!TipoCarnet
                Text2(4).Text = IIf(Val(Data3.Recordset!TipoCarnet) = 2, "Cualificado", "Básico")
            End If
        End If
        
        
        Text3(28).Text = DBLet(Data3.Recordset!refcontrato, "T")
        Text3(29).Text = DBLet(Data3.Recordset!refpedido, "T")
    
        
        
        
        'Datos de la tabla slipre
        CargaGrid DataGrid1, Data2, True
    Else
        For i = 0 To Text3.Count - 1
            Text3(i).Text = ""
        Next i
        For i = 0 To 4
            Text2(i).Text = ""
        Next i
        Text2(18).Text = "" 'nomdirenvio
        chkEnvio.Value = 0
        chkPedxCli.Value = 0
        'Datos de la tabla slipre
        CargaGrid DataGrid1, Data2, False
    End If
    
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    
    'Viene de DblClick en frmAlmMovimArticulos y carga el form con los valores
    If UnaVez Then
        UnaVez = False
        If hcoCodMovim <> "" Then
            If Data1.Recordset.EOF Then
                PonerCadenaBusqueda
            Else
                PonerCampos
                DataGrid1_RowColChange 0, 0
            End If
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub Form_Load()
Dim RefeGrande As Boolean
Dim B1 As Boolean
    UnaVez = True
    'Icono del formulario
    Me.Icon = frmPpal.Icon
    
     'Icono de busqueda
    For kCampo = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(kCampo).Picture = frmPpal.imgListComun.ListImages(19).Picture
    Next kCampo



    



    ' ICONITOS DE LA BARRA
    btnPrimero = 23
    With Me.Toolbar1
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 1   'Botón Buscar
        .Buttons(2).Image = 2   'Botón Todos
        .Buttons(5).Image = 4   'Modificar
        .Buttons(6).Image = 5   'Borrar
        .Buttons(9).Image = 10 'Mto Lineas
        .Buttons(10).Image = 16 'Imprimir
        .Buttons(11).Image = 40 'Imprimir albaran
        .Buttons(13).Image = 43 'Asignar Numeros de lote
        
        .Buttons(14).Image = 48 'Campos
        .Buttons(15).Image = 45 'Tipo precio
        
        .Buttons(16).Image = 51 'Modificar fecha - Deshacer factura8llevar a albarnes
        .Buttons(18).Image = 31 'Valoracion
        .Buttons(19).Image = 54  'SIGNOTEC
        
        
        .Buttons(21).Image = 15  'Salir
        .Buttons(btnPrimero).Image = 6  'Primero
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
        .Buttons(btnPrimero + 3).Image = 9 'Último
    End With
    
    'Antes Octubre 2015
    'Toolbar1.Buttons(14).visible = vParamAplic.Ariagro <> ""
    'mnEditarCampos.visible = vParamAplic.Ariagro <> ""
    'Ahora
    SolapaCamposFito = vParamAplic.Ariagro <> "" Or vParamAplic.ManipuladorFitosanitarios2
    Toolbar1.Buttons(14).visible = SolapaCamposFito
    mnEditarCampos.visible = SolapaCamposFito
    
    If SolapaCamposFito Then
        cmdCarnetManipulador.visible = vUsu.Nivel = 0
        cmdMtoCampos(0).visible = vParamAplic.Ariagro <> ""
        cmdMtoCampos(1).visible = vParamAplic.Ariagro <> ""
        ListView1.visible = vParamAplic.Ariagro <> ""
        
        FrameCamposMani.visible = vParamAplic.ManipuladorFitosanitarios2
        If Not vParamAplic.ManipuladorFitosanitarios2 Then
            Me.FrameCampos.Caption = ""
            cmdMtoCampos(0).Left = 240
            cmdMtoCampos(1).Left = 240
            ListView1.Left = 640
        Else
            Me.FrameCampos.Caption = "Manipulador"
            FrameCamposMani.BorderStyle = 0
        End If
        If vParamAplic.Ariagro <> "" Then
            If Me.FrameCampos.Caption <> "" Then Me.FrameCampos.Caption = Me.FrameCampos.Caption & " / "
            Me.FrameCampos.Caption = Me.FrameCampos.Caption & "Campos"
        End If
    End If
    
    
    'El boton de imprimir campos lo usare en euler para los costes
    Me.SSTab1.TabVisible(2) = False
    Me.SSTab1.TabVisible(3) = False
    If InstalacionEsEulerTaxco Then
        Toolbar1.Buttons(14).Image = 36
        
        If vParamAplic.NumeroInstalacion = vbEuler Then
        
            Toolbar1.Buttons(14).visible = True
            Toolbar1.Buttons(14).ToolTipText = "Costes"
            Me.SSTab1.TabVisible(2) = True
            Me.SSTab1.TabVisible(3) = True 'vparamaplic.NumeroInstalacion
            Me.cmdLineasCostes(0).Picture = frmPpal.imgListComun.ListImages(3).Picture
            Me.cmdLineasCostes(1).Picture = frmPpal.imgListComun.ListImages(4).Picture
            Me.cmdLineasCostes(2).Picture = frmPpal.imgListComun.ListImages(14).Picture
            Me.cmdLineasImpresion(0).Picture = frmPpal.imgListComun.ListImages(3).Picture
            Me.cmdLineasImpresion(1).Picture = frmPpal.imgListComun.ListImages(4).Picture
            Me.cmdLineasImpresion(2).Picture = frmPpal.imgListComun.ListImages(14).Picture
            Me.cmdLineasImpresion(3).Picture = frmPpal.imgListComun.ListImages(40).Picture
        End If
        For kCampo = 9 To 13
            Text3(kCampo).Left = 240
            Text3(kCampo).Width = 7305
        Next
        FrameRecepMercan.BorderStyle = 0
        FrameRecepMercan.visible = True
        PrimeraVez = True
        CarpetaImagenesEULER = DevuelveDesdeBD(conAri, "pathDocs", "eulerparam", "1", "1")
        PonerImagenFirma
        PrimeraVez = False
        
         
        If vParamAplic.NumeroInstalacion = vbEuler Then
            FrameALE.Left = 240
        Else
            FrameTAXCO.BorderStyle = 0
            FrameTAXCO.visible = True
        
            lblSerie.visible = True
            lblSerie.Caption = ""
        End If
        
    End If
    
    EsDeVarios = False
    If vUsu.Nivel = 0 Then EsDeVarios = vParamAplic.GrabaModificarPrecioAlaBaja
    Toolbar1.Buttons(15).visible = EsDeVarios
    mnTipoPreciosLinea.visible = EsDeVarios
    
    
    Toolbar1.Buttons(16).visible = vUsu.Nivel = 0
    
    Toolbar1.Buttons(19).visible = vParamAplic.PathFirmasAlbaran <> "" Or vParamAplic.PathFirmasFacturas <> ""
    
    Me.SSTab1.Tab = 0
    LimpiarCampos   'Limpia los campos TextBox
    CargaCombo
    
    
    'cargar icono de observaciones de los albaranes de factura
    Me.cmdObserva3.Picture = frmPpal.imgListComun.ListImages(41).Picture
    cmdObserva3.Tag = 0
'    CargarICO Me.cmdObserva, "message.ico"
    Me.FrameObserva.visible = False
    Me.cmdObserva3.ToolTipText = "ver observaciones albaran"
    FrameALE.BorderStyle = 0
    
    VieneDeBuscar = False
    
    'Comprobar si es Departamento o Direccion
    Me.Label1(1).Caption = DevuelveTextoDepto(True)
    
    'Direcion envio SOLO si esta en parametros
    Label1(48).visible = vParamAplic.DireccionesEnvio
    imgBuscar(10).visible = vParamAplic.DireccionesEnvio
    Text3(18).visible = vParamAplic.DireccionesEnvio
    Text2(18).visible = vParamAplic.DireccionesEnvio
        
        
    Me.Label1(45).visible = vParamAplic.ctaAportacion <> ""
    Text1(45).visible = vParamAplic.ctaAportacion <> ""
        
        
    If vEmpresa.TieneAnalitica Then
        txtAux(9).Tag = "Cod. centro coste|T|S|||slifac|codccost|||"
        Label1(46).Caption = "Centro coste"
    Else
        
        B1 = False
        If vParamAplic.NumeroInstalacion = 2 Then If vUsu.Nivel = 0 Then B1 = True
        
        If B1 Then
            txtAux(9).Tag = "Cod. Proveedor|N|N|||slifac|comisionagente|#0.00||"
        Else
            txtAux(9).Tag = "Cod. Proveedor|N|N|||slifac|codprovex|0||"
        End If
        
        
        Label1(46).Caption = "Proveedor"
    End If
        
        
    'FECHA ENVIO.
    'Sera fechaliqu para SAIL
    'Fecha liq.
    If vParamAplic.TipoFormularioClientes = 0 Then
        Label1(47).Caption = "F. envio"
        
    Else
        'SAIL
        Label1(47).Caption = "F. liquid."
        FrameReparEuler.BorderStyle = 0
        imgBuscarEULER.Picture = frmPpal.imgListComun.ListImages(19).Picture
        
    End If
        
    
        
    'Referencia cliente
    RefeGrande = True
    If vParamAplic.NumeroInstalacion = 0 Then
        RefeGrande = False
    Else
        If vParamAplic.NumeroInstalacion = 3 Or vParamAplic.NumeroInstalacion = 2 Then RefeGrande = False
    End If
    If RefeGrande Then
        Text3(16).Width = IIf(vParamAplic.ctaAportacion <> "", 3765, 5325)
        Text3(16).MaxLength = 255
    End If
        
        
    '## A mano
    NombreTabla = "scafac"
    NomTablaLineas = "slifac" 'Tabla lineas de Facturacion
    Ordenacion = " ORDER BY scafac.codtipom, scafac.numfactu, scafac.fecfactu "
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    'ASignamos un SQL al DATA1
    Dim T1 As Single
    T1 = Timer
    CadenaConsulta = "Select * from " & NombreTabla
    If hcoCodMovim <> "" Then
        'Se llama desde Dobleclick en frmAlmMovimArticulos
        'como tenemos aqui el nº de albaran, buscar a que factura corresponde
        'en la scafac1
        CadenaConsulta = CadenaConsulta & ObtenerSelFactura
'        CadenaConsulta = CadenaConsulta & " WHERE codtipom='" & hcoCodTipoM & "' AND numalbar= " & hcoCodMovim
    Else
        'CadenaConsulta = CadenaConsulta & " where numfactu=-1"
        'Cambio sugerido por Msoler
        'CadenaConsulta = CadenaConsulta & " WHERE codtipom is null and numfactu is null and fecfactu is null "
        CadenaConsulta = CadenaConsulta & " WHERE false"
    End If
    
    Data1.ConnectionString = conn
    Data1.RecordSource = CadenaConsulta
    
    'ANTE
    'If hcoCodMovim <> "" Then Data1.Refresh
    Data1.Refresh
    
    
    'Cargar el DataGrid de lineas de Revisiones inicialmente a nada DATA2
    PrimeraVez = True
    
    If hcoCodMovim = "" Then
        If DatosADevolverBusqueda = "" Then
            PonerModo 0
        Else
            BotonBuscar
        End If
'        CargaGrid DataGrid1, Data2, False
        'Poner los grid sin apuntar a nada
        LimpiarDataGrids
        PrimeraVez = False
    Else
        If Data1.Recordset.EOF Then
            PonerModo 0
        Else
            PonerModo 2
            
          
            
        End If
    End If
    
End Sub


Private Sub LimpiarCampos()
On Error Resume Next
    limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador.Caption = ""
    'Aqui va el especifico de cada form es
    '### a mano
    Me.Check1.Value = 0
    Check2.Value = 0 'facturae
    Me.Combo1.ListIndex = -1
    chkEnvio.Value = 0
    chkPedxCli.Value = 0
    
    If vParamAplic.Ariagro <> "" Then Me.ListView1.ListItems.Clear
    If vParamAplic.TieneTelefonia2 > 0 Then
        Me.ListView2.ListItems.Clear
        Me.ListView3.ListItems.Clear
    End If
    If InstalacionEsEulerTaxco Then
        lwCostes.ListItems.Clear
        lwEulerLineas.ListItems.Clear
        imgFirmaRecep.visible = False
    End If
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
End Sub


Private Sub frmA_DatoSeleccionado(CadenaSeleccion As String)
'Form Agentes
Dim Indice As Byte
    Indice = 14
    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod agente
    FormateaCampo Text1(Indice)
    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom agente
End Sub

Private Sub frmAlb_DatoSeleccionado(CadenaSeleccion As String)
   BuscaChekc = CadenaSeleccion
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
Dim cadB As String
Dim Aux As String
      
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        If EsCabecera2 = 0 Then 'Llama desde VerTodos del Form
            cadB = ""
            Aux = ValorDevueltoFormGrid(Text1(1), CadenaDevuelta, 1)
            cadB = Aux
            Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 2)
            cadB = cadB & " and " & Aux
            Aux = ValorDevueltoFormGrid(Text1(2), CadenaDevuelta, 3)
            cadB = cadB & " and " & Aux
            CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
        Else
            If EsCabecera2 = 1 Then
                'Llama desde Prismatico Direcciones/Departamentos
                Text1(12).Text = Format(RecuperaValor(CadenaDevuelta, 1), "000")
                Text1(13).Text = RecuperaValor(CadenaDevuelta, 2)
            Else
                Text3(18).Text = Format(RecuperaValor(CadenaDevuelta, 1), "000")
                Text2(18).Text = RecuperaValor(CadenaDevuelta, 2)
            End If
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub frmC_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Clientes
    Text1(4).Text = RecuperaValor(CadenaSeleccion, 1)  'Cod Clien
End Sub


Private Sub frmCliV_DatoSeleccionado(CadenaSeleccion As String)
    BuscaChekc = RecuperaValor(CadenaSeleccion, 1)
End Sub

Private Sub frmCP_DatoSeleccionado(CadenaSeleccion As String)
'Formulario Mantenimiento C. Postales
Dim Indice As Byte
Dim devuelve As String

    Indice = 9
    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'CPostal
    Text1(Indice + 1).Text = ObtenerPoblacion(Text1(Indice).Text, devuelve)  'Poblacion
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


Private Sub frmFE_DatoSeleccionado(CadenaSeleccion As String)
'Formas de Envio
Dim Indice As Byte
    Indice = 3
    Text3(Indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000") 'Cod Forma Envio
    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Forma Envio
End Sub

Private Sub frmFP_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Formas de Pago
Dim Indice As Byte
    Indice = 15
    Text1(Indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000") 'Cod Forma Pago
    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Forma Pago
End Sub


Private Sub frmP_DatoSeleccionado(CadenaSeleccion As String)
    txtAux(9).Text = RecuperaValor(CadenaSeleccion, 1)
    txtAux2(9).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmT_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Trabajadores
Dim Indice As Byte
    Indice = Val(Me.imgBuscar(3).Tag)
    Text1(Indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000") 'Cod Trabajador
    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Trabajador
End Sub


Private Sub imgBuscar_Click(Index As Integer)
Dim Indice As Byte

    If Modo = 0 Then Exit Sub
    If Modo = 2 And Index <> 11 Then Exit Sub
    
    Screen.MousePointer = vbHourglass

    Select Case Index
        Case 0 'Cod. Cliente
            PonerFoco Text1(4)
            
            Set frmC = New frmBasico2
            AyudaClientes frmC, Text1(4).Text
            Set frmC = Nothing

            
            Indice = 5
            PonerFoco Text1(Indice)
            
        Case 1 'NIF para cliente de Varios
'            Set frmCV = New frmFacClientesV
'            frmCV.DatosADevolverBusqueda = "0"
'            frmCV.Show vbModal
'            Set frmCV = Nothing
            Indice = 6
            Set frmCV = New frmBasico2
            AyudaClientesV frmCV, Text1(Indice)
            Set frmCV = Nothing
            
            PonerFoco Text1(Indice)
            
        Case 2 'Cod. Postal
            Set frmCP = New frmCPostal
            frmCP.DatosADevolverBusqueda = "0"
            frmCP.Show vbModal
            Set frmCP = Nothing
            Indice = 9
            VieneDeBuscar = True
            PonerFoco Text1(Indice)
        
        Case 3 'Cod. Direc.
             'Mostrar las Direc. o Dptos del cliente seleccionado
             If Trim(Text1(4).Text) = "" Then
                MsgBox "Debe seleccionar un cliente.", vbInformation
                Screen.MousePointer = vbDefault
                Exit Sub
             Else
                EsCabecera2 = 1
                MandaBusquedaPrevia " codclien= " & Val(Text1(4).Text)
                Indice = 12
             End If
             PonerFoco Text1(Indice)
             
        Case 4 'Agente
            Indice = 14
            PonerFoco Text1(Indice)
'            Set frmA = New frmFacAgentesCom
'            frmA.DatosADevolverBusqueda = "0"
'            frmA.Show vbModal
            Set frmA = New frmBasico2
            AyudaAgentesComerciales frmA, Text1(Indice), , True
            Set frmA = Nothing
            
         Case 5 'Forma de Pago
            Indice = 15
'            PonerFoco Text1(Indice)
'            Set frmFP = New frmFacFormasPago
'            frmFP.DatosADevolverBusqueda = "0"
'            frmFP.Show vbModal
            Set frmFP = New frmBasico2
            AyudaFormasPago frmFP, Text1(Indice)
            Set frmFP = Nothing
            PonerFoco Text1(Indice)
            
        Case 6, 7, 8 'Realizada Por Trabajador (Pedido, Albaran, Preparador Material
            Indice = Index - 6
            Me.imgBuscar(3).Tag = Indice
'            Set frmT = New frmAdmTrabajadores
'            frmT.DatosADevolverBusqueda = "0"
'            frmT.Show vbModal
            Set frmT = New frmBasico2
            AyudaTrabajadores frmT, Text3(Indice)
            Set frmT = Nothing

            PonerFoco Text3(Indice)
       
        Case 9 'Cod Envio
            Indice = 3
            PonerFoco Text3(Indice)
            Set frmFE = New frmFacFormasEnvio
            frmFE.DatosADevolverBusqueda = "0"
            frmFE.Show vbModal
            Set frmFE = Nothing
            PonerFoco Text3(Indice)
            
            
        Case 10
             'Direcciones envio
             If Trim(Text1(4).Text) = "" Then
                MsgBox "Debe seleccionar un cliente.", vbInformation
                Screen.MousePointer = vbDefault
                Exit Sub
             Else
                EsCabecera2 = 2
                MandaBusquedaPrevia " codclien= " & Val(Text1(4).Text)
                
             End If
             PonerFoco Text3(18)
             
        Case 11
                If Modo = 0 Then Exit Sub
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
    End Select
    
    Screen.MousePointer = vbDefault
End Sub


Private Sub imgBuscarEULER_Click()
Dim Aux As String
    If Modo <> 1 Then Exit Sub
    
    
    
    CadenaDesdeOtroForm = ""
    frmFacEulerDatosRep.Buscar = True
    frmFacEulerDatosRep.Show vbModal
    If CadenaDesdeOtroForm <> "" Then
        
        Aux = " (scafac.codtipom,scafac.numfactu,scafac.fecfactu) IN (" & CadenaDesdeOtroForm & ")"
        If chkVistaPrevia = 1 Then
            EsCabecera2 = 0
            MandaBusquedaPrevia Aux
        ElseIf Aux <> "" Then
            'Se muestran en el mismo form
    '        cadB = cadB & " and scafac.codtipom='" & CodTipoMov & "'" 'Solo seleccionamos los del Movimiento, aqui los FAV
            CadenaConsulta = "select scafac.* from " & NombreTabla & " INNER JOIN scafac1 ON scafac.codtipom=scafac1.codtipom AND scafac.numfactu=scafac1.numfactu AND scafac.fecfactu=scafac1.fecfactu "
            CadenaConsulta = CadenaConsulta & " WHERE " & Aux & " GROUP BY scafac.codtipom,scafac.numfactu,scafac.fecfactu " & Ordenacion
            PonerCadenaBusqueda
        End If
        
        
        VisualizarPorTipoAlbaran False
        
        
    End If
    
End Sub



Private Sub imgFirmaRecep_Click()
    If Modo <> 2 Then Exit Sub
    If imgFirmaRecep.Tag = "" Then Exit Sub
    
    LanzaVisorMimeDocumento Me.hwnd, imgFirmaRecep.Tag
End Sub

Private Sub imgGeolocalizacion_Click()
    If Modo <> 2 Then Exit Sub
     If Text3(26).Text <> "" And Text3(27).Text <> "" Then
        BuscaChekc = TransformaComasPuntos(Text3(26).Text) & "," & TransformaComasPuntos(Text3(27).Text)
        AbrirGeolocalizacion BuscaChekc
        
        BuscaChekc = ""
    End If
End Sub

Private Sub lwCostes_DblClick()
Dim C1 As String
Dim prov As String
Dim i As Integer

    If lwCostes.SelectedItem Is Nothing Then Exit Sub
    
    If lwCostes.SelectedItem.Text <> "PRO" Then Exit Sub
    
    C1 = lwCostes.SelectedItem.SubItems(1)
    i = InStr(1, C1, "(")
    If i > 0 Then
        C1 = Mid(C1, i + 1)
        i = InStr(1, C1, ")")
        If i > 0 Then prov = Mid(C1, 1, i - 1)
    End If
    
    If i = 0 Then
        MsgBox "Imposible encontrar proveedor", vbExclamation
        Exit Sub
    End If
    
    C1 = lwCostes.SelectedItem.SubItems(2)
    i = 0
    If Mid(C1, 1, 3) = "ALC" Then
        'Buscaremos por albaran
        
        C1 = Trim(Mid(C1, 5))
        C1 = "numalbar = " & DBSet(C1, "T") & " AND fechaalb =" & DBSet(lwCostes.SelectedItem.SubItems(3), "F") & " AND codprove "
        C1 = DevuelveDesdeBD(conAri, "concat(numalbar,'|',fechaalb)", "scaalp", C1, prov)
        If C1 <> "" Then
            'Esta todavia en albaranes
            C1 = C1 & "|" & prov & "|"
        
        Else
            'Veamos si esta facturado
            C1 = Trim(Mid(lwCostes.SelectedItem.SubItems(2), 5))
            C1 = "numalbar = " & DBSet(C1, "T") & " AND fechaalb =" & DBSet(lwCostes.SelectedItem.SubItems(3), "F") & " AND codprove "
            C1 = DevuelveDesdeBD(conAri, "concat(numalbar,'|',fechaalb)", "scafpa", C1, prov)
            If C1 = "" Then
                MsgBox "Imposible localizar albaran compra de factura: " & lwCostes.SelectedItem.SubItems(2), vbExclamation
                Exit Sub
            End If
            C1 = C1 & "|" & prov & "|"
            i = 2 'No hace falta qwue busque la factura, para despues sacar el albaran. YA lo tengo
                
        End If
    Else
        i = 1
    End If
    
    If i = 1 Then
        'Buscamos la factura
        C1 = Trim(Mid(C1, 5))
        C1 = "numfactu = " & DBSet(C1, "T") & " AND fecfactu =" & DBSet(lwCostes.SelectedItem.SubItems(3), "F") & " AND codprove "
        C1 = DevuelveDesdeBD(conAri, "concat(numalbar,'|',fechaalb)", "scafpa", C1, prov)
        If C1 = "" Then
            MsgBox "Imposible localizar albaran compra de factura: " & lwCostes.SelectedItem.SubItems(2), vbExclamation
            Exit Sub
        End If
        C1 = C1 & "|" & prov & "|"
    End If
    If i = 0 Then
    
      'IT.Tag = "numalbar =" & DBSet(Rs!NUmAlbar, "T") & " AND  fechaalb =" & DBSet(Rs!FechaAlb, "F") & " AND codprove =" & Rs!Codprove
       With frmComEntAlbaranSA
            .hcoCodMovim = RecuperaValor(C1, 1)
            .hcoFechaMovim = RecuperaValor(C1, 2)
            .hcoCodProve = RecuperaValor(C1, 3)
            .EsHistorico = False
            .Show vbModal
        End With
    
    Else
        
        
         With frmComHcoFacturSA
            .hcoCodMovim = RecuperaValor(C1, 1)
            .hcoFechaMovim = RecuperaValor(C1, 2)
            .hcoCodProve = RecuperaValor(C1, 3)
            .Show vbModal
        End With
    End If
    
    
End Sub


Private Sub lwEulerLineas_DblClick()
    cmdLineasImpresion_Click 1
End Sub

Private Sub mnBuscar_Click()
    Me.SSTab1.Tab = 0
    BotonBuscar
End Sub


Private Sub mnEditarCampos_Click()
    
    If Modo <> 2 Then Exit Sub
    
    If Val(Me.cmdObserva3.Tag) <> 2 Then
        cmdObserva3.Tag = 1
        cmdObserva3_Click
    End If
    
    If Val(Me.cmdObserva3.Tag) <> 2 Then
        MsgBox "Visualice los campos", vbExclamation
        Exit Sub
    End If
    
    
    
    
        If BLOQUEADesdeFormulario(Me) Then
            PonerModo 6
            PonerBotonCabecera True
            
            Me.cmdMtoCampos(0).visible = True
            Me.cmdMtoCampos(1).visible = True
        End If
    
End Sub

Private Sub mnEliminar_Click()
    If Modo = 5 Then 'Eliminar lineas de Pedido
'         BotonEliminarLinea
    Else   'Eliminar Pedido
    
        If vParamAplic.NumeroInstalacion = vbFenollar Then
           
                MsgBox "Reestablezca el albaran", vbExclamation
                Exit Sub
            
        End If
        BotonEliminar
    End If
End Sub


Private Sub mnImprimir_Click()
Dim Indice As Byte

'Imprimir Factura
    
    
    If Data1.Recordset.EOF Then Exit Sub
    
    If Data1.Recordset!codtipom = "FTI" Then 'ticket de venta del TPV
        BotonImprimirTicket
    Else
        If EsFraTelefono Then
            ImprimirFraTelefonia
        
        Else
            If CInt(DBLet(Data3.Recordset!NumTermi, "N")) > 0 Then
                'Es factura del TPV
                BotonImprimir 63
            Else
                'Impresion normal
                Indice = 53  '53: Informe de Facturas
                If vParamAplic.NumeroInstalacion = vbTaxco And Data3.Recordset!Codtipoa = "ALO" Then
                    If MsgBox("¿ Impresion extendida ?", vbQuestion + vbYesNoCancel) = vbYes Then Indice = 94
                End If
                BotonImprimir (Indice)
            End If
        End If
    End If
End Sub

Private Function EsFraTelefono() As Boolean
    EsFraTelefono = False
    
    If Data1.Recordset!codtipom = "FAT" Then
        If vParamAplic.TieneTelefonia2 = 1 Then EsFraTelefono = True             'ALZIRA
    ElseIf Data1.Recordset!codtipom = "FAI" Then
        If vParamAplic.TieneTelefonia2 > 0 Then EsFraTelefono = Me.ListView2.ListItems.Count > 0
    End If
    
End Function


Private Sub mnImprimirAlbaran_Click()
Dim Seguir As Boolean
Dim TipoA As String
    If Me.Data1.Recordset.EOF Then Exit Sub
    If Me.Data3.Recordset.EOF Then Exit Sub
    
    
    'Albaranes que no se pueden montar
    Seguir = False
    If Not IsNull(Data3.Recordset!Codtipoa) Then
        If Data3.Recordset!Codtipoa <> "" Then
            TipoA = CStr(Data3.Recordset!Codtipoa)
            If TipoA = "FTI" Or TipoA = "ALM" Then
                Seguir = False
            Else
                Seguir = True
            End If
        End If
    End If
    If Not Seguir Then
        MsgBox "No se puede imprimir el albaran seleccionado", vbExclamation
        Exit Sub
    End If
    
    
    
    If Val(Data3.Recordset!Numalbar) = 0 Then
        MsgBox "No se puede imprimir el albaran seleccionado", vbExclamation
        Exit Sub
    End If
    
    
    If Data2.Recordset.EOF Then
        MsgBox "Albaran no tiene lineas", vbExclamation
        Exit Sub
    End If
    
    ImprimirAlbaran 1
    
    
End Sub

Private Sub mnLineas_Click()
    BotonMtoLineas 1, "Facturas"
End Sub


Private Sub mnModificar_Click()

    If vUsu.Nivel > 1 Then
        MsgBox "No tiene permiso para realizar la accion", vbExclamation
        Exit Sub
    End If

    If Modo = 5 Then 'Modificar lineas
        'bloquea la tabla cabecera de factura: scafac
        If BLOQUEADesdeFormulario(Me) Then
            'bloquear la tabla cabecera de albaranes de la factura: scafac1
            If BloqueaAlbxFac Then
                If BloqueaLineasFac Then BotonModificarLinea
            End If
        End If
         
    Else   'Modificar Pedido
        'bloquea la tabla cabecera de factura: scafac
        If BLOQUEADesdeFormulario(Me) Then
            'bloquear la tabla cabecera de albaranes de la factura: scafac1
            If BloqueaAlbxFac Then BotonModificar
        End If
    End If
End Sub


Private Function BloqueaAlbxFac() As Boolean
'bloquea todos los albaranes de la factura
Dim Sql As String

    On Error GoTo EBloqueaAlb
    
    BloqueaAlbxFac = False
    'bloquear cabecera albaranes x factura
    Sql = "select * FROM scafac1 "
    Sql = Sql & ObtenerWhereCP(True) & " FOR UPDATE"
    conn.Execute Sql, , adCmdText
    BloqueaAlbxFac = True

EBloqueaAlb:
    If Err.Number <> 0 Then BloqueaAlbxFac = False
End Function


Private Function BloqueaLineasFac() As Boolean
'bloquea todas las lineas de la factura
Dim Sql As String

    On Error GoTo EBloqueaLin

    BloqueaLineasFac = False
    'bloquear cabecera albaranes x factura
    Sql = "select * FROM slifac "
    Sql = Sql & ObtenerWhereCP(True) & " FOR UPDATE"
    conn.Execute Sql, , adCmdText
    BloqueaLineasFac = True

EBloqueaLin:
    If Err.Number <> 0 Then BloqueaLineasFac = False
End Function


Private Sub mnModLotes_Click()
    
    'Si no es EOF.... bla bla bla
    SSTab1.Tab = 1
    
    
    If Data1.Recordset.EOF Then Exit Sub
    If Data2.Recordset.EOF Then Exit Sub
    If Data3.Recordset.EOF Then Exit Sub
    
    
    'Si no es fra venta... salimos
    If Text1(1).Text <> "FAV" And Text1(1).Text <> "FTI" Then
        MsgBox "Movimiento debe ser FAV/FTI", vbExclamation
        Exit Sub
    End If
    
    If DBLet(Data3.Recordset!Codtipoa, "T") = "" Then
        MsgBox "Tipo albaran incorrecto", vbExclamation
        Exit Sub
    End If
    
    If vParamAplic.ManipuladorFitosanitarios2 Then
        ''Llamaremos a la funcion de carga de numeros de lote
        
    Else
        HacerNumerosLote_
    End If
    'Cargamos lineas otra vez
    CargaGrid DataGrid1, Data2, True
End Sub
    
Private Sub HacerNumerosLote_()
Dim vWhere As String

    On Error GoTo EPedirNLotes
        
        
    'aqui aqui aqui aqui aqui aqui###
    DescargarDatosTMPNumLotes "tmpnlotes", "codusu = " & vUsu.Codigo
    
    
    
    
    vWhere = ObtenerWhereCP(True)
    vWhere = vWhere & " AND codtipoa='" & Data3.Recordset.Fields!Codtipoa & "' "
    vWhere = vWhere & " AND numalbar=" & Data3.Recordset.Fields!Numalbar
    vWhere = " FROM slifac " & vWhere
    'tmpnlotes codusu,numalbar,fechaalb,codprove,numlinea,codartic,codalmac,nomartic,cantidad,numlotes
    vWhere = ",numlinea, codArtic, codAlmac, NomArtic, Cantidad, numlote " & vWhere
    
    vWhere = "Select " & vUsu.Codigo & "," & DBSet(Data3.Recordset!Numalbar, "N") & "," & DBSet(Data3.Recordset!FechaAlb, "F") & "," & DBSet(Data2.Recordset!Numfactu, "N") & vWhere
    
    vWhere = "INSERT INTO tmpnlotes(codusu,numalbar,fechaalb,codprove,numlinea,codartic,codalmac,nomartic,cantidad,numlotes) " & vWhere
    
    conn.Execute vWhere
    
    
    
        Set frmNLote = New frmAlmCargarNLote
        'EN esta cadena ira para el SQL
        vWhere = ObtenerWhereCP(True)
        vWhere = vWhere & " AND codtipoa='" & Data3.Recordset.Fields!Codtipoa & "' "
        vWhere = vWhere & " AND numalbar=" & Data3.Recordset.Fields!Numalbar
        frmNLote.Desde2 = vWhere
        'Para el select del frm
        vWhere = "numalbar=" & DBSet(Data3.Recordset!Numalbar, "N") & " AND fechaalb=" & DBSet(Data3.Recordset!FechaAlb, "F") & " AND codprove=" & DBSet(Data2.Recordset!Numfactu, "N")
        frmNLote.parSelSQL = vWhere
        frmNLote.Show vbModal
        Set frmNLote = Nothing
        
        
     'Eliminar de la tabla temporal tmpnlotes los lotes introducidos
    DescargarDatosTMPNumLotes "tmpnlotes", "codusu = " & vUsu.Codigo
        
EPedirNLotes:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
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


Private Sub mnTipoPreciosLinea_Click()
     If Modo <> 2 Then Exit Sub
     If vUsu.Nivel > 0 Then Exit Sub 'por si acaso
     If Data1.Recordset.EOF Then Exit Sub
     Screen.MousePointer = vbHourglass
     BuscaChekc = "Factura: " & Me.Data1.Recordset!codtipom & Format(Me.Data1.Recordset!Numfactu, "000000") & " de " & Format(Me.Data1.Recordset!FecFactu, "dd/mm/yyyy") & "|"
     BuscaChekc = BuscaChekc & "codtipom='" & Data1.Recordset!codtipom & "' AND numfactu="
     BuscaChekc = BuscaChekc & Data1.Recordset!Numfactu & " AND fecfactu=" & DBSet(Data1.Recordset!FecFactu, "F") & "|"
     
     frmListado4.vCadena = BuscaChekc
     frmListado4.Opcion = 6
     frmListado4.Show vbModal
     CargaGrid DataGrid1, Data2, True
     BuscaChekc = ""
     
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub


Private Sub SSTab1_Click(PreviousTab As Integer)
    Me.Label1(35).visible = Me.SSTab1.Tab = 1 And Me.DataGrid1.visible
    Me.Text2(16).visible = Me.SSTab1.Tab = 1 And Me.DataGrid1.visible
    Me.Label1(46).visible = (Modo = 5) And Me.DataGrid1.visible And Me.SSTab1.Tab = 1 And (vEmpresa.TieneAnalitica)
    Me.txtAux2(9).visible = (Modo = 5) And Me.DataGrid1.visible And Me.SSTab1.Tab = 1 And (vEmpresa.TieneAnalitica)
    Me.imgBuscar(11).visible = Me.SSTab1.Tab = 1 And Me.DataGrid1.visible
End Sub

Private Sub Text1_Change(Index As Integer)
    If Index = 9 Then HaCambiadoCP = True 'Cod. Postal
End Sub

'### A mano
'Los metodos del text tendran que estar
'Los descomentamos cuando esten puestos ya los controles
Private Sub Text1_GotFocus(Index As Integer)
    kCampo = Index
    If Index = 9 Then HaCambiadoCP = False 'CPostal
    If Index = 1 And Modo = 1 Then
        SendKeys "{tab}"
        Exit Sub
    End If
    ConseguirFoco Text1(Index), Modo
End Sub


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    If KeyCode <> 13 Then KEYdown KeyCode
End Sub


Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
    If KeyAscii = 13 Then KeyAscii = 0
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
Dim ImporteDto As Currency
    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
        
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
          
    'Si queremos hacer algo ..
    Select Case Index
        Case 2 'Fecha factura
                If Text1(Index).Text <> "" Then PonerFormatoFecha Text1(Index)
                
        Case 3, 27, 28 'Cod Vendedor
'                Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "straba", "nomtraba", "codtraba")

        Case 4 'Cod. Cliente
            If Modo = 1 Then 'Modo=1 Busqueda
                '-- Laura 12/01/2007
                'Text1(5).Text = PonerNombreDeCod(Text1(Index), conAri, NombreTabla, "nomclien")
                Text1(5).Text = PonerNombreDeCod(Text1(Index), conAri, "sclien", "nomclien")
                '--
            Else
                PonerDatosCliente (Text1(Index).Text)
            End If
        
        Case 6 'NIF
            If Not EsDeVarios Then Exit Sub
            If Modo = 4 Then 'Modificar
                'si no se ha modificado el nif del cliente no hacer nada
                If Text1(6).Text = Data1.Recordset!nifClien Then Exit Sub
            End If
            PonerDatosClienteVario (Text1(Index).Text)
        
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
        
        Case 12 'Cod. Direc
            If Modo = 1 Then Exit Sub
            If PonerFormatoEntero(Text1(Index)) Then
                'Comprobar que el cliente seleccionada tiene esa direccion
                If PonerDptoEnCliente Then
                    'Comprobar que el cliente tiene mantenimientos en esa direc/dpto
                    devuelve = DevuelveDesdeBDNew(conAri, "scaman", "nummante", "codclien", Text1(4).Text, "N", , "coddirec", Text1(12).Text, "N")
                    If devuelve <> "" And Text1(Index).Locked = False Then
                        devuelve = "El cliente tiene Mantenimientos."
                        MsgBox devuelve, vbInformation
                    End If
                Else
                    PonerFoco Text1(Index)
                End If
            Else
                Text1(Index + 1).Text = ""
            End If
            
        Case 14 'Cod. Agente
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "sagent", "nomagent")
            Else
                Text2(Index).Text = ""
            End If
        
        Case 15 'Forma de Pago
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "sforpa", "nomforpa")
            Else
                Text2(Index).Text = ""
            End If
            
        Case 16, 17 'Descuentos
            If PonerFormatoDecimal(Text1(Index), 4) Then   'Tipo 4: Decimal(4,2)
                If Modo = 4 Then
                    
                    devuelve = ""
                    ImporteDto = ImporteFormateado(Text1(Index).Text)
                    If Index = 16 Then
                        If DBLet(Data1.Recordset!DtoPPago, "N") <> ImporteDto Then devuelve = "S"
                    Else
                        If DBLet(Data1.Recordset!DtoGnral, "N") <> ImporteDto Then devuelve = "S"
                    End If
                    If devuelve <> "" Then CalcularDatosFactura
                End If
            End If
            
        Case 18 To 21 'banco, sucursal
            PonerFormatoEntero Text1(Index)
        Case 29 'Cod envio
'            Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "senvio", "nomenvio")
    End Select
End Sub


Private Sub HacerBusqueda()
Dim cadB As String
Dim cadAux As String
Dim OtraBusq As String

    '--- Laura 12/01/2007
    cadAux = Text1(5).Text
    If Text1(4).Text <> "" Then Text1(5).Text = ""
    
    '---
    If Combo1.ListIndex >= 0 Then Text1(1).Text = Mid(Combo1.List(Combo1.ListIndex), 1, 3)
    
    
    
    cadB = ObtenerBusqueda(Me, False, BuscaChekc)
    If Combo1.ListIndex < 0 Then
        If vParamAplic.NumeroInstalacion = 2 Then
            'NO ha selecionado ningun tipo de movimiento
            If (vUsu.Codigo Mod 1000) > 0 Then
                If cadB <> "" Then cadB = cadB & " AND "
                cadB = cadB & " scafac.codtipom "
                If Val(vUsu.AlmacenPorDefecto2) = vParamAplic.AlmacenB Then
                    cadB = cadB & " = "
                Else
                    cadB = cadB & " <> "
                End If
                cadB = cadB & " 'FAZ'"
            End If
            
        ElseIf vParamAplic.NumeroInstalacion = vbFenollar Then
            If Not HaMostradoCanal2_El_B Then
                If cadB <> "" Then cadB = cadB & " AND "
                cadB = cadB & "scafac.codtipom<>'FAZ'"
            End If
        End If
    End If
    
    If vParamAplic.NumeroInstalacion = vbHerbelca Then
        If vUsu.CodigoAgente > 0 Then
            If cadB <> "" Then cadB = cadB & " AND "
            cadB = cadB & " (codagent = " & vUsu.CodigoAgente
            If vUsu.ClientesEnQueAgenteEsVisitador <> "" Then cadB = cadB & " OR scafac.codclien in (" & vUsu.ClientesEnQueAgenteEsVisitador & ")"
            cadB = cadB & ")"
        End If
    End If
    
    If InstalacionEsEulerTaxco Then
        If vParamAplic.NumeroInstalacion = vbEuler Then
            OtraBusq = DevuelveBusquedaCostesEuler
            If OtraBusq <> "" Then
                If cadB <> "" Then cadB = cadB & " AND "
                cadB = cadB & " (scafac.codtipom,scafac.numfactu,scafac.fecfactu) IN (Select  distinct codtipom,numfactu,fecfactu  FROM  slifac_eu  where " & OtraBusq & ")"
            End If
    
        ElseIf vParamAplic.NumeroInstalacion = vbTaxco Then
            OtraBusq = DevuelveBusquedaTaxco
            
            If OtraBusq <> "" Then
                If cadB <> "" Then cadB = cadB & " AND "
                cadB = cadB & " (scafac.codtipom,scafac.numfactu,scafac.fecfactu) IN (Select  distinct codtipom,numfactu,fecfactu  FROM  scafac_eu  where " & OtraBusq & ")"
            End If
        End If
        
    
    
    End If
    '--- Laura 12/01/2007
    Text1(5).Text = cadAux
    '---
    
    If chkVistaPrevia = 1 Then
        EsCabecera2 = 0
        MandaBusquedaPrevia cadB
    ElseIf cadB <> "" Then
        'Se muestran en el mismo form
'        cadB = cadB & " and scafac.codtipom='" & CodTipoMov & "'" 'Solo seleccionamos los del Movimiento, aqui los FAV
        CadenaConsulta = "select scafac.* from " & NombreTabla & " INNER JOIN scafac1 ON scafac.codtipom=scafac1.codtipom AND scafac.numfactu=scafac1.numfactu AND scafac.fecfactu=scafac1.fecfactu "
        CadenaConsulta = CadenaConsulta & " WHERE " & cadB & " GROUP BY scafac.codtipom,scafac.numfactu,scafac.fecfactu " & Ordenacion
        PonerCadenaBusqueda
    End If
    
    
End Sub


Private Sub MandaBusquedaPrevia(cadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim Cad As String
Dim tabla As String
Dim Titulo As String
Dim Desc As String, devuelve As String
    'Llamamos a al form
    '##A mano
    Cad = ""
    If EsCabecera2 = 0 Then
        Cad = Cad & ParaGrid(Text1(1), 10, "Tipo Fac.")
        Cad = Cad & ParaGrid(Text1(0), 15, "Nº Factura")
        Cad = Cad & ParaGrid(Text1(2), 15, "Fecha Fac.")
        Cad = Cad & ParaGrid(Text1(4), 10, "Cliente")
        Cad = Cad & ParaGrid(Text1(5), 50, "Nombre Cliente")
        tabla = NombreTabla & " INNER JOIN scafac1 ON scafac.codtipom=scafac1.codtipom AND scafac.numfactu=scafac1.numfactu AND scafac.fecfactu=scafac1.fecfactu "
        'CadenaConsulta = "select scafac.* from " & NombreTabla & " INNER JOIN scafac1 ON scafac.codtipom=scafac1.codtipom AND scafac.numfactu=scafac1.numfactu AND scafac.fecfactu=scafac1.fecfactu "
        'CadenaConsulta = CadenaConsulta & " WHERE " & cadB & " GROUP BY scafac.codtipom,scafac.numfactu,scafac.fecfactu " & Ordenacion
        
        Titulo = "Facturas"
        devuelve = "0|1|2|"
    Else
        If EsCabecera2 = 1 Then
            'DEPARTAMENTO    DIRECCION
            If vParamAplic.HayDeparNuevo = 1 Then
                Titulo = "Dptos Cliente: "
                Desc = "Dpto."
            ElseIf vParamAplic.HayDeparNuevo = 0 Then
                Titulo = "Direc. Cliente: "
                Desc = "Direc."
            Else
                Titulo = "Obras Cliente: "
                Desc = "Obra"
            End If
            Titulo = Titulo & Text1(4).Text & " - " & Text1(5).Text
            Cad = Cad & "Cod. " & Desc & "|sdirec|coddirec|N||15·"
            Cad = Cad & "Desc. " & Desc & "|sdirec|nomdirec|T||35·"
            tabla = "sdirec"
            devuelve = "0|1|"
        Else
            'DIRECCION ENVIO
            Titulo = "Dir. envio cliente: " & Text1(4).Text & " - " & Text1(5).Text
            Cad = Cad & "Codigo|sdirenvio|coddiren|N||15·"
            Cad = Cad & "Descricpion|sdirenvio|nomdiren|T||35·"
            tabla = "sdirenvio"
            devuelve = "0|1|"
        End If
    End If
           
    If Cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = Cad
        frmB.vTabla = tabla
        frmB.vSQL = cadB
        HaDevueltoDatos = False
        '###A mano
'        frmB.vDevuelve = "0|1|"
        frmB.vDevuelve = devuelve
        frmB.vTitulo = Titulo
        frmB.vselElem = 0
        If vParamAplic.NumeroInstalacion = vbFenollar Then
            frmB.vselElem = 2
            frmB.vDescendente = True
        End If
        
        frmB.vConexionGrid = conAri  'Conexión a BD: Ariges
        If EsCabecera2 > 0 Then frmB.Label1.FontSize = 11
'        frmB.vBuscaPrevia = chkVistaPrevia
        '#
        frmB.Show vbModal
        Set frmB = Nothing
        If EsCabecera2 = 0 Then
            PonerCadenaBusqueda
            Text1(0).Text = Format(Text1(0).Text, "0000000")
        End If

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
            PonerFoco Text1(kCampo)
'            Text1(0).BackColor = vbYellow
        End If
        lblIndicador.Caption = ""
        Exit Sub
    Else
        Data1.Recordset.MoveFirst
        PonerModo 2
        
        LLamaLineas Modo, 0, "DataGrid2"
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
'Carga el grid de los AlbaranesxFactura, es decir, la tabla scafac1 de la factura seleccionada
Dim B As Boolean
Dim b2 As Boolean

    On Error GoTo EPonerLineas

    Screen.MousePointer = vbHourglass
    
    'Datos de la tabla albaranes x factura: scafac1
    CargaGrid DataGrid2, Data3, True
    
    'Comprobar si el albaran de la factura viene de una venta de ticket del TPV
    B = False
    b2 = False
    If Not Data3.Recordset.EOF Then
        If Not IsNull(Data3.Recordset!NumVenta) Then
            B = True
            If Data3.Recordset!codtipom = "FAV" And Data3.Recordset!Codtipoa <> "FTI" Then b2 = True
        End If
    End If
    
    'Visualizar los campos de Oferta y Pedido si es una Factura q no es de venta TPV
    'o visulaizar numventa, numtermi si es una Factura de venta del TPV
    Label1(6).Caption = "Nº Pedido"
    Label1(18).Caption = "Fecha Pedido"
    If B Then
        If b2 Then
            Label1(6).Caption = "Nº Ticket"
            Label1(18).Caption = "Fecha Ticket"
        End If
        Label1(40).Caption = "Nº Terminal"
        Label1(22).Caption = "Nº Venta"
    Else
        Label1(40).Caption = "Nª Oferta"
        Label1(22).Caption = "Fecha Oferta"
    End If
    'sem. entrega
    Label1(2).visible = Not (B And b2)
    Text3(8).visible = Not (B And b2)
    'OFERTA
    Text3(4).visible = Not B
    Text3(5).visible = Not B
    'VENTA
    Text3(14).visible = B
    Text3(15).visible = B
    
    
    
    
    If vParamAplic.Ariagro <> "" Then CargaDatosCampos
    If vParamAplic.TieneTelefonia2 > 0 Then CargaDatosTelefonia
    If InstalacionEsEulerTaxco Then PonerCamposFicha True
        
        
    'Poner la referencia del cliente
  '  If Not data3.Recordset.EOF Then Text1(3).Text = DBLet(data3.Recordset.Fields!referenc, "T")
    
    Screen.MousePointer = vbDefault
    Exit Sub
    
EPonerLineas:
    MuestraError Err.Number, "PonerCamposLineas"
    PonerModo 2
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerCampos()
Dim BrutoFac As Single

    On Error Resume Next

    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma Me, Data1
    
  
    
    If lblSerie.visible Then
        If lblSerie.Tag <> Data1.Recordset!codtipom Then
             lblSerie.Caption = "Serie " & DevuelveDesdeBD(conAri, "letraser", "stipom", "codtipom", Text1(1).Text, "T")
             lblSerie.Tag = Text1(1).Text
        End If
    End If
        
    
    If Text1(16).Text = "0,00" Then Text1(16).Text = ""
    If Text1(17).Text = "0,00" Then Text1(17).Text = ""
    
    'Poner la base imponible (impbruto - dtoppago - dtognral
    BrutoFac = CSng(Text1(22).Text) - CSng(Text1(23).Text) - CSng(Text1(24).Text)
    Text1(25).Text = Format(BrutoFac, FormatoImporte)
    
    FormatoDatosTotales
    
    'poner descripcion campos
    Modo = 4
    Text1_LostFocus (12) 'direc./dpto
    Text1_LostFocus (14) 'agente
    Text1_LostFocus (15) 'forma de pago
    Modo = 2
    
    PonerCamposLineas '
    
    
    
    
    
    
    
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
Dim B As Boolean

    On Error GoTo EPonerModo

    'Actualiza Iconos Insertar,Modificar,Eliminar
    '## No tiene el boton modificar y no utiliza la funcion general
    ActualizarToolbar Modo, Kmodo
    Text1(3).visible = False  'SIEMPRE VISIBLE FALSE
    Modo = Kmodo
    PonerIndicador lblIndicador, Modo
    
    

    
    'Modo 2. Hay datos y estamos visualizandolos
    '=========================================
    B = (Modo = 2)
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.visible = B
    Else
        cmdRegresar.visible = False
    End If
        
    
    'En EULER, en modo=2 aparece el boton de añadir albaranes
    If vParamAplic.NumeroInstalacion = vbEuler Then cmdInsertaAlbaranCostes.visible = B
    
    
    'Poner Flechas de desplazamiento visibles
    NumReg = 1
    If Modo = 2 Then
        If Not Data1.Recordset.EOF Then
            If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
        End If
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero, B, NumReg
          
          
        
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1
    'si estamos en modificar bloquea las compos que son clave primaria
    BloquearText1 Me, Modo
    
    'marzo 22
    If Modo = 4 Then
        
        '       0.- Se puede TODO. NO esta contabilizada
        '       1.- Se puede "parcial", lo que hacia antes
        '       2.- Solo se puede trabajador, zona,referencias.....
        '       3.- NADA. Fuera ejercicios
        If TipoModificacionFactura > 1 Then
            'No dejamos NI modificar dtos
            BloquearTxt Text1(4), True
            BloquearTxt Text1(16), True
            BloquearTxt Text1(17), True
        End If
    End If
    
    Me.Check1.Enabled = (Modo = 1)
    Me.Check2.Enabled = (Modo = 1)
    
    B = (Modo <> 1)
    'Campos Nº Factura bloqueado y en azul
    BloquearTxt Text1(0), B, True
    BloquearTxt Text1(3), B 'referencia
    
    
    'Importes siempre bloqueados, excepto para busquedas. ivas y aportacion tb bloqueado
    For i = 22 To 45
        BloquearTxt Text1(i), (Modo <> 1)
    Next i
    'Aportacion
    BloquearTxt Text1(25), True
    Text1(25).BackColor = &HFFFFC0
    
    If Modo <> 1 Then
        Text1(35).BackColor = &HFFFFC0
        Text1(36).BackColor = &HFFFFC0
        Text1(37).BackColor = &HFFFFC0
        Text1(38).BackColor = &HC0FFC0
    End If
    
    'bloquear los Text3 que son las lineas de scafac1
    BloqueaText3
    
    'Si no es modo lineas Boquear los TxtAux
    For i = 0 To txtAux.Count - 1
        BloquearTxt txtAux(i), (Modo <> 5)
    Next i
    BloquearTxt txtAux(8), True
    BloquearTxt txtAux(10), True
    BloquearTxt txtAux(11), True
    
    'Si no es modo Busqueda Bloquear los TxtAux3 (son los txtaux de los albaranes de factura)
    For i = 0 To txtAux3.Count - 1
        BloquearTxt txtAux3(i), (Modo <> 1)
    Next i
    
    'ampliacion linea
    B = Me.DataGrid1.visible And Me.SSTab1.Tab = 1
    
    Me.Label1(35).visible = B
    Me.Text2(16).visible = B
    
    BloquearTxt Text2(16), (Modo <> 5) Or (Modo = 5 And ModificaLineas <> 1)
    'nombre Proveedor
    Me.Label1(46).visible = (Modo = 5) And B
    Me.txtAux2(9).visible = (Modo = 5) And B
    
    imgBuscarEULER.visible = Modo = 1 And vParamAplic.NumeroInstalacion = vbEuler
    
    If vParamAplic.NumeroInstalacion = vbTaxco Then
        TextmatriculaTaxco.visible = Modo = 1
        Label1(60).visible = Modo = 1
        FrameTAXCO.visible = Modo = 1
        lblSerie.visible = Modo = 2
        lblSerie.Tag = ""
    End If
    
    Me.Combo1.visible = (Modo = 1)

    '---------------------------------------------
    B = (Modo <> 0 And Modo <> 2 And Modo < 5)
    cmdCancelar.visible = B
    cmdAceptar.visible = B
    
    
    For i = 0 To 5
        Me.imgBuscar(i).Enabled = B
    Next i
    For i = 6 To 9
        Me.imgBuscar(i).Enabled = B And (Modo <> 1)
    Next i
    
    Me.imgBuscar(1).visible = False
                    
    If InstalacionEsEulerTaxco Then
        If Modo = 1 Then FrameEuler.Enabled = True
        ModoBusquedaCostes Modo = 1
    End If
                    
    'trampa
    If Modo = 1 Then
       Me.chkEnvio.Tag = "Documento ar|N|S|||scafac1|docarchiv|||"
       chkPedxCli.Tag = "Ped|N|S|||scafac1|PideCliente|||"
    Else
        chkEnvio.Tag = ""
        chkPedxCli.Tag = ""
    End If
                    
                    
   ' St op
                    
    'Sept 2020
    'Si el usuario no es nivel admin y esta modificando
    If vUsu.Nivel > 0 And Modo = 4 Then
        'Aqui NO dejamos cambiar ciertas cosas
        ' Codclien, aportacion, dto pp y pie
        NumReg = 2
        BuscaChekc = "4|45|16|17|"
        While BuscaChekc <> ""
            i = InStr(1, BuscaChekc, "|")
            If i = 0 Then
                BuscaChekc = ""
            Else
                NumReg = CByte(Val(Mid(BuscaChekc, 1, i - 1)))
                BuscaChekc = Mid(BuscaChekc, i + 1)
                BloquearTxt Text1(NumReg), True   'YA sabemos que modo=4
            End If
        Wend
        imgBuscar(0).Enabled = False
        BuscaChekc = ""
    End If
                    
                    
                     
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
'Comprobar que los datos de la cabecera son correctos antes de Insertar o Modificar
'la cabecera del Pedido
Dim B As Boolean
Dim bT As Boolean

    On Error GoTo EDatosOK

    DatosOk = False
    
    ComprobarDatosTotales
    
    'comprobamos datos OK de la tabla scafac
    B = CompForm(Me, 1) 'Comprobar formato datos ok de la cabecera: opcion=1
    If Not B Then Exit Function
    
    
    
    'Lleva direcciones de envio. Comprobamos que la que ha puesto existe...
    If B Then
        If vParamAplic.DireccionesEnvio Then
            If Text3(18).Text = "" Xor Text2(18).Text = "" Then
                MsgBox "Dirección de envio INCORRECTA", vbExclamation
                B = False
                PonerFoco Text3(18)
            End If
            'Ha puesto un codenvio y parece ser que existe... LO COMPURBEO que no hay referenciales
            If B And Text3(18).Text <> "" Then
                BuscaChekc = DevuelveDesdeBDNew(1, "sdirenvio", "nomdiren", "codclien", Text1(4).Text, "N", "", "coddiren", Text3(18).Text, "N")
                If BuscaChekc = "" Then
                    MsgBox "NO existe la dirección de envio: " & Text3(18).Text, vbExclamation
                    PonerFoco Text3(18)
                    B = False
                End If
                BuscaChekc = ""
            End If
         End If 'de direnvii
    End If
    
    
    'MARZO 2013
    '----------
    ' Si es FAI y tiene telefonia, o
    ' no puede modificar la referencia proveedor si es relacionada con un archivo
    ' procesado
    If vParamAplic.TieneTelefonia2 > 0 Then
        bT = False
        If Text1(1).Text = "FAI" Then
            If DBLet(Data3.Recordset!referenc, "T") <> "" Then bT = True
        Else
            If Text1(1).Text = "FAT" Then bT = True
        End If
            
        
        If bT Then
            If DBLet(Data3.Recordset!referenc, "T") <> Text3(16).Text Then
                'OK, ha cambiado la referencia
                BuscaChekc = DevuelveDesdeBD(conAri, "count(*)", "tel_fichtraspasados", "fichero", Data3.Recordset!referenc, "T")
                If BuscaChekc <> "" Then
                    If Val(BuscaChekc) > 0 Then
                        MsgBox "No puede cambiar la referencia de una factura interna de telefonia", vbExclamation
                        Text3(16).Text = Data3.Recordset!referenc
                        PonerFoco Text3(16)
                        B = False
                    End If
                End If
                BuscaChekc = ""
            End If
        End If
    End If
        
        
        
    If InstalacionEsEulerTaxco Then
        
    
    End If
        
        
    DatosOk = B
    
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Function DatosOkLinea() As Boolean
Dim B As Boolean
Dim i As Byte

    On Error GoTo EDatosOkLinea

    DatosOkLinea = False
    B = True

    For i = 0 To txtAux.Count - 1
        If i = 4 Or i = 6 Or i = 7 Then
            If txtAux(i).Text = "" Then
                MsgBox "El campo " & RecuperaValor(txtAux(i).Tag, 1) & " no puede ser nulo", vbExclamation
                B = False
                PonerFoco txtAux(i)
                Exit Function
            End If
        End If
    Next i
            
            
    'PRoveedor
    If txtAux(9).Text <> "" And txtAux2(9).Text = "" Then
        MsgBox "Codigo proveedor/CC incorrecto", vbExclamation
        PonerFoco txtAux(9)
        B = False
        Exit Function
    End If
            
    DatosOkLinea = B
    
EDatosOkLinea:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Sub Text2_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 16 And KeyCode = 40 Then 'campo Amliacion Linea y Flecha hacia abajo
        PonerFocoBtn Me.cmdAceptar
    Else
        KEYdown KeyCode
    End If
End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 16 And KeyAscii = 13 Then 'campo Amliacion Linea y ENTER
        PonerFocoBtn Me.cmdAceptar
    End If
End Sub


Private Sub Text3_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text3_LostFocus(Index As Integer)
    If Modo = 1 Then Exit Sub
    Select Case Index
        Case 0, 1, 2 'trabajador
            Text2(Index).Text = PonerNombreDeCod(Text3(Index), conAri, "straba", "nomtraba", "codtraba", "Cod. Trabajador", "N")
        Case 3 'cod. envio
            Text2(Index).Text = PonerNombreDeCod(Text3(Index), conAri, "senvio", "nomenvio", "codenvio", "Cod. Envio", "N")
            If Screen.ActiveControl.TabIndex <> 27 Then PonerFocoBtn Me.cmdAceptar
        
        Case 13 'observa 5
            PonerFocoBtn Me.cmdAceptar
            
        Case 17
           ' If PonerFormatoFecha(Text3(17)) Then PonerFoco Text3(17)
            
         Case 18
            If Text3(18).Text = "" Then
                Text2(18).Text = ""
            Else
                If PonerFormatoEntero(Text3(18)) Then
                    Text2(18).Text = PonerNombreDeCod(Text3(18), conAri, "sdirenvio", "nomdiren", "codclien = " & Val(Text1(4).Text) & " AND coddiren", "N")
                    If Text2(18).Text = "" Then MsgBox "No existe la direccion de envio", vbExclamation
                Else
                    'Form
                    Text2(18).Text = ""
                End If
                
                If Text3(18).Text <> "" And Text2(18).Text = "" Then
                    Text3(18).Text = ""
                    PonerFoco Text3(18)
                End If
            End If
        Case 23
            'FH
            If Not EsFechaHoraOK(Text3(Index)) Then
                                
            End If
        Case 26, 27
            If Not PonerFormatoDecimal(Text3(Index), 8) Then Text3(Index).Text = ""
    End Select
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1: mnBuscar_Click  'Buscar
        Case 2: BotonVerTodos  'Todos
            

        Case 5: mnModificar_Click  'Modificar
        Case 6: mnEliminar_Click  'Borrar
        
        Case 9: mnLineas_Click  'Lineas
        Case 10: mnImprimir_Click 'Imprimir Albaran
        
        Case 11: mnImprimirAlbaran_Click
            
        Case 13: mnModLotes_Click
        
        Case 14:
            If InstalacionEsEulerTaxco Then
                ImprimirCostesEuler
            Else
                mnEditarCampos_Click
            End If
        
        Case 15: mnTipoPreciosLinea_Click
         
        Case 16:
                EliminarCambiarFechaFactura
         
         
        Case 18
                If Modo = 5 Then
                    'Ajustar loeste fitosanitarios
                    ModificaLote
                Else
                    ImprimirValoracionOferta
                End If
        Case 19
                If Modo <> 2 Then Exit Sub
                AbrirPDFs
                
        Case 21: mnSalir_Click    'Salir
            
        Case btnPrimero To btnPrimero + 3 'Flechas Desplazamiento
            Desplazamiento (Button.Index - btnPrimero)
    End Select
End Sub


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub ActualizarToolbar(Modo As Byte, Kmodo As Byte)
'Modo: Modo antiguo
'Kmodo: Modo que se va a poner

    If (Modo = 5) And (Kmodo <> 5) Then
        'El modo antigu era modificando las lineas
        'Luego hay que reestablecer los dibujitos y los TIPS
        '-- insertar
'        Toolbar1.Buttons(5).Image = 3
'        Toolbar1.Buttons(5).ToolTipText = "Nuevo Albaran"
        '-- Modificar
        Toolbar1.Buttons(5).Image = 4
        Toolbar1.Buttons(5).ToolTipText = "Modificar Factura"
        '-- eliminar
        Toolbar1.Buttons(6).Image = 5
        Toolbar1.Buttons(6).ToolTipText = "Eliminar Factura"
        
        '
        Toolbar1.Buttons(6).Image = 5
        Toolbar1.Buttons(6).ToolTipText = "Eliminar Factura"
        
        
        
        'Noviembre 2015
        If vParamAplic.ManipuladorFitosanitarios2 Then
            Toolbar1.Buttons(18).Image = 31
            Toolbar1.Buttons(18).ToolTipText = "Valoración factura"
        End If
    End If
    If Kmodo = 5 Then
        'Ponemos nuevos dibujitos y tal y tal
        'Luego hay que reestablecer los dibujitos y los TIPS
        '-- insertar
'        Toolbar1.Buttons(5).Image = 12
'        Toolbar1.Buttons(5).ToolTipText = "Nueva linea"
        '-- Modificar
        Toolbar1.Buttons(5).Image = 13
        Toolbar1.Buttons(5).ToolTipText = "Modificar linea factura"
        '-- eliminar
        Toolbar1.Buttons(6).Image = 14
        Toolbar1.Buttons(6).ToolTipText = "Eliminar linea factura"
        
        'Oct 2015
        If vParamAplic.ManipuladorFitosanitarios2 Then
            Toolbar1.Buttons(18).Image = 48
            Toolbar1.Buttons(18).ToolTipText = "Lotes asignados"
        End If
        
    End If
End Sub
    
    
Private Function ModificarLinea() As Boolean
'Modifica un registro en la tabla de lineas de Albaran: slialb
Dim Sql As String
Dim vWhere As String
Dim B As Boolean

    On Error GoTo EModificarLinea

    ModificarLinea = False
    If Data2.Recordset.EOF Then Exit Function
    
    vWhere = ObtenerWhereCP(True)
    vWhere = vWhere & " AND codtipoa='" & Data3.Recordset.Fields!Codtipoa & "' "
    vWhere = vWhere & " AND numalbar=" & Data3.Recordset.Fields!Numalbar
    vWhere = vWhere & " AND numlinea=" & Data2.Recordset.Fields!numlinea
    
    If DatosOkLinea() Then
        Sql = "UPDATE slifac SET "
        
        
        'Si le articulo era de varios, podiamos cambiar el texto
        If txtAux(2).visible Then Sql = Sql & " nomartic=" & DBSet(txtAux(2).Text, "T") & ", "
        
        Sql = Sql & " ampliaci=" & DBSet(Text2(16).Text, "T") & ", "
        Sql = Sql & "precioar = " & DBSet(txtAux(4).Text, "N") & ", "
        Sql = Sql & "dtoline1= " & DBSet(txtAux(6).Text, "N") & ", dtoline2= " & DBSet(txtAux(7).Text, "N") & ", "
        Sql = Sql & "importel = " & DBSet(txtAux(8).Text, "N") & ", "
        Sql = Sql & "origpre='" & txtAux(5) & "'"
        'TRAZA
        If vParamAplic.NumeroInstalacion = 2 Then
            'NADA
        Else
            If vEmpresa.TieneAnalitica Then
                Sql = Sql & ",codccost= " & DBSet(txtAux(9).Text, "T", "S")
            Else
                Sql = Sql & ",codprovex= " & DBSet(txtAux(9).Text, "N", "S")
            End If
        End If
        Sql = Sql & " " & vWhere
    End If
    
    If Sql <> "" Then
        'actualizar la factura y vencimientos
        B = ModificarFactura(Sql)
        
        
        If B Then
            'Noviembre 2020
            'Acutalizamos en smoval el importe
            If ImporteFormateado(txtAux(8).Text) <> Data2.Recordset!ImporteL Then
                Sql = "UPDATE smoval SET impormov = " & DBSet(txtAux(8).Text, "N")
                Sql = Sql & " WHERE codartic =" & DBSet(Data2.Recordset!codArtic, "T")
                Sql = Sql & " AND codalmac =" & DBSet(CStr(Data2.Recordset!codAlmac), "T") & " AND detamovi =" & DBSet(CStr(Data3.Recordset!Codtipoa), "T")
                Sql = Sql & " AND fechamov = " & DBSet(CStr(Data3.Recordset!FechaAlb), "F") & " AND  document= " & Format(Data2.Recordset!Numalbar, "0000000")
                Sql = Sql & " AND  numlinea =" & DBSet(Data2.Recordset!numlinea, "N")
                ejecutar Sql, True
            End If
        End If
        ModificarLinea = B
    End If
    
EModificarLinea:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Modificar Lineas Factura" & vbCrLf & Err.Description
        B = False
    End If
    ModificarLinea = B
End Function


Private Sub PonerBotonCabecera(B As Boolean)
'Pone el boton de Regresar a la Cabecera si pasamos a MAntenimiento de Lineas
'o Pone los botones de Aceptar y cancelar en Insert,update o delete lineas
    On Error Resume Next

    Me.cmdAceptar.visible = Not B
    Me.cmdCancelar.visible = Not B
    Me.cmdRegresar.visible = B
    Me.cmdRegresar.Caption = "Cabecera"
    If B Then
        Me.lblIndicador.Caption = "Líneas " & TituloLinea
        PonerFocoBtn Me.cmdRegresar
        cmdRegresar.Cancel = True
        cmdAceptar.Caption = "&Aceptar"
    Else
        Me.cmdCancelar.Cancel = True
        Me.cmdAceptar.Caption = "Aceptar"
    End If
    'Habilitar las opciones correctas del menu segun Modo
    PonerModoOpcionesMenu (Modo)
    PonerOpcionesMenu 'Habilitar las opciones correctas del menu segun Nivel de Acceso
    DataGrid2.Enabled = Not B
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub CargaGrid(ByRef vDataGrid As DataGrid, ByRef vData As Adodc, enlaza As Boolean)
Dim B As Boolean
Dim Opcion As Byte
Dim Sql As String

    On Error GoTo ECargaGrid

    B = DataGrid1.Enabled
    If vDataGrid.Name = "DataGrid1" Then
        Opcion = 1
    Else
        Opcion = 2
    End If
    Sql = MontaSQLCarga(enlaza, Opcion)
    CargaGridGnral vDataGrid, vData, Sql, PrimeraVez
    
    vDataGrid.RowHeight = 270
    
    CargaGrid2 vDataGrid, vData
    vDataGrid.ScrollBars = dbgAutomatic
    
     B = (Modo = 5) And (ModificaLineas = 1 Or ModificaLineas = 2)
     vDataGrid.Enabled = Not B
    
    Exit Sub
    
ECargaGrid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub


Private Sub CargaGrid2(ByRef vDataGrid As DataGrid, ByRef vData As Adodc)
Dim tots As String
Dim B1 As Boolean

    On Error GoTo ECargaGrid

    Select Case vDataGrid.Name
        Case "DataGrid1" 'Cod. Almacen
            'SQL = "SELECT codtipom, numfactu, fecfactu, numalbar, numlinea,
            'codalmac, codartic, nomartic, ampliaci, cantidad,numbultos, precioar, origpre, dtoline1, dtoline2, importel "
            tots = "N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;"
            tots = tots & "S|txtAux(0)|T|Alm.|500|;S|txtAux(1)|T|Artículo|1600|;S|txtAux(2)|T|Nombre Art.|3300|;"
            tots = tots & "N||||0|;S|txtAux(3)|T|Cantidad|900|;S|txtAux(11)|T|Bultos|700|;S|txtAux(4)|T|Precio|1200|;S|txtAux(5)|T|OP|400|;S|txtAux(6)|T|Dto 1|600|;S|txtAux(7)|T|Dto 2|600|;S|txtAux(8)|T|Importe|1240|;"
            'TRAZA
'            tots = tots & "S|txtAux(9)|T|Prov.|750|;S|cmdaux|B|||;S|txtAux(10)|T|Nom. prov.|2000|;"
            If vEmpresa.TieneAnalitica Then
                'codprove,nomprove, codccost
                tots = tots & "N||||0|;S|cmdaux|B|||;N||||0|;S|txtAux(9)|T|CCoste|750|;"

            Else
                B1 = False
                If vParamAplic.NumeroInstalacion = 2 Then If vUsu.Nivel = 0 Then B1 = True
                If B1 Then
                    'herbelca
                    tots = tots & "S|txtAux(9)|T|Comis.|750|;S|cmdaux|B|||;N||||0|;N||||0|;"
                Else
                    'resto
                    tots = tots & "S|txtAux(9)|T|Prov.|750|;S|cmdaux|B|||;N||||0|;N||||0|;"
                End If
            End If
            'numlote
            tots = tots & "S|txtAux(10)|T|Nº Lote|1300|;"
            
            
            arregla tots, DataGrid1, Me
            DataGrid1.Columns(9).Alignment = dbgRight
            DataGrid1.Columns(10).Alignment = dbgRight
            DataGrid1.Columns(12).Alignment = dbgCenter
            DataGrid1.Columns(13).Alignment = dbgRight
            DataGrid1.Columns(14).Alignment = dbgRight
            DataGrid1.Columns(15).Alignment = dbgRight
                       
         Case "DataGrid2" 'albaranes x articulo
'             SQL = "SELECT codtipom,numfactu,fecfactu,codtipoa,numalbar, fechaalb,"
             'numpedcl,fecpedcl,sementre,numofert,fecofert, referenc, codenvio,codtraba, codtrab1, codtrab2,observa1,observa2,observa3,observa4,observa5,numtermi,numventa  "
            tots = "N||||0|;N||||0|;N||||0|;"
            tots = tots & "S|txtAux3(0)|T|Tipo|600|;S|txtAux3(1)|T|Albaran|1100|;S|txtAux3(2)|T|Fecha|1200|;"
            tots = tots & "N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;"
            tots = tots & "N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;"
            If vParamAplic.DireccionesEnvio Then tots = tots & "N||||0|;"
            tots = tots & "N||||0|;"  'docarchivado
            
            'Mani`pulador fitosantiarios  pidecliente
            tots = tots & "N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;"
            '                                                       fechaent,perrecep,dnirecep,latitud,longitud"
            If InstalacionEsEulerTaxco Then tots = tots & "N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;"
            'refcontrato refpedido
            tots = tots & "N||||0|;N||||0|;"
            
            arregla tots, DataGrid2, Me
                     
            DataGrid2_RowColChange 1, 1
    End Select
    
    vDataGrid.HoldFields
    Exit Sub
    
ECargaGrid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub




Private Sub TxtAux_Change(Index As Integer)
    If Index = 6 And ModificaLineas = 2 Then 'Precio y Modo Borrar Lineas
        txtAux(5).Text = "M"
    End If
End Sub

Private Sub txtAux_GotFocus(Index As Integer)
    ConseguirFocoLin txtAux(Index)
End Sub

Private Sub TxtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
End Sub


Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub txtAux_LostFocus(Index As Integer)

    'Quitar espacios en blanco
    If Not PerderFocoGnralLineas(txtAux(Index), ModificaLineas) Then Exit Sub
    
    Select Case Index
        Case 4 'Precio
             'Tipo 2: Decimal(10,4)
             If txtAux(Index).Text <> "" Then PonerFormatoDecimal txtAux(Index), 2
            
        Case 6, 7 'Descuentos
            PonerFormatoDecimal txtAux(Index), 4 'Tipo 4: Decimal(4,2)
            If Index = 7 Then PonerFoco Me.Text2(16)
            
        Case 8 'Importe Linea
            PonerFormatoDecimal txtAux(Index), 3 'Tipo 3: Decimal(10,2)
        Case 9
              txtAux(9).Text = Trim(txtAux(9).Text)
'              txtAux(10).Tag = ""
              If txtAux(9).Text <> "" Then
                    If vEmpresa.TieneAnalitica Then
                        txtAux(9).Text = UCase(txtAux(9).Text)
                        If vParamAplic.ContabilidadNueva Then
                            txtAux2(Index).Text = DevuelveDesdeBD(conConta, "nomccost", "ccoste", "codccost", txtAux(9).Text, "T")
                        Else
                            txtAux2(Index).Text = DevuelveDesdeBD(conConta, "nomccost", "cabccost", "codccost", txtAux(9).Text, "T")
                        End If
                        If txtAux2(Index).Text = "" Then
                            MsgBox "No existe centro de coste: " & txtAux(9).Text, vbExclamation
                            txtAux(9).Text = ""
                            PonerFoco txtAux(9)
                        End If
                    
                    
                    Else
                        If Not IsNumeric(txtAux(9).Text) Then
                            MsgBox "Campo proveedor debe ser numérico", vbExclamation
                        Else
                            txtAux2(Index).Text = DevuelveDesdeBD(conAri, "nomprove", "sprove", "codprove", txtAux(9).Text)
                            If txtAux2(Index).Text = "" Then
                                MsgBox "No existe proveedor: " & txtAux(9).Text, vbExclamation
                                txtAux(9).Text = ""
                                PonerFoco txtAux(9)
                            End If
                        End If
                    End If
                End If
'                txtAux(10).Text = txtAux(10).Tag
'                txtAux(10).Tag = ""
                
    End Select
    
    If (Index = 3 Or Index = 4 Or Index = 6 Or Index = 7) Then 'Cant., Precio, Dto1, Dto2
        If txtAux(1).Text = "" Then Exit Sub
        txtAux(8).Text = CalcularImporte(txtAux(3).Text, txtAux(4).Text, txtAux(6).Text, txtAux(7).Text, vParamAplic.TipoDtos)
        PonerFormatoDecimal txtAux(8), 1
    End If
End Sub


Private Sub BotonMtoLineas(numTab As Integer, Cad As String)
    Me.SSTab1.Tab = numTab
    If Me.DataGrid1.visible Then
        If Me.Data2.Recordset.RecordCount < 1 Then
            MsgBox "La factura no tiene lineas.", vbInformation
            Exit Sub
        End If
        TituloLinea = Cad
    End If
    If vUsu.Nivel >= 1 Then
        MsgBox "No tiene suficientes privilegios. Consulte al administrador del sistema. ", vbExclamation
        Exit Sub
    End If
    If Me.cmdObserva3.Tag <> 0 Then
        'Debe poner las lineas
        MsgBox "Visualize las lineas de la factura", vbExclamation
        Exit Sub
    End If
    
    ModificaLineas = 0
    PonerModo 5
    PonerBotonCabecera True
End Sub


Private Function Eliminar() As Boolean
Dim Sql As String, LEtra As String
Dim B As Boolean
Dim vTipoMov As CTiposMov
Dim cContaFra As cContabilizarFacturas

    On Error GoTo FinEliminar

    B = False
    If Data1.Recordset.EOF Then Exit Function
        
    conn.BeginTrans
        
    'Eliminar en las tablas de la Contabilidad
    '------------------------------------------
    LEtra = ObtenerLetraSerie(Data1.Recordset!codtipom)
    
    Set cContaFra = New cContabilizarFacturas
    
    
    If LEtra <> "" Then

        Sql = "DELETE FROM "
        If vParamAplic.ContabilidadNueva Then
            If Data1.Recordset!codtipom = "FAZ" Then Sql = Sql & "ariconta" & vParamAplic.ContabilidadB & "."
            Sql = Sql & "cobros WHERE numserie='" & LEtra & "' AND numfactu=" & Data1.Recordset.Fields!Numfactu
            Sql = Sql & " AND fecfactu='" & Format(Data1.Recordset.Fields!FecFactu, FormatoFecha) & "'"
        Else
            Sql = Sql & " scobro WHERE numserie='" & LEtra & "' AND codfaccl=" & Data1.Recordset.Fields!Numfactu
            Sql = Sql & " AND fecfaccl='" & Format(Data1.Recordset.Fields!FecFactu, FormatoFecha) & "'"
        End If
        ConnConta.Execute Sql
        B = True
    Else
        B = False
    End If

    'Eliminar en tablas de factura de Ariges
    '------------------------------------------
    If B Then
        Sql = " " & ObtenerWhereCP(True)
    
        'Lineas de facturas (slifac)
        conn.Execute "Delete from slifac " & Sql
    
    
        'Lineas lotes
        conn.Execute "Delete from slifaclotes  " & Sql
        
       
        If InstalacionEsEulerTaxco Then
            conn.Execute "Delete from slifac_eu " & Sql
            
            conn.Execute "Delete from scafac_eu " & Sql
        End If
        
        'Campos
        conn.Execute "Delete from slifaccampos " & Sql
    
        'Lineas de cabeceras de albaranes de la factura
        conn.Execute "Delete from scafac1 " & Sql
        
        'Eliminar los vencimientos
        conn.Execute "Delete from svenci " & Sql
        
        'Cabecera de facturas (scafac)
        conn.Execute "Delete from " & NombreTabla & Sql
        
        'Decrementar contador si borramos la ult. factura
        Set vTipoMov = New CTiposMov
        vTipoMov.DevolverContador Data1.Recordset!codtipom, Val(Text1(0).Text)
        Set vTipoMov = Nothing
    End If
    
    B = True
    
FinEliminar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar Factura", Err.Description
        B = False
    End If
    If Not B Then
        conn.RollbackTrans
        Eliminar = False
    Else
        
        If LEtra <> "" Then
            'Preparao para eliminar
            If cContaFra.EstablecerValoresInciales(ConnConta) Then
                Sql = CStr(Data1.Recordset!FecFactu)
                cContaFra.FijarNumeroFactura CLng(Data1.Recordset!Numfactu), Year(Data1.Recordset!FecFactu), LEtra
            End If
        End If
        
        
        'De ARIGES
        conn.CommitTrans
        
        If cContaFra.RealizarContabilizacion Then
            If Data1.Recordset!codtipom <> "FAZ" Then
                ConnConta.BeginTrans
                'YA HE FIJADO LOS VALORES. En sql tengo la fecha factura
                If cContaFra.EliminarFRACLIcontab(True, CDate(Sql)) Then
                    ConnConta.CommitTrans
                Else
                    ConnConta.RollbackTrans
                End If
            End If 'FAZ
        End If
        Set cContaFra = Nothing
        Eliminar = True
    End If
End Function


Private Sub LimpiarDataGrids()
'Pone los Grids sin datos, apuntando a ningún registro
On Error Resume Next

    CargaGrid DataGrid2, Data3, False
    CargaGrid DataGrid1, Data2, False
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub PosicionarData()
Dim Indicador As String
Dim vWhere As String

    If Not Data1.Recordset.EOF Then
        'Hay datos en el Data1 bien porque se ha hecho VerTodos o una Busqueda
         vWhere = "(" & ObtenerWhereCP(False) & ")"
         If SituarDataMULTI(Data1, vWhere, Indicador) Then
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


Private Function ObtenerWhereCP(conWhere As Boolean) As String
Dim Sql As String

    On Error Resume Next
    
    Sql = " codtipom= '" & Text1(1).Text & "' and numfactu= " & Val(Text1(0).Text) & " and fecfactu='" & Format(Text1(2).Text, FormatoFecha) & "' "
    If conWhere Then Sql = " WHERE " & Sql
    ObtenerWhereCP = Sql
    
    If Err.Number <> 0 Then MuestraError Err.Number, "Obteniendo cadena WHERE.", Err.Description
End Function


Private Function MontaSQLCarga(enlaza As Boolean, Opcion As Byte) As String
'--------------------------------------------------------------------
' MontaSQlCarga:
'   Basándose en la información proporcionada por el vector de campos
'   crea un SQl para ejecutar una consulta sobre la base de datos que los
'   devuelva.
' Si ENLAZA -> Enlaza con el data1
'           -> Si no lo cargamos sin enlazar a ningun campo
'--------------------------------------------------------------------
Dim Sql As String
Dim B1 As Boolean
    
    If Opcion = 1 Then
        Sql = "SELECT codtipom, numfactu, fecfactu, numalbar, numlinea, codalmac, codartic, nomartic,"
        Sql = Sql & " ampliaci, cantidad,numbultos, precioar, origpre, dtoline1, dtoline2, importel ,"
        B1 = False
        If vParamAplic.NumeroInstalacion = 2 Then If vUsu.Nivel = 0 Then B1 = True
        If B1 Then
            'Diciembre 2014
            Sql = Sql & " comisionagente " ' if(pvpinferior=1,'Si','')"
        Else
            Sql = Sql & " codprovex"
        End If
        Sql = Sql & " codprovex, nomprove,codccost,numlote"
        Sql = Sql & " FROM slifac left join sprove on codprovex=codprove " 'lineas de factura
    ElseIf Opcion = 2 Then
        Sql = "SELECT codtipom,numfactu,fecfactu,codtipoa,numalbar, fechaalb, numpedcl,fecpedcl,sementre,numofert,fecofert, referenc, codenvio,codtraba, codtrab1, codtrab2,observa1,observa2,observa3,observa4,observa5,numtermi,numventa,fecenvio  "
        If vParamAplic.DireccionesEnvio Then Sql = Sql & ",coddiren"
        Sql = Sql & ",docarchiv "
        'Fitos
        
        Sql = Sql & ",ManipuladorNumCarnet,ManipuladorFecCaducidad,ManipuladorNombre,TipoCarnet,PideCliente "
        
        If InstalacionEsEulerTaxco Then Sql = Sql & ",fechaent,perrecep,dnirecep,latitud,longitud"
        Sql = Sql & ",refcontrato, refpedido"
        
        Sql = Sql & " FROM scafac1 " 'cabeceras albaranes de la factura
    End If
    
    If enlaza Then
        Sql = Sql & " " & ObtenerWhereCP(True)
        If Opcion = 1 Then Sql = Sql & " AND numalbar=" & Data3.Recordset.Fields!Numalbar
    Else
        'aNTES
        'SQL = SQL & " WHERE numfactu = -1 "
        'AHORA     Cambio sugerido por mangel para acelerar la entrada
        ' 2018 oCtubre.  Pongo where false. Es mas rapido que cualquier otra cosa
        If True Then
            Sql = Sql & " WHERE false"
        Else
            Sql = Sql & " WHERE codtipom is null and numfactu is null and fecfactu is null and codtipoa is null and numalbar is null "
            If Opcion = 1 Then Sql = Sql & " AND numlinea is null"
        End If
    End If
    Sql = Sql & " ORDER BY codtipom, numfactu, fecfactu,numalbar "
    If Opcion = 1 Then Sql = Sql & ", numlinea "
    MontaSQLCarga = Sql
End Function


Private Sub PonerModoOpcionesMenu(Modo As Byte)
'Activas unas Opciones de Menu y Toolbar según el modo en que estemos
Dim B As Boolean

        B = (Modo = 2) Or (Modo = 5 And ModificaLineas = 0)
        'Modificar
        Toolbar1.Buttons(5).Enabled = B
        Me.mnModificar.Enabled = B
        
        'eliminar
        Toolbar1.Buttons(6).Enabled = (Modo = 2) And vUsu.Nivel = 0     'False   'Marzo 2019. Lo pongo a false siempre
        Me.mnEliminar.Enabled = (Modo = 2) And vUsu.Nivel = 0     'False            ' ""
            
        B = (Modo = 2)
        'Mantenimiento lineas
        Toolbar1.Buttons(9).Enabled = B
        Me.mnLineas.Enabled = B
        
        
        'Cambiar numeros de lote
        Toolbar1.Buttons(13).Enabled = B
        Me.mnModLotes.Enabled = B
        
        If Toolbar1.Buttons(14).visible Then
            Toolbar1.Buttons(14).Enabled = B
            Me.mnEditarCampos.Enabled = B
        End If
        
        If Toolbar1.Buttons(15).visible Then
            Toolbar1.Buttons(15).Enabled = B
            Me.mnTipoPreciosLinea.Enabled = B
        End If

        Toolbar1.Buttons(16).Enabled = B
        Toolbar1.Buttons(18).Enabled = B
        If vParamAplic.ManipuladorFitosanitarios2 Then
            If Modo = 5 Then Toolbar1.Buttons(18).Enabled = True
        End If
        If Toolbar1.Buttons(19).visible Then Toolbar1.Buttons(19).Enabled = B
        
        
        'Imprimir
        Toolbar1.Buttons(10).Enabled = B
        Me.mnImprimir.Enabled = B
        Toolbar1.Buttons(11).Enabled = B
        mnImprimirAlbaran.Enabled = B
        
        
        B = (Modo >= 3) Or Modo = 1
        'Buscar
        Toolbar1.Buttons(1).Enabled = Not B
        Me.mnBuscar.Enabled = Not B
        'Ver Todos
        Toolbar1.Buttons(2).Enabled = Not B
        Me.mnVerTodos.Enabled = Not B
End Sub



Private Sub PonerDatosCliente(codClien As String, Optional nifClien As String)
Dim vCliente As CCliente
Dim Observaciones As String
Dim B As Boolean

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
            If vCliente.ClienteBloqueado(2, False) Then   '2: bloqueo es bloqueo
                If Modo = 3 Then
                    B = True
                ElseIf Modo = 4 Then
                     If (Val(Text1(4).Text) <> Val(Data1.Recordset!codClien)) Then B = True
                End If
                If B Then
                    LimpiarDatosCliente
                    Set vCliente = Nothing
                    Exit Sub
                End If
            End If
            
'            EsDeVarios = vCliente.EsClienteVarios(Text1(4).Text)
            EsDeVarios = vCliente.DeVarios
            BloquearDatosCliente (EsDeVarios)
        
            If Modo = 4 And EsDeVarios Then 'Modificar
                'si no se ha modificado el cliente no hacer nada
                If CLng(Text1(4).Text) = CLng(Data1.Recordset!codClien) Then
                    Set vCliente = Nothing
                    Exit Sub
                End If
            End If
        
        
'            If Actualizar = False And EsDeVarios = False Then Exit Sub
            
'            If (Not EsDeVarios) Or (EsDeVarios And modo = 3) Then
            Text1(4).Text = Format(vCliente.Codigo, "000000")
            If (Modo = 3) Or (Modo = 4) Then
                Text1(5).Text = vCliente.Nombre  'Nom clien
                Text1(8).Text = vCliente.Domicilio
                Text1(9).Text = vCliente.CPostal
                Text1(10).Text = vCliente.Poblacion
                Text1(11).Text = vCliente.Provincia
                Text1(6).Text = vCliente.NIF
                Text1(7).Text = DBLet(vCliente.TfnoClien, "T")
            End If
            
            'insertar
            If Modo = 3 Then Text1(15).Text = vCliente.ForPago

            Observaciones = DBLet(vCliente.Observaciones)
            If Observaciones <> "" Then
                MsgBox Observaciones, vbInformation, "Observaciones del cliente"
            End If
                
            'cuenta bancaria
            Text1(18).Text = vCliente.Banco
            FormateaCampo Text1(18)
            Text1(19).Text = vCliente.Sucursal
            FormateaCampo Text1(19)
            Text1(20).Text = vCliente.DigControl
            Text1(21).Text = vCliente.CuentaBan
            Text1(46).Text = vCliente.Iban
            'Comprobar si el cliente tiene cobros pendientes
            ComprobarCobrosCliente codClien, Text1(1).Text
        End If
    Else
        LimpiarDatosCliente
        PonerFoco Text1(4)
    End If
    Set vCliente = Nothing

EPonerDatos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poner Datos Cliente", Err.Description
End Sub


Private Sub PonerDatosClienteVario(nifClien As String)
Dim vCliente As CCliente
Dim B As Boolean
   
    If nifClien = "" Then Exit Sub
   
    Set vCliente = New CCliente
    B = vCliente.LeerDatosCliVario(nifClien)
    Text1(5).Text = vCliente.Nombre  'Nom clien
    Text1(8).Text = vCliente.Domicilio
    Text1(9).Text = vCliente.CPostal
    Text1(10).Text = vCliente.Poblacion
    Text1(11).Text = vCliente.Provincia
    Text1(7).Text = DBLet(vCliente.TfnoClien, "T")
            
    If Not B Then PonerFoco Text1(6)
    Set vCliente = Nothing
End Sub



Private Sub LimpiarDatosCliente()
Dim i As Byte



    For i = 4 To 13
        Text1(i).Text = ""
    Next i
    'If (Modo = 3 Or Modo = 4) Then PonerFoco Text1(4)

End Sub
    
    
Private Sub BotonImprimir(OpcionListado As Byte)
Dim cadFormula As String
Dim cadParam As String
Dim numParam As Byte
Dim cadSelect As String 'select para insertar en tabla temporal
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim ImprimeDirecto As Boolean
Dim NumCopias As Integer

    If Text1(0).Text = "" Then
        MsgBox "Debe seleccionar una Factura para Imprimir.", vbInformation
        Exit Sub
    End If
    
    cadFormula = ""
    cadParam = ""
    cadSelect = ""
    numParam = 0
    NumCopias = 1
    
    '===================================================
    '============ PARAMETROS ===========================
    If (OpcionListado = 53) Then
        If Text1(1).Text = "FAZ" Then
            'Factura B
            indRPT = 30
        
        'EULER
        ElseIf Text1(1).Text = "FAO" Then
            indRPT = 78 'Orden trabajo
        ElseIf Text1(1).Text = "FAE" Then
            indRPT = 79 'trabajo exterior
        'TELEFONIA
        ElseIf Text1(1).Text = "FAT" Then
            indRPT = 63 'Facturas telefonia
            
        ElseIf Text1(1).Text = "FPY" Then
              indRPT = 12 'Facturas Clientes
            If lwCostes.ListItems.Count > 0 Then
                If MsgBox("Impresion formato proyectos?", vbQuestion + vbYesNo) = vbYes Then indRPT = 89
            
            End If
        Else
            indRPT = 12 'Facturas Clientes
            
            'Si es rectificativa
            If Text1(1).Text = "FRT" Then NumCopias = vParamAplic.NumCop_FraRectifica
            
        End If
        
        'En taxco
        If vParamAplic.NumeroInstalacion = vbTaxco Then
            'Facturas alvic
            cadParam = "|" & Trim(Text1(1).Text) & "|"
            If InStr(1, "|FA1|FA2|FA3|FAB|FAD|", cadParam) > 0 Then indRPT = 93
            If InStr(1, "|F01|F02|F03|F11|F12|F13|F21|F22|F23|F31|F32|F33|", cadParam) > 0 Then indRPT = 93
            cadParam = ""
        End If
        
    Else
        If (OpcionListado = 89) Then
            indRPT = OpcionListado
    
        ElseIf (OpcionListado = 94) Then
            indRPT = OpcionListado
        Else
            'OpcionListado = 53
            '-----------------------------------------------
            indRPT = 18 'Facturas Clientes TPV
        End If
    End If
    If Not PonerParamRPT2(indRPT, cadParam, numParam, nomDocu, ImprimeDirecto, pPdfRpt, pRptvMultiInforme) Then Exit Sub
      
      
      
    'PUNTO VERDE
    '--------------------------------------------------------------------------
    If vParamAplic.ArtReciclado <> "" Then
        cadParam = cadParam & "PuntoVerde= """ & vParamAplic.ArtReciclado & """|"
        numParam = numParam + 1
    End If
      
    'Nombre fichero .rpt a Imprimir
    If Not ImprimeDirecto Then frmImprimir.NombreRPT = nomDocu
        
    '===================================================
    '================= FORMULA =========================
    'Cadena para seleccion Nº de Factura
    '---------------------------------------------------
    If Text1(0).Text <> "" Then
        'Cod Tipo Movimiento
        devuelve = "{" & NombreTabla & ".codtipom}='" & Text1(1).Text & "'"
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
        
        'Nº Factura
        devuelve = "{" & NombreTabla & ".numfactu}=" & Val(Text1(0).Text)
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
        
        cadSelect = cadFormula
        
        'Fecha Factura
        devuelve = "{" & NombreTabla & ".fecfactu}= Date(" & Year(Text1(2).Text) & "," & Month(Text1(2).Text) & "," & Day(Text1(2).Text) & ")"
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
        'Fecha Factura en cadSelect
        devuelve = "{" & NombreTabla & ".fecfactu}= '" & Format(Text1(2).Text, FormatoFecha) & "'"
        If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Sub
    End If
   
    If Not HayRegParaInforme(NombreTabla, cadSelect) Then Exit Sub
     
     
     If ImprimeDirecto Then
        'Imrpime directo
        If MsgBox("Imprimir la factura?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        ImprimirDirectoFact cadSelect
     Else
     
     
        
        devuelve = DevuelveDesdeBD(conAri, "letraser", "stipom", "codtipom", Text1(1).Text, "T")
     
        If vParamAplic.NumeroInstalacion = vbTaxco Then
            If OpcionListado = 53 Or OpcionListado = 94 Then
                'EN TAXCO añadimos la marca de agua
                cadParam = cadParam & "pDuplicado=1|"
                numParam = numParam + 1
                
            End If
        End If
             
         With frmImprimir
                'Nuevo. Febrero 2010
                .outClaveNombreArchiv = devuelve & Format(Text1(0).Text, "000")
                .outCodigoCliProv = Text1(4).Text
                .outTipoDocumento = 2
                .SeleccionaRPTCodigo = pRptvMultiInforme
                .FormulaSeleccion = cadFormula
                .OtrosParametros = cadParam
                .NumeroParametros = numParam
                .NombrePDF = pPdfRpt
                .SoloImprimir = False
                .EnvioEMail = False
                .NumeroCopias = NumCopias
                .Opcion = OpcionListado
                .Titulo = IIf(indRPT = 89, "Impresion lineas especiales", "")
                .Show vbModal
        End With
    End If
End Sub



Private Sub BotonImprimirTicket()

Dim cadImpresion As String, Sql As String



    cadImpresion = "{scafac.codtipom}='" & Text1(1).Text & "' and {scafac.numfactu}=" & Text1(0).Text
    Sql = cadImpresion & " and {scafac.fecfactu}=" & DBSet(Text1(2).Text, "F")
    cadImpresion = cadImpresion & " and {scafac.fecfactu}=Date(" & Year(CDate(Text1(2).Text)) & "," & Month(CDate(Text1(2).Text)) & "," & Day(CDate(Text1(2).Text)) & ")"
    
    If Not HayRegParaInforme("scafac", Sql) Then Exit Sub
    


    


    
'    cadImpresion = cadImpresion & " and {scafac.fecfactu}=Date(" &Year(RSVenta!fecventa) & "," & Month(RSVenta!fecventa) & "," & Day(RSVenta!fecventa) & ")"
    Sql = "spatpvg.codclien=sclien.codclien AND 1"
    Sql = Trim(DevuelveDesdeBD(conAri, "nifclien", "sclien,spatpvg", Sql, "1"))
    Sql = "|CIFvarios= """ & Sql & """|"

    With frmVisReport
        .CambiaODBC = False
        .FormulaSeleccion = cadImpresion
        .SoloImprimir = False
        .OtrosParametros = Sql
        .NumeroParametros = 1
        .MostrarTree = False
        
        Sql = DevuelveDesdeBD(conAri, "documrpt", "scryst", "codcryst", "66")
        If Sql = "" Then Sql = "rTPVTicket.rpt"
      
      
        
        
        .Informe = App.Path & "\Informes\" & Sql
        .ConSubInforme = False
        .Opcion = 93
        .ExportarPDF = False
        .Show vbModal
   End With
   
'   If bImpre Then
'        'volver la impresora a la predeterminada
'        EstablecerImpresora NomImpre
'   End If
   
End Sub




Private Function ModificaAlbxFac() As Boolean
Dim Sql As String
Dim B As Boolean
    
    On Error GoTo EModificaAlb
    
    ModificaAlbxFac = False
    'comprobar datos OK de la scafac1
     B = CompForm(Me, 2) 'Comprobar formato datos ok de la cabecera alb: opcion=2
    If Not B Then Exit Function
    
    Sql = "UPDATE scafac1 SET codenvio=" & Text3(3).Text & ", "
    Sql = Sql & "codtraba=" & Text3(0).Text & ", "
    Sql = Sql & "codtrab1=" & DBSet(Text3(1).Text, "N", "S") & ", " 'Trab. pedido
    Sql = Sql & "codtrab2=" & Text3(2).Text & ", " 'Trab. Prep. Material
    Sql = Sql & "referenc=" & DBSet(Text3(16).Text, "T", "S") 'referencia cliente
    'Si hubiera que updaear fechaenvio
    
    If Me.FrameObserva.visible Then
        Sql = Sql & ", observa1=" & DBSet(Text3(9).Text, "T")
        Sql = Sql & ", observa2=" & DBSet(Text3(10).Text, "T")
        Sql = Sql & ", observa3=" & DBSet(Text3(11).Text, "T")
        Sql = Sql & ", observa4=" & DBSet(Text3(12).Text, "T")
        Sql = Sql & ", observa5=" & DBSet(Text3(13).Text, "T")
    End If
    Sql = Sql & ", docarchiv = " & Me.chkEnvio.Value
    Sql = Sql & ", PideCliente = " & Me.chkPedxCli.Value
    If vParamAplic.DireccionesEnvio Then Sql = Sql & ", coddiren=" & DBSet(Text3(18).Text, "N", "S")   'Direnvio
        
    If InstalacionEsEulerTaxco Then
        For kCampo = 23 To 27
            Sql = Sql & "," & RecuperaValor("fechaent|perrecep|dnirecep|latitud|longitud|", kCampo - 22)
            Sql = Sql & " = " & DBSet(Text3(kCampo).Text, IIf(kCampo = 23, "FH", IIf(kCampo > 25, "N", "T")), "S")
        Next kCampo
        kCampo = 0
       
    End If

    Sql = Sql & ", refcontrato=" & DBSet(Text3(28).Text, "T", "S")
    Sql = Sql & ", refpedido=" & DBSet(Text3(29).Text, "T", "S")
    
    Sql = Sql & ObtenerWhereCP(True)
    Sql = Sql & " AND codtipoa='" & Data3.Recordset.Fields!Codtipoa & "' AND numalbar=" & Data3.Recordset.Fields!Numalbar
    conn.Execute Sql
    ModificaAlbxFac = True
    
EModificaAlb:
    If Err.Number <> 0 Then MuestraError Err.Number, "Modificar Albaranes de factura", Err.Description
End Function


Private Function ModificarFactura(Optional sqlLineas As String) As Boolean
'si se ha modificado la linea de slifac, añadir a la transaccion la modificación de la linea y recalcular
Dim bol As Boolean
Dim MenError As String
Dim Sql As String, LEtra As String
Dim vFactura As CFactura
Dim recalcular As Boolean
Dim RecalDesdeRecFinan As Boolean
Dim CliVar As Boolean
Dim NoTocarEnTesoreria As Boolean
Dim TocarEnTesoreria As Boolean
    On Error GoTo EModFact

    
    'Comprobar si hay que recalcular la factura
    recalcular = False
    If sqlLineas <> "" Then
        'comprobamos si se ha modificado la linea del albaran (precio y descuentos)
        recalcular = True
        
    ElseIf CSng(Data1.Recordset!DtoPPago) <> CSng(DBSet(Text1(16).Text, "N")) Then
        'si se ha cambiado el dto ppago
        recalcular = True
    ElseIf CSng(Data1.Recordset!DtoGnral) <> CSng(DBSet(Text1(17).Text, "N")) Then
        'si se ha cambiado el descuento general
        recalcular = True
    ElseIf CSng(Data1.Recordset!TotalFac) <> CSng(Text1(38).Text) Then
        recalcular = True
    ElseIf CLng(Data1.Recordset!codClien) <> CLng(Text1(4).Text) Then
        'si se ha cambiado el cliente (bonificarab o no)
        recalcular = TieneBonificaciones
        
        
    'Abril 2015
    'Dejo el ultimo el de la forma de pago.
    'Si es tiket y solo cambia la forma de pago NO recalculo
     ElseIf CInt(Data1.Recordset!codforpa) <> CInt(Text1(15).Text) Then
        'si se ha cambiado la forma de pago
        'Octubre 2022
        'SIEMPRE , si cambia forpa, recalculamos
        'If Me.Data3.Recordset!Codtipoa <> "ATI" Then recalcular = True
        recalcular = True
    End If
    
    
    bol = True
    conn.BeginTrans
    ConnConta.BeginTrans
    
    
    'Marzo 2011
    bol = True
    If sqlLineas <> "" Then
        'actualizar el importe de la linea modificada
        MenError = "Modificando lineas de Factura."
        conn.Execute sqlLineas
    End If
    
    
    If vParamAplic.ArticuloRecargoFinanciero <> "" Then
        MenError = "Tratar recargo financiero"
        bol = TratarRecargoFinanciero(RecalDesdeRecFinan)
        If Not bol Then
            'NO ha ido bien. Cancelara la modificacion
            recalcular = False 'para que no haga nada y cancele todo
        Else
            'Si ha ido bien , y didec que hay que recalcular... pues a recalcular
            If RecalDesdeRecFinan Then recalcular = True
        End If
    End If
        
    
    
    
    If recalcular Then
        
        
        'recalcular las bases imponibles x IVA
        MenError = "Recalcular importes IVA"
        bol = CalcularDatosFactura
        
    
    End If
    
    
    
    If bol Then
'        ComprobarDatosTotales
        
        'modificamos la scafac
        MenError = "Modificando cabecera de factura"
        bol = ModificaDesdeFormulario(Me, 1)
        
        If bol Then
            'Si es cliente de varios actualizar datos cliente en tabla:sclvar
            MenError = "Modificando datos cliente varios"
            bol = ActualizarClienteVarios(Text1(4).Text, Text1(6).Text)
        End If
        
        If bol Then
            MenError = "Modificando albaranes de factura"
            'modificar la tabla: scafac1
            bol = ModificaAlbxFac
            
            If bol And Not recalcular Then
                'No hay que recalcular, pero HAY que volver a generar el cobro ya que ha cambiado de cliente
                If Val(Text1(4).Text) <> Val(Data1.Recordset!codClien) Then recalcular = True
            End If
                
            
            
            
            If bol And recalcular Then 'si se ha modificado la factura
                MenError = "Actualizando en Tesoreria"
                
                
                
                'borrar los vencimientos de ariges.svenci
                'y eliminar de tesoreria conta.scobros los registros de la factura(si existen en Tesoreria)
                
                'Eliminar los vencimientos
                '----------------------------------------
                Sql = ObtenerWhereCP(True)
                conn.Execute "Delete from svenci " & Sql
                
                'Eliminar de Tesoreria
                '----------------------------------------
'                SQL = ObtenerLetraSerie(Text1(1).Text)
'                SQL = "SELECT COUNT(*) FROM scobro WHERE numserie='" & SQL & "' and codfaccl=" & Text1(0).Text
'                SQL = SQL & " AND fecfaccl=" & DBSet(Text1(2).Text, "F")
'
'                If RegistrosAListar(SQL, conConta) Then
                    'antes de Eliminar en las tablas de la Contabilidad
                Set vFactura = New CFactura
                If vFactura.LeerDatos(Text1(1).Text, Text1(0).Text, Text1(2).Text) Then
                Else
                  bol = False
                End If
              
                If Text1(1).Text = "FAZ" Then bol = False
              
                If bol Then
                    
                    TocarEnTesoreria = True
                    If vParamAplic.ContabilidadNueva Then
                        Sql = DevuelveDesdeBD(conAri, "letraser", "stipom", "codtipom", Text1(1).Text, "T")
                        Sql = " numserie='" & Sql & "'"
                        Sql = Sql & " AND numfactu =" & Text1(0).Text
                        Sql = Sql & " AND fecfactu =" & DBSet(Text1(2).Text, "F") & " AND 1"
                        Sql = DevuelveDesdeBD(conConta, "impcobro", "cobros", Sql, "1")
                        If Sql <> "" Then
                            If CCur(Sql) <> 0 Then TocarEnTesoreria = False
                        End If
                    End If
                    
                    If Not TocarEnTesoreria Then vFactura.CuentaPrev = "" 'SI HACEMOS ESTO, NO GENERA EN tesoreria
                    
                    vFactura.NIF = Text1(6).Text
                    vFactura.NombreClien = Text1(5).Text
                    vFactura.DomicilioClien = Text1(8).Text
                    vFactura.CPostal = Text1(9).Text
                    vFactura.Poblacion = Text1(10).Text
                    vFactura.Provincia = Text1(11).Text
                    vFactura.Telefono = Text1(7).Text

                    
                    
                    
                
                    'Eliminar de la scobro
                    If vParamAplic.ContabilidadNueva Then
                        Sql = " cobros WHERE numserie='" & vFactura.LetraSerie & "' AND numfactu=" & Data1.Recordset.Fields!Numfactu
                        Sql = Sql & " AND fecfactu='" & Format(Data1.Recordset.Fields!FecFactu, FormatoFecha) & "'"
                    Else
                        Sql = " scobro WHERE numserie='" & vFactura.LetraSerie & "' AND codfaccl=" & Data1.Recordset.Fields!Numfactu
                        Sql = Sql & " AND fecfaccl='" & Format(Data1.Recordset.Fields!FecFactu, FormatoFecha) & "'"
                    End If
                    
                    If TocarEnTesoreria Then ConnConta.Execute "Delete from " & Sql
                    bol = True

                    'Volvemos a Insertar los Vencimientos de la Factura. Tabla: svenci
                    'Grabar en TESORERIA. Tabla de Contabilidad: sconta.scobros
                    If bol Then
                        vFactura.Agente = Text1(14).Text
                        bol = vFactura.InsertarEnTesoreria("", MenError, True)
                    End If
                End If
                Set vFactura = Nothing
                
                'pongo bol a true para que siga
                If Text1(1).Text = "FAZ" Then bol = True
                
            End If
'            End If
        End If
    End If

EModFact:
     If Err.Number <> 0 Then
        Screen.MousePointer = vbDefault
        bol = False
    End If
    If bol Then
        conn.CommitTrans
        ConnConta.CommitTrans
        ModificarFactura = True
    Else
        conn.RollbackTrans
        ConnConta.RollbackTrans
        ModificarFactura = False
        MenError = "Actualizando Factura." & vbCrLf & "----------------------------" & vbCrLf & MenError
        MuestraError Err.Number, MenError, Err.Description
    End If
End Function


Private Function CalcularDatosFactura() As Boolean
Dim i As Integer
Dim vFactu As CFactura
Dim FacOK As Boolean
Dim CambiaIVA As Boolean
Dim C As String

    'Limpiar en el form los datos calculados de la factura
    'y volvemos a recalcular
    For i = 22 To 38
         Text1(i).Text = ""
    Next i
    
    
    Set vFactu = New CFactura
    vFactu.DtoPPago = CCur(ComprobarCero(Text1(16).Text))
    vFactu.DtoGnral = CCur(ComprobarCero(Text1(17).Text))
    vFactu.Cliente = Text1(4).Text
    
    CambiaIVA = False
    If Text1(1).Text = "FRT" Then
        'Facturas rectificativas. EXISTE la posibilidad que haya cambio de IVA en funcion de la fecha
        'a la factura que rectifica
        'Vamos a intentar sacar la fecha
        If Not Data3.Recordset Is Nothing Then
            If Not Data3.Recordset.EOF Then
                C = DBLet(Data3.Recordset!observa1, "T")
                If C <> "" Then
                    If Len(C) > 10 Then
                        'Esto es un poco A PIÑON
                        C = Right(C, 10) 'En observa1 estan las 10 ultimas posiciones para la fecha de la factura que rectigfico en su momento
                        If IsDate(C) Then
                            If CDate(C) < vParamAplic.FechaCambioIva Then CambiaIVA = True
                        End If
                    End If
                End If
            End If
        End If
    Else
        If CDate(Text1(2).Text) < vParamAplic.FechaCambioIva Then CambiaIVA = True
    End If
    vFactu.codtipom = Text1(1).Text  'abril 2015
    
    'Diciembre 2016
    
    
                                                                              'Que coja el IVA wque tiene, sin cambios
    If vFactu.CalcularDatosFactura(ObtenerWhereCP(False), NombreTabla, NomTablaLineas, CambiaIVA) Then
        
        FacOK = True
        Text1(22).Text = vFactu.BrutoFac
        Text1(23).Text = vFactu.ImpPPago
        Text1(24).Text = vFactu.ImpGnral
        Text1(25).Text = vFactu.BaseImp
        Text1(26).Text = QuitarCero(vFactu.TipoIVA1)
        Text1(27).Text = QuitarCero(vFactu.TipoIVA2)
        Text1(28).Text = QuitarCero(vFactu.TipoIVA3)
        Text1(29).Text = vFactu.PorceIVA1
        Text1(30).Text = vFactu.PorceIVA2
        Text1(31).Text = vFactu.PorceIVA3
        Text1(32).Text = vFactu.BaseIVA1
        Text1(33).Text = vFactu.BaseIVA2
        Text1(34).Text = vFactu.BaseIVA3
        Text1(35).Text = vFactu.ImpIVA1
        Text1(36).Text = vFactu.ImpIVA2
        Text1(37).Text = vFactu.ImpIVA3
        Text1(38).Text = vFactu.TotalFac
        
        'Sept 2012
        'Los ivas con RE
        Text1(39).Text = vFactu.ImpIVA1RE
        Text1(40).Text = vFactu.PorceIVA1RE
        Text1(41).Text = vFactu.ImpIVA2RE
        Text1(42).Text = vFactu.PorceIVA2RE
        Text1(43).Text = vFactu.ImpIVA2RE
        Text1(44).Text = vFactu.PorceIVA3RE
        
        
        FormatoDatosTotales
    Else
        FacOK = False
        MuestraError Err.Number, "Calculando Totales", Err.Description
    End If
    Set vFactu = Nothing
    CalcularDatosFactura = FacOK
End Function


Private Sub FormatoDatosTotales()
Dim i As Byte
Dim L As Boolean
Dim N As Byte

    For i = 22 To 25
        Text1(i).Text = QuitarCero(Text1(i).Text)
        Text1(i).Text = Format(Text1(i).Text, FormatoImporte)
    Next i
    
    'Desglose B.Imponible por IVA
    For i = 32 To 34
        L = True
        'Para el RE equivalencia
        If i = 32 Then
            N = 7
        Else
            If i = 33 Then
                N = 8
            Else
                N = 9
            End If
        End If
        
        
        If Text1(i).Text <> "" Then
             If CSng(Text1(i).Text) = 0 And Text1(i - 6).Text = "" Then
                Text1(i).Text = QuitarCero(Text1(i).Text)
                Text1(i - 3).Text = QuitarCero(Text1(i - 3).Text)
                Text1(i - 6).Text = QuitarCero(Text1(i - 6).Text)
                Text1(i + 3).Text = QuitarCero(Text1(i).Text)
            Else
                Text1(i).Text = Format(Text1(i).Text, FormatoImporte)
                Text1(i - 3) = Format(Text1(i - 3).Text, FormatoDescuento)
    '            Text3(i - 6) = Format(Text3(i - 6).Text, "000")
                Text1(i + 3).Text = Format(Text1(i + 3).Text, FormatoImporte)
            End If
            
            'IVA RE
           
            If Text1(i).Text <> "" Then  'Si lleva base imponimbe
                If Text1(i + N + 1).Text <> "" Then
                    If CSng(Text1(i + N + 1).Text) <> 0 Then L = False
                End If
            End If 'de si lleva base imponible
        End If
        
        
        
        If L Then
        
            Text1(i + N).Text = QuitarCero(Text1(i + N).Text)
            Text1(i + N + 1).Text = QuitarCero(Text1(i + N + 1).Text)
        Else
            Text1(i + N).Text = Format(Text1(i + N).Text, FormatoImporte)
            Text1(i + N + 1).Text = Format(Text1(i + N + 1).Text, FormatoImporte)
        End If


            
        
    Next i
End Sub



Private Sub ComprobarDatosTotales()
Dim i As Byte

    For i = 22 To 25
        Text1(i).Text = ComprobarCero(Text1(i).Text)
    Next i
End Sub

'   FactContabilizada_:
'       0.- Se puede TODO. NO esta contabilizada
'       1.- Se puede "parcial", lo que hacia antes
'       2.- Solo se puede trabajador, zona,referencias.....
'       3.- NADA
Private Function FactContabilizada(ByRef EstaEnTesoreria As String) As Boolean
Dim LEtra As String, numasien As String
Dim cControlFra As CControlFacturaContab
    On Error GoTo EContab
    
    TipoModificacionFactura = 3
    
    If vUsu.Nivel > 0 And Val(Data1.Recordset!intconta) = 1 Then
        MsgBox "Factura contabilizada", vbExclamation
        FactContabilizada = True
        Exit Function
    End If
    
    
    'Cojo la letra de serie
    LEtra = ObtenerLetraSerie(Text1(1).Text)
    
    
    
    
        Set cControlFra = New CControlFacturaContab
        numasien = ""
        
        'Con estos dos NO dejo pasar
        BuscaChekc = cControlFra.FechaCorrectaContabilizazion(ConnConta, Text1(2))
        If BuscaChekc <> "" Then numasien = numasien & "- " & BuscaChekc & vbCrLf
        If numasien <> "" Then
            FactContabilizada = True
            MsgBox numasien, vbExclamation
            Exit Function
        End If
        
        'Dejare modificar SIN importes
        BuscaChekc = cControlFra.FechaCorrectaIVA(ConnConta, Text1(2))
        If BuscaChekc <> "" Then numasien = numasien & "- " & BuscaChekc & vbCrLf
        Set cControlFra = Nothing
        
        If numasien <> "" Then
            FactContabilizada = False
            TipoModificacionFactura = 2
            MsgBox numasien, vbExclamation
            Exit Function
        End If
        numasien = ""
    
    
    
        'Primero comprobaremos que esta el cobro en contabilidad
        If "FAZ" <> Text1(1).Text Then
            EstaEnTesoreria = ""
            If Not ComprobarCobroArimoney(EstaEnTesoreria, LEtra, CLng(Text1(0).Text), CDate(Text1(2).Text)) Then
                TipoModificacionFactura = 1
                FactContabilizada = False
                Exit Function
            End If
        
        Else
            MsgBox "Compruebe vencimientos", vbExclamation
        End If

    'comprabar que se puede modificar/eliminar la factura
    If Me.Check1.Value = 1 And "FAZ" <> Text1(1).Text Then 'si esta contabilizada
        'comprobar en la contabilidad si esta contabilizada
      
        If LEtra <> "" Then
        
            If vParamAplic.ContabilidadNueva Then
                'Aunque en la nueva contabiliad SIEMPRE esta con apunte.
                numasien = DevuelveDesdeBDNew(conConta, "factcli", "numasien", "numserie", LEtra, "T", , "numfactu", Text1(0).Text, "N", "anofactu", Year(Text1(2).Text), "N")
            Else
                numasien = DevuelveDesdeBDNew(conConta, "cabfact", "numasien", "numserie", LEtra, "T", , "codfaccl", Text1(0).Text, "N", "anofaccl", Year(Text1(2).Text), "N")
            End If
            If Val(ComprobarCero(numasien)) <> 0 Then
'                FactContabilizada = True
'                MsgBox "La factura esta contabilizada y no se puede modificar.", vbInformation
'                Exit Function
                
            Else
                numasien = ""
            End If
            
            
            
            
        Else
'            MsgBox "Las factura de venta no tienen asignada una letra de serie", vbInformation
            numasien = ""
        End If
        
        LEtra = "La factura esta en la contabilidad"
        If numasien <> "" Then LEtra = LEtra & vbCrLf & "Nº asiento: " & numasien
        LEtra = LEtra & vbCrLf & vbCrLf & "¿Continuar?"
        
        numasien = String(50, "*") & vbCrLf
        numasien = numasien & numasien & vbCrLf & vbCrLf
        LEtra = numasien & LEtra & vbCrLf & vbCrLf & numasien
        If MsgBox(LEtra, vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
            FactContabilizada = False
            TipoModificacionFactura = 1
        Else
            FactContabilizada = True
            
        End If
    Else
        FactContabilizada = False
        TipoModificacionFactura = 0
    End If
    
    
    
EContab:
    If Err.Number <> 0 Then MuestraError Err.Number, "Comprobar Factura contabilizada", Err.Description
End Function


Private Sub TxtAux3_GotFocus(Index As Integer)
    ConseguirFoco txtAux3(Index), Modo
End Sub

Private Sub TxtAux3_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index <> 0 And KeyCode <> 38 Then KEYdown KeyCode
End Sub

Private Sub TxtAux3_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub TxtAux3_LostFocus(Index As Integer)
    If Not PerderFocoGnral(txtAux3(Index), Modo) Then Exit Sub
End Sub


Private Sub BloquearDatosCliente(bol As Boolean)
Dim i As Byte

    'bloquear/desbloquear campos de datos segun sea de varios o no
    If Modo <> 5 Then
        Me.imgBuscar(1).visible = bol
        Me.imgBuscar(1).Enabled = bol
        Me.imgBuscar(2).Enabled = bol
        
        For i = 5 To 11 'si no es de varios no se pueden modificar los datos
            BloquearTxt Text1(i), Not bol
        Next i
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



Private Function ObtenerSelFactura() As String
Dim Cad As String
Dim Rs As ADODB.Recordset

    On Error Resume Next

    Cad = ""
    If Me.DesdeFichaCliente Then
        '
        Cad = " WHERE codtipom='" & hcoCodTipoM & "' AND numfactu= " & hcoCodMovim & " AND fecfactu=" & DBSet(hcoFechaMov, "F")
        
    Else
        'Tengo YA el codigo de la factura
                '******************************************************
                'laura: esto se puede comentar, ya no hay movimiento FTI en la smoval
                If hcoCodTipoM = "FTI" Then
                    'no hay albaran directamente va a factura de ticket
                    
                    'ver si lo encontramos como factura: codtipom, numfactu,fecfactu
                    Cad = "SELECT COUNT(*) FROM scafac "
                    Cad = Cad & " WHERE codtipom='" & hcoCodTipoM & "' AND numfactu= " & hcoCodMovim & " AND fecfactu=" & DBSet(hcoFechaMov, "F")
                    If RegistrosAListar(Cad) > 0 Then
                        Cad = " WHERE codtipom='" & hcoCodTipoM & "' AND numfactu= " & hcoCodMovim & " AND fecfactu=" & DBSet(hcoFechaMov, "F")
                    Else
                        Cad = ""
                    End If
                Else
                    If hcoCodTipoM = "FAM" Then
                        Cad = "  WHERE codtipom='" & hcoCodTipoM & "' AND numfactu= " & hcoCodMovim & " AND fecfactu=" & DBSet(hcoFechaMov, "F")
                    End If
                End If
                
                
                '******************************************************
                If Cad = "" Then
                    'En la smoval estaba e mov. de ALbaran
                    Cad = "SELECT codtipom,numfactu,fecfactu FROM scafac1 "
                    Cad = Cad & " WHERE codtipoa=" & DBSet(hcoCodTipoM, "T") & " AND numalbar=" & hcoCodMovim & " AND fechaalb=" & DBSet(hcoFechaMov, "F")
                    
                    Set Rs = New ADODB.Recordset
                    Rs.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                    If Not Rs.EOF Then 'where para la factura
                        Cad = " WHERE codtipom='" & Rs!codtipom & "' AND numfactu= " & Rs!Numfactu & " AND fecfactu=" & DBSet(Rs!FecFactu, "F")
                    Else
                        Cad = " WHERE false "
                    End If
                    Rs.Close
                    Set Rs = Nothing
                End If
    
    End If
    ObtenerSelFactura = Cad
End Function



Private Function PonerDptoEnCliente() As Boolean
Dim vClien As CCliente
Dim NomDpto As String

    Set vClien = New CCliente
    vClien.Codigo = Text1(4).Text
    'si existe el departamento para el cliente
    If vClien.DptoCliente(Text1(12).Text, NomDpto) Then
        Text1(13).Text = NomDpto
        PonerDptoEnCliente = True
    Else
        PonerDptoEnCliente = False
    End If
    Set vClien = Nothing
End Function


Private Sub CargaCombo()
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim i As Byte
    
    Combo1.Clear
    
    Sql = "SELECT codtipom,nomtipom FROM stipom WHERE codtipom LIKE 'F%'"
    
    If vParamAplic.NumeroInstalacion = vbFenollar Then
    
        If Not HaMostradoCanal2_El_B Then Sql = Sql & " AND codtipom <>'FAZ'"
        
    Else
        'Para cualquiera menos root
        If (vUsu.Codigo Mod 1000) > 0 Then
            Sql = Sql & " AND codtipom"
            If Val(vUsu.AlmacenPorDefecto2) = vParamAplic.AlmacenB Then
                Sql = Sql & " = "
            Else
                Sql = Sql & "<>"
            End If
            Sql = Sql & "'FAZ'"
        End If
    End If
        
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not Rs.EOF
        Sql = Rs!nomtipom
        Sql = Replace(Sql, "Factura", "")
        Combo1.AddItem Rs!codtipom & "-" & Sql
        Combo1.ItemData(Combo1.NewIndex) = i
        i = i + 1
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
End Sub


Private Sub ImprimirAlbaran(OpcionListado As Byte)
Dim cadFormula As String
Dim cadParam As String
Dim numParam As Byte
Dim cadSelect As String 'select para insertar en tabla temporal
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String



    
    cadFormula = ""
    cadParam = ""
    cadSelect = ""
    numParam = 0
    
    '===================================================
    '============ PARAMETROS ===========================
    indRPT = 42
    If Not PonerParamRPT2(indRPT, cadParam, numParam, nomDocu, False, pPdfRpt, pRptvMultiInforme) Then Exit Sub
      
      
      
    'PUNTO VERDE
    '--------------------------------------------------------------------------
    If vParamAplic.ArtReciclado <> "" Then
        cadParam = cadParam & "PuntoVerde= """ & vParamAplic.ArtReciclado & """|"
        numParam = numParam + 1
    End If
      
    'Nombre fichero .rpt a Imprimir
    frmImprimir.NombreRPT = nomDocu
        
    '===================================================
    '================= FORMULA =========================
    'Cadena para seleccion Nº de Factura
    '---------------------------------------------------
    
        'Cod Tipo Movimiento
        devuelve = "{" & NombreTabla & ".codtipom}='" & Text1(1).Text & "'"
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
        
        'Nº Factura
        devuelve = "{" & NombreTabla & ".numfactu}=" & Val(Text1(0).Text)
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
        
        
        
        'cODTIPOA
        devuelve = "{scafac1.codtipoa}=" & DBSet(Data3.Recordset!Codtipoa, "T")
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
        
        'Numalbar
        devuelve = "{scafac1.numalbar}=" & DBSet(Data3.Recordset!Numalbar, "N")
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
        
        
        
        cadSelect = cadFormula
        'Fecha Factura
        devuelve = "{" & NombreTabla & ".fecfactu}= Date(" & Year(Text1(2).Text) & "," & Month(Text1(2).Text) & "," & Day(Text1(2).Text) & ")"
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
        'Fecha Factura en cadSelect
        devuelve = "{" & NombreTabla & ".fecfactu}= '" & Format(Text1(2).Text, FormatoFecha) & "'"
        If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Sub
    
   
        'If Not HayRegParaInforme(NombreTabla, cadSelect) Then Exit Sub
        '=========================================================================
        'ipo de IVA
        'que se aplica a ese cliente
        devuelve = DevuelveDesdeBDNew(conAri, "sclien", "tipoiva", "codclien", Text1(4).Text, "N")
        If devuelve <> "" Then
            cadParam = cadParam & "pTipoIVA= " & devuelve & "|"
            numParam = numParam + 1
        End If
         
     
     
        
        devuelve = DevuelveDesdeBD(conAri, "letraser", "stipom", "codtipom", Text1(1).Text, "T")
     
     
         With frmImprimir
                'Nuevo. Febrero 2010
                '.outClaveNombreArchiv = devuelve & Format(Text1(0).Text, "000")
                '.outCodigoCliProv = Text1(4).Text
                '.outTipoDocumento = 2
                
                .outClaveNombreArchiv = Data3.Recordset!Codtipoa & Format(Data3.Recordset!Numalbar, "0000000")
                .outCodigoCliProv = Text1(4).Text
                .outTipoDocumento = 7
                .SeleccionaRPTCodigo = pRptvMultiInforme
                
                
                .FormulaSeleccion = cadFormula
                .OtrosParametros = cadParam
                .NumeroParametros = numParam
                .NombrePDF = pPdfRpt
                .SoloImprimir = False
                .EnvioEMail = False
                .Opcion = 45
                .Titulo = "Albarán facturado"
                .Show vbModal
        End With
    
End Sub


'En vTesoreria pondremos como estaba el recibo
'Es decir. El  msgbox que pondra al final lo guardo en esta variable
Private Function ComprobarCobroArimoney(vTesoreria As String, LEtra As String, Codfaccl As Long, Fecha As Date) As Boolean
Dim vR As ADODB.Recordset
Dim Cad As String


On Error GoTo EComprobarCobroArimoney
    ComprobarCobroArimoney = False
    Set vR = New ADODB.Recordset
    
    If vParamAplic.ContabilidadNueva Then
        Cad = "Select * from cobros WHERE numserie='" & LEtra & "'"
        Cad = Cad & " AND numfactu =" & Codfaccl
        Cad = Cad & " AND fecfactu =" & DBSet(Fecha, "F")
    Else
        Cad = "Select * from scobro WHERE numserie='" & LEtra & "'"
        Cad = Cad & " AND codfaccl =" & Codfaccl
        Cad = Cad & " AND fecfaccl =" & DBSet(Fecha, "F")
    
    End If
    

    '
    vTesoreria = ""
    vR.Open Cad, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cad = ""
    If vR.EOF Then
        vTesoreria = "NO se ha encotrado ningun vencimiento en la tesoreria"
    Else
        While Not vR.EOF
            Cad = ""
            If DBLet(vR!codrem, "T") <> "" Then
                Cad = "El cobro asociado a la factura esta remesado(" & vR!codrem & ")"
            Else
                If DBLet(vR!recedocu, "N") = 1 Then
                    Cad = "Documento recibido"
                Else
                    
                        If DBLet(vR!transfer, "N") = 1 Then
                            Cad = "Esta en una transferencia"
                        Else
                           If DBLet(vR!impcobro, "N") <> 0 Then Cad = "Tiene cobro realizado: " & vR!impcobro
                        
                            
                                    'Si hubeira que poner mas coas iria aqui
                        End If 'transfer
                    
                End If 'recdedocu
            End If 'remesado
            If Cad <> "" Then vTesoreria = vTesoreria & "Vto: " & vR!numorden & "      " & Cad & vbCrLf
            vR.MoveNext
        Wend
    End If
    vR.Close
    
    
    
    If vTesoreria <> "" Then
        Cad = vTesoreria & vbCrLf & vbCrLf
        If vUsu.Nivel >= 1 Then
            MsgBox Cad, vbExclamation
        Else
            Cad = Cad & vbCrLf & vbCrLf & "Debe revisar la tesorería"
            Cad = Cad & "¿Continuar?"
            If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then ComprobarCobroArimoney = True
        End If
    Else
        ComprobarCobroArimoney = True
    End If
            
EComprobarCobroArimoney:
    If Err.Number <> 0 Then MuestraError Err.Number
    Set vR = Nothing
End Function


Private Function TieneBonificaciones() As Boolean
Dim Cad As String

    On Error GoTo ETieneBonificaciones
    TieneBonificaciones = False
        
    
        Cad = ObtenerWhereCP(True)
        Cad = Cad & " AND numalbar=" & Data3.Recordset.Fields!Numalbar
        Cad = "codartic in (Select codartic from slifac " & Cad & ") AND 1"
        
        
        Cad = DevuelveDesdeBD(conAri, "count(*)", "sbonif", Cad, "1")
        If Cad = "" Then Cad = "0"
        If Val(Cad) > 0 Then TieneBonificaciones = True
        
        
        Exit Function
ETieneBonificaciones:
    MuestraError Err.Number, "Comprobando bonificaciones"
    
    
End Function






Private Function TratarRecargoFinanciero(ByRef HayQueRecalcularElImporte As Boolean) As Boolean
Dim RecargoFinanciero As Boolean
Dim Aux As String
Dim C2 As String
Dim Raux As ADODB.Recordset
Dim PorceRecargo As Currency
Dim vFactu As CFactura
Dim CambiaIVA As Boolean

        On Error GoTo eTratarRecargoFinanciero
        TratarRecargoFinanciero = True
        HayQueRecalcularElImporte = False
        
        
        'Comprobamos si el "moviemiento" lleva recargo financiero. Si no me salgo y lo dejo to tal y como esta
        RecargoFinanciero = True
        If Data1.Recordset!codtipom = "FRT" Or Data1.Recordset!codtipom = "FAZ" Or Data1.Recordset!codtipom = "FAI" Then Exit Function
            
        
        'VEo si tiene la factura recargo financiero y me cargo la linea
        Aux = ObtenerWhereCP(False)
        Aux = Replace(Aux, "scafac", "slifac")
        Aux = DevuelveDesdeBD(conAri, "count(*)", "slifac", Aux & " AND codartic ", vParamAplic.ArticuloRecargoFinanciero, "T")
        If Aux = "" Then Aux = "0"
        If Val(Aux) > 0 Then
            Aux = ObtenerWhereCP(True)
            Aux = Aux & " AND codartic = " & DBSet(vParamAplic.ArticuloRecargoFinanciero, "T")
            conn.Execute "DELETE from slifac " & Aux
            Espera 0.2
            HayQueRecalcularElImporte = True
        End If
        
    
        
        
        Aux = DevuelveDesdeBD(conAri, "porgasfi", "sforpa", "codforpa", Text1(15).Text)
        If Aux = "" Then Aux = "0"
        If CCur(Aux) = 0 Then
            RecargoFinanciero = False 'NO METO Recargo financiero
        Else
            PorceRecargo = CCur(Aux)
        End If
        If Not RecargoFinanciero Then Exit Function   'me salgo

        
        'Compruebamos si el cliente tiene recargo financiero.
        'Si lleva recargo financiero, pero el cliente no se le aplica..
        Aux = DevuelveDesdeBD(conAri, "Recargofinanciero", "sclien", "codclien", Text1(4).Text)
        If Aux = "" Then Aux = "0"
        If Aux = "0" Then RecargoFinanciero = False
        If Not RecargoFinanciero Then Exit Function
        
        
        'Sept 2012
        'Habra que ver si hace cambio IVA
        CambiaIVA = False
        If Text1(1).Text = "FRT" Then
            'Facturas rectificativas. EXISTE la posibilidad que haya cambio de IVA en funcion de la fecha
            'a la factura que rectifica
            'Vamos a intentar sacar la fecha
            If Not Data3.Recordset Is Nothing Then
                If Not Data3.Recordset.EOF Then
                    Aux = DBLet(Data3.Recordset!observa1, "T")
                    If Aux <> "" Then
                        If Len(Aux) > 10 Then
                            'Esto es un poco A PIÑON
                            Aux = Right(Aux, 10) 'En observa1 estan las 10 ultimas posiciones para la fecha de la factura que rectigfico en su momento
                            If IsDate(Aux) Then
                                If CDate(Aux) < vParamAplic.FechaCambioIva Then CambiaIVA = True
                            End If
                        End If
                    End If
                End If
            End If
        Else
            If CDate(Text1(2).Text) < vParamAplic.FechaCambioIva Then CambiaIVA = True
        End If
        
        
        
        
        
        
        
        'Hay calcular el RE financiero
        Set vFactu = New CFactura
        If Not vFactu.CalcularDatosFactura(ObtenerWhereCP(False), NombreTabla, NomTablaLineas, CambiaIVA) Then
            'Error calculando el total factura
            TratarRecargoFinanciero = False
            
        
        
        Else
            HayQueRecalcularElImporte = True
            Set Raux = New ADODB.Recordset
            
            Aux = ObtenerWhereCP(False)
            Aux = "select codtipom,numfactu,fecfactu,codtipoa,numalbar from scafac1 where " & Aux & " order by codtipoa asc,numalbar desc"
            Raux.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Raux.EOF Then
                Raux.Close
                Err.Raise 513, "Error obteniedo albaran de factura para el rec. financiero"
            Else
                'montanmos el SQL del insert en buscchec
                'insert into `slifac` (`codtipom`,`numfactu`,`fecfactu`,`codtipoa`,`numalbar`,`numlinea`,
                BuscaChekc = "(" & DBSet(Raux!codtipom, "T") & "," & DBSet(Raux!Numfactu, "N") & "," & DBSet(Raux!FecFactu, "F") & ","
                BuscaChekc = BuscaChekc & DBSet(Raux!Codtipoa, "T") & "," & DBSet(Raux!Numalbar, "N") & ","
                Aux = ObtenerWhereCP(True)
                Aux = Aux & " AND codtipoa = " & DBSet(Raux!Codtipoa, "T") & " AND numalbar = " & Raux!Numalbar
            End If
            Raux.Close
            Aux = "Select max(numlinea),min(codalmac) FROM slifac " & Aux
            Raux.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            Aux = "0"
            C2 = "1"  'Almacen
            If Not Raux.EOF Then
                Aux = DBLet(Raux.Fields(0), "N")
                'no deberia ser Null, pero por si acaso
                If Not IsNull(Raux.Fields(1)) Then C2 = Raux.Fields(1)
            End If
            Raux.Close
            Set Raux = Nothing
            ''insert into `slifac` (`.......... ,`numlinea`,codalmac`
            BuscaChekc = BuscaChekc & Val(Aux) + 1 & "," & C2 & ","
        
            Aux = DBSet(vParamAplic.ArticuloRecargoFinanciero, "T") & ","
            ',`codartic`,`nomartic`,`ampliaci`,`cantidad`,`numbultos`,`precioar`,`dtoline1`,`dtoline2`,`importel`,`origpre`
            PorceRecargo = Round((PorceRecargo * CCur(vFactu.TotalFac)) / 100, 2)
            C2 = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", vParamAplic.ArticuloRecargoFinanciero, "T")
            If C2 = "" Then C2 = "** SIN REFERENCIA **"
            Aux = Aux & DBSet(C2, "T") & ",NULL,1,0," & TransformaComasPuntos(CStr(PorceRecargo)) & ",0,0," & TransformaComasPuntos(CStr(PorceRecargo)) & ",'M',"
            BuscaChekc = BuscaChekc & Aux
            
            'insert into `slifac` (`codtipom`,`numfactu`,`fecfactu`,`codtipoa`,`numalbar`,`numlinea`,
            '`codalmac`,`codartic`,`nomartic`,`ampliaci`,`cantidad`,`numbultos`,`precioar`,`dtoline1`,`dtoline2`,`importel`,`origpre`
            '`precioiv`,`preciomp`,`preciost`,`preciouc`,`codproveX)
            BuscaChekc = BuscaChekc & "NULL,0,0,0,0)"
            
            
            'au="inser.... + VALUE ()
            C2 = "insert into `slifac` (`codtipom`,`numfactu`,`fecfactu`,`codtipoa`,`numalbar`,`numlinea`,`codalmac`,`codartic`,`nomartic`,`ampliaci`,`cantidad`,`numbultos`,`precioar`,`dtoline1`,`dtoline2`,`importel`,`origpre`,`precioiv`,`preciomp`,`preciost`,`preciouc`,`codproveX`)"
            C2 = C2 & " VALUES " & BuscaChekc
            conn.Execute C2
            
            Espera 0.3
            BuscaChekc = ""
    End If
    Set vFactu = Nothing



    Exit Function
eTratarRecargoFinanciero:
    MuestraError Err.Number, Err.Description
    TratarRecargoFinanciero = False
    Set Raux = Nothing
    BuscaChekc = ""
End Function



'**********************************************************************************
'Campos ALZIRA MOIXENT

Private Sub MultiInsercionCampos()
Dim i As Integer
Dim C As String
Dim VariedadPartida As String

        'Quito el indicador # de multi campo
        If InStr(1, CadenaDesdeOtroForm, 1) > 0 Then CadenaDesdeOtroForm = Mid(CadenaDesdeOtroForm, 2)



'        BuscaChekc = BuscaChekc & "(select codcampo from slifaccampos where numalbar=" & Data1.Recordset!NumAlbar
'        BuscaChekc = BuscaChekc & " AND "


        BuscaChekc = ObtenerWhereCP(False) & " AND 1"
        BuscaChekc = DevuelveDesdeBD(conAri, "max(numlinea)", "slifaccampos", BuscaChekc, "1", "N")
        NumRegElim = 0
        If BuscaChekc <> "" Then NumRegElim = Val(BuscaChekc)
        NumRegElim = NumRegElim + 1
        C = ""
        While CadenaDesdeOtroForm <> ""
            i = InStr(1, CadenaDesdeOtroForm, "·#")

            If i = 0 Then
                CadenaDesdeOtroForm = ""
            Else
                BuscaChekc = Mid(CadenaDesdeOtroForm, 1, i - 1)
                CadenaDesdeOtroForm = Mid(CadenaDesdeOtroForm, i + 2)
                VariedadPartida = "," & DBSet(RecuperaValor(BuscaChekc, 2), "T", "S") & "," & DBSet(RecuperaValor(BuscaChekc, 3), "T", "S")
                BuscaChekc = RecuperaValor(BuscaChekc, 1) 'cdocampo

                For i = 1 To Me.ListView1.ListItems.Count
                    'Si no lo ha insertado YA
                    If Val(Me.ListView1.ListItems(i).Text) = Val(BuscaChekc) Then
                        BuscaChekc = ""
                        Exit For
                    End If

                Next i

                If BuscaChekc <> "" Then

                        '  slifaccampos(codtipom,numfactu,fecfactu,codtipoa,numalbar,,numlinea,codcampo)
                        C = C & ", (" & DBSet(Data1.Recordset!codtipom, "T") & "," & Data1.Recordset!Numfactu
                        C = C & "," & DBSet(Data3.Recordset!FecFactu, "F") & "," & DBSet(Data3.Recordset!Codtipoa, "T")
                        C = C & "," & DBSet(Data3.Recordset!Numalbar, "N") & "," & NumRegElim & "," & BuscaChekc & "," & DBSet(Now, "FH")
                        C = C & VariedadPartida & ")" ' ",NULL,NULL" & ")"   ',nomvarie , nompartida
                        NumRegElim = NumRegElim + 1
                End If
            End If
        Wend
        If C <> "" Then
            C = Mid(C, 2) 'quito la primera coma
            '
            C = "INSERT INTO slifaccampos(codtipom,numfactu,fecfactu,codtipoa,numalbar,numlinea,codcampo,fechahora,nomvarie , nompartida) VALUES " & C
            If ejecutar(C, False) Then
                'Hay que refrescar y boton anyadir

            End If
        End If

        C = ""
        BuscaChekc = ""

        '
        
End Sub


Private Sub CargaDatosCampos()
Dim IT


    On Error GoTo eCargaDatosCampos

    'Para no meter MUCHOS ariagro.ss
    'Pongo @# y luego lo reemplazo por vparamaplic.Ariagro.
'    SQL = "select rcampos.codcampo, rpartida.nomparti, variedades.nomvarie"
'    SQL = SQL & " from (@#rcampos inner join @#rpartida on rcampos.codparti = rpartida.codparti)"
'    SQL = SQL & " inner join @#variedades on rcampos.codvarie = variedades.codvarie"
'    'where socio
'    SQL = Replace(SQL, "@#", vParamAplic.Ariagro & ".")
'
    
    
    
    BuscaChekc = "select rcampos.codcampo, rpartida.nomparti, variedades.nomvarie,rcampos.codclien,rsocios.codsocio,rsocios.nomsocio,rcampos.codsitua"
    BuscaChekc = BuscaChekc & " from ((@#rcampos inner join @#rpartida on rcampos.codparti = rpartida.codparti)"
    BuscaChekc = BuscaChekc & " inner join @#variedades on rcampos.codvarie = variedades.codvarie)"
    BuscaChekc = BuscaChekc & " inner join @#rsocios on rsocios.codsocio=rcampos.codsocio"
    
    BuscaChekc = Replace(BuscaChekc, "@#", vParamAplic.Ariagro & ".")
    
    BuscaChekc = BuscaChekc & " WHERE codcampo IN "
    BuscaChekc = BuscaChekc & "(select codcampo from slifaccampos  "
    BuscaChekc = BuscaChekc & ObtenerWhereCP(True)
    BuscaChekc = BuscaChekc & ")"
    ListView1.ListItems.Clear
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open BuscaChekc, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    While Not miRsAux.EOF
        Set IT = ListView1.ListItems.Add()
        IT.Text = Format(miRsAux!codCampo, "000000")
        IT.SubItems(1) = DBLet(miRsAux!nomparti, "T")
        IT.SubItems(2) = DBLet(miRsAux!nomvarie, "T")
        IT.SubItems(3) = Format(DBLet(miRsAux!codsocio, "N"), "00000")
        IT.SubItems(4) = DBLet(miRsAux!nomsocio, "T")
        IT.Tag = miRsAux!codCampo
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    BuscaChekc = ""
  
    Exit Sub
    
eCargaDatosCampos:
    MuestraError Err.Number, "Cargando datos ariagro", Err.Description
  
End Sub




Private Sub ImprimirFraTelefonia()
Dim Aux As String
Dim Cadpa As String
Dim nPar As Byte
    If Modo <> 2 Then Exit Sub
    If Me.Data1.Recordset Is Nothing Then Exit Sub
    If Me.Data1.Recordset.EOF Then Exit Sub

   
    NumRegElim = 63
    'Aux = DevuelveDesdeBDNew(conAri, "scryst", "documrpt", "codcryst", CStr(NumRegElim), "N")
    If Not PonerParamRPT2(CByte(NumRegElim), Cadpa, nPar, Aux, False, "", False) Then Exit Sub
    
    If vParamAplic.NumeroInstalacion = vbTaxco Then
        'EN TAXCO añadimos la marca de agua
        Cadpa = Cadpa & "pDuplicado=1|"
        nPar = nPar + 1
    End If
    
            
            
   
    With frmImprimir
        .NombreRPT = Aux
        
            Aux = DevuelveDesdeBDNew(conAri, "stipom", "letraser", "codtipom", Data1.Recordset!codtipom, "T")  'LEtra de serie
            Aux = "{tel_cab_factura.Serie} ='" & Aux & "' and " & _
                                            "{tel_cab_factura.Ano} =" & Year(Data1.Recordset!FecFactu) & " and {tel_cab_factura.NumFact} ="
            Aux = Aux & Data1.Recordset!Numfactu
            
            
            'SEPTIEMBRE 2013
            'Tel_cab_factura y scafac1 estan enlazdas
            
                'Cod Tipo Movimiento
                pPdfRpt = ""
                Aux = "{" & NombreTabla & ".codtipom}='" & Text1(1).Text & "'"
                If Not AnyadirAFormula(pPdfRpt, Aux) Then Exit Sub
        
                'Nº Factura
                Aux = "{" & NombreTabla & ".numfactu}=" & Val(Text1(0).Text)
                If Not AnyadirAFormula(pPdfRpt, Aux) Then Exit Sub
        
                'Fecha Factura
                Aux = "{" & NombreTabla & ".fecfactu}= Date(" & Year(Text1(2).Text) & "," & Month(Text1(2).Text) & "," & Day(Text1(2).Text) & ")"
                If Not AnyadirAFormula(pPdfRpt, Aux) Then Exit Sub

    

        .FormulaSeleccion = pPdfRpt
        .OtrosParametros = Cadpa
        .NumeroParametros = nPar
        .Titulo = "Factura telefonía"
        .SoloImprimir = False
        .EnvioEMail = False
        .Opcion = 2000 '2000 generico
       
        .ConSubInforme = True
        .Show vbModal
    End With
    

End Sub


Private Sub CargaDatosTelefonia()
Dim Cad As String
Dim IT As ListItem

    Me.ListView2.ListItems.Clear
    Me.ListView3.ListItems.Clear
    
    If LetrasFraTelefonia = "" Then
        'Voy a cargar las letras de talefonia
        LetrasFraTelefonia = DevuelveDesdeBD(conAri, "letraser", "stipom", "codtipom", "FAT", "T")
        Cad = DevuelveDesdeBD(conAri, "letraser", "stipom", "codtipom", "FAI", "T")
        LetrasFraTelefonia = LetrasFraTelefonia & "|" & Cad & "|"
    End If
    If Me.Data1.Recordset!codtipom = "FAV" Or Me.Data1.Recordset!codtipom = "FMO" Then
        Cad = ""  'Es las mas normal
    Else
        If Me.Data1.Recordset!codtipom = "FAT" Then
            Cad = RecuperaValor(LetrasFraTelefonia, 1) & "|" & Year(Data1.Recordset!FecFactu) & "|" & Data1.Recordset!Numfactu & "|"
        Else
            If Data1.Recordset!codtipom = "FAI" Then
                'Puede ser, o no, un telefonia
                
                Cad = RecuperaValor(LetrasFraTelefonia, 2) & "|" & Year(Data1.Recordset!FecFactu) & "|" & Data3.Recordset!Numfactu & "|"   'NUMALBAR
            Else
                Cad = ""
            End If
        End If
    End If
    If Cad = "" Then Exit Sub
    
    
    CargaLwTelefonia ListView2, RecuperaValor(Cad, 1), RecuperaValor(Cad, 2), RecuperaValor(Cad, 3), FormatoImporte, True
    If Me.ListView2.ListItems.Count > 0 Then
        Cad = "SELECT Fichero, Numero_de_telefono, Descripcion_tipo_de_llamada, Tipo_destino, Numero_llamado, "
        Cad = Cad & " Fecha, Hora_inicio, Cantidad_medida_originada, Importe, Unidad_de_medida"
        Cad = Cad & " FROM   Telefono.detalle_de_llamadas  where Fichero='" & Text3(16).Text
        Cad = Cad & "' and Numero_de_telefono='" & Text1(7).Text & "'"
        Cad = Cad & " ORDER BY detalle_de_llamadas.Fichero, detalle_de_llamadas.Numero_de_telefono, detalle_de_llamadas.Fecha, detalle_de_llamadas.Hora_inicio"
        Set miRsAux = New ADODB.Recordset
        miRsAux.Open Cad, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        While Not miRsAux.EOF
            
            Cad = Trim(DBLet(miRsAux!Descripcion_tipo_de_llamada, "T"))
            If Cad <> "" Then
                Set IT = Me.ListView3.ListItems.Add()
                IT.Text = Cad
                IT.SubItems(1) = Trim(DBLet(miRsAux!Numero_llamado, "T"))
                IT.SubItems(2) = Mid(miRsAux!Fecha, 3, 2) & "/" & Mid(miRsAux!Fecha, 1, 2)
                IT.SubItems(3) = DBLet(miRsAux!Importe, "N")
            End If
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        Set miRsAux = Nothing
    End If
    
End Sub



Private Sub EliminarCambiarFechaFactura()
Dim miSQL As String
Dim cControlFra As CControlFacturaContab
Dim LEtra As String
Dim CambiarFecha As Boolean
Dim B1 As Boolean

Dim PuedeModificarCobrosContabilidad As Boolean


    'Primera comprobacion. La factura es "actual"
    'Es decir del periodo de IVA, veremos que si no es la ultima
    PuedeModificarCobrosContabilidad = True
    LEtra = ObtenerLetraSerie(Data1.Recordset!codtipom)
    If LEtra = "" Then Exit Sub
    
        Set cControlFra = New CControlFacturaContab
        miSQL = ""
        
        'Con estos dos NO dejo pasar
        BuscaChekc = cControlFra.FechaCorrectaContabilizazion(ConnConta, Text1(2))
        If BuscaChekc <> "" Then miSQL = miSQL & "- " & BuscaChekc & vbCrLf
        BuscaChekc = cControlFra.FechaCorrectaIVA(ConnConta, Text1(2))
        If BuscaChekc <> "" Then miSQL = miSQL & "- " & BuscaChekc & vbCrLf
            
            
        If DBLet(Data1.Recordset!intconta, "N") = 1 Then miSQL = "- Factura contabilizada." & vbCrLf & miSQL
        
            
        If miSQL <> "" Then
            B1 = True 'mostrar msg
            
            If vParamAplic.PuedeModificarAriconta Then
                If CDate(Text1(2).Text) < vEmpresa.FechaIni Then
                    B1 = True 'Fecha anterior a fecha ejercicio. NO se toca
                Else
                    B1 = False
                End If
            End If
            
            If B1 Then
            
            
                MsgBox miSQL, vbExclamation
                Set cControlFra = Nothing
                Exit Sub
            
              
                
            End If
        End If
        
        PuedeModificarCobrosContabilidad = True
        
        
        If cControlFra.FechaMenorUltimaFacturaCliente(ConnConta, Text1(2), LEtra) Then
            If BuscaChekc <> "" Then miSQL = miSQL & "- Hay facturas contabilizada con fechas posterior" & vbCrLf
        End If
        Set cControlFra = Nothing
        
        
        Dim C2 As String
        
        If vParamAplic.ContabilidadNueva Then
            C2 = " numserie='" & LEtra & "' AND numfactu=" & Data1.Recordset.Fields!Numfactu
            C2 = C2 & " AND fecfactu='" & Format(Data1.Recordset.Fields!FecFactu, FormatoFecha) & "' AND 1"
            C2 = DevuelveDesdeBD(conConta, "sum(impvenci + coalesce(gastos,0) - coalesce(impcobro,0))", "cobros", C2, "1")
            If C2 <> "" Then
                If CCur(C2) = 0 Then miSQL = miSQL & "- Cobrada totalmente" & vbCrLf: PuedeModificarCobrosContabilidad = False
                If CCur(C2) <> Data1.Recordset!TotalFac Then miSQL = miSQL & "- Cobrada parcialmente: " & C2 & " // " & DBLet(Data1.Recordset!TotalFac, "T") & vbCrLf: PuedeModificarCobrosContabilidad = False
            End If
        Else
            'SQL = SQL & " scobro WHERE numserie='" & LEtra & "' AND codfaccl=" & Data1.Recordset.Fields!Numfactu
            'SQL = SQL & " AND fecfaccl='" & Format(Data1.Recordset.Fields!FecFactu, FormatoFecha) & "'"
        End If
        
        
        
        
   
        
        
        If miSQL <> "" Then
            miSQL = miSQL & vbCrLf & "¿Seguro que desea continuar el proceso?"
            
            If MsgBox(miSQL, vbExclamation + vbYesNo) <> vbYes Then Exit Sub
        
        End If
    
        
    CadenaDesdeOtroForm = "0"
    If CStr(Data1.Recordset!codtipom) = "FAT" Then CadenaDesdeOtroForm = "1"
    CadenaDesdeOtroForm = Text1(2).Text & "|" & LEtra & "|" & CadenaDesdeOtroForm & "|"
    frmVarios.Opcion = 13
    frmVarios.Show vbModal
    
    
    If CadenaDesdeOtroForm <> "" Then
        Screen.MousePointer = vbHourglass
        
        CambiarFecha = Not CadenaDesdeOtroForm = "OK"
        
        conn.BeginTrans
        ConnConta.BeginTrans
        If HacerAccionesModFechaElimFra(CambiarFecha, PuedeModificarCobrosContabilidad) Then
            conn.CommitTrans
            ConnConta.CommitTrans
            Espera 0.25
            
            If CambiarFecha Then Text1(2).Text = CadenaDesdeOtroForm
            CadenaDesdeOtroForm = "scafac.codtipom= '" & Text1(1).Text & "' and scafac.numfactu= " & Val(Text1(0).Text) & " and scafac.fecfactu='" & Format(Text1(2).Text, FormatoFecha) & "' "
        
            
            
            Set LOG = New cLOG
            If CambiarFecha Then
                LEtra = "Nueva fecha: " & Text1(2).Text & vbCrLf & CadenaDesdeOtroForm
                LOG.Insertar 25, vUsu, LEtra
            Else
                LEtra = "Eliminar reestb factura:" & CadenaDesdeOtroForm
                LOG.Insertar 26, vUsu, LEtra
            End If
            
            Set LOG = Nothing
            
            If CambiarFecha Then
                CadenaConsulta = "select scafac.* from " & NombreTabla & " INNER JOIN scafac1 ON scafac.codtipom=scafac1.codtipom AND scafac.numfactu=scafac1.numfactu AND scafac.fecfactu=scafac1.fecfactu "
                CadenaConsulta = CadenaConsulta & " WHERE " & CadenaDesdeOtroForm & " GROUP BY scafac.codtipom,scafac.numfactu,scafac.fecfactu " & Ordenacion
                PonerCadenaBusqueda
                
                'Si ha cambiado fecha, vemos de calcular de nuevo los vencimientos
                RecalculaSvenciDespuesMofificarFecha
                
                
                
                
                
            Else
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
            
            
            
            
            
        Else
            conn.RollbackTrans
            ConnConta.RollbackTrans
        End If
        Screen.MousePointer = vbDefault
    End If
End Sub



Private Function HacerAccionesModFechaElimFra(CambiarFecha As Boolean, PuedeModificarCobrosEnAriconta As Boolean) As Boolean
Dim Sql As String
Dim C2 As String
Dim c3 As String
Dim LEtra As String
Dim RA As ADODB.Recordset
Dim TienAsiente As Boolean



    On Error GoTo eHacerAccionesModFechaElimFra

    HacerAccionesModFechaElimFra = False
    
    BuscaChekc = ObtenerWhereCP(True)
    If BuscaChekc = "" Then Exit Function
    
        
    conn.Execute "SET FOREIGN_KEY_CHECKS=0"
    ConnConta.Execute "SET FOREIGN_KEY_CHECKS=0"
    

    If CambiarFecha Then BuscaChekc = " set fecfactu=" & DBSet(CadenaDesdeOtroForm, "F") & " " & BuscaChekc
    
        
    If CambiarFecha Then
        conn.Execute "UPDATE slifac " & BuscaChekc
        
        If InstalacionEsEulerTaxco Then
            conn.Execute "UPDATE slifac_eu " & BuscaChekc
            conn.Execute "UPDATE slifac_eu2 " & BuscaChekc
        End If
        
        
        'Campos
        conn.Execute "UPDATE slifaccampos " & BuscaChekc
    
        'Lineas de cabeceras de albaranes de la factura
        conn.Execute "UPDATE scafac1 " & BuscaChekc
            
        'Lineas de cabeceras de albaranes de la factura
        conn.Execute "UPDATE scafacportes " & BuscaChekc
            
            
        'Eliminar los vencimientos
        conn.Execute "UPDATE svenci " & BuscaChekc
        
        'Cabecera de facturas (scafac)
        conn.Execute "UPDATE " & NombreTabla & BuscaChekc
        
        
        If vParamAplic.PuedeModificarAriconta Then
            Set RA = New ADODB.Recordset
            
            If Val(Data1.Recordset!intconta) = 1 Then
            
                LEtra = DevuelveDesdeBD(conAri, "letraser", "stipom", "codtipom", Text1(1).Text, "T")
                If LEtra = "" Then Err.Raise 513, , "Error obteniendo letra contabilidad"
                C2 = "numserie= " & DBSet(LEtra, "T") & " AND numfactu= " & Val(Text1(0).Text) & " AND fecfactu='" & Format(Text1(2).Text, FormatoFecha) & "' "
            
                RA.Open "Select numasien,fechaent,numdiari,anofactu,numserie,numfactu FROM factcli WHERE " & C2, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
                If RA.EOF Then
                    c3 = "No existe factura " & LEtra & "  " & Text1(0).Text & " de " & Text1(2).Text & " en contabilidad"
                    Err.Raise 513, , c3
                Else
                    
                    If DBLet(RA!numasien, "N") = 0 Then
                        MsgBox "Factura traspasada", vbInformation
                        TienAsiente = False
                    Else
                        c3 = ""
                        If IsNull(RA!FechaEnt) Then c3 = "Error fecha asiento nula "
                        If IsNull(RA!numdiari) Then c3 = "Error numero diario nulo "
                        If c3 <> "" Then Err.Raise 513, , c3
                        
                        c3 = "set fechaent = " & DBSet(CadenaDesdeOtroForm, "F") & " WHERE numasien = " & RA!numasien & " AND numdiari =" & RA!numdiari & " AND fechaent = " & DBSet(RA!FechaEnt, "F")
                        ConnConta.Execute "UPDATE hlinapu " & c3
                        ConnConta.Execute "UPDATE hcabapu " & c3
                        TienAsiente = True
                        
                        
                    End If
                    Sql = ""
                    If TienAsiente Then Sql = "  fechaent =" & DBSet(CadenaDesdeOtroForm, "F") & ", "
                    c3 = " fecfactu = " & DBSet(CadenaDesdeOtroForm, "F") & " WHERE anofactu = " & RA!anofactu & " AND numserie =" & DBSet(RA!numSerie, "T") & " AND numfactu = " & RA!Numfactu
            
                    ConnConta.Execute "UPDATE factcli SET " & Sql & c3
                    ConnConta.Execute "UPDATE factcli_lineas  SET " & c3
                    ConnConta.Execute "UPDATE factcli_totales SET " & c3
                    
                    
                    
                    
                End If 'ra.eof
                RA.Close
                    
                    
            End If  'If Val(Data1.Recordset!intconta) = 1
                    
           '// hasta aqui         que puede modificar en ariconta.factura

       End If
    
    
    Else
        '#cabecera de albaran
        Sql = "INSERT INTO scaalb(codtipom,numalbar,fechaalb,factursn,codclien,nomclien,domclien,codpobla,pobclien,proclien,"
        Sql = Sql & "nifclien,telclien,coddirec,nomdirec,referenc,facturkm,cantidkm,codtraba,codtrab1,codtrab2,"
        Sql = Sql & "codagent,codforpa,codenvio,dtoppago,dtognral,tipofact,observa01,observa02,observa03,observa04,observa05,"
        Sql = Sql & "numofert,fecofert,numpedcl,fecpedcl,fecentre,sementre,codtipmf,numfactu,fecfactu,esticket,numtermi,numventa,aportacion,pesoalba,portes,"
        Sql = Sql & "fecenvio,docarchiv,tipliquid,actuacion,tipoimp,origdat,coddiren,tipAlbaran,albImpreso,codzonas,observacrm"
        Sql = Sql & ", ManipuladorNumCarnet , ManipuladorFecCaducidad , ManipuladorNombre,TipoCarnet"
        Sql = Sql & ", PideCliente,numbultos,fechaAux,puntos"
        'sql = sql & ", codinter,codnatura,notasportes "
        Sql = Sql & ")   SELECT codtipoa,numalbar,fechaalb,1 factursn, codclien,nomclien,domclien,codpobla,pobclien,proclien,"
        Sql = Sql & "nifclien,telclien,coddirec,nomdirec,referenc,"
        Sql = Sql & "0 facturakm ,0 cuantoskm,codtraba,codtrab1,codtrab2,"
        Sql = Sql & "codagent,codforpa,codenvio,dtoppago,dtognral,0 tipofac, observa1,observa2,observa3,observa4,observa5,"
        Sql = Sql & "numofert,fecofert,numpedcl,fecpedcl,fecpedcl,sementre,"
        Sql = Sql & "NULL codtipmf, NULL numfactu,NULL fecfactu ,0 esticket, numtermi,numventa,aportacion,pesoalba,portes,"
        Sql = Sql & "fecenvio,docarchiv,NULL tipliquid,actuacion,tipoimp,origdat,"
        Sql = Sql & "coddiren,tipAlbaran,0 albImpreso,1 codzona,NULL observacrm "
        Sql = Sql & ", ManipuladorNumCarnet , ManipuladorFecCaducidad , ManipuladorNombre,TipoCarnet"
        Sql = Sql & ", PideCliente,numbultos,fechaAux,puntos"
        'sql = sql & ", codinter,codnatura,notasportes "
        
        Sql = Sql & " FROM scafac,scafac1 Where scafac.NumFactu = scafac1.NumFactu And scafac.codtipom = scafac1.codtipom"
        ' SQL = " codtipom= '" & Text1(1).Text & "' and numfactu= " & Val(Text1(0).Text) & " and fecfactu='" & Format(Text1(2).Text, FormatoFecha) & "' "
        Sql = Sql & " AND scafac.fecfactu=scafac1.fecfactu AND scafac.numfactu =" & Val(Text1(0).Text)
        Sql = Sql & " AND scafac.fecfactu=" & DBSet(Text1(2).Text, "F") & " AND scafac.codtipom =" & DBSet(Text1(1).Text, "T")
        conn.Execute Sql
        
        '#Lineas albaran
        Sql = "INSERT INTO slialb (codtipom ,numalbar,numlinea,codalmac,codartic,nomartic,ampliaci,cantidad,numbultos,precioar,dtoline1,dtoline2,importel,origpre,"
        Sql = Sql & "codproveX,numlote,codccost,codtipor,codcapit,precoste,codtraba,pvpInferior,comisionagente,idL,ordenlin)"
        Sql = Sql & " SELECT scafac1.codtipoa,scafac1.numalbar,numlinea,codalmac,codartic,nomartic,ampliaci,cantidad,slifac.numbultos,precioar,dtoline1,dtoline2,importel,origpre,"
        Sql = Sql & "codproveX , numLote, CodCCost, codtipor, codcapit, precoste, slifac.CodTraba,slifac.pvpInferior,slifac.comisionagente,slifac.idL,slifac.ordenlin "
        Sql = Sql & "FROM scafac,scafac1,slifac WHERE scafac.codtipom=scafac1.codtipom AND scafac.numfactu=scafac1.numfactu"
        Sql = Sql & " AND scafac.fecfactu=scafac1.fecfactu AND"
        Sql = Sql & " scafac.codtipom = slifac.codtipom And scafac.NumFactu = slifac.NumFactu"
        Sql = Sql & " AND scafac.fecfactu=slifac.fecfactu AND"
        Sql = Sql & " scafac1.codtipoa = slifac.codtipoa And scafac1.NumAlbar = slifac.NumAlbar"
        Sql = Sql & " AND scafac.numfactu =" & Val(Text1(0).Text)
        Sql = Sql & " AND scafac.fecfactu=" & DBSet(Text1(2).Text, "F") & " AND scafac.codtipom =" & DBSet(Text1(1).Text, "T")
        conn.Execute Sql
    
        'Los costes
        If InstalacionEsEulerTaxco Then
            Sql = "INSERT INTO slialb_eu(codtipom,numalbar,numlinea,fechamov,codalmac,codartic,nomartic,cantidad,precioar)"
            Sql = Sql & " select scafac1.codtipoa,scafac1.numalbar,numlinea,fechamov,slifac_eu.codalmac,slifac_eu.codartic,slifac_eu.nomartic,slifac_eu.cantidad,precioar"
            Sql = Sql & " FROM scafac,scafac1,slifac_eu WHERE scafac.codtipom=scafac1.codtipom AND scafac.numfactu=scafac1.numfactu"
            Sql = Sql & " AND scafac.fecfactu=scafac1.fecfactu AND"
            Sql = Sql & " scafac.codtipom = slifac_eu.codtipom And scafac.NumFactu = slifac_eu.NumFactu"
            Sql = Sql & " AND scafac.fecfactu=slifac_eu.fecfactu AND"
            Sql = Sql & " scafac1.codtipoa = slifac_eu.codtipoa And scafac1.NumAlbar = slifac_eu.NumAlbar"
            Sql = Sql & " AND scafac.numfactu =" & Val(Text1(0).Text)
            Sql = Sql & " AND scafac.fecfactu=" & DBSet(Text1(2).Text, "F") & " AND scafac.codtipom =" & DBSet(Text1(1).Text, "T")
            Sql = Sql & " AND tipo=1" 'los otros los genera al pasar alba->fra
            conn.Execute Sql
            
            
            Sql = "INSERT INTO slialb_eu2 (codtipom,numalbar,numlinea,articulo,descrarticulo,cantidad,precioar,dtoline1,importel)"
            Sql = Sql & " select scafac1.codtipoa,scafac1.numalbar,numlinea,articulo,descrarticulo,cantidad,precioar,dtoline1,importel"
            Sql = Sql & " FROM scafac,scafac1,slifac_eu2 WHERE scafac.codtipom=scafac1.codtipom AND scafac.numfactu=scafac1.numfactu"
            Sql = Sql & " AND scafac.fecfactu=scafac1.fecfactu AND"
            Sql = Sql & " scafac.codtipom = slifac_eu2.codtipom And scafac.NumFactu = slifac_eu2.NumFactu"
            Sql = Sql & " AND scafac.fecfactu=slifac_eu2.fecfactu AND"
            Sql = Sql & " scafac1.codtipoa = slifac_eu2.codtipoa And scafac1.NumAlbar = slifac_eu2.NumAlbar"
            Sql = Sql & " AND scafac.numfactu =" & Val(Text1(0).Text)
            Sql = Sql & " AND scafac.fecfactu=" & DBSet(Text1(2).Text, "F") & " AND scafac.codtipom =" & DBSet(Text1(1).Text, "T")
            conn.Execute Sql
            
            
        End If
    
        If vParamAplic.CartaPortes Then
            Sql = "INSERT INTO scaalb_portes(codtipom,numalbar,matricula,descr)"
    
            Sql = " SELECT codtipoa,numalbar,matricula,descr"
            
            Sql = Sql & " FROM scafac,scafacportes Where scafac.NumFactu = scafacportes.NumFactu And scafac.codtipom = scafacportes.codtipom"
            Sql = Sql & " AND scafac.fecfactu=scafacportes.fecfactu AND scafac.numfactu =" & Val(Text1(0).Text)
            Sql = Sql & " AND scafac.fecfactu=" & DBSet(Text1(2).Text, "F") & " AND scafac.codtipom =" & DBSet(Text1(1).Text, "T")
            conn.Execute Sql
            
                
            conn.Execute "Delete from scafacportes " & BuscaChekc
        End If
        
        
        LEtra = DevuelveDesdeBD(conAri, "letraser", "stipom", "codtipom", Text1(1).Text, "T")
        Sql = "DELETE FROM "
        If vParamAplic.ContabilidadNueva Then
            Sql = Sql & " cobros WHERE numserie='" & LEtra & "' AND numfactu=" & Data1.Recordset.Fields!Numfactu
            Sql = Sql & " AND fecfactu='" & Format(Data1.Recordset.Fields!FecFactu, FormatoFecha) & "'"
        Else
            Sql = Sql & " scobro WHERE numserie='" & LEtra & "' AND codfaccl=" & Data1.Recordset.Fields!Numfactu
            Sql = Sql & " AND fecfaccl='" & Format(Data1.Recordset.Fields!FecFactu, FormatoFecha) & "'"
        End If
        ConnConta.Execute Sql
        
        
        'Si es proyecto, abro el proyecto
        If Text1(1).Text = "FPY" Then
            'Por si acaso, la fecha fue modificada o algo
            Sql = "UPDATE sproyecto SET numfactu=NULL, fecfinal=NULL WHERE year(fecfinal)=" & Year(Data1.Recordset.Fields!FecFactu) & " AND numfactu=" & Data1.Recordset.Fields!Numfactu
            conn.Execute Sql
        End If
        
        'La borramos
        conn.Execute "Delete from slifac " & BuscaChekc
        If InstalacionEsEulerTaxco Then
            conn.Execute "Delete from slifac_eu " & BuscaChekc
            conn.Execute "Delete from slifac_eu2 " & BuscaChekc
        End If
        
                
        
        
        'Campos
        conn.Execute "Delete from slifaccampos " & BuscaChekc
    
        'Lineas de cabeceras de albaranes de la factura
        conn.Execute "Delete from scafac1 " & BuscaChekc
            
        'Eliminar los vencimientos
        conn.Execute "Delete from svenci " & BuscaChekc
        
        'Cabecera de facturas (scafac)
        conn.Execute "Delete from " & NombreTabla & BuscaChekc
            
            
            
            
        If vParamAplic.PuedeModificarAriconta Then
            Set RA = New ADODB.Recordset
            
            If Val(Data1.Recordset!intconta) = 1 Then
            
                LEtra = DevuelveDesdeBD(conAri, "letraser", "stipom", "codtipom", Text1(1).Text, "T")
                If LEtra = "" Then Err.Raise 513, , "Error obteniendo letra contabilidad"
                C2 = "numserie= " & DBSet(LEtra, "T") & " AND numfactu= " & Val(Text1(0).Text) & " AND fecfactu='" & Format(Text1(2).Text, FormatoFecha) & "' "
            
                RA.Open "Select numasien,fechaent,numdiari,anofactu,numserie,numfactu FROM factcli WHERE " & C2, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
                If RA.EOF Then
                    c3 = "No existe factura " & LEtra & "  " & Text1(0).Text & " de " & Text1(2).Text & " en contabilidad"
                    MsgBox c3, vbInformation
                   ' Err.Raise 513, , C3
                Else
                    
                    If DBLet(RA!numasien, "N") = 0 Then
                       ' MsgBox "Factura traspasada", vbInformation
                        TienAsiente = False
                    Else
                        c3 = ""
                        If IsNull(RA!FechaEnt) Then c3 = "Error fecha asiento nula "
                        If IsNull(RA!numdiari) Then c3 = "Error numero diario nulo "
                        If c3 <> "" Then Err.Raise 513, , c3
                        
                        c3 = " WHERE numasien = " & RA!numasien & " AND numdiari =" & RA!numdiari & " AND fechaent = " & DBSet(RA!FechaEnt, "F")
                        ConnConta.Execute "DELETE FROM hlinapu " & c3
                        ConnConta.Execute "DELETE FROM hcabapu " & c3
                        TienAsiente = True
                        
                        
                    End If
                    c3 = " WHERE anofactu = " & RA!anofactu & " AND numserie =" & DBSet(RA!numSerie, "T") & " AND numfactu = " & RA!Numfactu
                               
                    ConnConta.Execute "DELETE FROM factcli  " & c3
                    ConnConta.Execute "DELETE FROM factcli_lineas   " & c3
                    ConnConta.Execute "DELETE FROM factcli_totales  " & c3
                    
                    
                    
                    MsgBox "Revisar cobros", vbExclamation
                    'cOBROS
                End If 'ra.eof
                RA.Close
                    
                    
            End If  'If Val(Data1.Recordset!intconta) = 1
                    
                    
         End If
                    
            
            
            
            
            
            
    End If
    
    

    conn.Execute "SET FOREIGN_KEY_CHECKS=1"
    ConnConta.Execute "SET FOREIGN_KEY_CHECKS=1"
    
    
    
    If PuedeModificarCobrosEnAriconta Then
        Screen.MousePointer = vbHourglass
        Espera 1
        
        
        If vParamAplic.ContabilidadNueva Then
            LEtra = DevuelveDesdeBD(conAri, "letraser", "stipom", "codtipom", Text1(1).Text, "T")
            Sql = " WHERE numserie='" & LEtra & "' AND numfactu=" & Data1.Recordset.Fields!Numfactu
            Sql = Sql & " AND fecfactu='" & Format(Data1.Recordset.Fields!FecFactu, FormatoFecha) & "'"
            
            If LEtra <> "" Then
    
                If CambiarFecha Then
                    Sql = "UPDATE cobros set fecfactu= " & DBSet(CadenaDesdeOtroForm, "F") & " " & Sql
                Else
                    Sql = "DELETE FROM cobros " & Sql
                End If

                EjecutaEnConConta (Sql)
            End If
        
        End If
    
    End If
    
    
    
    
    
    
    
    
    HacerAccionesModFechaElimFra = True
    Set RA = Nothing
    Exit Function
    
eHacerAccionesModFechaElimFra:
    MuestraError Err.Number, Err.Description
    conn.Execute "SET FOREIGN_KEY_CHECKS=1"
    ConnConta.Execute "SET FOREIGN_KEY_CHECKS=1"
    
    Set RA = Nothing
End Function

Private Sub RecalculaSvenciDespuesMofificarFecha()
Dim vFactura As CFactura


    Set vFactura = New CFactura
    If Text1(1).Text <> "FAZ" Then
        Espera 0.25
        If vFactura.LeerDatos(Text1(1).Text, Text1(0).Text, Text1(2).Text) Then    'CadenaDesdeOtroForm: lleva la nueva fecha
    
                    
            vFactura.CuentaPrev = "" 'SI HACEMOS ESTO, NO GENERA EN tesoreria
        
            vFactura.Agente = Text1(14).Text
            ejecutar "DELETE FROM svenci " & ObtenerWhereCP(True), False
            If vFactura.InsertarEnTesoreria("", "", False) Then   'false: no inserte an ariconta
                MsgBox "Vencimientos ARIGES modificados.  Revise tesoreria.", vbInformation
            End If
        End If
    End If
    
    Set vFactura = Nothing
     
End Sub

Private Function EjecutaEnConConta(Sql) As Boolean
    On Error Resume Next
    ConnConta.Execute Sql
    If Err.Number <> 0 Then
        MuestraError Err.Number, , Err.Description
        EjecutaEnConConta = False
    Else
        EjecutaEnConConta = True
    End If
    
End Function



'******************************************************************************************************
'******************************************************************************************************
'******************************************************************************************************
'EULER

Private Sub PonerCamposFicha(Todo As Boolean)   'Todo=False   Solo lineas facturas euler
Dim N As Byte
Dim Sql As String
Dim Cad2 As String
Dim N2 As Integer
Dim ImpMano As Currency
Dim impVeh As Currency
Dim total As Currency
Dim Impo As Currency

    Set miRsAux = New ADODB.Recordset

    If Todo Then
    
    
        If vParamAplic.NumeroInstalacion = vbEuler Then
            Me.FrameALE.visible = Data3.Recordset!Codtipoa = "ALO"     'Text1(1).Text = "FAE"
        Else
            FrameALE.visible = False
            Me.FrameTAXCO.visible = Data3.Recordset!Codtipoa = "ALO"     'Text1(1).Text = "FAE"
        End If
        Me.FrameReparEuler.visible = Data3.Recordset!Codtipoa = "ALR"      'Text1(1).Text = "FAE"
        
        Sql = "ReferPedido,FechaPed,bombamarca,bombaModelo,motormarca,motorModelo"
        Sql = Sql & ",TrabajoExterior,observaciones,TipoPortes"
        
        Sql = Sql & ",NumParteTrabajo,NumTrabajExterno,RecepAgenClien,RecepPortes,RecepAgenCliMat,RecpNumExp,FechaAlb,TipoBombResSuperHor"
        Sql = Sql & ",TipoBombResSuperVer,TipoBombLimSuperHor,TipoBombLimSuperVer,TipoBombResSumPoz,TipoBombLimSumPoz,TipoBombResSumVer"
        Sql = Sql & ",TipoBombLimSumVer,TipoBomAgitadorRes,TipoBomAgitadorLim,TipoBomResOtrosEqu,TipoBomLimOtrosEqu,DatosBommarca,DatosBomNumCurva"
        Sql = Sql & ",DatosBomModelo,DatosBomNumSerie,DatosBomAno,DatosBomH,DatosBomTipoRodete,DatosBomCaudal,DatosBomUdCaudal,DatosMotorMarca"
        Sql = Sql & ",DatosMotorModelo , DatosMotorNumSerie, DatosMotorV, DatosMotorI, DatosMotorCV, DatosMotorKw, DatosMotorrpm, NumParteTrabajo, NumTrabajExterno"
        Sql = Sql & ",numrepar"
        lwCostes.ListItems.Add , , "Leyendo"
        
        
        Sql = "Select " & Sql & " FROM scafac_eu "
        Sql = Sql & ObtenerWhereCP(True)
            
        
        miRsAux.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If miRsAux.EOF Then
            LimpiarFichaTecnica False
            
        Else
            
            
            'cboEulerT.ListIndex = DBLet(miRsAux!partetrabajo)  '0 1
            
            'EL SQL estara montaddo para que coincida el orden del columna con el index
            
            
            For N = 0 To txtEuler.Count - 1
                txtEuler(N).Text = DBLet(miRsAux.Fields(CInt(N)), "T")
            Next
        
            'Relamente , lo de arriba no haria falta en TAXCO
            If vParamAplic.NumeroInstalacion = vbTaxco Then
            
                N = 0
                If Data1.Recordset!codtipom = "FAO" And FrameTAXCO.visible Then N = 1
                If Data1.Recordset!codtipom = "FA5" And FrameTAXCO.visible Then N = 1
                
                If N = 1 Then
                    'bombamarca -Matricula
                    'bombaModelo -bastidor
                    'motorModelo -marca / modelo
                    'motormarca -motor
                    'ReferPedido -neumaticos
                    'RecepAgenCliMat -licencia
                    'RecpNumExp -taximetro
                    'numrepar -kms
                    
                    Me.txtTaxco(0).Text = DBLet(miRsAux!bombamarca, "T")
                    Me.txtTaxco(1).Text = DBLet(miRsAux!bombaModelo, "T")
                    Me.txtTaxco(2).Text = DBLet(miRsAux!motorModelo, "T")
                    Me.txtTaxco(3).Text = DBLet(miRsAux!motormarca, "T")
                    Me.txtTaxco(4).Text = DBLet(miRsAux!ReferPedido, "T")
                    Me.txtTaxco(5).Text = DBLet(miRsAux!RecepAgenCliMat, "T")
                    Me.txtTaxco(6).Text = DBLet(miRsAux!RecpNumExp, "T")
                    Me.txtTaxco(7).Text = DBLet(miRsAux!numrepar, "T")
                    
                Else
                    For N = 0 To txtTaxco.Count - 1
                        txtTaxco(N).Text = ""
                    Next
        
                End If
            End If
        
            'Agencia cliente
            N = 1
            If DBLet(miRsAux!TipoPortes, "N") = 0 Then N = 0
            optEuler(N).Value = True
            
           
            
            ''Empieza en la 20
            'For N = 1 To Me.chkEuler.Count
            '    chkEuler(N - 1).Value = DBLet(miRsAux.Fields(CInt(N) + 19), "N")
            'Next
            
            txtEuler(8).Text = ""
            If Data3.Recordset!Codtipoa = "ALR" Then
                
                Sql = ""
                Cad2 = DBLet(miRsAux!NumParteTrabajo, "T")
                If Cad2 <> "" Then Sql = Sql & "Orden de trabajo: " & Cad2
                
                Cad2 = DBLet(miRsAux!NumTrabajExterno, "T")
                If Cad2 <> "" Then Sql = Sql & "Trabajo exterior: " & Cad2
                
                Cad2 = DBLet(miRsAux!RecepAgenCliMat, "T")
                If Cad2 <> "" Then
                    Sql = Sql & vbCrLf & "Agen/Cli/Matr: " '& cad2
                    Cad2 = Cad2 & "  [" & IIf(DBLet(miRsAux!RecepAgenClien, "T") = 0, "Agencia", "Cliente") & "]"
                    
                    If Not IsNull(miRsAux!RecpNumExp) Then Cad2 = Cad2 & "  Expediente: " & miRsAux!RecpNumExp & " " & DBLet(miRsAux!FechaAlb, "T")
                    If Not IsNull(miRsAux!RecepPortes) Then
                        Cad2 = Cad2 & vbCrLf & "Portes: "
                        If miRsAux!RecepPortes = 0 Then
                            Cad2 = Cad2 & "Debidos"
                        Else
                            Cad2 = Cad2 & "pagados"
                        End If
                    End If
                End If
                If Cad2 <> "" Then Sql = Sql & Cad2
                If Sql <> "" Then txtEuler(8).Text = "RECEPCION DEL EQUIPO" & vbCrLf & String(40, "-") & Sql
                
                
                ',TipoBombResSuperVer,TipoBombLimSuperHor,TipoBombLimSuperVer,TipoBombResSumPoz,TipoBombLimSumPoz,TipoBombResSumVer"
                'TipoBombLimSumVer,TipoBomAgitadorRes,TipoBomAgitadorLim,TipoBomResOtrosEqu,TipoBomLimOtrosEqu,"
                
                '
                Sql = ""
                Cad2 = ""
                If Not IsNull(miRsAux!DatosBommarca) Then Cad2 = Cad2 & "Marca: " & miRsAux!DatosBommarca & vbCrLf
                If Not IsNull(miRsAux!DatosBomNumCurva) Then Cad2 = Cad2 & "Curva: " & miRsAux!DatosBomNumCurva & vbCrLf
                If Not IsNull(miRsAux!DatosBomModelo) Then Cad2 = Cad2 & "Modelo: " & miRsAux!DatosBomModelo & vbCrLf
                If Not IsNull(miRsAux!DatosBomNumSerie) Then Cad2 = Cad2 & "Serie: " & miRsAux!DatosBomNumSerie & vbCrLf
                If Not IsNull(miRsAux!DatosBomAno) Then Cad2 = Cad2 & "Año: " & miRsAux!DatosBomAno & vbCrLf
                
        
                If Cad2 <> "" Then Sql = "Parte hidraulica" & vbCrLf & Cad2
                
                Cad2 = ""
                If Not IsNull(miRsAux!DatosMotorMarca) Then Cad2 = Cad2 & "Marca: " & miRsAux!DatosMotorMarca & vbCrLf
                If Not IsNull(miRsAux!DatosMotorModelo) Then Cad2 = Cad2 & "Modelo: " & miRsAux!DatosMotorModelo & vbCrLf
                If Not IsNull(miRsAux!DatosMotorNumSerie) Then Cad2 = Cad2 & "NºSerie: " & miRsAux!DatosMotorNumSerie & vbCrLf
                If Not IsNull(miRsAux!DatosMotorV) Then Cad2 = Cad2 & "V: " & miRsAux!DatosMotorV & vbCrLf
                If Not IsNull(miRsAux!DatosMotorI) Then Cad2 = Cad2 & "I(A): " & miRsAux!DatosMotorI & vbCrLf
                If Not IsNull(miRsAux!DatosMotorCV) Then Cad2 = Cad2 & "CV: " & miRsAux!DatosMotorCV & vbCrLf
                If Not IsNull(miRsAux!DatosMotorKw) Then Cad2 = Cad2 & "KW: " & miRsAux!DatosMotorKw & vbCrLf
                If Not IsNull(miRsAux!DatosMotorrpm) Then Cad2 = Cad2 & "RPM: " & miRsAux!DatosMotorrpm & vbCrLf
                
                'Tipo rodete
                If DBLet(miRsAux!DatosBomTipoRodete, "N") > 0 Then Cad2 = Cad2 & "Rodete: " & RecuperaValor("C|N|O|V|", miRsAux!DatosBomTipoRodete - 3) & vbCrLf
                
                If Cad2 <> "" Then
                    If Sql <> "" Then Sql = Sql & vbCrLf & vbCrLf
                    Sql = Sql & "MOTOR" & vbCrLf & Cad2
                End If
                
                If Sql <> "" Then txtEuler(8).Text = txtEuler(8).Text & vbCrLf & vbCrLf & "DATOS EQUIPO" & vbCrLf & String(40, "-") & vbCrLf & Sql
    
            End If  'de alr
        End If
        miRsAux.Close
    End If 'todo
    
    
    
    
    
    
    
    
    'Carga costes euler
    'ImpMano
    'Total
    If Todo Then
        Me.lwCostes.ListItems.Clear
        Sql = " *, if(tipo=0,0,if(tipo=5,2,tipo*10)) orden1  "   ' if(tipo=0,0,1) orden1 "
        Sql = "Select " & Sql & " FROM  slifac_eu " & ObtenerWhereCP(True)
        Sql = Sql & " order by orden1,fechamov"
        miRsAux.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        N2 = 0
        ImpMano = 0
        impVeh = 0
        Cad2 = ""
        total = 0
        While Not miRsAux.EOF
            N2 = N2 + 1
            
            Select Case miRsAux!Tipo
            Case 0
                'Horas
                lwCostes.ListItems.Add , , "HOR"
                lwCostes.ListItems(N2).SubItems(1) = "Horas trabajadas"
                lwCostes.ListItems(N2).SubItems(2) = " "
                lwCostes.ListItems(N2).SubItems(3) = " "
                lwCostes.ListItems(N2).SubItems(4) = " "
                
            Case 1, 2
                'Materia prima
                lwCostes.ListItems.Add , , IIf(miRsAux!Tipo = 1, "MAT", "ALB")
                lwCostes.ListItems(N2).SubItems(1) = IIf(miRsAux!Tipo = 1, "Material auxiliar", " ")
                lwCostes.ListItems(N2).SubItems(2) = " "
                lwCostes.ListItems(N2).SubItems(3) = " "
                
               
            Case 5
                'Vehiculo
                lwCostes.ListItems.Add , , "VEH"
                lwCostes.ListItems(N2).SubItems(1) = "Gastos vehiculo"
                lwCostes.ListItems(N2).SubItems(2) = " "
                lwCostes.ListItems(N2).SubItems(3) = miRsAux!NomArtic & " "
                lwCostes.ListItems(N2).SubItems(4) = " "
                
            Case Else
                'proveedor
                lwCostes.ListItems.Add , , "PRO"
                lwCostes.ListItems(N2).SubItems(1) = DBLet(miRsAux!Aux, "T") & " "
                lwCostes.ListItems(N2).SubItems(2) = DBLet(miRsAux!Documento, "T") & " "
                
                
            End Select
            
            If miRsAux!Tipo <> 0 Then
                lwCostes.ListItems(N2).SubItems(4) = miRsAux!NomArtic
                lwCostes.ListItems(N2).ListSubItems(4).ToolTipText = miRsAux!NomArtic
                lwCostes.ListItems(N2).SubItems(3) = miRsAux!FechaMov
                
            End If
            'Cantidad
            lwCostes.ListItems(N2).SubItems(5) = Format(miRsAux!cantidad, FormatoImporte)
            lwCostes.ListItems(N2).SubItems(6) = Format(miRsAux!precioar, FormatoImporte)
            Impo = Round2(miRsAux!precioar * miRsAux!cantidad, 2)
            lwCostes.ListItems(N2).SubItems(7) = Format(Impo, FormatoImporte)
            
            If miRsAux!Tipo = 0 Then
                ImpMano = ImpMano + Impo
            Else
                If miRsAux!Tipo = 5 Then
                    impVeh = impVeh + Impo
                Else
                    lwCostes.ListItems(N2).ListSubItems(7).Tag = DBLet(miRsAux!codArtic, "T")
                End If
            End If
            total = total + Impo
            
            'Para modificar y borrar
            'KEY
            'codtipom numfactu fecfactu codtipoa numalbar numlinea Tipo
            Sql = miRsAux!codtipom & "|" & miRsAux!Numfactu & "|" & miRsAux!FecFactu & "|" & miRsAux!Codtipoa & "|" & miRsAux!Numalbar & "|" & miRsAux!numlinea & "|" & miRsAux!Tipo & "|"
            lwCostes.ListItems(N2).Tag = Sql
            
            If miRsAux!Tipo >= 3 Then lwCostes.ListItems(N2).ListSubItems(2).ToolTipText = DBLet(miRsAux!Documento, "T")
            
            
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        If N2 > 0 Then
            If ImpMano <> 0 Then
                N2 = N2 + 1
                lwCostes.ListItems.Add , , " "
                For N = 1 To 5
                    lwCostes.ListItems(N2).SubItems(N) = " "
                Next
                lwCostes.ListItems(N2).SubItems(6) = "Mano obra"
                lwCostes.ListItems(N2).SubItems(7) = Format(ImpMano, FormatoImporte)
            End If
            
            If impVeh <> 0 Then
                N2 = N2 + 1
                lwCostes.ListItems.Add , , " "
                For N = 1 To 5
                    lwCostes.ListItems(N2).SubItems(N) = " "
                Next
                lwCostes.ListItems(N2).SubItems(6) = "Vehiculos"
                lwCostes.ListItems(N2).SubItems(7) = Format(impVeh, FormatoImporte)
            End If
            
            If total <> 0 Then Impo = total - ImpMano - impVeh
            
            If Impo <> 0 Then
                lwCostes.ListItems.Add , , " "
                N2 = N2 + 1
                For N = 1 To 5
                    lwCostes.ListItems(N2).SubItems(N) = " "
                Next
                lwCostes.ListItems(N2).SubItems(6) = "Materiales"
                lwCostes.ListItems(N2).SubItems(7) = Format(Impo, FormatoImporte)
                
            End If
                
            If total <> 0 Then
                lwCostes.ListItems.Add , , " "
                N2 = N2 + 1
                For N = 1 To 5
                    lwCostes.ListItems(N2).SubItems(N) = " "
                Next
                lwCostes.ListItems(N2).SubItems(6) = "  TOTAL"
                lwCostes.ListItems(N2).SubItems(7) = Format(total, FormatoImporte)
            End If
        End If
    End If 'todo
        
        
        
    Me.lwEulerLineas.ListItems.Clear
    lwEulerLineas.Tag = ""
    Sql = "Select codtipoa,numalbar,numlinea,articulo,descrarticulo,cantidad,precioar,dtoline1,importel FROM  slifac_eu2 " & ObtenerWhereCP(True)
    Sql = Sql & " order by numalbar,numlinea"
    miRsAux.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    ImpMano = 0
    N2 = 0
    While Not miRsAux.EOF
        N2 = N2 + 1
        lwEulerLineas.ListItems.Add , "k" & Format(miRsAux!numlinea, "000") & miRsAux!Numalbar, miRsAux!Articulo
        lwEulerLineas.ListItems(N2).SubItems(1) = Replace(miRsAux!descrarticulo, Chr(13), " ")
        lwEulerLineas.ListItems(N2).SubItems(2) = Format(miRsAux!cantidad, FormatoCantidad)
        lwEulerLineas.ListItems(N2).SubItems(3) = Format(miRsAux!precioar, FormatoPrecio)
        lwEulerLineas.ListItems(N2).SubItems(4) = Format(miRsAux!dtoline1, FormatoCantidad)
        lwEulerLineas.ListItems(N2).SubItems(5) = Format(miRsAux!ImporteL, FormatoCantidad)
        lwEulerLineas.ListItems(N2).ToolTipText = miRsAux!descrarticulo
        ImpMano = ImpMano + miRsAux!ImporteL
        
        'Agosto 2021
        'El tag llevará el SQL de seleccion del albaran
        lwEulerLineas.ListItems(N2).Tag = miRsAux!Codtipoa & "|" & miRsAux!Numalbar & "|"
        
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
        
    If N2 > 0 Then
        'Tiene lineas y NO suma el burto
        If Data1.Recordset!BrutoFac <> ImpMano Then
            Sql = "Bruto factura: " & Data1.Recordset!BrutoFac & vbCrLf
            Sql = Sql & "Suma lineas: " & ImpMano
            lwEulerLineas.Tag = Sql
            MsgBox Sql, vbExclamation
        End If
    End If
        
        
        
        
        
        
        
        
    Set miRsAux = Nothing
End Sub

Private Sub LimpiarFichaTecnica(SinTxts As Boolean)
Dim N As Byte
    

    If Not SinTxts Then
        For N = 0 To Me.txtEuler.Count - 1
            txtEuler(N).Text = ""
        Next
        
        
        
        If vParamAplic.NumeroInstalacion = vbTaxco Then
            For N = 0 To Me.txtTaxco.Count - 1
                txtTaxco(N).Text = ""
            Next
        End If
    End If
    
    Me.optEuler(0).Value = True
    Me.optEuler(0).Value = False  'Ninguno seleccionado
    
    
    lwCostes.ListItems.Clear
    
End Sub




Private Sub ImprimirValoracionOferta()
If Modo <> 2 Then Exit Sub
    If Me.Data1.Recordset Is Nothing Then Exit Sub
    If Me.Data1.Recordset.EOF Then Exit Sub

   
    NumRegElim = 81
    BuscaChekc = DevuelveDesdeBDNew(conAri, "scryst", "documrpt", "codcryst", CStr(NumRegElim), "N")
    
    If BuscaChekc = "" Then Exit Sub
    
    

    
    
    
    With frmImprimir
            'Cod Tipo Movimiento
            BuscaChekc = "{scafac.codtipom}='" & Text1(1).Text & "'"
            'Nº Factura
            BuscaChekc = BuscaChekc & " AND {scafac.numfactu}=" & Val(Text1(0).Text)
            'Fecha Factura
            BuscaChekc = BuscaChekc & " AND {scafac.fecfactu}= Date(" & Year(Text1(2).Text) & "," & Month(Text1(2).Text) & "," & Day(Text1(2).Text) & ")"
            
  
           
'           .outClaveNombreArchiv = devuelve & Format(Text1(0).Text, "000")
'           .outCodigoCliProv = Text1(4).Text
           .outTipoDocumento = 0
           .SeleccionaRPTCodigo = 0
           .FormulaSeleccion = BuscaChekc
           .OtrosParametros = "|pEmpresa=""" & vParam.NombreEmpresa & """|"
           .NumeroParametros = 2
           .NombreRPT = DevuelveDesdeBDNew(conAri, "scryst", "documrpt", "codcryst", CStr(NumRegElim), "N")
           .NombrePDF = .NombreRPT
           .SoloImprimir = False
           .EnvioEMail = False
           .Titulo = "Valoracion factura"
           .NumeroCopias = 1
           .Opcion = 2000
           
           .Show vbModal
    End With


    
End Sub



Private Sub ModificaLote()
Dim CadenaInsertTmpLotes As String
Dim J As Integer
Dim LotesArticulos
Dim IncioBusqueda As Integer
Dim fin As Boolean
Dim Sql As String
       
          
        If Not vParamAplic.ManipuladorFitosanitarios2 Then Exit Sub   'Por si acaso se ha metido aqui
                   
        Sql = DevuelveDesdeBD(conAri, "numserie", "sartic", "codartic", Data2.Recordset!codArtic, "T")
        If Sql = "" Then Exit Sub
        Set miRsAux = New ADODB.Recordset
        'codtipom numfactu fecfactu codtipoa numalbar numlinea
        CadenaInsertTmpLotes = "codtipom ='" & Data1.Recordset!codtipom & "' AND numfactu =" & Data1.Recordset!Numfactu
        CadenaInsertTmpLotes = CadenaInsertTmpLotes & " AND fecfactu='" & Format(Data1.Recordset!FecFactu, FormatoFecha) & "' AND codtipoa = '" & Data3.Recordset!Codtipoa
        CadenaInsertTmpLotes = CadenaInsertTmpLotes & "' AND numalbar = " & Data3.Recordset!Numalbar & " AND numlinea =" & Data2.Recordset!numlinea
        CadenaInsertTmpLotes = "Select numlote,cantidad,fecentra from slifaclotes  WHERE " & CadenaInsertTmpLotes & "  order by sublinea"
 
        miRsAux.Open CadenaInsertTmpLotes, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        LotesArticulos = "|"
        While Not miRsAux.EOF
            LotesArticulos = LotesArticulos & miRsAux!numLote & "#@#" & Format(miRsAux!fecentra, "dd/mm/yyyy") & Mid(miRsAux!cantidad & Space(10), 1, 10) & "|"
          
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        
            CadenaInsertTmpLotes = ""
            Sql = "select codartic,numlotes,fecentra,canentra-vendida disponible from slotes where "
            Sql = Sql & " codartic=" & DBSet(Data2.Recordset!codArtic, "T") & " and canentra-vendida >0 order by fecentra "
            
            miRsAux.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            NumRegElim = 0
            While Not miRsAux.EOF
                NumRegElim = NumRegElim + 1
                'insert into tmpnlotes(codusu,numlinea,fechaalb,codprove,cantidad,numlotes)
                CadenaInsertTmpLotes = CadenaInsertTmpLotes & ", (" & vUsu.Codigo & "," & DBSet(miRsAux!codArtic, "T") & "," & NumRegElim
                CadenaInsertTmpLotes = CadenaInsertTmpLotes & "," & DBSet(miRsAux!fecentra, "F")
                'CadenaInsertTmpLotes = CadenaInsertTmpLotes & "," & DBSet(txtAux(2).Text, "T") & "," & DBSet(txtAux2(2).Text, "T")
                CadenaInsertTmpLotes = CadenaInsertTmpLotes & "," & DBSet(miRsAux!disponible * 100, "N") & ","
                                
                Sql = "|" & miRsAux!numlotes & "#@#"
                fin = False
                IncioBusqueda = 1
                
                While Not fin
                    
                     
                    J = InStr(IncioBusqueda, LotesArticulos, Sql)
                    If J > 0 Then
                        J = J + Len(Sql)
                        Sql = Mid(LotesArticulos, J, 10)
                        
                        If Sql = Format(miRsAux!fecentra, "dd/mm/yyyy") Then
                            'OK, esta es la linea
                            Sql = Trim(Mid(LotesArticulos, J + 10, 10))
                            fin = True
                        Else
                            Sql = "|" & miRsAux!numlotes & "#@#"   'Vuelve a la busqueda
                            IncioBusqueda = InStr(J, LotesArticulos, "|")
                        End If
                    Else
                        Sql = "0"
                        fin = True
                    End If
                Wend
                CadenaInsertTmpLotes = CadenaInsertTmpLotes & Sql
                
                CadenaInsertTmpLotes = CadenaInsertTmpLotes & "," & DBSet(miRsAux!numlotes, "T") & ")"
                
                miRsAux.MoveNext
            Wend
            miRsAux.Close
            Set miRsAux = Nothing
                
                
            'Si hay mas de uno mostraremos cual y cuanto puede coger
            If NumRegElim = 0 Then
                MsgBox "Ningun lote disponible para el artículo", vbExclamation
              
            Else
              
                'Mas de un  lote disponible
                Screen.MousePointer = vbHourglass
                
                conn.Execute "DELETE FROM tmpnlotes where codusu =" & vUsu.Codigo
                Espera 0.3
                CadenaInsertTmpLotes = Mid(CadenaInsertTmpLotes, 2)
                CadenaInsertTmpLotes = "insert into tmpnlotes(codusu,codartic,numlinea,fechaalb,codprove,cantidad,numlotes) VALUES " & CadenaInsertTmpLotes
                conn.Execute CadenaInsertTmpLotes
                
                
                
              
                    CadenaDesdeOtroForm = ""
                    frmFacTPVLotes.TotalLineas = Data2.Recordset!cantidad
                    frmFacTPVLotes.NombreArticulo = Data2.Recordset!NomArtic
                    frmFacTPVLotes.Show vbModal
              
                    If CadenaDesdeOtroForm = "OK" Then
                    
                        'Primero devolveremos la cantidad que tenia la linea
                        ReestableceLotesArticulo
                        
                        'Borramos la linea de lotes
                        Sql = Sql_Lineas_Lotes
                        Sql = Mid(Sql, InStr(1, Sql, " WHERE "))
                        Sql = "DELETE FROM slifaclotes  " & Sql
                        conn.Execute Sql
                        Espera 0.4
                        
                        Sql = "INSERT INTO slifaclotes(codtipom,numfactu,fecfactu,codtipoa,numalbar,numlinea,sublinea,cantidad,numlote,fecentra,codartic)"
                        Sql = Sql & " SELECT '" & Data1.Recordset!codtipom & "'," & Data1.Recordset!Numfactu & ",'" & Format(Data1.Recordset!FecFactu, FormatoFecha) & "' ,'" & Data3.Recordset!Codtipoa
                        Sql = Sql & "'," & Data3.Recordset!Numalbar & "," & Data2.Recordset!numlinea
                        Sql = Sql & " , numlinea , Cantidad, numlotes,fechaalb,codartic "
                        Sql = Sql & " FROM tmpnlotes  WHERE codusu = " & vUsu.Codigo & " and cantidad <>0 "
            
                        conn.Execute Sql
                        
                        'Tengo que updatear la cantidad vendida
                        Set miRsAux = New ADODB.Recordset
                        miRsAux.Open "Select * FROM tmpnlotes  WHERE codusu = " & vUsu.Codigo & " and cantidad <>0 ", conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                        While Not miRsAux.EOF
                            If miRsAux!cantidad <> 0 Then
                                If miRsAux!cantidad > 0 Then
                                    Sql = "+"
                                Else
                                    Sql = "-"
                                End If
                                Sql = "UPDATE slotes SET vendida=vendida " & Sql & DBSet(Abs(miRsAux!cantidad), "N")
                                Sql = Sql & " WHERE numlotes =" & DBSet(miRsAux!numlotes, "T") & " AND codartic= " & DBSet(miRsAux!codArtic, "T")
                                Sql = Sql & " AND fecentra= " & DBSet(miRsAux!FechaAlb, "F")
                            
                                conn.Execute Sql
                            End If
                            miRsAux.MoveNext
                        Wend
                        miRsAux.Close
                    End If
            
            

                    Espera 0.3
                        
                        
                    
              
            End If


    

End Sub


Private Function Sql_Lineas_Lotes() As String
        Sql_Lineas_Lotes = "codtipom ='" & Data1.Recordset!codtipom & "' AND numfactu =" & Data1.Recordset!Numfactu
        Sql_Lineas_Lotes = Sql_Lineas_Lotes & " AND fecfactu='" & Format(Data1.Recordset!FecFactu, FormatoFecha) & "' AND codtipoa = '" & Data3.Recordset!Codtipoa
        Sql_Lineas_Lotes = Sql_Lineas_Lotes & "' AND numalbar = " & Data3.Recordset!Numalbar & " AND numlinea =" & Data2.Recordset!numlinea
        Sql_Lineas_Lotes = "Select * from slifaclotes  WHERE " & Sql_Lineas_Lotes
        Sql_Lineas_Lotes = Sql_Lineas_Lotes & " AND numlinea =" & Data2.Recordset!numlinea
        
End Function

Private Sub ReestableceLotesArticulo()
        
        BuscaChekc = DevuelveDesdeBD(conAri, "numserie", "sartic", "codartic", Data2.Recordset!codArtic, "T")
        If Trim(BuscaChekc) <> "" Then
            Set miRsAux = New ADODB.Recordset
            
            
            BuscaChekc = Sql_Lineas_Lotes
            miRsAux.Open BuscaChekc, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not miRsAux.EOF
                If miRsAux!cantidad <> 0 Then
                    'Vamos a actualizar VENDIDA
                    'Luego si estoy reestableciendo
                    If miRsAux!cantidad < 0 Then
                        BuscaChekc = "+"
                    Else
                        BuscaChekc = "-"
                    End If
                    BuscaChekc = "UPDATE slotes SET vendida = vendida " & BuscaChekc & " " & DBSet(Abs(miRsAux!cantidad), "N")
                    BuscaChekc = BuscaChekc & " WHERE numlotes= " & DBSet(miRsAux!numLote, "T")
                    BuscaChekc = BuscaChekc & " AND codArtic= " & DBSet(miRsAux!codArtic, "T")
                    BuscaChekc = BuscaChekc & " AND fecentra = " & DBSet(miRsAux!fecentra, "F")
                    conn.Execute BuscaChekc
                End If
                miRsAux.MoveNext
            Wend
            miRsAux.Close
            Set miRsAux = Nothing
        End If
End Sub



Private Sub ImprimirCostesEuler()

    On Error GoTo eImprimirCostesEuler
    If Modo <> 2 Then Exit Sub
    
    

    With frmImprimir
        .NombreRPT = "EULFacturaCostes.rpt"
        .FormulaSeleccion = "{" & NombreTabla & ".codtipom}='" & Text1(1).Text & "' AND {" & NombreTabla & ".numfactu}=" & Val(Text1(0).Text) & " AND {" & NombreTabla & ".fecfactu}= Date(" & Year(Text1(2).Text) & "," & Month(Text1(2).Text) & "," & Day(Text1(2).Text) & ")"
        .OtrosParametros = "|pCodUsu=" & vUsu.Codigo & "|"
        .NumeroParametros = 1
        .Titulo = "Costes EULER"
        .SoloImprimir = False
        .EnvioEMail = False
        .Opcion = 2000 '2000 generico
        .ConSubInforme = True
        .Show vbModal
    End With
    
    
    
    
    
    
    
eImprimirCostesEuler:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    
End Sub







'--------------------------------------------------------------------------------
'Private Function CargaCostesEuler2() As Boolean
'Dim C As String
'
'    On Error GoTo eCargaCostesEuler
'
'
'    C = "Select     "
'
'
'
'
'eCargaCostesEuler:
'    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
'
'    lblIndicador.Caption = lblIndicador.Tag
'    Screen.MousePointer = vbDefault
'End Function






'-------------------------------------
' Abrir PDFs vinculados
Private Sub AbrirPDFs()
    Screen.MousePointer = vbHourglass
    lblIndicador.Tag = lblIndicador.Caption
    
    'MOntamos la cadena con los pDFS para abrir en listview
    CadenaDesdeOtroForm = ""
    If vParamAplic.NumeroInstalacion = 1 Then
        'ALZIRA
        MontaPDFsAlzira
        
    Else
        'DE momento SOLO 4tonda
        
        MontaPDFs4Tonda
        
    End If
    
    If CadenaDesdeOtroForm = "" Then
        MsgBox "Ningun dato a reimprmir", vbExclamation
    Else
        frmListado5.OpcionListado = 21
        frmListado5.Show vbModal
        
    End If
    
    CadenaDesdeOtroForm = ""
    lblIndicador.Caption = lblIndicador.Tag
    lblIndicador.Tag = ""
    Screen.MousePointer = vbDefault
End Sub


Private Function MontaPDFsAlzira()
Dim C As String

    'Veremos si tiene
    On Error GoTo eAbrirPDFsAlzira
    
    
    'Factura
    If vParamAplic.PathFirmasFacturas <> "" Then
    
        C = DevuelveDesdeBD(conAri, "letraser", "stipom", "codtipom", CStr(Data1.Recordset!codtipom), "T")
        'S03-0215925_
        C = "_" & C & "-" & Format(Data1.Recordset!Numfactu, "0000000") & "_"
        
        C = Dir(vParamAplic.PathFirmasFacturas & "\" & Format(Data1.Recordset!FecFactu, FormatoFecha) & "\*" & C & "*.pdf", vbArchive)
        If C <> "" Then
            C = vParamAplic.PathFirmasFacturas & "\" & Format(Data1.Recordset!FecFactu, FormatoFecha) & "\" & C
            AnchoLogin = Format(Data1.Recordset!FecFactu, "dd/mm/yyyy") & Data1.Recordset!codtipom & Format(Data1.Recordset!Numfactu, "0000000") & "#"
            CadenaDesdeOtroForm = "@" & AnchoLogin & C & "@"
        
        End If
    End If
    
    
    
    
    If vParamAplic.PathFirmasAlbaran <> "" And CadenaDesdeOtroForm = "" Then
        C = ObtenerWhereCP(False)
        C = "Select numalbar,codtipoa,fechaalb FROM scafac1 where " & C & " ORDER BY fechaalb"
        Set miRsAux = New ADODB.Recordset
        miRsAux.Open C, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            
            '20762090J_ALV-0010699.pdf_       B96940374_ALV-0012320_16.47.47.5916
            C = DBLet(miRsAux!Codtipoa, "T")
            C = "_" & C & "-" & Format(miRsAux!Numalbar, "0000000")
            C = Dir(vParamAplic.PathFirmasAlbaran & "\" & Format(miRsAux!FechaAlb, FormatoFecha) & "\*" & C & "*.pdf", vbArchive)
            If C <> "" Then
                C = vParamAplic.PathFirmasAlbaran & "\" & Format(miRsAux!FechaAlb, FormatoFecha) & "\" & C
                AnchoLogin = Format(miRsAux!FechaAlb, "dd/mm/yyyy") & miRsAux!Codtipoa & Format(miRsAux!Numalbar, "0000000") & "#"
                CadenaDesdeOtroForm = CadenaDesdeOtroForm & AnchoLogin & C & "@"
            End If
        
            miRsAux.MoveNext
            
        Wend
        miRsAux.Close
        Set miRsAux = Nothing
    End If
    
    
   
    
    Exit Function
eAbrirPDFsAlzira:
    MuestraError Err.Number, Err.Description
    Set miRsAux = Nothing
End Function


Private Function MontaPDFs4Tonda()
Dim C As String

    'Veremos si tiene
    On Error GoTo eMontaPDFs4Tonda
    
    
    
    '4Tonda monta Path\AÑO\mes(00)\
    
    
    'Factura
    If vParamAplic.PathFirmasFacturas <> "" Then
    
        
    End If
    
    
    
    
    If vParamAplic.PathFirmasAlbaran <> "" And CadenaDesdeOtroForm = "" Then
        C = ObtenerWhereCP(False)
        C = "Select numalbar,codtipoa,fechaalb FROM scafac1 where " & C & " ORDER BY fechaalb"
        Set miRsAux = New ADODB.Recordset
        miRsAux.Open C, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            
            '20762090J_ALV-0010699.pdf_
            C = Year(miRsAux!FechaAlb) & "\" & Format(Month(miRsAux!FechaAlb), "00") & "\" & "A-" & Format(miRsAux!Numalbar, "0000000")
            C = Dir(vParamAplic.PathFirmasAlbaran & "\" & C & "*.pdf", vbArchive)
            If C <> "" Then
                C = vParamAplic.PathFirmasAlbaran & "\" & Year(miRsAux!FechaAlb) & "\" & Format(Month(miRsAux!FechaAlb), "00") & "\" & C
                AnchoLogin = Format(miRsAux!FechaAlb, "dd/mm/yyyy") & miRsAux!Codtipoa & Format(miRsAux!Numalbar, "0000000") & "#"
                CadenaDesdeOtroForm = CadenaDesdeOtroForm & AnchoLogin & C & "@"
            End If
        
            miRsAux.MoveNext
            
        Wend
        miRsAux.Close
        Set miRsAux = Nothing
    End If
    
    
   
   
    Exit Function
eMontaPDFs4Tonda:
    MuestraError Err.Number, Err.Description
    Set miRsAux = Nothing
End Function






Private Sub ModoBusquedaCostes(Busqueda As Boolean)
Dim i As Integer

    If Busqueda Then
        Me.lwCostes.Height = 3255
        Me.txtCostes(0).Left = Me.lwCostes.Left
        Me.txtCostes(0).Width = Me.lwCostes.ColumnHeaders(1).Width
    Else
        Me.lwCostes.Height = 3615
    End If
    
    txtCostes(0).visible = Busqueda
    
    For i = 1 To txtCostes.Count - 1
        txtCostes(i).visible = Busqueda
        If Busqueda Then
            txtCostes(i).Left = txtCostes(i - 1).Left + txtCostes(i - 1).Width
            txtCostes(i).Width = Me.lwCostes.ColumnHeaders(i + 1).Width
        End If
    Next i
     
End Sub


Private Function DevuelveBusquedaCostesEuler() As String
Dim i As Byte
Dim EsLike As Boolean
Dim Aux As String
Dim J As Integer

    DevuelveBusquedaCostesEuler = ""

    For i = 0 To Me.txtCostes.Count - 1
        Me.txtCostes(i).Text = Trim(Me.txtCostes(i).Text)
        If Me.txtCostes(i).Text <> "" Then


            'Los textos
            Select Case i
            Case 0
                'Si es PRO , HOR, ALB, MAT
                txtCostes(i).Text = UCase(txtCostes(i).Text)
                Aux = ""
                If txtCostes(i).Text = "HOR" Then
                    Aux = " = 0"
                Else
                    If txtCostes(i).Text = "MAT" Then
                        Aux = " = 1"
                    ElseIf txtCostes(i).Text = "ALB" Then
                        Aux = " = 2"
                    ElseIf txtCostes(i).Text = "PRO" Then
                        Aux = " > 2"
                    End If
                End If
                If Aux <> "" Then
                    DevuelveBusquedaCostesEuler = DevuelveBusquedaCostesEuler & " AND tipo " & Aux
                Else
                    txtCostes(i).Text = ""  'no me sirve lo que han puesto
                End If
                Aux = "" 'Ya hemos concatenado la cadena de busqueda. Para que no lo vuelva a hacer: ""
            Case 3
                If SeparaCampoBusqueda("F", "slifac_eu.fechamov", txtCostes(i).Text, Aux, False) > 0 Then
                    Aux = ""
                Else
                    Aux = " AND " & Aux
                End If
               
            
                
            Case 5, 6, 7

                If SeparaCampoBusqueda("N", RecuperaValor("slifac_eu.cantidad|slifac_eu.cantidad|(slifac_eu.cantidad * slifac_eu.precioar)|", i - 4), txtCostes(i).Text, Aux) > 0 Then
                    Aux = ""
                Else
                    Aux = " AND " & Aux
                End If
            
            Case Else
            
                If InStr(1, txtCostes(i).Text, "*") > 0 Then
                    Aux = " like " & DBSet(Replace(Me.txtCostes(i).Text, "*", "%"), "T")
                Else
                    Aux = " = " & DBSet(Me.txtCostes(i).Text, "T")
                End If
                Aux = " AND " & RecuperaValor("aux|documento||nomartic|", i + 0) & Aux
                
            End Select
            If Aux <> "" Then DevuelveBusquedaCostesEuler = DevuelveBusquedaCostesEuler & Aux
        End If
    Next

    If DevuelveBusquedaCostesEuler <> "" Then DevuelveBusquedaCostesEuler = Mid(DevuelveBusquedaCostesEuler, 5)        'quitamos el primer and


    
End Function



Private Function DevuelveBusquedaTaxco() As String
Dim i As Byte
Dim EsLike As Boolean
Dim Aux As String
Dim J As Integer

    DevuelveBusquedaTaxco = ""

    For i = 0 To Me.txtTaxco.Count - 1
        Me.txtTaxco(i).Text = Trim(Me.txtTaxco(i).Text)
        If Me.txtTaxco(i).Text <> "" Then

            
            If i = 7 Then
                'kilomnetros
                If SeparaCampoBusqueda("N", "numrepar", txtTaxco(i).Text, Aux) > 0 Then
                    Aux = ""
                Else
                    Aux = " AND " & Aux
                End If
            
            Else
                'resto camopos
                'bombamarca|bombaModelo|motorModelo|motormarca|ReferPedido|RecepAgenCliMat|RecpNumExp|
                If InStr(1, txtTaxco(i).Text, "*") > 0 Then
                    Aux = " like " & DBSet(Replace(Me.txtTaxco(i).Text, "*", "%"), "T")
                Else
                    Aux = " = " & DBSet(Me.txtTaxco(i).Text, "T")
                End If
                Aux = " AND " & RecuperaValor("bombamarca|bombaModelo|motorModelo|motormarca|ReferPedido|RecepAgenCliMat|RecpNumExp|", i + 1) & Aux
                
            End If
            If Aux <> "" Then DevuelveBusquedaTaxco = DevuelveBusquedaTaxco & Aux
        End If
    Next

    If DevuelveBusquedaTaxco <> "" Then DevuelveBusquedaTaxco = Mid(DevuelveBusquedaTaxco, 5)       'quitamos el primer and


    
End Function




Private Sub txtCostes_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtCostes_LostFocus(Index As Integer)
    txtCostes(Index).Text = Trim(txtCostes(Index).Text)
    If txtCostes(Index).Text = "" Then Exit Sub
    
    If Index = 0 Then
        If InStr(1, "PRO|HOR|ALB", UCase(txtCostes(Index).Text)) = 0 Then
            MsgBox "Valores posibles: HOR -PRO -ALB", vbExclamation
            txtCostes(Index).Text = ""
        End If
    End If
    
    
    
End Sub

Private Sub PonerImagenFirma()
Dim C As String
    On Error GoTo ePonerImagenFirma
    
    If CarpetaImagenesEULER = "" Then Exit Sub
    
  '  If PrimeraVez Then
    If UnaVez Then
        If Dir(CarpetaImagenesEULER, vbDirectory) = "" Then
            MsgBox "No existe carpeta: " & CarpetaImagenesEULER, vbExclamation
            CarpetaImagenesEULER = ""
        
        Else
            C = DevuelveDesdeBD(conAri, "carpetafirmas", "eulerparam", "1", "1")
            CarpetaImagenesEULER = CarpetaImagenesEULER & "\" & C
            
                
            If Dir(CarpetaImagenesEULER, vbDirectory) = "" Then
                MsgBox "No existe carpeta: " & CarpetaImagenesEULER, vbExclamation
                CarpetaImagenesEULER = ""
            End If
        End If
        Exit Sub
    End If
    imgFirmaRecep.visible = False
    If Modo <> 2 Then Exit Sub
    
    If Data3.Recordset.EOF Then
        C = ""
    Else
        C = CarpetaImagenesEULER & "\" & Mid(Data3.Recordset!Codtipoa & "   ", 1, 3) & Format(Data3.Recordset!Numalbar, "0000000") & ".jpg"
    End If
    If Dir(C, vbArchive) = "" Then C = ""
        
    If C <> "" Then
        imgFirmaRecep.visible = True
        imgFirmaRecep.Tag = C
    End If
    
    
    
    Exit Sub
ePonerImagenFirma:
    Err.Clear
    CarpetaImagenesEULER = ""
End Sub



Private Sub FijarCadenaModificaUsuarioNormal(Cambios As String)
Dim K As Integer
Dim cTag As cTag

    On Error GoTo eFijarCadenaModificaUsuarioNormal

    Cambios = ""
    TituloLinea = ""
    BuscaChekc = ""
    Set cTag = New cTag
    For K = 0 To Me.Text1.Count - 1
        If Text1(K).Tag <> "" Then
            If cTag.Cargar(Text1(K)) Then
                TituloLinea = cTag.columna
                BuscaChekc = ""
                If Not IsNull(Data1.Recordset.Fields(TituloLinea)) Then
                    If cTag.Formato <> "" Then
                        TituloLinea = cTag.columna
                        BuscaChekc = Format(Data1.Recordset.Fields(TituloLinea), cTag.Formato)
                    Else
                        If cTag.TipoDato = "F" Then
                           BuscaChekc = Format(Data1.Recordset.Fields(TituloLinea), "dd/mm/yyyy")
                        Else
                            BuscaChekc = Data1.Recordset.Fields(TituloLinea)
                        End If
                    End If
                End If
                
                'Peculiaridades
                If K = 23 Then If Data1.Recordset.Fields(TituloLinea) = 0 Then BuscaChekc = "0" 'dotpp
                If K = 24 Then If Data1.Recordset.Fields(TituloLinea) = 0 Then BuscaChekc = "0"  'dopie
                If K = 45 Then If Data1.Recordset.Fields(TituloLinea) = 0 Then BuscaChekc = ""   'aportacion
        
                
                
                If BuscaChekc <> Text1(K).Text Then Cambios = Cambios & cTag.Nombre & ": " & BuscaChekc & "   --- Ant: " & Text1(K).Text & vbCrLf
                
                
            End If
        End If
    Next
    
    
eFijarCadenaModificaUsuarioNormal:
    If Err.Number <> 0 Then
        Cambios = Err.Description & vbCrLf & Cambios
        Err.Clear
    End If
    Set cTag = Nothing
End Sub




Private Function HacerAnyadirAlbaranCosteEuler() As Boolean
Dim Sql As String
Dim Suma As Currency
    
    On Error GoTo eHacerAnyadirAlbaranCosteEuler
    
    HacerAnyadirAlbaranCosteEuler = False
    
    Set miRsAux = New ADODB.Recordset
    NumRegElim = 0
    
    'QUE NO TIENE lineas
    Sql = "SELECT * from slialb WHERE codtipom = '" & RecuperaValor(BuscaChekc, 1) & "' and numalbar =" & RecuperaValor(BuscaChekc, 2)
    miRsAux.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then Sql = ""
    miRsAux.Close
    
    If Sql = "" Then Err.Raise 513, , "Albaran con lineas de articulos"
    
    
    
    Sql = "select sreloj.codtraba,nomtraba,fecha,sreloj.codtipor,nomtipor,horainicio,horafin,calculadas from sreloj left join stipor on sreloj.codtipor=stipor.codtipor"
    Sql = Sql & " left join straba on straba.codtraba=sreloj.codtraba"
    Sql = Sql & " WHERE codtipom = '" & RecuperaValor(BuscaChekc, 1) & "' and numalbar =" & RecuperaValor(BuscaChekc, 2)
    Sql = Sql & " ORDER BY fecha,horainicio"
    Suma = 0
    TituloLinea = ""
    
    miRsAux.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        If IsNull(miRsAux!HoraFin) Then
            TituloLinea = TituloLinea & " - " & miRsAux!NomTipor & "    " & miRsAux!horainicio
        Else
            Suma = Suma + miRsAux!calculadas
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Sql = ""
    If TituloLinea <> "" Then Err.Raise 513, , "Tareas sin cerrar " & vbCrLf & TituloLinea
        
    'YA tenemos las horas
    TituloLinea = "Trabajos realizados : " & Format(Suma, FormatoImporte) & " horas"
    
    Sql = "SELECT count(*),round(cantidad * precioar,2),max(numlinea) numlinea from slialb_eu WHERE codtipom = '" & RecuperaValor(BuscaChekc, 1) & "' and numalbar =" & RecuperaValor(BuscaChekc, 2)
    miRsAux.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Sql = "0"
    If Not miRsAux.EOF Then
        TituloLinea = TituloLinea & vbCrLf & "Costes.  Lineas " & miRsAux.Fields(0) & "     Importe: " & Format(miRsAux.Fields(1), FormatoImporte)
        Sql = DBLet(miRsAux!numlinea, "N")
    End If
    miRsAux.Close
    
    
    TituloLinea = "Albaran: " & Replace(BuscaChekc, "|", "") & vbCrLf & TituloLinea
    TituloLinea = TituloLinea & vbCrLf & vbCrLf & "¿Insertar ?"
    If MsgBox(TituloLinea, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Function
    
    
    
    
    'OK. Va beenne
    'Sera una transaccion
    ' Insertar en lineas costes el nuevo coste
    ' Pasar de albaran a esta factura
    ' Borrar albaran
    conn.BeginTrans
    NumRegElim = Val(Sql) + 1   'Si numregelim >0 estamos dentro de la transaccion
    
    
    'Insertamos el albaran en scafac1
    Sql = "INSERT INTO scafac1(codtipom,numfactu,fecfactu,codtipoa,numalbar,fechaalb,codenvio,codtraba,codtrab1,codtrab2,"
    Sql = Sql & " observa1,observa2,observa5,coddiren,tipAlbaran,TipoCarnet,fechaAux) "
    Sql = Sql & " SELECT '" & Data1.Recordset!codtipom & "' as codtipom," & DBSet(Data1.Recordset!Numfactu, "N") & " as numfactu," & DBSet(Data1.Recordset!FecFactu, "F") & " as fecfactu," & "codtipom as codtipoa "
    Sql = Sql & " ,numalbar,fechaalb,codenvio,codtraba,codtrab1,codtrab2, observa01,observa02, "
    Sql = Sql & DBSet("Facturada.  " & vUsu.Login & "  " & Format(Now, "dd/mm/yy hh:nn"), "T") & " ,coddiren,tipAlbaran,TipoCarnet,fechaAux "
    Sql = Sql & " FROM scaalb WHERE codtipom = '" & RecuperaValor(BuscaChekc, 1) & "' and numalbar =" & RecuperaValor(BuscaChekc, 2)
    conn.Execute Sql
    
    
    'Insertamos slifac eu en dese slialb_eu
    Sql = "INSERT INTO slifac_eu(codtipom,numfactu,fecfactu,codtipoa,numalbar,numlinea,fechamov,codalmac,codartic,nomartic,cantidad,precioar,tipo) "
    Sql = Sql & " SELECT '" & Data1.Recordset!codtipom & "' as codtipom," & DBSet(Data1.Recordset!Numfactu, "N") & " as numfactu," & DBSet(Data1.Recordset!FecFactu, "F") & " as fecfactu," & "codtipom as codtipoa,  numalbar,numlinea,fechamov,codalmac,codartic,nomartic,cantidad,precioar,1 tipo"  'costes materiales
    Sql = Sql & " FROM slialb_eu WHERE codtipom = '" & RecuperaValor(BuscaChekc, 1) & "' and numalbar =" & RecuperaValor(BuscaChekc, 2)
    conn.Execute Sql
    
    
    'Insertamos slifaceu el coste horas
    If Suma > 0 Then
        Sql = "INSERT INTO slifac_eu(codtipom,numfactu,fecfactu,codtipoa,numalbar,numlinea,fechamov,codalmac,codartic,nomartic,cantidad,precioar,tipo) "
        Sql = Sql & " SELECT '" & Data1.Recordset!codtipom & "' as codtipom," & DBSet(Data1.Recordset!Numfactu, "N") & " as numfactu," & DBSet(Data1.Recordset!FecFactu, "F") & " as fecfactu,"
        Sql = Sql & DBSet(RecuperaValor(BuscaChekc, 1), "T") & " as codtipoa, " & RecuperaValor(BuscaChekc, 2) & " as numalbar, " & NumRegElim & " as numlinea"
        Sql = Sql & ", " & Format(Now, "dd/mm/yyyy") & " fechamov,1 codalmac,'' codartic,'Horas trabajadas ' nomartic,"
        Sql = Sql & TransformaComasPuntos(CStr(Suma)) & " cantidad," & TransformaComasPuntos(vParamAplic.PrecioHoraCosteEUL) & " precioar,0 tipo"
        conn.Execute Sql
        
    End If
    
    
    
    'Borramos
    Sql = " WHERE codtipom = '" & RecuperaValor(BuscaChekc, 1) & "' and numalbar =" & RecuperaValor(BuscaChekc, 2)
    conn.Execute "DELETE FROM slialb_eu " & Sql
    conn.Execute "DELETE FROM slialb_eu2 " & Sql
    conn.Execute "DELETE FROM slialb " & Sql
    conn.Execute "DELETE FROM scaalb " & Sql
    
        
            
    'HacerAnyadirAlbaranCosteEuler=true
    
    Sql = "Añadir albaran a factura.  " & vbCrLf & "Factura: " & Data1.Recordset!codtipom & Data1.Recordset!Numfactu & " " & Data1.Recordset!FecFactu
    Sql = Sql & vbCrLf & "Albaran " & RecuperaValor(BuscaChekc, 1) & "  " & RecuperaValor(BuscaChekc, 2)
    Set LOG = New cLOG
    LOG.Insertar 34, vUsu, Sql
    Set LOG = Nothing
        
        
    NumRegElim = 0 'para que no de msgerror
    HacerAnyadirAlbaranCosteEuler = True
    conn.CommitTrans
    
eHacerAnyadirAlbaranCosteEuler:
    
    If Err.Number <> 0 Then MuestraError Err.Number, , "Insertando albaran costes" & vbCrLf & Err.Description & vbCrLf & Sql
    If NumRegElim > 0 Then conn.RollbackTrans
    
    Set miRsAux = Nothing
End Function
