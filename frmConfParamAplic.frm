VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmConfParamAplic 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Par�metros"
   ClientHeight    =   9180
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13290
   Icon            =   "frmConfParamAplic.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9180
   ScaleWidth      =   13290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   135
      TabIndex        =   284
      Top             =   135
      Width           =   1050
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   240
         TabIndex        =   285
         Top             =   180
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Modificar"
               Object.Tag             =   "0"
            EndProperty
         EndProperty
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
      Left            =   12060
      TabIndex        =   108
      Top             =   8505
      Width           =   1065
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Left            =   120
      TabIndex        =   110
      Top             =   8430
      Width           =   3000
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
         Height          =   255
         Left            =   120
         TabIndex        =   111
         Top             =   210
         Width           =   2280
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
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
      Left            =   10860
      TabIndex        =   107
      Top             =   8505
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
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
      Left            =   12060
      TabIndex        =   109
      Top             =   8505
      Width           =   1065
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   3240
      Top             =   7920
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      Height          =   7275
      Left            =   135
      TabIndex        =   112
      Top             =   960
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   12832
      _Version        =   393216
      Tabs            =   7
      TabsPerRow      =   7
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Datos Varios 1"
      TabPicture(0)   =   "frmConfParamAplic.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(6)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(5)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(4)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(1)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(14)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1(59)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1(63)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "imgayuda(2)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "imgBuscar(1)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Text1(0)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "FrameDiasMante"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "FrameOpciones"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Text1(5)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Text1(4)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "FramePrecioKm"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Text1(1)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "cboTipodtos"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "cboOrdenDtos"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Text2(1)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "cboCreaTarifa"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Frame13"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "cboDpto"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).ControlCount=   23
      TabCaption(1)   =   "Facturaci�n"
      TabPicture(1)   =   "frmConfParamAplic.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1(86)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "FrameTelefoniaArtic"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Frame9"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Frame10"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Frame12"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Frame15"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Text1(76)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).ControlCount=   8
      TabCaption(2)   =   "Compras / Copias"
      TabPicture(2)   =   "frmConfParamAplic.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label1(67)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "imgBuscar(65)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "imgayuda(6)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label1(75)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Label1(87)"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Label1(88)"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Label1(89)"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "imgBuscar(79)"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Frame3"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "Frame11"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "FrameNumCopias"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "Frame14"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "Frame4"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "Frame16"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "Text2(65)"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "Text1(65)"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).Control(16)=   "Text1(72)"
      Tab(2).Control(16).Enabled=   0   'False
      Tab(2).Control(17)=   "FrameSepOfertas"
      Tab(2).Control(17).Enabled=   0   'False
      Tab(2).Control(18)=   "Text1(77)"
      Tab(2).Control(18).Enabled=   0   'False
      Tab(2).Control(19)=   "Text1(78)"
      Tab(2).Control(19).Enabled=   0   'False
      Tab(2).Control(20)=   "Text1(79)"
      Tab(2).Control(20).Enabled=   0   'False
      Tab(2).Control(21)=   "Text2(79)"
      Tab(2).Control(21).Enabled=   0   'False
      Tab(2).ControlCount=   22
      TabCaption(3)   =   "Contabilidad "
      TabPicture(3)   =   "frmConfParamAplic.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label1(15)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Label1(17)"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Label1(18)"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Label1(19)"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "imgBuscar(39)"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "Label1(48)"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "imgBuscar(40)"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "Label2(7)"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "Label2(6)"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).Control(9)=   "Label1(49)"
      Tab(3).Control(9).Enabled=   0   'False
      Tab(3).Control(10)=   "Label1(50)"
      Tab(3).Control(10).Enabled=   0   'False
      Tab(3).Control(11)=   "imgBuscar(41)"
      Tab(3).Control(11).Enabled=   0   'False
      Tab(3).Control(12)=   "Label1(52)"
      Tab(3).Control(12).Enabled=   0   'False
      Tab(3).Control(13)=   "Label1(53)"
      Tab(3).Control(13).Enabled=   0   'False
      Tab(3).Control(14)=   "imgBuscar(45)"
      Tab(3).Control(14).Enabled=   0   'False
      Tab(3).Control(15)=   "Label1(47)"
      Tab(3).Control(15).Enabled=   0   'False
      Tab(3).Control(16)=   "Label1(51)"
      Tab(3).Control(16).Enabled=   0   'False
      Tab(3).Control(17)=   "Label1(58)"
      Tab(3).Control(17).Enabled=   0   'False
      Tab(3).Control(18)=   "Label2(4)"
      Tab(3).Control(18).Enabled=   0   'False
      Tab(3).Control(19)=   "imgBuscar(67)"
      Tab(3).Control(19).Enabled=   0   'False
      Tab(3).Control(20)=   "Label1(72)"
      Tab(3).Control(20).Enabled=   0   'False
      Tab(3).Control(21)=   "Label1(73)"
      Tab(3).Control(21).Enabled=   0   'False
      Tab(3).Control(22)=   "imgBuscar(66)"
      Tab(3).Control(22).Enabled=   0   'False
      Tab(3).Control(23)=   "Label1(74)"
      Tab(3).Control(23).Enabled=   0   'False
      Tab(3).Control(24)=   "imgayuda(8)"
      Tab(3).Control(24).Enabled=   0   'False
      Tab(3).Control(25)=   "imgayuda(9)"
      Tab(3).Control(25).Enabled=   0   'False
      Tab(3).Control(26)=   "Text1(20)"
      Tab(3).Control(26).Enabled=   0   'False
      Tab(3).Control(27)=   "Text1(21)"
      Tab(3).Control(27).Enabled=   0   'False
      Tab(3).Control(28)=   "Text1(22)"
      Tab(3).Control(28).Enabled=   0   'False
      Tab(3).Control(29)=   "Text1(23)"
      Tab(3).Control(29).Enabled=   0   'False
      Tab(3).Control(30)=   "Text2(46)"
      Tab(3).Control(30).Enabled=   0   'False
      Tab(3).Control(31)=   "Text1(46)"
      Tab(3).Control(31).Enabled=   0   'False
      Tab(3).Control(32)=   "Text1(47)"
      Tab(3).Control(32).Enabled=   0   'False
      Tab(3).Control(33)=   "Text2(47)"
      Tab(3).Control(33).Enabled=   0   'False
      Tab(3).Control(34)=   "Text1(49)"
      Tab(3).Control(34).Enabled=   0   'False
      Tab(3).Control(35)=   "Text1(48)"
      Tab(3).Control(35).Enabled=   0   'False
      Tab(3).Control(36)=   "Text2(48)"
      Tab(3).Control(36).Enabled=   0   'False
      Tab(3).Control(37)=   "cboObsFactura"
      Tab(3).Control(37).Enabled=   0   'False
      Tab(3).Control(38)=   "Text2(52)"
      Tab(3).Control(38).Enabled=   0   'False
      Tab(3).Control(39)=   "Text1(52)"
      Tab(3).Control(39).Enabled=   0   'False
      Tab(3).Control(40)=   "Text1(50)"
      Tab(3).Control(40).Enabled=   0   'False
      Tab(3).Control(41)=   "Frame8"
      Tab(3).Control(41).Enabled=   0   'False
      Tab(3).Control(42)=   "Text2(70)"
      Tab(3).Control(42).Enabled=   0   'False
      Tab(3).Control(43)=   "Text1(70)"
      Tab(3).Control(43).Enabled=   0   'False
      Tab(3).Control(44)=   "chkContabIntraCom"
      Tab(3).Control(44).Enabled=   0   'False
      Tab(3).Control(45)=   "Text2(71)"
      Tab(3).Control(45).Enabled=   0   'False
      Tab(3).Control(46)=   "Text1(71)"
      Tab(3).Control(46).Enabled=   0   'False
      Tab(3).Control(47)=   "CboModAnalitica"
      Tab(3).Control(47).Enabled=   0   'False
      Tab(3).ControlCount=   48
      TabCaption(4)   =   "Internet"
      TabPicture(4)   =   "frmConfParamAplic.frx":007C
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "FrameEMail"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "FrameSoporte"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).ControlCount=   2
      TabCaption(5)   =   "Datos Varios 2"
      TabPicture(5)   =   "frmConfParamAplic.frx":0098
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame5"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "Frame7"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).ControlCount=   2
      TabCaption(6)   =   "Disponible"
      TabPicture(6)   =   "frmConfParamAplic.frx":00B4
      Tab(6).ControlEnabled=   0   'False
      Tab(6).ControlCount=   0
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
         Index           =   79
         Left            =   -69720
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   282
         Text            =   "Text2"
         Top             =   4200
         Width           =   3105
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
         Index           =   79
         Left            =   -70560
         MaxLength       =   4
         TabIndex        =   262
         Tag             =   "Situacion riesgo sin bloqueo|N|S|0|9999|spara1|situ_riesgo_Nobloq|||"
         Text            =   "Text1"
         Top             =   4200
         Width           =   855
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
         Index           =   78
         Left            =   -72705
         MaxLength       =   255
         TabIndex        =   266
         Tag             =   "F.e.|T|S|||spara1|PathFirmasAlbaran|||"
         Text            =   "Text1 "
         Top             =   5760
         Width           =   6135
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
         Index           =   77
         Left            =   -72705
         MaxLength       =   255
         TabIndex        =   265
         Tag             =   "F.e.|T|S|||spara1|PathFirmasFacturas|||"
         Text            =   "Text1 "
         Top             =   5280
         Width           =   6135
      End
      Begin VB.ComboBox CboModAnalitica 
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
         Left            =   -64755
         Style           =   2  'Dropdown List
         TabIndex        =   279
         Tag             =   "Modo anal�tica|N|N|0|9|spara1|modanalitica|||"
         Top             =   1920
         Width           =   1935
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
         Index           =   76
         Left            =   -72825
         MaxLength       =   16
         TabIndex        =   277
         Tag             =   "R|N|S||1000|spara1|preciohoracoste|0.0000||"
         Text            =   "Text1 "
         Top             =   6465
         Width           =   1455
      End
      Begin VB.Frame FrameSepOfertas 
         Caption         =   "Separador ofertas"
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
         Height          =   735
         Left            =   -74670
         TabIndex        =   268
         Top             =   6240
         Width           =   11265
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
            Index           =   73
            Left            =   1950
            MaxLength       =   16
            TabIndex        =   270
            Tag             =   "Reci. |T|S|||spara1|artSeparador|||"
            Text            =   "Text1 "
            Top             =   297
            Width           =   2070
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
            Index           =   73
            Left            =   4065
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   269
            Text            =   "Text2"
            Top             =   300
            Width           =   6555
         End
         Begin VB.Label Label1 
            Caption         =   "Art�culo "
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   83
            Left            =   240
            TabIndex        =   271
            Top             =   315
            Width           =   780
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   73
            Left            =   1590
            Tag             =   "-1"
            ToolTipText     =   "Buscar art�culo"
            Top             =   330
            Width           =   240
         End
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
         Index           =   72
         Left            =   -72705
         MaxLength       =   255
         TabIndex        =   264
         Tag             =   "F.e.|T|S|||spara1|PathFacturaE|||"
         Text            =   "Text1 "
         Top             =   4800
         Width           =   6135
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
         Index           =   65
         Left            =   -74640
         MaxLength       =   4
         TabIndex        =   261
         Tag             =   "Situacion bloqueo|N|S|0|9999|spara1|situbloq|||"
         Text            =   "Text1"
         Top             =   4200
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
         Index           =   65
         Left            =   -73755
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   260
         Text            =   "Text2"
         Top             =   4200
         Width           =   3105
      End
      Begin VB.Frame Frame16 
         Caption         =   "Pagos comisiones agentes"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   735
         Left            =   -66210
         TabIndex        =   257
         Top             =   4680
         Width           =   3915
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
            Index           =   74
            Left            =   2655
            TabIndex        =   258
            Tag             =   "%pago comisiones|N|S|0|100|spara1|PorcenpagAgenPag|||"
            Text            =   "Text1"
            Top             =   270
            Width           =   795
         End
         Begin VB.Image imgayuda 
            Height          =   255
            Index           =   10
            Left            =   2835
            ToolTipText     =   "Pago comisiones"
            Top             =   0
            Width           =   255
         End
         Begin VB.Label Label1 
            Caption         =   "% dias para impagados"
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
            Index           =   84
            Left            =   240
            TabIndex        =   259
            Top             =   300
            Width           =   2475
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Cheques  regalo"
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
         Height          =   735
         Left            =   -74640
         TabIndex        =   253
         Top             =   3000
         Width           =   8085
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
            Left            =   1770
            MaxLength       =   3
            TabIndex        =   255
            Tag             =   "Forma de pago para cheque regalo |N|S|0|999|spara1|codforpa|000||"
            Text            =   "Tex"
            Top             =   237
            Width           =   780
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
            Left            =   2610
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   254
            Text            =   "Text2"
            Top             =   240
            Width           =   4290
         End
         Begin VB.Label Label1 
            Caption         =   "Forma pago "
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
            Left            =   120
            TabIndex        =   256
            Top             =   270
            Width           =   1200
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   24
            Left            =   1410
            Tag             =   "-1"
            ToolTipText     =   "Buscar forma pago"
            Top             =   270
            Width           =   240
         End
      End
      Begin VB.Frame Frame14 
         Height          =   735
         Left            =   -66180
         TabIndex        =   251
         Top             =   480
         Width           =   2880
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
            Height          =   320
            Index           =   63
            Left            =   2025
            MaxLength       =   2
            TabIndex        =   47
            Tag             =   "Copias facturacion|N|S|1||spara1|numcopias|||"
            Text            =   "Text1"
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Copias facturaci�n"
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
            Index           =   64
            Left            =   120
            TabIndex        =   252
            Top             =   270
            Width           =   1920
         End
      End
      Begin VB.Frame FrameNumCopias 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   3255
         Left            =   -66180
         TabIndex        =   243
         Top             =   1320
         Width           =   2880
         Begin VB.TextBox txtNumCopias 
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
            Left            =   2130
            TabIndex        =   58
            Text            =   "Text3"
            Top             =   2790
            Width           =   510
         End
         Begin VB.TextBox txtNumCopias 
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
            Left            =   2130
            TabIndex        =   57
            Text            =   "Text3"
            Top             =   2400
            Width           =   510
         End
         Begin VB.TextBox txtNumCopias 
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
            Left            =   2130
            TabIndex        =   56
            Text            =   "Text3"
            Top             =   1968
            Width           =   510
         End
         Begin VB.TextBox txtNumCopias 
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
            Left            =   2130
            TabIndex        =   55
            Text            =   "Text3"
            Top             =   1536
            Width           =   510
         End
         Begin VB.TextBox txtNumCopias 
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
            Left            =   2130
            TabIndex        =   54
            Text            =   "Text3"
            Top             =   1104
            Width           =   510
         End
         Begin VB.TextBox txtNumCopias 
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
            Left            =   2130
            TabIndex        =   53
            Text            =   "Text3"
            Top             =   672
            Width           =   510
         End
         Begin VB.TextBox txtNumCopias 
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
            Left            =   2130
            TabIndex        =   52
            Text            =   "Text3"
            Top             =   240
            Width           =   510
         End
         Begin VB.Label Label1 
            Caption         =   "Fras.mostrador"
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
            Index           =   90
            Left            =   240
            TabIndex        =   286
            Top             =   2790
            Width           =   1890
         End
         Begin VB.Label Label1 
            Caption         =   "N� copias"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004000&
            Height          =   255
            Index           =   82
            Left            =   165
            TabIndex        =   250
            Top             =   0
            Width           =   1035
         End
         Begin VB.Label Label1 
            Caption         =   "Fras.Rectificativas"
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
            Index           =   81
            Left            =   240
            TabIndex        =   249
            Top             =   2400
            Width           =   1890
         End
         Begin VB.Label Label1 
            Caption         =   "Albar�n ruta"
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
            Index           =   80
            Left            =   240
            TabIndex        =   248
            Top             =   1965
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Albar�n venta"
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
            Index           =   79
            Left            =   240
            TabIndex        =   247
            Top             =   1530
            Width           =   1545
         End
         Begin VB.Label Label1 
            Caption         =   "Pedido x zonas"
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
            Index           =   78
            Left            =   240
            TabIndex        =   246
            Top             =   1110
            Width           =   1650
         End
         Begin VB.Label Label1 
            Caption         =   "Pedidos"
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
            Index           =   77
            Left            =   240
            TabIndex        =   245
            Top             =   672
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Ofertas"
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
            Index           =   76
            Left            =   240
            TabIndex        =   244
            Top             =   240
            Width           =   855
         End
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
         Index           =   71
         Left            =   -67650
         MaxLength       =   2
         TabIndex        =   71
         Tag             =   "IVAexento|N|S|0||spara1|IvaIntracomAdicional|||"
         Text            =   "Text1"
         Top             =   2760
         Width           =   615
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
         Index           =   71
         Left            =   -66975
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   241
         Text            =   "Text2"
         Top             =   2760
         Width           =   2745
      End
      Begin VB.CheckBox chkContabIntraCom 
         Caption         =   "Permitir contabilizacion"
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
         Left            =   -67440
         TabIndex        =   69
         Tag             =   "Mantenimientos|N|N|||spara1|IntracomAdicionContab|||"
         Top             =   1920
         Width           =   2625
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
         Index           =   70
         Left            =   -73155
         MaxLength       =   10
         TabIndex        =   68
         Tag             =   "Cta intracom|N|S|||spara1|CtaContabIntracom|||"
         Text            =   "3"
         Top             =   1920
         Width           =   1340
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
         Index           =   70
         Left            =   -71745
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   237
         Text            =   "Text2"
         Top             =   1920
         Width           =   3030
      End
      Begin VB.Frame Frame11 
         Caption         =   "Rotaci�n"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   1575
         Left            =   -74640
         TabIndex        =   232
         Top             =   1440
         Width           =   7095
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
            Index           =   69
            Left            =   5925
            TabIndex        =   51
            Tag             =   "Rotacion. Maximo|N|S|0||spara1|consummax|0.00||"
            Text            =   "Text1"
            Top             =   960
            Width           =   705
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
            Index           =   68
            Left            =   2310
            TabIndex        =   50
            Tag             =   "Rotacion. Minimo|N|S|0||spara1|consummin|0.00||"
            Text            =   "Text1"
            Top             =   960
            Width           =   750
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
            Index           =   67
            Left            =   5925
            TabIndex        =   49
            Tag             =   "Rotacion. Mes2|N|S|0|31|spara1|mesconant2|||"
            Text            =   "Text1"
            Top             =   327
            Width           =   705
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
            Index           =   66
            Left            =   2310
            TabIndex        =   48
            Tag             =   "Rotacion. Mes1|N|S|0||spara1|mesconant1|||"
            Text            =   "Text1"
            Top             =   327
            Width           =   750
         End
         Begin VB.Image imgayuda 
            Height          =   255
            Index           =   7
            Left            =   1080
            ToolTipText     =   "Rotaci�n"
            Top             =   0
            Width           =   255
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Meses aprov. m�ximo"
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
            Index           =   71
            Left            =   3720
            TabIndex        =   236
            Top             =   990
            Width           =   2115
         End
         Begin VB.Label Label1 
            Caption         =   "Meses aprov. minimo"
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
            Index           =   70
            Left            =   150
            TabIndex        =   235
            Top             =   990
            Width           =   2145
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Meses de consumo 2"
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
            Index           =   69
            Left            =   3720
            TabIndex        =   234
            Top             =   360
            Width           =   2070
         End
         Begin VB.Label Label1 
            Caption         =   "Meses de consumo 1"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   68
            Left            =   150
            TabIndex        =   233
            Top             =   360
            Width           =   2100
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Compras"
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
         Height          =   855
         Left            =   -74640
         TabIndex        =   229
         Top             =   480
         Width           =   7095
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
            Index           =   16
            Left            =   1470
            MaxLength       =   2
            TabIndex        =   43
            Tag             =   "Dia 1 de pago compras|N|S|0|31|spara1|diapago1|||"
            Text            =   "Text1"
            Top             =   360
            Width           =   660
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
            Index           =   17
            Left            =   2190
            MaxLength       =   2
            TabIndex        =   44
            Tag             =   "Dia 2 de pago compras|N|S|0|31|spara1|diapago2|||"
            Text            =   "Text1"
            Top             =   360
            Width           =   660
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
            Index           =   18
            Left            =   2910
            MaxLength       =   2
            TabIndex        =   45
            Tag             =   "Dia 3 de pago compras|N|S|0|31|spara1|diapago3|||"
            Text            =   "Text1"
            Top             =   360
            Width           =   660
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
            Index           =   19
            Left            =   5295
            MaxLength       =   2
            TabIndex        =   46
            Tag             =   "Mes a no girar|N|S|0|12|spara1|mesnogir|||"
            Text            =   "Text1"
            Top             =   360
            Width           =   795
         End
         Begin VB.Label Label1 
            Caption         =   "D�as de pago"
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
            Left            =   120
            TabIndex        =   231
            Top             =   360
            Width           =   1320
         End
         Begin VB.Label Label1 
            Caption         =   "Mes a no girar"
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
            Index           =   13
            Left            =   3720
            TabIndex        =   230
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.Frame FrameSoporte 
         ForeColor       =   &H00972E0B&
         Height          =   1635
         Left            =   -73695
         TabIndex        =   224
         Top             =   4800
         Width           =   11100
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
            Index           =   14
            Left            =   2220
            MaxLength       =   100
            TabIndex        =   92
            Tag             =   "Version Web|T|S|||spara1|webversion|||"
            Text            =   "3"
            Top             =   1080
            Width           =   7770
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
            Index           =   13
            Left            =   2220
            MaxLength       =   100
            TabIndex        =   91
            Tag             =   "Mail de Soporte|T|S|||spara1|mailsoporte|||"
            Text            =   "3"
            Top             =   690
            Width           =   7770
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
            Index           =   12
            Left            =   2220
            MaxLength       =   100
            TabIndex        =   90
            Tag             =   "Web de Soporte|T|S|||spara1|websoporte|||"
            Text            =   "3"
            Top             =   300
            Width           =   7770
         End
         Begin VB.Label Label8 
            Caption         =   "Soporte"
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
            Height          =   285
            Index           =   1
            Left            =   360
            TabIndex        =   228
            Top             =   0
            Width           =   840
         End
         Begin VB.Label Label1 
            Caption         =   "Web check version"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   16
            Left            =   300
            TabIndex        =   227
            Top             =   1140
            Width           =   1905
         End
         Begin VB.Label Label1 
            Caption         =   "Mail soporte"
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
            Index           =   12
            Left            =   300
            TabIndex        =   226
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Web de soporte"
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
            Index           =   9
            Left            =   300
            TabIndex        =   225
            Top             =   360
            Width           =   1590
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Avisos"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   2415
         Left            =   -74160
         TabIndex        =   212
         Top             =   3840
         Width           =   11595
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
            Index           =   39
            Left            =   4050
            MaxLength       =   2
            TabIndex        =   105
            Tag             =   "avi.repa.|N|S|0||spara1|avirepara|||"
            Text            =   "3"
            Top             =   1635
            Width           =   825
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
            Index           =   38
            Left            =   4050
            MaxLength       =   2
            TabIndex        =   106
            Tag             =   "avi.avisos|N|S|0||spara1|aviavios|||"
            Text            =   "3"
            Top             =   1995
            Width           =   825
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
            Index           =   37
            Left            =   4050
            MaxLength       =   2
            TabIndex        =   104
            Tag             =   "avi.mante|N|S|0||spara1|avimanteni|||"
            Text            =   "3"
            Top             =   1275
            Width           =   825
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
            Index           =   36
            Left            =   7215
            MaxLength       =   2
            TabIndex        =   103
            Tag             =   "alb.pro.|N|S|0||spara1|avialbpro|||"
            Text            =   "3"
            Top             =   720
            Width           =   780
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
            Index           =   35
            Left            =   4050
            MaxLength       =   2
            TabIndex        =   102
            Tag             =   "alb.cli.|N|S|0||spara1|avialbcli|||"
            Text            =   "3"
            Top             =   720
            Width           =   825
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
            Index           =   34
            Left            =   7215
            MaxLength       =   2
            TabIndex        =   101
            Tag             =   "ped.pro.|N|S|0||spara1|avipedpro|||"
            Text            =   "3"
            Top             =   315
            Width           =   780
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
            Index           =   33
            Left            =   4050
            MaxLength       =   2
            TabIndex        =   100
            Tag             =   "ped. cli|N|S|0||spara1|avipedcli|||"
            Text            =   "3"
            Top             =   315
            Width           =   825
         End
         Begin VB.Label Label1 
            Caption         =   "Abiertos"
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
            Index           =   43
            Left            =   4950
            TabIndex        =   223
            Top             =   2040
            Width           =   2955
         End
         Begin VB.Label Label1 
            Caption         =   "Pendientes de reparar sin motivo de reparaci�n"
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
            Index           =   42
            Left            =   4950
            TabIndex        =   222
            Top             =   1680
            Width           =   3555
         End
         Begin VB.Label Label1 
            Caption         =   "No facturados"
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
            Index           =   41
            Left            =   4950
            TabIndex        =   221
            Top             =   1320
            Width           =   2955
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Avisos "
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
            Index           =   39
            Left            =   2175
            TabIndex        =   219
            Top             =   2040
            Width           =   705
         End
         Begin VB.Label Label1 
            Caption         =   "Reparaciones"
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
            Index           =   38
            Left            =   2175
            TabIndex        =   218
            Top             =   1680
            Width           =   1425
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Mantenimientos"
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
            Index           =   37
            Left            =   2175
            TabIndex        =   217
            Top             =   1320
            Width           =   1545
         End
         Begin VB.Label Label1 
            Caption         =   "Albaranes proveedores"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   36
            Left            =   4950
            TabIndex        =   216
            Top             =   765
            Width           =   2265
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Albaranes clientes"
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
            Index           =   35
            Left            =   2175
            TabIndex        =   215
            Top             =   765
            Width           =   1800
         End
         Begin VB.Label Label1 
            Caption         =   "Pedidos proveedores"
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
            Index           =   34
            Left            =   4950
            TabIndex        =   214
            Top             =   360
            Width           =   2130
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Pedidos clientes"
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
            Index           =   33
            Left            =   2175
            TabIndex        =   213
            Top             =   360
            Width           =   1590
         End
         Begin VB.Label Label1 
            Caption         =   "Dias desde la fecha"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   40
            Left            =   120
            TabIndex        =   220
            Top             =   360
            Width           =   7275
         End
      End
      Begin VB.Frame FrameEMail 
         Height          =   2895
         Left            =   -73680
         TabIndex        =   205
         Top             =   1080
         Width           =   11055
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
            IMEMode         =   3  'DISABLE
            Index           =   57
            Left            =   3030
            MaxLength       =   30
            TabIndex        =   89
            Tag             =   "LanzaMailOutlook|T|S|||spara1|arigesmail|||"
            Text            =   "3"
            Top             =   2400
            Width           =   1665
         End
         Begin VB.CheckBox chkOutlook 
            Caption         =   "Enviar desde Outlook"
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
            Left            =   5880
            TabIndex        =   88
            Tag             =   "Outlook|N|N|||spara1|EnvioDesdeOutlook|||"
            Top             =   1560
            Width           =   2490
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
            IMEMode         =   3  'DISABLE
            Index           =   11
            Left            =   1440
            MaxLength       =   30
            PasswordChar    =   "*"
            TabIndex        =   87
            Tag             =   "Password SMTP|T|S|||spara1|smtppass|||"
            Text            =   "3"
            Top             =   1560
            Width           =   4305
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
            Left            =   1440
            MaxLength       =   50
            TabIndex        =   86
            Tag             =   "Usuario SMTP|T|S|||spara1|smtpuser|||"
            Text            =   "3"
            Top             =   1180
            Width           =   4305
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
            Left            =   1440
            MaxLength       =   50
            TabIndex        =   85
            Tag             =   "Servidor SMTP|T|S|||spara1|smtphost|||"
            Text            =   "3"
            Top             =   800
            Width           =   8445
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
            Left            =   1440
            MaxLength       =   50
            TabIndex        =   84
            Tag             =   "Direccion e-mail|T|S|||spara1|diremail|||"
            Text            =   "3"
            Top             =   420
            Width           =   8445
         End
         Begin VB.Line Line1 
            X1              =   240
            X2              =   8040
            Y1              =   2160
            Y2              =   2160
         End
         Begin VB.Label Label1 
            Caption         =   "Lanza pantalla mail outlook"
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
            Index           =   60
            Left            =   240
            TabIndex        =   211
            Top             =   2460
            Width           =   2730
         End
         Begin VB.Label Label8 
            Caption         =   "Envio E-Mail"
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
            Height          =   195
            Index           =   0
            Left            =   360
            TabIndex        =   210
            Top             =   0
            Width           =   1320
         End
         Begin VB.Label Label1 
            Caption         =   "Password"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   23
            Left            =   300
            TabIndex        =   209
            Top             =   1620
            Width           =   1065
         End
         Begin VB.Label Label1 
            Caption         =   "Usuario"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   22
            Left            =   300
            TabIndex        =   208
            Top             =   1260
            Width           =   780
         End
         Begin VB.Label Label1 
            Caption         =   "Servidor SMTP"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   21
            Left            =   300
            TabIndex        =   207
            Top             =   840
            Width           =   1140
         End
         Begin VB.Label Label1 
            Caption         =   "E-Mail"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   20
            Left            =   300
            TabIndex        =   206
            Top             =   480
            Width           =   900
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Clientes"
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
         Height          =   3120
         Left            =   -74160
         TabIndex        =   190
         Top             =   600
         Width           =   11595
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
            Index           =   25
            Left            =   960
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   197
            Text            =   "Text2"
            Top             =   480
            Width           =   4770
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
            Left            =   120
            MaxLength       =   3
            TabIndex        =   93
            Tag             =   "Actividad|N|S|0||spara1|defactividad|000||"
            Text            =   "Tex"
            Top             =   480
            Width           =   780
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
            Left            =   6720
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   196
            Text            =   "Text2"
            Top             =   480
            Width           =   4635
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
            Left            =   5880
            MaxLength       =   3
            TabIndex        =   94
            Tag             =   "Envio|N|S|0|999|spara1|defenvio|000||"
            Text            =   "Tex"
            Top             =   480
            Width           =   735
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
            Index           =   27
            Left            =   960
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   195
            Text            =   "Text2"
            Top             =   1080
            Width           =   4770
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
            Left            =   120
            MaxLength       =   3
            TabIndex        =   95
            Tag             =   "Zona|N|S|0|999|spara1|defzona|000||"
            Text            =   "Tex"
            Top             =   1080
            Width           =   735
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
            Index           =   28
            Left            =   6720
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   194
            Text            =   "Text2"
            Top             =   1080
            Width           =   4635
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
            Index           =   28
            Left            =   5880
            MaxLength       =   3
            TabIndex        =   96
            Tag             =   "Ruta|N|S|0|999|spara1|defruta|000||"
            Text            =   "Tex"
            Top             =   1080
            Width           =   735
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
            Index           =   29
            Left            =   960
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   193
            Text            =   "Text2"
            Top             =   1800
            Width           =   4770
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
            Index           =   29
            Left            =   120
            MaxLength       =   3
            TabIndex        =   97
            Tag             =   "Situacion|N|S|0|999|spara1|defstituacion|000||"
            Text            =   "Tex"
            Top             =   1800
            Width           =   780
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
            Index           =   30
            Left            =   6720
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   192
            Text            =   "Text2"
            Top             =   1800
            Width           =   4635
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
            Index           =   30
            Left            =   5880
            MaxLength       =   3
            TabIndex        =   98
            Tag             =   "Tarifa|N|S|0|999|spara1|deftarifa|000||"
            Text            =   "Tex"
            Top             =   1800
            Width           =   735
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
            Index           =   31
            Left            =   960
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   191
            Text            =   "Text2"
            Top             =   2520
            Width           =   4770
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
            Index           =   31
            Left            =   120
            MaxLength       =   3
            TabIndex        =   99
            Tag             =   "Agente|N|S|0|999|spara1|defagente|000||"
            Text            =   "Tex"
            Top             =   2520
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Actividad"
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
            Left            =   120
            TabIndex        =   204
            Top             =   240
            Width           =   1005
         End
         Begin VB.Label Label1 
            Caption         =   "Envio"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   26
            Left            =   5880
            TabIndex        =   203
            Top             =   240
            Width           =   630
         End
         Begin VB.Label Label1 
            Caption         =   "Zona"
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
            Left            =   120
            TabIndex        =   202
            Top             =   840
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "Ruta"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   28
            Left            =   5880
            TabIndex        =   201
            Top             =   840
            Width           =   525
         End
         Begin VB.Label Label1 
            Caption         =   "Situaci�n"
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
            Index           =   29
            Left            =   120
            TabIndex        =   200
            Top             =   1560
            Width           =   945
         End
         Begin VB.Label Label1 
            Caption         =   "Tarifa"
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
            Index           =   30
            Left            =   5880
            TabIndex        =   199
            Top             =   1560
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Agente"
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
            Index           =   31
            Left            =   120
            TabIndex        =   198
            Top             =   2280
            Width           =   735
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   26
            Left            =   6540
            ToolTipText     =   "Buscar forma de envio"
            Top             =   240
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   27
            Left            =   660
            ToolTipText     =   "Buscar zona"
            Top             =   840
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   28
            Left            =   6540
            ToolTipText     =   "Buscar ruta"
            Top             =   840
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   25
            Left            =   1155
            Tag             =   "-1"
            ToolTipText     =   "Buscar actividad"
            Top             =   240
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   29
            Left            =   1155
            ToolTipText     =   "Buscar situacion"
            Top             =   1560
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   30
            Left            =   6540
            ToolTipText     =   "Buscar tarifa"
            Top             =   1560
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   31
            Left            =   930
            ToolTipText     =   "Buscar agente"
            Top             =   2280
            Width           =   240
         End
      End
      Begin VB.Frame Frame15 
         Caption         =   "Recargo financiero"
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
         Height          =   750
         Left            =   -74865
         TabIndex        =   186
         Top             =   2505
         Width           =   12705
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
            Index           =   64
            Left            =   2400
            MaxLength       =   16
            TabIndex        =   40
            Tag             =   "Recar |T|S|||spara1|artRecargoFina|||"
            Text            =   "Text1 "
            Top             =   297
            Width           =   2070
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
            Index           =   64
            Left            =   4515
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   187
            Text            =   "Text2"
            Top             =   300
            Width           =   5700
         End
         Begin VB.Label Label1 
            Caption         =   "Art�culo "
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   65
            Left            =   150
            TabIndex        =   188
            Top             =   315
            Width           =   1515
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   64
            Left            =   2130
            Tag             =   "-1"
            ToolTipText     =   "Buscar art�culo"
            Top             =   330
            Width           =   240
         End
      End
      Begin VB.ComboBox cboDpto 
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
         ItemData        =   "frmConfParamAplic.frx":00D0
         Left            =   2460
         List            =   "frmConfParamAplic.frx":00DD
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Tag             =   "Dep-direc-obras|N|N|||spara1|haydepar|||"
         Top             =   3360
         Width           =   1815
      End
      Begin VB.Frame Frame8 
         Caption         =   "I. V. A."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   2895
         Left            =   -74880
         TabIndex        =   138
         Top             =   3600
         Width           =   10215
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
            Index           =   62
            Left            =   4320
            MaxLength       =   2
            TabIndex        =   83
            Tag             =   "IVA2|N|S|0|99|spara1|iva_oldre2|||"
            Text            =   "Text1"
            Top             =   2400
            Width           =   495
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
            Index           =   61
            Left            =   1200
            MaxLength       =   2
            TabIndex        =   81
            Tag             =   "IVA1|N|S|0|99|spara1|iva_oldre1|||"
            Text            =   "Text1"
            Top             =   2400
            Width           =   495
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
            Index           =   61
            Left            =   1680
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   182
            Text            =   "Text2"
            Top             =   2400
            Width           =   2145
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
            Index           =   62
            Left            =   4800
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   181
            Text            =   "Text2"
            Top             =   2400
            Width           =   2145
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
            Index           =   60
            Left            =   4320
            MaxLength       =   2
            TabIndex        =   82
            Tag             =   "IVA2|N|S|0|99|spara1|iva_old2|||"
            Text            =   "Text1"
            Top             =   1920
            Width           =   495
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
            Index           =   59
            Left            =   1200
            MaxLength       =   2
            TabIndex        =   80
            Tag             =   "IVA1|N|S|0|99|spara1|iva_old1|||"
            Text            =   "Text1"
            Top             =   1920
            Width           =   495
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
            Index           =   59
            Left            =   1680
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   180
            Text            =   "Text2"
            Top             =   1920
            Width           =   2145
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
            Index           =   60
            Left            =   4800
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   179
            Text            =   "Text2"
            Top             =   1920
            Width           =   2145
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
            Index           =   42
            Left            =   7920
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   147
            Text            =   "Text2"
            Top             =   960
            Width           =   2145
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
            Index           =   45
            Left            =   7920
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   146
            Text            =   "Text2"
            Top             =   480
            Width           =   2145
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
            Index           =   41
            Left            =   4800
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   145
            Text            =   "Text2"
            Top             =   960
            Width           =   2145
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
            Index           =   44
            Left            =   4800
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   144
            Text            =   "Text2"
            Top             =   480
            Width           =   2145
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
            Index           =   40
            Left            =   1680
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   143
            Text            =   "Text2"
            Top             =   960
            Width           =   2145
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
            Index           =   43
            Left            =   1680
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   142
            Text            =   "Text2"
            Top             =   480
            Width           =   2145
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
            Index           =   42
            Left            =   7440
            MaxLength       =   2
            TabIndex        =   79
            Tag             =   "IVRE3|N|S|0|99|spara1|ivare3eq|||"
            Text            =   "Text1"
            Top             =   960
            Width           =   495
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
            Index           =   41
            Left            =   4320
            MaxLength       =   2
            TabIndex        =   77
            Tag             =   "IVRE2|N|S|0|99|spara1|ivare2eq|||"
            Text            =   "Text1"
            Top             =   960
            Width           =   495
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
            Index           =   40
            Left            =   1200
            MaxLength       =   2
            TabIndex        =   75
            Tag             =   "IVRE1|N|S|0|99|spara1|ivare1eq|||"
            Text            =   "Text1"
            Top             =   960
            Width           =   495
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
            Index           =   43
            Left            =   1200
            MaxLength       =   2
            TabIndex        =   74
            Tag             =   "IVA1|N|S|0|99|spara1|ivare1|||"
            Text            =   "Text1"
            Top             =   480
            Width           =   495
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
            Index           =   44
            Left            =   4320
            MaxLength       =   2
            TabIndex        =   76
            Tag             =   "IVA2|N|S|0|99|spara1|ivare2|||"
            Text            =   "Text1"
            Top             =   480
            Width           =   495
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
            Index           =   45
            Left            =   7440
            MaxLength       =   2
            TabIndex        =   78
            Tag             =   "IVA3|N|S|0|99|spara1|ivare3|||"
            Text            =   "Text1"
            Top             =   480
            Width           =   495
         End
         Begin VB.Label Label2 
            Caption         =   "Normal"
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
            Left            =   240
            TabIndex        =   184
            Top             =   1920
            Width           =   675
         End
         Begin VB.Label Label2 
            Caption         =   "R.E."
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
            Left            =   240
            TabIndex        =   183
            Top             =   2400
            Width           =   495
         End
         Begin VB.Line Line4 
            BorderColor     =   &H00808080&
            X1              =   7080
            X2              =   7080
            Y1              =   480
            Y2              =   2760
         End
         Begin VB.Line Line3 
            BorderColor     =   &H00808080&
            X1              =   3960
            X2              =   3960
            Y1              =   480
            Y2              =   2760
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   49
            Left            =   4080
            Tag             =   "-1"
            ToolTipText     =   "Buscar I.V.A."
            Top             =   2400
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   48
            Left            =   960
            Tag             =   "-1"
            ToolTipText     =   "Buscar I.V.A."
            Top             =   2400
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   47
            Left            =   4080
            Tag             =   "-1"
            ToolTipText     =   "Buscar I.V.A."
            Top             =   1920
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   46
            Left            =   960
            Tag             =   "-1"
            ToolTipText     =   "Buscar I.V.A."
            Top             =   1920
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "IVA antiguo"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004000&
            Height          =   195
            Index           =   61
            Left            =   120
            TabIndex        =   178
            Top             =   1650
            Width           =   1275
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00808080&
            X1              =   240
            X2              =   10080
            Y1              =   1560
            Y2              =   1560
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   35
            Left            =   7200
            Tag             =   "-1"
            ToolTipText     =   "Buscar I.V.A."
            Top             =   960
            Width           =   240
         End
         Begin VB.Label Label2 
            Caption         =   "R.E."
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
            Left            =   240
            TabIndex        =   149
            Top             =   960
            Width           =   630
         End
         Begin VB.Label Label2 
            Caption         =   "Normal"
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
            Left            =   240
            TabIndex        =   148
            Top             =   480
            Width           =   720
         End
         Begin VB.Label Label1 
            Caption         =   "General"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   45
            Left            =   1080
            TabIndex        =   141
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Reducido"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   44
            Left            =   4200
            TabIndex        =   140
            Top             =   240
            Width           =   960
         End
         Begin VB.Label Label1 
            Caption         =   "Super-Reducido"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   46
            Left            =   7320
            TabIndex        =   139
            Top             =   240
            Width           =   1605
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   36
            Left            =   960
            Tag             =   "-1"
            ToolTipText     =   "Buscar I.V.A."
            Top             =   480
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   33
            Left            =   960
            Tag             =   "-1"
            ToolTipText     =   "Buscar I.V.A."
            Top             =   960
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   37
            Left            =   4080
            Tag             =   "-1"
            ToolTipText     =   "Buscar I.V.A."
            Top             =   480
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   34
            Left            =   4080
            Tag             =   "-1"
            ToolTipText     =   "Buscar I.V.A."
            Top             =   960
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   38
            Left            =   7200
            Tag             =   "-1"
            ToolTipText     =   "Buscar I.V.A."
            Top             =   480
            Width           =   240
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   "Garantia de reparaci�n"
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
         Height          =   735
         Left            =   240
         TabIndex        =   176
         Top             =   6120
         Width           =   4320
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
            Height          =   320
            Index           =   58
            Left            =   2280
            MaxLength       =   4
            TabIndex        =   11
            Tag             =   "Dias de garantia de Reparacion|N|S|0|9999|spara1|diasgaranrepa|||"
            Text            =   "Text1"
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Dias"
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
            Index           =   62
            Left            =   240
            TabIndex        =   177
            Top             =   360
            Width           =   1920
         End
      End
      Begin VB.ComboBox cboCreaTarifa 
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
         ItemData        =   "frmConfParamAplic.frx":0104
         Left            =   2460
         List            =   "frmConfParamAplic.frx":0111
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Tag             =   "Tipo Descuentos|N|N|||spara1|creatarifart|||"
         Top             =   2940
         Width           =   1815
      End
      Begin VB.Frame Frame12 
         Caption         =   "Portes"
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
         Height          =   1155
         Left            =   -74880
         TabIndex        =   169
         Top             =   450
         Width           =   12735
         Begin VB.ComboBox cboPortes 
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
            ItemData        =   "frmConfParamAplic.frx":0130
            Left            =   1320
            List            =   "frmConfParamAplic.frx":013D
            Style           =   2  'Dropdown List
            TabIndex        =   36
            Tag             =   "Portes|N|S|||spara1|tipoportes|||"
            Top             =   690
            Width           =   2295
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
            Index           =   56
            Left            =   6000
            MaxLength       =   16
            TabIndex        =   37
            Tag             =   "R|N|S||10000|spara1|impminped|#,##0.00||"
            Text            =   "Text1 "
            Top             =   630
            Width           =   1500
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
            Index           =   55
            Left            =   8865
            MaxLength       =   16
            TabIndex        =   38
            Tag             =   "i |N|S|||spara1|abonokilos|#,##0.0000||"
            Text            =   "Text1 "
            Top             =   630
            Width           =   1500
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
            Index           =   54
            Left            =   3420
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   171
            Text            =   "Text2"
            Top             =   240
            Width           =   5190
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
            Index           =   54
            Left            =   1320
            MaxLength       =   16
            TabIndex        =   35
            Tag             =   "Reci. |T|S|||spara1|ArticuloPortes|||"
            Text            =   "Text1 "
            Top             =   240
            Width           =   2070
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo portes"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   66
            Left            =   120
            TabIndex        =   189
            Top             =   690
            Width           =   1140
         End
         Begin VB.Image imgayuda 
            Height          =   255
            Index           =   0
            Left            =   12165
            ToolTipText     =   "Portes"
            Top             =   240
            Width           =   285
         End
         Begin VB.Label Label1 
            Caption         =   "Importe minimo"
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
            Index           =   57
            Left            =   4440
            TabIndex        =   173
            Top             =   690
            Width           =   1665
         End
         Begin VB.Label Label1 
            Caption         =   "Abono kilos"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   56
            Left            =   7680
            TabIndex        =   172
            Top             =   690
            Width           =   1140
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   54
            Left            =   960
            Tag             =   "-1"
            ToolTipText     =   "Buscar art�culo"
            Top             =   285
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Articulo"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   55
            Left            =   120
            TabIndex        =   170
            Top             =   300
            Width           =   780
         End
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
         Index           =   50
         Left            =   -66120
         MaxLength       =   2
         TabIndex        =   64
         Tag             =   "N�Conta|N|S|1|99|spara1|conta_B|||"
         Text            =   "Text1"
         Top             =   600
         Width           =   495
      End
      Begin VB.Frame Frame10 
         Caption         =   "Reciclado / Punto verde"
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
         Height          =   735
         Left            =   -74865
         TabIndex        =   165
         Top             =   4185
         Width           =   12705
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
            Index           =   53
            Left            =   2400
            MaxLength       =   16
            TabIndex        =   42
            Tag             =   "Reci. |T|S|||spara1|ArtReciclado|||"
            Text            =   "Text1 "
            Top             =   297
            Width           =   2070
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
            Index           =   53
            Left            =   4515
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   166
            Text            =   "Text2"
            Top             =   300
            Width           =   5700
         End
         Begin VB.Label Label1 
            Caption         =   "Art�culo "
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   54
            Left            =   150
            TabIndex        =   167
            Top             =   315
            Width           =   780
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   53
            Left            =   2130
            Tag             =   "-1"
            ToolTipText     =   "Buscar art�culo"
            Top             =   330
            Width           =   240
         End
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
         Index           =   52
         Left            =   -72855
         MaxLength       =   2
         TabIndex        =   70
         Tag             =   "IVAexento|N|S|0||spara1|IvaIntracom|||"
         Text            =   "Text1"
         Top             =   2760
         Width           =   615
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
         Index           =   52
         Left            =   -72180
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   163
         Text            =   "Text2"
         Top             =   2760
         Width           =   2265
      End
      Begin VB.ComboBox cboObsFactura 
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
         ItemData        =   "frmConfParamAplic.frx":0168
         Left            =   -74760
         List            =   "frmConfParamAplic.frx":016A
         Style           =   2  'Dropdown List
         TabIndex        =   65
         Tag             =   "Orden Descuentos|N|S|||spara1|obsfactura|||"
         Top             =   1200
         Width           =   2655
      End
      Begin VB.Frame Frame9 
         Caption         =   "Aportaci�n en facturas"
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
         Height          =   723
         Left            =   -74865
         TabIndex        =   158
         Top             =   3345
         Width           =   12705
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
            Index           =   51
            Left            =   2400
            MaxLength       =   10
            TabIndex        =   41
            Tag             =   "Cta aportacion|N|S|||spara1|ctaaportacion|||"
            Text            =   "3"
            Top             =   240
            Width           =   1340
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
            Index           =   51
            Left            =   3750
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   159
            Text            =   "Text2"
            Top             =   240
            Width           =   6465
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   42
            Left            =   2130
            Tag             =   "-1"
            ToolTipText     =   "Buscar cuenta contable"
            Top             =   240
            Width           =   240
         End
         Begin VB.Label Label2 
            Caption         =   "Cuenta"
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
            Index           =   8
            Left            =   150
            TabIndex        =   160
            Top             =   285
            Width           =   1110
         End
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
         Index           =   48
         Left            =   -66975
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   156
         Text            =   "Text2"
         Top             =   3120
         Width           =   2745
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
         Index           =   48
         Left            =   -67650
         MaxLength       =   2
         TabIndex        =   73
         Tag             =   "IVAexento|N|S|0||spara1|ivaexento|||"
         Text            =   "Text1"
         Top             =   3120
         Width           =   615
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
         Index           =   49
         Left            =   -65520
         MaxLength       =   5
         TabIndex        =   67
         Tag             =   "N� Contabilidad|N|S|||spara1|porreten|||"
         Text            =   "3"
         Top             =   1200
         Width           =   660
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
         Index           =   47
         Left            =   -69330
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   152
         Text            =   "Text2"
         Top             =   1200
         Width           =   3105
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
         Index           =   47
         Left            =   -70680
         MaxLength       =   10
         TabIndex        =   66
         Tag             =   "Cta retencion|N|S|||spara1|ctareten|||"
         Text            =   "3"
         Top             =   1200
         Width           =   1340
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
         Index           =   46
         Left            =   -72855
         MaxLength       =   2
         TabIndex        =   72
         Tag             =   "REA|N|S|0||spara1|iva_rea|||"
         Text            =   "Text1"
         Top             =   3120
         Width           =   615
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
         Index           =   46
         Left            =   -72180
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   150
         Text            =   "Text2"
         Top             =   3120
         Width           =   2265
      End
      Begin VB.Frame FrameTelefoniaArtic 
         Caption         =   "Telefon�a"
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
         Height          =   1215
         Left            =   -74865
         TabIndex        =   135
         Top             =   5025
         Width           =   12705
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
            Index           =   75
            Left            =   2400
            MaxLength       =   16
            TabIndex        =   273
            Tag             =   "Tfni |T|S|||spara1|artTfoniaIvaExento|||"
            Text            =   "Text1 "
            Top             =   720
            Width           =   2070
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
            Index           =   75
            Left            =   4515
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   272
            Text            =   "Text2"
            Top             =   720
            Width           =   5700
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
            Index           =   32
            Left            =   4515
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   136
            Text            =   "Text2"
            Top             =   300
            Width           =   5700
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
            Index           =   32
            Left            =   2400
            MaxLength       =   16
            TabIndex        =   59
            Tag             =   "Recar |T|S|||spara1|codartictel|||"
            Text            =   "Text1 "
            Top             =   297
            Width           =   2070
         End
         Begin VB.Label Label1 
            Caption         =   "Art�culo IVA exento"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   85
            Left            =   150
            TabIndex        =   274
            Top             =   780
            Width           =   1920
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   75
            Left            =   2130
            Tag             =   "-1"
            ToolTipText     =   "Buscar art�culo"
            Top             =   765
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   32
            Left            =   2130
            Tag             =   "-1"
            ToolTipText     =   "Buscar art�culo"
            Top             =   330
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Art�culo a facturar"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   32
            Left            =   150
            TabIndex        =   137
            Top             =   360
            Width           =   1875
         End
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
         Index           =   1
         Left            =   3180
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   134
         Text            =   "Text2"
         Top             =   960
         Width           =   4155
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
         Index           =   23
         Left            =   -73800
         MaxLength       =   30
         TabIndex        =   60
         Tag             =   "Servidor Contabilidad|T|S|||spara1|serconta|||"
         Text            =   "3wwwwwwwwwwwwwwwwwwwwwwwwwwwww"
         Top             =   600
         Width           =   1665
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
         Index           =   22
         Left            =   -67560
         MaxLength       =   2
         TabIndex        =   63
         Tag             =   "N� Contabilidad|N|S|||spara1|numconta|||"
         Text            =   "3"
         Top             =   600
         Width           =   345
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
         Index           =   21
         Left            =   -71280
         MaxLength       =   20
         TabIndex        =   61
         Tag             =   "Usuario Contabilidad|T|S|||spara1|usuconta|||"
         Text            =   "3wwwwwwwwwwwwwwwwwww"
         Top             =   600
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
         IMEMode         =   3  'DISABLE
         Index           =   20
         Left            =   -69600
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   62
         Tag             =   "Password Contabilidad|T|S|||spara1|pasconta|||"
         Text            =   "3"
         Top             =   600
         Width           =   1185
      End
      Begin VB.ComboBox cboOrdenDtos 
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
         Left            =   2460
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Tag             =   "Orden Descuentos|N|N|||spara1|ordendto|||"
         Top             =   2100
         Width           =   1815
      End
      Begin VB.Frame Frame2 
         Caption         =   "Desplazamientos"
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
         Height          =   750
         Left            =   -74865
         TabIndex        =   126
         Top             =   1665
         Width           =   12705
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
            Left            =   2400
            MaxLength       =   16
            TabIndex        =   39
            Tag             =   "Art�culo para facturar desplazamientos |T|S|||spara1|codartid|||"
            Text            =   "Text1 "
            Top             =   327
            Width           =   2070
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
            Index           =   15
            Left            =   4515
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   127
            Text            =   "Text2"
            Top             =   330
            Width           =   5700
         End
         Begin VB.Label Label1 
            Caption         =   "Art�culo a facturar"
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
            Left            =   150
            TabIndex        =   128
            Top             =   360
            Width           =   1860
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   15
            Left            =   2130
            Tag             =   "-1"
            ToolTipText     =   "Buscar art�culo"
            Top             =   360
            Width           =   240
         End
      End
      Begin VB.ComboBox cboTipodtos 
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
         Left            =   2460
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Tag             =   "Tipo Descuentos|N|N|||spara1|tipodtos|||"
         Top             =   2520
         Width           =   1815
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
         Left            =   2460
         MaxLength       =   3
         TabIndex        =   2
         Tag             =   "C�digo Tarifa PVP|N|N|||spara1|codtarif|000||"
         Text            =   "Text1"
         Top             =   960
         Width           =   705
      End
      Begin VB.Frame FramePrecioKm 
         Caption         =   "Precio Km"
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
         Height          =   1070
         Left            =   240
         TabIndex        =   117
         Top             =   3840
         Width           =   4365
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
            Left            =   2970
            MaxLength       =   10
            TabIndex        =   7
            Tag             =   "Precio Km desplaz. Clientes|N|S|0|9999.0000|spara1|preukmcl|#,##0.0000||"
            Text            =   "Text1"
            Top             =   240
            Width           =   1020
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
            Left            =   2970
            MaxLength       =   10
            TabIndex        =   8
            Tag             =   "Precio Km desplaz. T�cnicos|N|S|0|9999.0000|spara1|preukmtc|#,##0.0000||"
            Text            =   "Text1"
            Top             =   660
            Width           =   1020
         End
         Begin VB.Label Label1 
            Caption         =   "Desplazamiento Clientes"
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
            Left            =   120
            TabIndex        =   119
            Top             =   255
            Width           =   2445
         End
         Begin VB.Label Label1 
            Caption         =   "Desplazamiento T�cnicos"
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
            Left            =   120
            TabIndex        =   118
            Top             =   660
            Width           =   2565
         End
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
         Index           =   4
         Left            =   2085
         MaxLength       =   35
         TabIndex        =   0
         Tag             =   "Nombre Director Gerente|T|S|||spara1|nomgeren|||"
         Text            =   "Text1"
         Top             =   480
         Width           =   3240
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
         Left            =   7560
         MaxLength       =   35
         TabIndex        =   1
         Tag             =   "Nombre responsable Admon|T|S|||spara1|nomadmin|||"
         Text            =   "Text1"
         Top             =   480
         Width           =   4890
      End
      Begin VB.Frame FrameOpciones 
         Caption         =   "Opciones"
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
         Height          =   5295
         Left            =   4725
         TabIndex        =   116
         Top             =   1440
         Width           =   8025
         Begin VB.CheckBox chkVarios 
            Caption         =   "Inventario x c�digo art�culo"
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
            Index           =   5
            Left            =   240
            TabIndex        =   276
            Tag             =   "Inv. x nombre articulo|N|N|||spara1|InventarioPorCodigo|||"
            Top             =   3312
            Width           =   3045
         End
         Begin VB.CheckBox chkVarios 
            Caption         =   "Rec�lculo del margen"
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
            Index           =   4
            Left            =   4155
            TabIndex        =   275
            Tag             =   "Rapida|N|N|||spara1|ActualizaMargen|||"
            Top             =   4440
            Width           =   3195
         End
         Begin VB.CheckBox chkVarios 
            Caption         =   "Actualiza precio especial"
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
            Index           =   2
            Left            =   4155
            TabIndex        =   33
            Tag             =   "Pr.esp.|N|N|||spara1|ActualizaPrecioEspecial|||"
            Top             =   3696
            Width           =   3315
         End
         Begin VB.CheckBox chkVarios 
            Caption         =   "Inicializar stock en inventario"
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
            Index           =   1
            Left            =   4155
            TabIndex        =   32
            Tag             =   "Inistock|N|N|||spara1|IncializarStokInv|||"
            Top             =   3312
            Width           =   3315
         End
         Begin VB.CheckBox chkVarios 
            Caption         =   "Recepcion mercan. solo ppal"
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
            Index           =   0
            Left            =   4155
            TabIndex        =   31
            Tag             =   "Merca. solo ppal|N|N|||spara1|RecMercanciaSoloPpal|||"
            Top             =   2928
            Width           =   3315
         End
         Begin VB.CheckBox chkPrecioMinimo 
            Caption         =   "Precio minimo"
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
            Left            =   4155
            TabIndex        =   30
            Tag             =   "Precio minimo|N|N|||spara1|preciominimo|||"
            Top             =   2550
            Width           =   3075
         End
         Begin VB.CheckBox chkGrabaLogPredto 
            Caption         =   "Graba log precio /dtos"
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
            Left            =   4155
            TabIndex        =   29
            Tag             =   "Operaciones aseguradas|N|N|||spara1|LogCambioPrecDto|||"
            Top             =   2160
            Width           =   3075
         End
         Begin VB.CheckBox chkTicketsAgrupads 
            Caption         =   "Contabilizar ticket TPV agrupados"
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
            Left            =   4155
            TabIndex        =   28
            Tag             =   "Tickets agrupadsos|N|N|||spara1|conttickagrupado|||"
            Top             =   1776
            Width           =   3675
         End
         Begin VB.CheckBox chkDireccionEnvio 
            Caption         =   "Direcciones de envio"
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
            Left            =   4155
            TabIndex        =   25
            Tag             =   "Dir envio|N|N|||spara1|DirecEnvio|||"
            Top             =   624
            Width           =   2835
         End
         Begin VB.CheckBox ChkDtoxCantidad 
            Caption         =   "Hay Dtos por cantidad"
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
            Left            =   4155
            TabIndex        =   24
            Tag             =   "Hay Dtos por cantidad|N|N|||spara1|dtoxcanti|||"
            Top             =   240
            Width           =   2835
         End
         Begin VB.CheckBox chkMataPrimaPorcen 
            Caption         =   "Materia prima como porcentaje"
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
            Left            =   240
            TabIndex        =   22
            Tag             =   "Descriptores|N|N|||spara1|compoporcen|||"
            Top             =   4080
            Width           =   3675
         End
         Begin VB.CheckBox chkDescriptores 
            Caption         =   "Usa descriptores especiales"
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
            Left            =   240
            TabIndex        =   21
            Tag             =   "Descriptores|N|N|||spara1|descriptores|||"
            Top             =   3696
            Width           =   3675
         End
         Begin VB.CheckBox chkProduccion 
            Caption         =   "Tiene producci�n"
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
            Left            =   240
            TabIndex        =   20
            Tag             =   "Tiene produccion|N|N|||spara1|produccion|||"
            Top             =   4800
            Width           =   3675
         End
         Begin VB.CheckBox chkHayServicio 
            Caption         =   "Hay Servicios"
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
            Left            =   240
            TabIndex        =   16
            Tag             =   "Hay Servicios|N|N|||spara1|hayservicio|||"
            Top             =   1776
            Width           =   3075
         End
         Begin VB.CheckBox chkCajacomp 
            Caption         =   "Cajas completas precios"
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
            Left            =   240
            TabIndex        =   12
            Tag             =   "Cajas Completas Precios|N|N|||spara1|cajacomp|||"
            Top             =   240
            Width           =   3075
         End
         Begin VB.CheckBox chkHaymante 
            Caption         =   "Realiza Mantenimientos"
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
            Left            =   240
            TabIndex        =   13
            Tag             =   "Mantenimientos|N|N|||spara1|haymante|||"
            Top             =   624
            Width           =   3075
         End
         Begin VB.CheckBox chkHayrepar 
            Caption         =   "Realiza Reparaciones"
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
            Left            =   240
            TabIndex        =   14
            Tag             =   "Reparaciones|N|N|||spara1|hayrepar|||"
            Top             =   1008
            Width           =   2850
         End
         Begin VB.CheckBox chkHayfrecu 
            Caption         =   "Hay Frecuencias"
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
            Left            =   240
            TabIndex        =   15
            Tag             =   "Hay Frecuencias|N|N|||spara1|hayfrecu|||"
            Top             =   1392
            Width           =   2850
         End
         Begin VB.CheckBox chkctrstock 
            Caption         =   "Control de Stock estricto"
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
            Left            =   240
            TabIndex        =   18
            Tag             =   "Control de Stock|N|N|||spara1|ctrstock|||"
            Top             =   2544
            Width           =   3675
         End
         Begin VB.CheckBox chkInventar 
            Caption         =   "Realiza Inventario x Proveedor"
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
            Left            =   240
            TabIndex        =   19
            Tag             =   "Inventarios por Proveedor|N|N|||spara1|inventar|||"
            Top             =   2928
            Width           =   3675
         End
         Begin VB.CheckBox chkHaynserie 
            Caption         =   "Hay N� Serie en Compras"
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
            Left            =   240
            TabIndex        =   17
            Tag             =   "Hay N� Serie en Compras|N|N|||spara1|haynserie|||"
            Top             =   2160
            Width           =   3075
         End
         Begin VB.CheckBox chkFraMost 
            Caption         =   "Fact. mostrador separada"
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
            Left            =   4155
            TabIndex        =   26
            Tag             =   "Dir envio|N|N|||spara1|FraMostra|||"
            Top             =   1008
            Width           =   3195
         End
         Begin VB.CheckBox chkMarcarParaFacturar 
            Caption         =   "Marcar albar�n para facturar"
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
            Left            =   4155
            TabIndex        =   27
            Tag             =   "Albfra|N|N|||spara1|AlbParaFcturar|||"
            Top             =   1392
            Width           =   3315
         End
         Begin VB.CheckBox chkAseguradas 
            Caption         =   "Operaciones aseguradas"
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
            Left            =   4155
            TabIndex        =   34
            Tag             =   "Operaciones aseguradas|N|N|||spara1|OperAseguradas|||"
            Top             =   4080
            Width           =   3075
         End
         Begin VB.CheckBox chkVarios 
            Caption         =   "Entrada r�pida en Fras. mostrador"
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
            Index           =   3
            Left            =   240
            TabIndex        =   23
            Tag             =   "Rapida|N|N|||spara1|FrasMostradorRapida|||"
            Top             =   4440
            Width           =   3795
         End
         Begin VB.Image imgayuda 
            Height          =   255
            Index           =   14
            Left            =   7575
            ToolTipText     =   "Operaciones aseguradas"
            Top             =   4560
            Width           =   255
         End
         Begin VB.Image imgayuda 
            Height          =   255
            Index           =   13
            Left            =   7575
            ToolTipText     =   "Actualiza precio especial"
            Top             =   3720
            Width           =   255
         End
         Begin VB.Image imgayuda 
            Height          =   255
            Index           =   12
            Left            =   7575
            ToolTipText     =   "Incializar stock"
            Top             =   3360
            Width           =   255
         End
         Begin VB.Image imgayuda 
            Height          =   255
            Index           =   11
            Left            =   7575
            ToolTipText     =   "Recepcion mercancia"
            Top             =   3000
            Width           =   255
         End
         Begin VB.Image imgayuda 
            Height          =   255
            Index           =   5
            Left            =   7575
            ToolTipText     =   "Operaciones aseguradas"
            Top             =   4110
            Width           =   255
         End
         Begin VB.Image imgayuda 
            Height          =   255
            Index           =   4
            Left            =   7575
            ToolTipText     =   "Alb. para facturar"
            Top             =   1440
            Width           =   255
         End
         Begin VB.Image imgayuda 
            Height          =   255
            Index           =   3
            Left            =   7575
            ToolTipText     =   "Fra. mostrador"
            Top             =   1080
            Width           =   255
         End
         Begin VB.Image imgayuda 
            Height          =   255
            Index           =   1
            Left            =   7575
            ToolTipText     =   "Dir. envio"
            Top             =   720
            Width           =   255
         End
      End
      Begin VB.Frame FrameDiasMante 
         Caption         =   "D�as Reparaci�n"
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
         Height          =   1095
         Left            =   240
         TabIndex        =   113
         Top             =   4920
         Width           =   4365
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
            Index           =   6
            Left            =   2955
            MaxLength       =   4
            TabIndex        =   10
            Tag             =   "Dias Repar. sin Mantenimiento|N|N|0|9999|spara1|diasnoman|||"
            Text            =   "Text"
            Top             =   680
            Width           =   1035
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
            Index           =   7
            Left            =   2955
            MaxLength       =   4
            TabIndex        =   9
            Tag             =   "Dias Repar. con Mantenimiento|N|N|0|9999|spara1|diassiman|||"
            Text            =   "Text1"
            Top             =   300
            Width           =   1035
         End
         Begin VB.Label Label1 
            Caption         =   "Sin Mantenimiento"
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
            TabIndex        =   115
            Top             =   675
            Width           =   1920
         End
         Begin VB.Label Label1 
            Caption         =   "Con Mantenimiento"
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
            Index           =   8
            Left            =   120
            TabIndex        =   114
            Top             =   300
            Width           =   2040
         End
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   3960
         MaxLength       =   15
         TabIndex        =   124
         Tag             =   "C�digo Par�metros Aplic|N|N|||spara1|codigo||S|"
         Text            =   "Text1"
         Top             =   480
         Width           =   645
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   2205
         Picture         =   "frmConfParamAplic.frx":016C
         Tag             =   "-1"
         ToolTipText     =   "Buscar tarifa"
         Top             =   990
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   79
         Left            =   -67680
         ToolTipText     =   "Buscar situacion"
         Top             =   3960
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Situaci�n riesgo sin bloqueo"
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
         Index           =   89
         Left            =   -70560
         TabIndex        =   283
         Top             =   3915
         Width           =   2850
      End
      Begin VB.Label Label1 
         Caption         =   "Albaranes firmados"
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
         Index           =   88
         Left            =   -74640
         TabIndex        =   281
         Top             =   5760
         Width           =   1905
      End
      Begin VB.Label Label1 
         Caption         =   "Facturas firmadas"
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
         Index           =   87
         Left            =   -74640
         TabIndex        =   280
         Top             =   5280
         Width           =   1770
      End
      Begin VB.Label Label1 
         Caption         =   "Precio hora (EUL)"
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
         Index           =   86
         Left            =   -74715
         TabIndex        =   278
         Top             =   6525
         Width           =   1755
      End
      Begin VB.Label Label1 
         Caption         =   "Path facturae"
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
         Index           =   75
         Left            =   -74640
         TabIndex        =   267
         Top             =   4800
         Width           =   1455
      End
      Begin VB.Image imgayuda 
         Height          =   255
         Index           =   6
         Left            =   -74880
         ToolTipText     =   "Operaciones aseguradas"
         Top             =   3840
         Width           =   255
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   65
         Left            =   -71625
         ToolTipText     =   "Buscar situacion"
         Top             =   3960
         Width           =   240
      End
      Begin VB.Image imgayuda 
         Height          =   255
         Index           =   9
         Left            =   -64215
         ToolTipText     =   "Compras intracomunitarias"
         Top             =   2760
         Width           =   255
      End
      Begin VB.Image imgayuda 
         Height          =   255
         Index           =   8
         Left            =   -71400
         ToolTipText     =   "Compras intracomunitarias"
         Top             =   1680
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "Intracom prov"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   74
         Left            =   -69525
         TabIndex        =   242
         Top             =   2760
         Width           =   1590
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   66
         Left            =   -73395
         Tag             =   "-1"
         ToolTipText     =   "Buscar I.V.A."
         Top             =   1920
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "IVA especiales"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   73
         Left            =   -74760
         TabIndex        =   240
         Top             =   2400
         Width           =   1725
      End
      Begin VB.Label Label1 
         Caption         =   "Contabi. fras prov.  intracomumitarias"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   72
         Left            =   -74760
         TabIndex        =   239
         Top             =   1680
         Width           =   3300
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   67
         Left            =   -67935
         Tag             =   "-1"
         ToolTipText     =   "Buscar cuenta contable"
         Top             =   2760
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Cuenta extra"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   -74760
         TabIndex        =   238
         Top             =   1920
         Width           =   1350
      End
      Begin VB.Image imgayuda 
         Height          =   255
         Index           =   2
         Left            =   4305
         ToolTipText     =   "Departamento-Direcciones"
         Top             =   3360
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Departamento"
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
         Index           =   63
         Left            =   360
         TabIndex        =   185
         Top             =   3390
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Crear tarifas"
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
         Index           =   59
         Left            =   360
         TabIndex        =   175
         Top             =   2970
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Modo anal�tica"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   58
         Left            =   -64755
         TabIndex        =   174
         Top             =   1680
         Width           =   1515
      End
      Begin VB.Label Label1 
         Caption         =   "N� Con presu *"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   51
         Left            =   -66600
         TabIndex        =   168
         Top             =   405
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Intracomun."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   47
         Left            =   -74520
         TabIndex        =   164
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   45
         Left            =   -73140
         Tag             =   "-1"
         ToolTipText     =   "Buscar I.V.A."
         Top             =   2760
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Integ.Fras.Observaciones "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   53
         Left            =   -74760
         TabIndex        =   162
         Top             =   960
         Width           =   2640
      End
      Begin VB.Label Label1 
         Caption         =   "R.E.A."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   52
         Left            =   -74520
         TabIndex        =   161
         Top             =   3120
         Width           =   555
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   41
         Left            =   -67935
         Tag             =   "-1"
         ToolTipText     =   "Buscar I.V.A."
         Top             =   3120
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Exento"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   50
         Left            =   -69525
         TabIndex        =   157
         Top             =   3120
         Width           =   1470
      End
      Begin VB.Label Label1 
         Caption         =   "Contabilidad"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   49
         Left            =   -74760
         TabIndex        =   155
         Top             =   360
         Width           =   1275
      End
      Begin VB.Label Label2 
         Caption         =   "%"
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
         Index           =   6
         Left            =   -65880
         TabIndex        =   154
         Top             =   1200
         Width           =   255
      End
      Begin VB.Label Label2 
         Caption         =   "Cuenta"
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
         Left            =   -71760
         TabIndex        =   153
         Top             =   1200
         Width           =   705
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   40
         Left            =   -70950
         Tag             =   "-1"
         ToolTipText     =   "Buscar cuenta contable"
         Top             =   1200
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Retenci�n"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   48
         Left            =   -71910
         TabIndex        =   151
         Top             =   960
         Width           =   1155
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   39
         Left            =   -73140
         Tag             =   "-1"
         ToolTipText     =   "Buscar I.V.A."
         Top             =   3120
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Servidor"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   19
         Left            =   -74760
         TabIndex        =   133
         Top             =   600
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "N� conta"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   18
         Left            =   -68280
         TabIndex        =   132
         Top             =   600
         Width           =   660
      End
      Begin VB.Label Label1 
         Caption         =   "Usuario"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   17
         Left            =   -72000
         TabIndex        =   131
         Top             =   600
         Width           =   780
      End
      Begin VB.Label Label1 
         Caption         =   "Pass."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   15
         Left            =   -70200
         TabIndex        =   130
         Top             =   600
         Width           =   840
      End
      Begin VB.Label Label1 
         Caption         =   "Orden Descuentos"
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
         Left            =   360
         TabIndex        =   129
         Top             =   2130
         Width           =   1845
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo de Descuentos"
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
         Left            =   360
         TabIndex        =   123
         Top             =   2550
         Width           =   2070
      End
      Begin VB.Label Label1 
         Caption         =   "C�digo Tarifa PVP"
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
         Left            =   360
         TabIndex        =   122
         Top             =   960
         Width           =   1845
      End
      Begin VB.Label Label1 
         Caption         =   "Director Gerente"
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
         Index           =   4
         Left            =   360
         TabIndex        =   121
         Top             =   510
         Width           =   1860
      End
      Begin VB.Label Label1 
         Caption         =   "Responsable Admon."
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
         Index           =   5
         Left            =   5475
         TabIndex        =   120
         Top             =   540
         Width           =   2100
      End
      Begin VB.Label Label1 
         Caption         =   "C�digo"
         Height          =   255
         Index           =   6
         Left            =   3120
         TabIndex        =   125
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Situaci�n riesgo con  bloqueo "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   67
         Left            =   -74640
         TabIndex        =   263
         Top             =   3915
         Width           =   2940
      End
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Enabled         =   0   'False
      Visible         =   0   'False
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
Attribute VB_Name = "frmConfParamAplic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents frmMtoArt As frmBasico2
Attribute frmMtoArt.VB_VarHelpID = -1
Private WithEvents frmFP As frmBasico2 'frmFacFormasPago 'Form Mto Formas de Pago
Attribute frmFP.VB_VarHelpID = -1


Private WithEvents frmA As frmFacActividades
Attribute frmA.VB_VarHelpID = -1
Private WithEvents frmE As frmFacFormasEnvio
Attribute frmE.VB_VarHelpID = -1
Private WithEvents frmZ As frmFacZonas
Attribute frmZ.VB_VarHelpID = -1
Private WithEvents frmR As frmFacRutas
Attribute frmR.VB_VarHelpID = -1
Private WithEvents frmAc As frmBasico2 '%=%=frmFacAgentesCom
Attribute frmAc.VB_VarHelpID = -1
Private WithEvents frmT As frmFacTarifas
Attribute frmT.VB_VarHelpID = -1
Private WithEvents frmS As frmFacSituaciones
Attribute frmS.VB_VarHelpID = -1


Private WithEvents frmB As frmBasico2 'BuscaGrid
Attribute frmB.VB_VarHelpID = -1


Private NombreTabla As String  'Nombre de la tabla o de la
Private CadenaConsulta As String



Dim PrimeraVez As Boolean
Dim Modo As Byte
'0: Inicial
'2: Visualizacion
'3: A�adir
'4: Modificar






Private Sub cboCreaTarifa_KeyPress(KeyAscii As Integer)
 KEYpress KeyAscii
End Sub



Private Sub cboDpto_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cboObsFactura_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cboOrdenDtos_KeyPress(KeyAscii As Integer)
      KEYpress KeyAscii
End Sub


Private Sub cboPortes_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cboTipodtos_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub





Private Sub Check1_Click()

End Sub

Private Sub chkAseguradas_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub chkAseguradas_KeyPress(KeyAscii As Integer)
     KEYpress KeyAscii
End Sub

Private Sub chkCajacomp_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub chkCajacomp_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkContabIntraCom_GotFocus()
 ConseguirfocoChk Modo
End Sub

Private Sub chkContabIntraCom_KeyPress(KeyAscii As Integer)
KEYpress KeyAscii
End Sub

Private Sub chkctrstock_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub chkctrstock_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub








Private Sub chkDireccionEnvio_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub chkDireccionEnvio_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub ChkDtoxCantidad_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub ChkDtoxCantidad_KeyPress(KeyAscii As Integer)
     KEYpress KeyAscii
End Sub

Private Sub chkHaydepar_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub chkHaydepar_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub



Private Sub chkFraMost_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub chkFraMost_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub



Private Sub chkGrabaLogPredto_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub chkGrabaLogPredto_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkHayfrecu_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub chkHayfrecu_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub chkHaymante_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub chkHaymante_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub chkHaynserie_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub chkHaynserie_KeyPress(KeyAscii As Integer)
     KEYpress KeyAscii
End Sub

Private Sub chkHayrepar_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub chkHayrepar_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkHayServicio_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub chkHayServicio_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkInventar_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub chkInventar_KeyDown(KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub chkInventar_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub






Private Sub chkMarcarParaFacturar_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub chkMarcarParaFacturar_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkMataPrimaPorcen_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub chkMataPrimaPorcen_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub



Private Sub chkOutlook_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub chkOutlook_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkPrecioMinimo_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub chkPrecioMinimo_KeyPress(KeyAscii As Integer)
     KEYpress KeyAscii
End Sub

Private Sub chkProduccion_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub chkProduccion_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkDescriptores_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub chkDescriptores_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub





Private Sub chkRenting_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub chkRenting_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkTicketsAgrupads_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub chkTicketsAgrupads_KeyPress(KeyAscii As Integer)
  KEYpress KeyAscii
End Sub




Private Sub chkVarios_GotFocus(Index As Integer)
    ConseguirfocoChk Modo
End Sub

Private Sub chkVarios_KeyPress(Index As Integer, KeyAscii As Integer)
     KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()
Dim actualiza As Boolean
Dim kms As Currency



    If Modo = 4 Then 'MODIFICAR
        If DatosOk Then
            vParamAplic.TipoDtos = Me.cboTipodtos.ListIndex
            vParamAplic.OrdenDtos = Me.cboOrdenDtos.ListIndex
            vParamAplic.ObsFactura = Me.cboObsFactura.ListIndex
            vParamAplic.CodTarifa = Text1(1).Text
            vParamAplic.NomGerente = Text1(4).Text
            vParamAplic.NomRespAdmin = Text1(5).Text
            kms = ImporteFormateado(ComprobarCero(Text1(2).Text))
            vParamAplic.PrecioKmClientes = CSng(CStr(kms))
            kms = ImporteFormateado(ComprobarCero(Text1(3).Text))
            vParamAplic.PrecioKmTecnicos = CSng(CStr(kms))
            vParamAplic.CajasCompletas = Me.chkCajacomp.Value
            vParamAplic.Mantenimientos = Me.chkHaymante.Value
            vParamAplic.Reparaciones = Me.chkHayrepar.Value
            vParamAplic.Frecuencias = Me.chkHayfrecu.Value
            vParamAplic.Servicios = Me.chkHayServicio.Value
            
            'Julio 2010
            vParamAplic.HayDeparNuevo = Me.cboDpto.ItemData(cboDpto.ListIndex)
            
            vParamAplic.ControlStock = Me.chkctrstock.Value
            vParamAplic.InventarioxProv = Me.chkInventar.Value
            vParamAplic.NumSeries = Me.chkHaynserie.Value  'Hay N� Serie en Compras?
            vParamAplic.DiasSiMante = Me.Text1(7).Text 'Dias Rep. con Mantenimiento
            vParamAplic.DiasNoMante = Me.Text1(6).Text 'Dias Rep. sin Mantenimiento
            
            'Articulo para facturar mantenimientos
            vParamAplic.ArticDesplaz = Me.Text1(15).Text
            'dias de pago para compras
            vParamAplic.DiaPago1 = CByte(DBLet(ComprobarCero(Text1(16).Text), "N"))
            vParamAplic.DiaPago2 = CByte(DBSet(Text1(17).Text, "N"))
            vParamAplic.DiaPago3 = CByte(DBSet(Text1(18).Text, "N"))
            vParamAplic.MesNoGirar = CByte(DBSet(Text1(19).Text, "N"))
            vParamAplic.ForPagoChequeRegalo = Me.Text1(24).Text
            
            vParamAplic.DireMail = Text1(8).Text 'Direccion email
            vParamAplic.SMTPhost = Text1(9).Text
            vParamAplic.SMTPuser = Text1(10).Text
            vParamAplic.SMTPpass = Text1(11).Text
            vParamAplic.WebSoporte = Text1(12).Text
            vParamAplic.MailSoporte = Text1(13).Text
            vParamAplic.WebVersion = Text1(14).Text
            
            'Datos contabilidad
            vParamAplic.ServidorConta = Text1(23).Text
            vParamAplic.UsuarioConta = Text1(21).Text
            vParamAplic.PasswordConta = Text1(20).Text
            vParamAplic.NumeroConta = ComprobarCero(Text1(22).Text)
            
            'Valores por defecto
            vParamAplic.PorDefecto_Activ = ComprobarCero(Text1(25).Text)
            vParamAplic.PorDefecto_Envio = ComprobarCero(Text1(26).Text)
            vParamAplic.PorDefecto_Zona = ComprobarCero(Text1(27).Text)
            vParamAplic.PorDefecto_Ruta = ComprobarCero(Text1(28).Text)
            vParamAplic.PorDefecto_Situ = ComprobarCero(Text1(29).Text)
            vParamAplic.PorDefecto_Tarifa = ComprobarCero(Text1(30).Text)
            vParamAplic.PorDefecto_Agente = ComprobarCero(Text1(31).Text)
            
            'Telefonia  2013. Ya no se utliza
            'vParamAplic.CodarticTfnia = Me.Text1(32).Text
            
            'Los avisos
            vParamAplic.avipedcli = ComprobarCero(Text1(33).Text)
            vParamAplic.avipedpro = ComprobarCero(Text1(34).Text)
            vParamAplic.avialbcli = ComprobarCero(Text1(35).Text)
            vParamAplic.avialbpro = ComprobarCero(Text1(36).Text)
            vParamAplic.avimanteni = ComprobarCero(Text1(37).Text)
            vParamAplic.aviavisos = ComprobarCero(Text1(38).Text)
            vParamAplic.avirepara = ComprobarCero(Text1(39).Text)
            
            
            'Los tipos de IVA
            vParamAplic.TipoIVAre1 = ComprobarCero(Text1(40).Text)
            vParamAplic.TipoIVAre2 = ComprobarCero(Text1(41).Text)
            vParamAplic.TipoIVAre3 = ComprobarCero(Text1(42).Text)
             
            vParamAplic.TipoIVA1 = ComprobarCero(Text1(43).Text)
            vParamAplic.TipoIVA2 = ComprobarCero(Text1(44).Text)
            vParamAplic.TipoIVA3 = ComprobarCero(Text1(45).Text)
            
            
            'Los tipos de IVA   antiguos     JUNIO 2010

             
            vParamAplic.OLDIVA1 = ComprobarCero(Text1(59).Text)
            vParamAplic.OLDIVA2 = ComprobarCero(Text1(60).Text)
           ' vParamAplic.OLDIVA3 = ComprobarCero(Text1(61).Text)
            
            vParamAplic.OLDIVAre1 = ComprobarCero(Text1(61).Text)
            vParamAplic.OLDIVAre2 = ComprobarCero(Text1(62).Text)
           ' vParamAplic.OLDIVAre3 = ComprobarCero(Text1(64).Text)
            
            'REtencion y REA
            vParamAplic.IVA_REA = ComprobarCero(Text1(46).Text)
            vParamAplic.CtaReten = ComprobarCero(Text1(47).Text)
            vParamAplic.PorReten = ComprobarCero(Text1(49).Text)
            
            'IVA exento
            vParamAplic.IVA_Exento2 = ComprobarCero(Text1(48).Text)
            vParamAplic.IVA_Intracomunitario = ComprobarCero(Text1(52).Text)

            
            'Tickets acgrupados
            vParamAplic.ContabilizarTicketAgrupados = Me.chkTicketsAgrupads.Value
            
            vParamAplic.ContabilidadB = ComprobarCero(Text1(50).Text)
            vParamAplic.ctaAportacion = Text1(51).Text
            
            vParamAplic.Produccion = Me.chkProduccion.Value
            vParamAplic.Descriptores = Me.chkDescriptores.Value
            
            vParamAplic.ArtReciclado = Text1(53).Text
            
            'Portes(FOntenas)
            vParamAplic.ArtPortesN = Text1(54).Text
            vParamAplic.AbonoKilos = ComprobarCero(Text1(55).Text)
            vParamAplic.ImporteMinimo = ComprobarCero(Text1(56).Text)
            
            vParamAplic.ComponentePorcentaje = Me.chkMataPrimaPorcen.Value
            
            ' ---- [14/09/2009] (LAURA)
            vParamAplic.DtoxCantidad = Me.ChkDtoxCantidad.Value
            vParamAplic.CreaTarifasArticulo = Me.cboCreaTarifa.ItemData(cboCreaTarifa.ListIndex)
            
            ' ----
            
            ' ---- [19/10/2009] [LAURA]: a�adir campo modo analitica
            If Me.CboModAnalitica.ListIndex >= 0 Then
                vParamAplic.ModoAnalitica = Me.CboModAnalitica.ListIndex
            End If
            
            vParamAplic.EnvioDesdeOutlook = Me.chkOutlook.Value
            
            
            vParamAplic.ExeEnvioMail = Trim(Text1(57).Text)
            vParamAplic.DiasGarantia = ComprobarCero(Text1(58).Text)
            
            
            vParamAplic.NumCopiasFacturacion = ComprobarCero(Text1(63).Text)
            If vParamAplic.NumCopiasFacturacion = 0 Then vParamAplic.NumCopiasFacturacion = 1
            
            
            vParamAplic.ArticuloRecargoFinanciero = Text1(64).Text
            'Portes(FOntenas)
            vParamAplic.TipoPortes = cboPortes.ListIndex
            'Direcciones de envio(ademas de departamento-direccion)
            vParamAplic.DireccionesEnvio = Me.chkDireccionEnvio.Value
            'Fras mostrador
            vParamAplic.FrasMostradorSerieDistinta = Me.chkFraMost.Value
            vParamAplic.MarcarAlbaranFacturar = Me.chkMarcarParaFacturar
            
            
            'SOLO LO MODIFICO POR EL YOG  30Dic2010
            'vParamAplic.OperacionesAseguradas = Me.chkAseguradas.Value
            vParamAplic.SituacionBloqueoOpAseg = ComprobarCero(Text1(65).Text)
            vParamAplic.SituacionBloqueoOpAsegSinbloq = ComprobarCero(Text1(79).Text)
            
            
            vParamAplic.Rot_ConsumMes1 = ComprobarCero(Text1(66).Text)
            vParamAplic.Rot_ConsumMes2 = ComprobarCero(Text1(67).Text)
            vParamAplic.Rot_ConsumMesMin = ComprobarCero(Text1(68).Text)
            vParamAplic.Rot_ConsumMesMax = ComprobarCero(Text1(69).Text)
            
            
            vParamAplic.LogCambioPrecDto = Me.chkGrabaLogPredto.Value
            
            
            'Febrero 2011
            vParamAplic.IvaIntracomAdicional = ComprobarCero(Text1(71).Text)
            vParamAplic.CtaContabIntracom = Text1(70).Text
            vParamAplic.IntracomAdicionalContab = Me.chkContabIntraCom.Value
                        
                        
            vParamAplic.PathFacturaE = Text1(72).Text
            vParamAplic.PrecioMinimo = Me.chkPrecioMinimo.Value
                        
            vParamAplic.ArtSeparador = Text1(73).Text
         
            vParamAplic.PorcenPagoAgentTalPag = ComprobarCero(Text1(74).Text)
         
            'Telefonia
            vParamAplic.ArtiTelefonia = Text1(32).Text
            
            vParamAplic.RecMercanciaSoloPpal = chkVarios(0).Value
            vParamAplic.ArtTfniaIvaExento = Text1(75).Text
            
            
            vParamAplic.IncializarStockEnInventario = chkVarios(1).Value
            
            vParamAplic.ActualizaPrecioEspecial = chkVarios(2).Value
            
            vParamAplic.EntradaRapidaFacturasMostrador = chkVarios(3).Value
            
            vParamAplic.RecalculoMargen = chkVarios(4).Value
            
            vParamAplic.InventarioCodigoArticulo = chkVarios(5).Value
            
            
            vParamAplic.PrecioHoraCosteEUL = ImporteFormateado(ComprobarCero(Text1(76).Text))
            
            
            vParamAplic.PathFirmasFacturas = Text1(77).Text
            vParamAplic.PathFirmasAlbaran = Text1(78).Text
            
            
            
            
            
            AsignarNumeroDeCopias
                        
            actualiza = vParamAplic.Modificar(Text1(0).Text)
            TerminaBloquear

            vParamAplic.ComprobarProgramaEnvioMail


            If actualiza Then  'Inserta o Modifica
                'Abrir la conexion a la conta q hemos modificado
                CerrarConexionConta
                If vParamAplic.NumeroConta <> 0 Then
                    If Not AbrirConexionConta(False) Then End
                End If
                PonerModo 2
                PonerFocoBtn Me.cmdSalir
                If vUsu.Skin >= 0 Then MsgBox "Deberia reiniciar el Ariges", vbInformation
            End If
        End If
    End If
End Sub


Private Sub cmdCancelar_Click()
    TerminaBloquear
    If Data1.Recordset.EOF Then
        PonerModo 0
        LimpiarCampos
    Else
        PonerCampos
        PonerModo 2
    End If
End Sub


Private Sub cmdSalir_Click()
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        PonerModo 0
    Else
        If Modo <> 4 Then PonerCadenaBusqueda
        PonerFoco Text1(0)
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub CargaIconosAyuda()
Dim Ima As Image
    On Error Resume Next 'mejor que no diera errores, pero bien, tampoco vamos a enfadarnos
    For Each Ima In Me.imgayuda
        Ima.Picture = frmPpal.imgListComun.ListImages(46).Picture
    Next
    Err.Clear
End Sub

Private Sub Form_Load()
Dim Im
    'Icono del formulario
    Me.Icon = frmPpal.Icon

    'ICONOS AYUDA
    CargaIconosAyuda
    ' ICONITOS DE LA BARRA
'    With Me.Toolbar1
'        .ImageList = frmPpal.imgListComun
''        .Buttons(1).Image = 3   'Anyadir
'        .Buttons(1).Image = 4   'Modificar
'        .Buttons(4).Image = 15  'Salir
'    End With
    
    With Me.Toolbar1
        .HotImageList = frmPpal.imgListComun_OM2
        .DisabledImageList = frmPpal.imgListComun_BN2
        .ImageList = frmPpal.ImgListComun2
        .Buttons(1).Image = 4
    End With
    
    'cargar iconos de busqueda
    For Each Im In Me.imgBuscar
        Im.Picture = imgBuscar(1).Picture 'frmPpal.imgListComun.ListImages(19).Picture
    Next
    
    'imgBuscar(1).Picture = frmPpal.imgListComun.ListImages(19).Picture
    'imgBuscar(15).Picture = frmPpal.imgListComun.ListImages(19).Picture
   '
   ' For NumRegElim = 24 To 42
   '     Me.imgBuscar(NumRegElim).Picture = frmPpal.imgListComun.ListImages(19).Picture
   ' Next NumRegElim
    
    LimpiarCampos   'Limpia los campos TextBox
    Me.SSTab1.Tab = 0
    
    SSTab1.TabVisible(6) = False
    
    CargarComboTipoDtos
    CargarComboOrdenDtos
    CargaComoboObsFactura
    CargarComboModoAnalitica
    
    ' ---- [21/10/2009] [LAURA]
    '-- modo analitica si contabilidad lleva analitica
    If vEmpresa.LeerNiveles Then
        Label1(58).visible = vEmpresa.TieneAnalitica
        Me.CboModAnalitica.visible = vEmpresa.TieneAnalitica
    End If
    
    
    '-- Abril 2016
    'Contabilizacion INTRACOMUNITARIAS. YA no llevan la extra, ni nada. Entran con el IVA que le corresponda, pero sin SUMAR
    imgayuda(8).visible = False
    chkContabIntraCom.visible = False
    Label1(72).visible = False
    Label2(4).visible = False
    Text1(70).visible = False
    Text2(70).visible = False
    
    Label1(74).visible = False
    Text1(71).visible = False
    Text2(71).visible = False
    
    CboModAnalitica.Top = 1600
    CboModAnalitica.Left = 1840
    Label1(58).Left = 240
    
    NombreTabla = "spara1"
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = conn
    CadenaConsulta = "Select * from " & NombreTabla
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    FrameTelefoniaArtic.visible = Val(DBLet(Data1.Recordset!Telefonia, "N")) >= 1
    
    '20 Dic 2010
    'Ope. aseguradas lo habilitamos por el YOG. Auqi solo puede cambiar la situbloq
    chkAseguradas.visible = False
    imgayuda(5).visible = False
    
    
    
    
    
    
    PonerModo 0
End Sub


Private Sub PonerCadenaBusqueda()
On Error GoTo EEPonerBusq
    Screen.MousePointer = vbHourglass

    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        LimpiarCampos
        MsgBox "No hay ning�n registro en la tabla " & NombreTabla, vbInformation
        Screen.MousePointer = vbDefault
        Me.Toolbar1.Buttons(1).Enabled = False 'Modificar
    Else
        Data1.Recordset.MoveFirst
        PonerCampos
    End If
    
    Screen.MousePointer = vbDefault
    Exit Sub
EEPonerBusq:
    MuestraError Err.Number, "PonerCadenaBusqueda"
    Screen.MousePointer = vbDefault
End Sub






Private Sub frmA_DatoSeleccionado(CadenaSeleccion As String)
    'ENVIO
    Text1(25).Text = RecuperaValor(CadenaSeleccion, 1)
    Text2(25).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmAc_DatoSeleccionado(CadenaSeleccion As String)
    'agentes
    Text1(31).Text = RecuperaValor(CadenaSeleccion, 1)
    Text2(31).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmB_DatoSeleccionado(CadenaSeleccion As String)
    CadenaDesdeOtroForm = CadenaSeleccion
End Sub

Private Sub frmE_DatoSeleccionado(CadenaSeleccion As String)
    'ENVIO
    Text1(26).Text = RecuperaValor(CadenaSeleccion, 1)
    Text2(26).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmFP_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Formas de Pago
Dim Indice As Byte

    Indice = 24
    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod Forma Pago
    FormateaCampo Text1(Indice)
    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Forma Pago
End Sub

Private Sub frmMtoArt_DatoSeleccionado(CadenaSeleccion As String)
'Mantenimiento de Articulos
    
    Text1(CInt(imgBuscar(1).Tag)).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod articulo
    Text2(CInt(imgBuscar(1).Tag)).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre articulo
End Sub

Private Sub frmR_DatoSeleccionado(CadenaSeleccion As String)
    'RUTA
    Text1(28).Text = RecuperaValor(CadenaSeleccion, 1)
    Text2(28).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmS_DatoSeleccionado(CadenaSeleccion As String)
    'SITUACION
    Text1(CInt(imgBuscar(1).Tag)).Text = RecuperaValor(CadenaSeleccion, 1)
    Text2(CInt(imgBuscar(1).Tag)).Text = RecuperaValor(CadenaSeleccion, 2)
    
    
End Sub

Private Sub frmT_DatoSeleccionado(CadenaSeleccion As String)
    'TARIFA
    If Not IsNumeric(Me.imgBuscar(1).Tag) Then Exit Sub
    
    If CInt(Me.imgBuscar(1).Tag) = 1 Then
        Text1(1).Text = RecuperaValor(CadenaSeleccion, 1)
        Text2(1).Text = RecuperaValor(CadenaSeleccion, 2)
    Else
        Text1(30).Text = RecuperaValor(CadenaSeleccion, 1)
        Text2(30).Text = RecuperaValor(CadenaSeleccion, 2)
    End If
End Sub

Private Sub frmZ_DatoSeleccionado(CadenaSeleccion As String)
    'ZONA
    Text1(27).Text = RecuperaValor(CadenaSeleccion, 1)
    Text2(27).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub imgAyuda_Click(Index As Integer)
Dim Ayuda As String

    'Sera las ayuda. Tampoco queiero la biblia, pero,
    'si un "pelin" de ayuda no me vendria mal a mi, imaginemos a el cliente final
    Select Case Index
    Case 0
        Ayuda = "Tipo portes: " & vbCrLf & "- Por albar�n/pedido.  A partir de un importe minimo y unos pesos por albar�n(pedido) se le sumaran los portes " & vbCrLf
        Ayuda = Ayuda & "- Por factura.   Portes tipo empresa HERBELCA. Entra fechaenvio, resto pedidos, importe minimo..."
    
    Case 1, 2
        Ayuda = "Departamento" & vbCrLf & "   - Entra en el proceso de facturaci�n. Podr� tener los valores: Departamento/Direcci�n/obra" & vbCrLf
        Ayuda = Ayuda & vbCrLf & "Lleva direciones envio." & vbCrLf & "   -No entra en proceso facturaci�n.  Si se habilita tendremos "
        Ayuda = Ayuda & vbCrLf & "        un campo mas en oferta/pedido/albaran donde indicar la direcci�n de envio."
        
    Case 3
        Ayuda = vbCrLf & "   Lleva contadores de facturas de venta separados:" & vbCrLf
        Ayuda = Ayuda & vbCrLf & " -Uno para facturas mostrador(FMO)"
        Ayuda = Ayuda & vbCrLf & " -Uno para facturas albar�n(FAV)"
    Case 4
        Ayuda = vbCrLf & " Cuando creamos un albar�n nuevo marcar�, o no, la opci�n ""facturar"" seg�n el valor de esta casilla"
    Case 5, 6
        Ayuda = vbCrLf & " Si marcamos el check de operaciones aseguradas, si supera el riesgo pondr� la siguientes sitauciones:"
        Ayuda = Ayuda & vbCrLf & " Prioriad cliente->"
        Ayuda = Ayuda & vbCrLf & "              Riesgo con bloqueo"
        Ayuda = Ayuda & vbCrLf & "                     La indicada en el campo ""situacion riesgo con bloqueo"""
        Ayuda = Ayuda & vbCrLf & "              Riesgo sin bloqueo"
        Ayuda = Ayuda & vbCrLf & "                      La indicada en el campo ""Situacion riesgo SIN bloqueo """
        
    Case 7
        Ayuda = vbCrLf & "Mirar� consumo del art�culo en los meses anteirores especificados en Meses de consumo 1 y 2"
        Ayuda = Ayuda & vbCrLf & " El listado dar� una cantidad orientativa para aprovisionarse del art�culo para los pr�ximos meses"
        Ayuda = Ayuda & vbCrLf & " especificados en los campos."
        
    Case 8, 9
        Ayuda = vbCrLf & "La contabilizaci�n de fras. proveedor intracomunitarias generar�, ademas de la propia factura proveedor,"
        Ayuda = Ayuda & vbCrLf & "dos facturas mas en contabilidad. Una de clientes y otra de proveedores a la cuenta especificada en 'cuenta extra'"
        Ayuda = Ayuda & vbCrLf & vbCrLf & "Las facturas 'extra' tendr�n como IVA el indicado en 'Intracom prov'"
        Ayuda = Ayuda & vbCrLf & "Entrar�n  para que se contabiliza si asi lo indica el parametro 'Permitir contabilizaci�n'"
        Ayuda = Ayuda & vbCrLf & "Para la factura de cliente cogera el contador y la serie del tipo de movimiento: 'CFI'"
        Ayuda = Ayuda & vbCrLf & "la de proveedor extra tendra el mismo numero que la de cliente"
    
    
        
        Ayuda = vbCrLf & "Abril 2016"
        Ayuda = Ayuda & vbCrLf & "La factura entrar� con el iva marcado en IVA Especiales-> intracomu"
        Ayuda = Ayuda & vbCrLf & "Pero este IVA no sumar� al total factura, y si que realizara los "
        Ayuda = Ayuda & vbCrLf & "apuntes a las cuentas de IVA cuando se contabilice. "
        
        
    
    
    
    
    Case 10
        Ayuda = vbCrLf & "Porcentaje.   Cuando el porcentaje de dias entre dias que hay desde fecha factura "
        Ayuda = Ayuda & vbCrLf & " hasta el vencimiento y dias desde la fecha de generacion de comisiones hasta el vencimiento"
        Ayuda = Ayuda & vbCrLf & " sea mayor que este valor se considerara como devuelto"
    
        Ayuda = Ayuda & vbCrLf & vbCrLf & "Si no se indica nada es de 75%"
    Case 11
        Ayuda = vbCrLf & "Cuando recepcionamos mercancia y va a buscar los pedidos de cliente"
        Ayuda = Ayuda & vbCrLf & "si busca en cualquiera de los almacenes o solo en el uno."
    Case 12
        Ayuda = vbCrLf & "Cuando ponemos a inventariar articulos, podremos elegir si "
        Ayuda = Ayuda & vbCrLf & "pone el stock a cero o deja el stock actual(>0)." & vbCrLf & vbCrLf
        Ayuda = Ayuda & vbCrLf & "Marcar ordenar por codigo. Tanto los listados como la introducci�n de existencias iran por ese orden"
        Ayuda = Ayuda & vbCrLf & "En caso de no marcar ir� por nombre/descripcion de art�culo."
    Case 13
        Ayuda = vbCrLf & "Cuando actualizamos el precio de un art�culo, si actualiza tambi�n, o no, el precio especial "
    Case 14
        Ayuda = vbCrLf & "Al cambiar el precio de compra, recalcular el margen del art�culo"
    End Select
    
    
    Ayuda = imgayuda(Index).ToolTipText & vbCrLf & String(45, "=") & vbCrLf & Ayuda
    MsgBox Ayuda, vbInformation
End Sub

Private Sub imgBuscar_Click(Index As Integer)
Dim i As Integer
    Screen.MousePointer = vbHourglass

    Select Case Index
        Case 15, 32, 53, 54, 64, 73, 75, 80, 81, 82 'cod. articulo
            Me.imgBuscar(1).Tag = Index
'            Set frmMtoArt = New frmAlmArticulosGr
'            frmMtoArt.DatosADevolverBusqueda = "@1@"
'            frmMtoArt.Show vbModal
            Set frmMtoArt = New frmBasico2
            AyudaArticulos frmMtoArt, Text1(Index)
            Set frmMtoArt = Nothing
            
        Case 24 'forma de pago
            If Modo = 4 Then TerminaBloquear
'            Set frmFP = New frmFacFormasPago
'            frmFP.DatosADevolverBusqueda = "0"
'            frmFP.Show vbModal
            Set frmFP = New frmBasico2
            AyudaFormasPago frmFP, Text1(Index)
            Set frmFP = Nothing
            If Modo = 4 Then
                If Not BLOQUEADesdeFormulario(Me) Then cmdCancelar_Click
            End If
    
        Case 25 'Codigo Actividad
            Set frmA = New frmFacActividades
            frmA.DatosADevolverBusqueda = "0"
            If Not IsNumeric(Text1(Index)) Then Text1(Index).Text = ""
            frmA.Show vbModal
            Set frmA = Nothing
            
        Case 26  'Cod. Envio
            Set frmE = New frmFacFormasEnvio
            frmE.DatosADevolverBusqueda = "0"
            If Not IsNumeric(Text1(Index)) Then Text1(Index).Text = ""
            frmE.Show vbModal
            Set frmE = Nothing
            
        Case 27  'Cod. Zona
            Set frmZ = New frmFacZonas
            frmZ.DatosADevolverBusqueda = "0"
            If Not IsNumeric(Text1(Index)) Then Text1(Index).Text = ""
            frmZ.Show vbModal
            Set frmZ = Nothing
            
        Case 28  'Cod. Ruta
            Set frmR = New frmFacRutas
            frmR.DatosADevolverBusqueda = "0"
            If Not IsNumeric(Text1(Index)) Then Text1(Index).Text = ""
            frmR.Show vbModal
            Set frmR = Nothing
            
        Case 4  'Cod. Forma de Pago
'            Set frmFP = New frmFacFormasPago
'            frmFP.DatosADevolverBusqueda = "0"
            If Not IsNumeric(Text1(Index)) Then Text1(Index).Text = ""
'            frmFP.Show vbModal
            Set frmFP = New frmBasico2
            AyudaFormasPago frmFP, Text1(Index)
            Set frmFP = Nothing
            
        Case 31 'C�digo de Agente
'            Set frmAc = New frmFacAgentesCom
'            frmAc.DatosADevolverBusqueda = "0"
            If Not IsNumeric(Text1(Index)) Then Text1(Index).Text = ""
            Set frmAc = New frmBasico2
            AyudaAgentesComerciales frmAc, Text1(Index), , True
            Set frmAc = Nothing
            
'            frmAc.Show vbModal
'            Set frmAc = Nothing
            
        Case 1, 30 'C�digo de Tarifa
            Me.imgBuscar(1).Tag = Index
            Set frmT = New frmFacTarifas
            frmT.DatosADevolverBusqueda = "0"
            If Not IsNumeric(Text1(Index)) Then Text1(Index).Text = ""
            frmT.Show vbModal
            Set frmT = Nothing
            
        Case 29, 65, 79 'C�digo de Situaci�n
            Me.imgBuscar(1).Tag = Index
            Set frmS = New frmFacSituaciones
            frmS.DatosADevolverBusqueda = "0"
            If Not IsNumeric(Text1(Index)) Then Text1(Index).Text = ""
            frmS.Show vbModal
            Set frmS = Nothing
            
            
        Case 33 To 42, 45 To 51, 66, 67 'Todos los ivas y la Cta de retencion, y cuenta aportacion TERMINAL
            CadenaDesdeOtroForm = ""
                        
            BuscaBuscaGRid2 (Index <> 40 And Index <> 42 And Index <> 66)
            If CadenaDesdeOtroForm <> "" Then
                Select Case Index
                Case 42
                    i = 9 'Para la cta aportacion
                Case 33 To 41, 45
                        i = 7
                Case 66, 67
                    i = 4
                Case Else
                    'IVAS antiguos
                    i = 13
                End Select
                Text1(Index + i).Text = RecuperaValor(CadenaDesdeOtroForm, 1)
                Text2(Index + i).Text = RecuperaValor(CadenaDesdeOtroForm, 2)
            End If
        
            
    End Select
    PonerFoco Text1(Index)
End Sub


Private Sub BuscaBuscaGRid2(EsIVa As Boolean)


    Screen.MousePointer = vbHourglass
'        Set frmB = New frmBuscaGrid
'        If EsIVa Then
'            'Busco IVAS
'            frmB.vCampos = "C�digo|tiposiva|codigiva|N||20�Denominacion|tiposiva|nombriva|T||70�"
'            frmB.vTabla = "tiposiva"
'            frmB.vTitulo = "Tipos de IVA"
'        Else
'
'            frmB.vCampos = "C�digo|cuentas|codmacta|T||20�Denominacion|cuentas|nommacta|T||70�"
'            frmB.vTabla = "cuentas"
'            frmB.vTitulo = "Cta contable"
'            frmB.vSQL = "apudirec = 'S'"
'
'        End If
'        frmB.vDevuelve = "0|1|"
'        frmB.vselElem = 1
'        frmB.vConexionGrid = conConta
'
'        frmB.vCargaFrame = False
'
'        frmB.Show vbModal
'        Set frmB = Nothing
    
    Set frmB = New frmBasico2
    If EsIVa Then
        AyudaTIvaContabilidad frmB
    Else
        AyudaCtasContables frmB
    End If
    Set frmB = Nothing


    Screen.MousePointer = vbDefault

End Sub


Private Sub mnModificar_Click()
    If BLOQUEADesdeFormulario(Me) Then BotonModificar
End Sub

Private Sub mnSalir_Click()
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub SSTab1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    ConseguirFoco Text1(Index), Modo
End Sub


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress (KeyAscii)
End Sub


Private Sub Text1_LostFocus(Index As Integer)
'    If Text1(Index).Text = "" Then Exit Sub

    'Quitar espacios en blanco
    Text1(Index).Text = Trim(Text1(Index).Text)
    
    Select Case Index
        Case 1 'tarifa de PVP
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "starif", "nomlista", "codlista", , "N")
            Else
                Text2(Index).Text = ""
            End If
        
        Case 2 'Km desplaz clientes
            PonerFormatoDecimal Text1(Index), 5 'Tipo 4: Decimal(8,4)
        Case 3, 76 'Km desplaz tecnicos
            PonerFormatoDecimal Text1(Index), 5 'Tipo 4: Decimal(8,4)
            
'        Case 6, 7 'Dias Reparacion con/sin mantenimiento
'            If Not EsNumerico(Text1(Index).Text) Then
'                Text1(Index).Text = ""
'                PonerFoco Text1(Index)
'            End If
        Case 14
            'PonerFocoBtn Me.cmdAceptar
            
        Case 15, 32, 53, 54, 64, 73, 75, 80, 81, 82  'cod. artic
            Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "sartic", "nomartic", "codartic", "Articulo")
            If Text2(Index).Text = "" Then Text1(Index).Text = ""
        Case 22 'n� conta
            'PonerFocoBtn Me.cmdAceptar
            
        Case 24 'FORMA DE PAGO
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "sforpa", "nomforpa")
            Else
                Text2(Index).Text = ""
            End If
            'PonerFocoBtn Me.cmdAceptar
            
            
        Case 25 To 31, 65, 79
            'Campos por defecto
            'Debug.Print Index & "-" & Text1(Index).Tag & ": " & Text1(Index).Text; ""
            If Text1(Index).Text = "" Then
                Text2(Index).Text = ""
            Else
                Select Case Index
                Case 25
                    Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "sactiv", "nomactiv", "codactiv")
                Case 26
                    Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "senvio", "nomenvio", "codenvio")
                Case 27
                    Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "szonas", "nomzonas", "Codzonas")
                Case 28
                    Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "srutas", "nomrutas", "codrutas")
                Case 29, 65, 79
                    Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "ssitua", "nomsitua", "codsitua")
                Case 30
                    Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "starif", "nomlista", "codlista")
                Case 31
                    Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "sagent", "nomagent", "codagent")
                End Select
            End If
            
        Case 40 To 46, 48, 52, 59 To 62, 71
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), conConta, "tiposiva", "nombriva", "codigiva")
            Else
                Text2(Index).Text = ""
            End If
        Case 47, 51, 70
            'Cta retencion y Cta aportacion al terminal   y para la cobntabili fras proveedor
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), conConta, "cuentas", "nommacta", "codmacta")
            Else
                Text2(Index).Text = ""
            End If
        Case 49, 74
            'pORCE RETENCION y procenta pago comisiones
            PonerFormatoDecimal Text1(Index), 4
        Case 50, 58, 66, 67
            PonerFormatoEntero Text1(Index)
            
        Case 55
            PonerFormatoDecimal Text1(Index), 5   'cuatro decimales
        Case 56, 68, 69
            PonerFormatoDecimal Text1(Index), 3
            
        
        Case 83
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "sclien", "nomclien", "codclien")
            Else
                Text2(Index).Text = ""
            End If
    End Select
End Sub


Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
    Select Case Index
        Case 6, 7, 16, 17, 18
            If Text1(Index).Text <> "" Then
                If Not EsNumerico(Text1(Index).Text) Then
                    Cancel = True
                    ConseguirFoco Text1(Index), Modo
                End If
            End If
    End Select
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
'        Case 1  'Anyadir
'            BotonAnyadir
        Case 1  'Modificar
            mnModificar_Click
        Case 4 'Salir
            mnSalir_Click
    End Select
End Sub


'Private Sub BotonAnyadir()
'    LimpiarCampos
'    PonerModo 3
'
'    Text1(0).Text = SugerirCodigoSiguienteStr("scryst", "codcryst")
'    PonerFoco Text1(0)
'End Sub


Private Sub BotonModificar()
    PonerModo 4
    
    Select Case Me.SSTab1.Tab
        Case 0:    PonerFoco Text1(4)
        Case 1: PonerFoco Text1(15)
        Case 2: PonerFoco Text1(8)
        Case 3: PonerFoco Text1(23)
    End Select
End Sub


Private Function DatosOk() As Boolean
Dim B As Boolean
Dim Cad As String

    On Error GoTo ErrOK

    DatosOk = False
    
    'Para que no de errores insesperados
    If Text1(6).Text = "" Then Text1(6).Text = "0"
    If Text1(7).Text = "" Then Text1(7).Text = "0"
    
    
    
    B = CompForm(Me, 1)
    
    '--- forma de pago de CHEQUE regalo
    'comprobar q el tipo de la forma de pago es EFECTIVO
    If B And Text1(24).Text <> "" Then
        If DevuelveDesdeBDNew(conAri, "sforpa", "tipforpa", "codforpa", Text1(24).Text, "N") <> "0" Then
            MsgBox "La forma de pago del cheque debe ser del tipo EFECTIVO", vbExclamation
            B = False
        End If
    End If
    
    If Text1(47).Text = "" Xor Text1(49).Text = "" Then
        MsgBox "Cta retenci�n o % retenci�n vacios", vbExclamation
        Exit Function
    End If
    
    
    If cboCreaTarifa.ListIndex < 0 Then
        MsgBox "Seleccion valor para crear tarifa", vbExclamation
        Exit Function
    End If
    
    
    
    
    
    If Text1(54).Text = "" Then
        If Me.cboPortes.ListIndex >= 0 Then
            MsgBox "Seleccione un articulo para los portes", vbExclamation
            cboPortes.ListIndex = -1
            PonerFoco Text1(54)
            B = False
        End If
    Else
        If cboPortes.ListIndex < 0 Then
            MsgBox "Seleccione un tipo de porte", vbExclamation
            PonerFocoCbo cboPortes
            B = False
        End If
    End If
    
    
    If Me.chkFraMost.Value = 1 Then
        'Ha seleccionado fra mostrador.  Vamos a comprobar que tien en codtipom el contador FMS de facturas de mostrador
        If DevuelveDesdeBDNew(conAri, "stipom", "codtipom", "codtipom", "FMO", "T") = "" Then
            MsgBox "Ha seleccionado contador separado para facturas mostrador y no existe la entrada en 'tipos de movimiento'", vbExclamation
            B = False
        End If
    End If
    
    
    
    'Si marca operaciones aseguradas debe marcar situbloq
    If B Then
        If vParamAplic.OperacionesAseguradas Then
            
            If Text1(65).Text = "" Then
                B = False
            Else
                If Text2(65).Text = "" Then B = False
            End If
            If Not B Then
                MsgBox "Si marca operaciones aseguradas debe marcar situacion de bloqueo", vbExclamation
                PonerFoco Text1(65)
            End If
        End If
    End If
            
            
   'Fras proveedor INTRACOM
   'Si indica la cuenta debe indicar el tipo de iva
'   If Me.Text1(70).Text <> "" Then
'        If Text1(71).Text = "" Then
'            MsgBox "Si indica la cta de contab. de fras proveedor 'extra' debe indicar el tipo de iva", vbExclamation
'            b = False
'        End If
'    Else
'        If Text1(71).Text <> "" Then
'            MsgBox "Ha indicado IVA para fra proveedor intracom, y no ha puesto la cuenta para las facturas 'extra'", vbExclamation
'            b = False
'        End If
'    End If
    
    
    
    If Text1(72).Text <> "" Then
        If Right(Text1(72).Text, 1) = "\" Then
            MsgBox "Carpeta integraci�n factura E no debe finalizar con \", vbExclamation
            B = False
        End If
    End If
    
    If vParamAplic.TieneTelefonia2 > 0 Then
        If Text1(32).Text = "" Then
            MsgBox "Indique el art�culo de telefon�a", vbExclamation
            B = False
        End If
    Else
        Text1(32).Text = ""
    End If
    
    
    
    'Si hay articulos inventariandose NO se puede cambiar la forma de inventariar
    
    
    If vParamAplic.IncializarStockEnInventario Xor Me.chkVarios(1).Value = 1 Then
        If Val(DevuelveDesdeBD(conAri, "count(*)", "salmac", "statusin", "1")) > 0 Then
            MsgBox "Articulos inventariandose. No puede modificar la forma de inventariar", vbExclamation
            B = False
        End If
    End If
    
    
    
    
    
    
    If Me.Text1(77).Text <> "" Then
        If Dir(Text1(77).Text, vbDirectory) = "" Then
            MsgBox "No existe carpeta facturas firmadas", vbExclamation
            B = False
        End If
    End If
    If Me.Text1(78).Text <> "" Then
        If Dir(Text1(78).Text, vbDirectory) = "" Then
            MsgBox "No existe carpeta albaranes firmados", vbExclamation
            B = False
        End If
    End If
    
    
    
    'Si tiene RIESGO tiene que tener las dos situaciones, aunque sea
    If vParamAplic.OperacionesAseguradas Then
        'Obligo a poner los dos situaciones de bloqueo aun sieondo las mismas
        Cad = "N"
        If Text1(65).Text <> "" And Text1(79).Text <> "" Then Cad = ""
        
        If Cad <> "" Then
            MsgBox "Las situaciones de bloqueo por superar riesgo son obligatorias", vbExclamation
            B = False
        End If
        
    End If
    
    
    
    DatosOk = B
    Exit Function
    
ErrOK:
    MuestraError Err.Number, "Comprobar datos", Err.Description
End Function


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub PonerBotonCabecera(B As Boolean)
    Me.cmdAceptar.visible = Not B
    Me.cmdCancelar.visible = Not B
    Me.cmdSalir.visible = B
    If B Then Me.lblIndicador.Caption = ""
End Sub


Private Sub PonerCampos()
On Error GoTo EPonerCampos

    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma Me, Data1
    'Portes
    If Text1(54).Text = "" Then cboPortes.ListIndex = -1
    
    'poner descripcion del articulo
    Text2(15).Text = PonerNombreDeCod(Text1(15), conAri, "sartic", "nomartic", "codartic", "Art�culos")
    Text2(32).Text = PonerNombreDeCod(Text1(32), conAri, "sartic", "nomartic", "codartic", "Art�culos")
    
    
    'poner descripcion de la forma de pago
    Text2(24).Text = PonerNombreDeCod(Text1(24), conAri, "sforpa", "nomforpa", "codforpa")
    
    'poner descripcion de la tarifa de PVP
    Text2(1).Text = PonerNombreDeCod(Text1(1), conAri, "starif", "nomlista", "codlista", , "N")
    
    
    For NumRegElim = 25 To 62
        If NumRegElim < 49 Or NumRegElim > 50 Then
            'If Text1(NumRegElim).Text <> "" Then Text1_LostFocus CInt(NumRegElim)
            Text1_LostFocus CInt(NumRegElim)
        End If
    Next NumRegElim

    Text1_LostFocus 64   'Recargo fianciero
    Text1_LostFocus 65   'Bloq riesgo
    
    Text1_LostFocus 70   'cta contbilizacion fras proveedores intracomunitarias
    Text1_LostFocus 71   'iva para dichas facturas
    
    Text1_LostFocus 73   'articulo separacion en ofertas
    
    Text1_LostFocus 75   'articulo separacion en ofertas
    
    Text1_LostFocus 79   'siuacion riesgo no bloqueante
    
    'Numeros de copia
    PonerValoresNumerosDeCopia
    
  
    
    
    
    
    NumRegElim = 0
    
    
    
    BloquearChecks Me, Modo
    
    Exit Sub
    
EPonerCampos:
    MuestraError Err.Number, "Poniendo Campos", Err.Description
    NumRegElim = 0
End Sub


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub


Private Sub LimpiarCampos()
    limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador.Caption = ""
    'Aqui va el especifico de cada form es
    '### a mano
End Sub


'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
Private Sub PonerModo(Kmodo As Byte)
Dim B As Boolean
   
    Modo = Kmodo
    PonerIndicador lblIndicador, Modo
      
    '------------------------------------------------------
    'Modo insertar o modificar
    B = (Kmodo >= 3) '-->Luego not b sera kmodo<3
    If cmdCancelar.visible Then
        cmdCancelar.Cancel = True
    Else
        cmdCancelar.Cancel = False
    End If
    PonerBotonCabecera Not B
       
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar adem�s limpia los campos Text1 y bloquea la clave primaria
    BloquearText1 Me, Modo
    
    'Bloquear los checkbox
    BloquearChecks Me, Modo

    'Bloquear el combobox
    B = Modo = 4
    Me.cboTipodtos.Enabled = B
    Me.cboOrdenDtos.Enabled = B
    Me.cboObsFactura.Enabled = B
    Me.cboCreaTarifa.Enabled = B
    Me.cboDpto.Enabled = B
    BloquearCmb Me.CboModAnalitica, Not B
    BloquearCmb Me.cboPortes, Not B
    
    
    'El frame del articulo de separacion en ofertas solo sera habilitado para el usuario ROOT
    FrameSepOfertas.Enabled = B And vUsu.Nivel = 0
    
    'FrameNumCopias.Enabled = B
    For NumRegElim = 0 To txtNumCopias.Count - 1
        BloquearTxt txtNumCopias(NumRegElim), Not B
    Next
    'Bloquear imagen de Busqueda
    Dim img As Image
    For Each img In Me.imgBuscar
        BloquearImg img, Not B
    Next
    
    'IVA intracom
    imgBuscar(66).visible = False
    imgBuscar(67).visible = False
    
    
    
'    BloquearImg Me.imgBuscar(1), (Modo <> 4)
'    BloquearImg Me.imgBuscar(15), (Modo <> 4)
'    For NumRegElim = 24 To 42
'        BloquearImg Me.imgBuscar(NumRegElim), (Modo <> 4)
'    Next NumRegElim
'    NumRegElim = 0
    
    PonerModoOpcionesMenu 'Activar opciones de menu seg�n el Modo
    PonerOpcionesMenu   'Activar opciones de menu seg�n nivel
                        'de permisos del usuario
End Sub




Private Sub PonerModoOpcionesMenu()
Dim B As Boolean
    B = (Modo = 3) Or (Modo = 4)
    Me.Toolbar1.Buttons(1).Enabled = Not B 'Modificar
    Me.mnModificar.Enabled = Not B
'    Me.Toolbar1.Buttons(2).Enabled = (Not b) 'Modificar
End Sub


Private Sub CargarComboTipoDtos()
'### Combo Tipo Descuento
'Cargaremos el combo, o bien desde una tabla o con valores fijos o como
'se quiera, la cuestion es cargarlo
' El estilo del combo debe de ser 2 - Dropdown List
' Si queremos que este ordenado, o lo ordenamos por la sentencia sql
' o marcamos la opcion sorted del combo
'0-Aditivo, 1-sobre Resto

    cboTipodtos.Clear
    cboTipodtos.AddItem "Aditivo"
    cboTipodtos.ItemData(cboTipodtos.NewIndex) = 0
    
    cboTipodtos.AddItem "sobre Resto"
    cboTipodtos.ItemData(cboTipodtos.NewIndex) = 1
End Sub


Private Sub CargarComboOrdenDtos()
'### Combo Tipo Descuento
'Cargaremos el combo, o bien desde una tabla o con valores fijos o como
'se quiera, la cuestion es cargarlo
' El estilo del combo debe de ser 2 - Dropdown List
' Si queremos que este ordenado, o lo ordenamos por la sentencia sql
' o marcamos la opcion sorted del combo
'0-Aditivo, 1-sobre Resto

    Me.cboOrdenDtos.Clear
    Me.cboOrdenDtos.AddItem "Familia/Marca"
    cboOrdenDtos.ItemData(cboOrdenDtos.NewIndex) = 0
    
    cboOrdenDtos.AddItem "Marca/Familia"
    cboOrdenDtos.ItemData(cboOrdenDtos.NewIndex) = 1
End Sub

Private Sub CargaComoboObsFactura()
'## Cuando contabilice, que valor pondra en el campo observaciones del
'   la factura, tanto cliente como de proveedores

    Me.cboObsFactura.Clear
    Me.cboObsFactura.AddItem "Sin observaciones"
    cboObsFactura.ItemData(cboObsFactura.NewIndex) = 0
    
    cboObsFactura.AddItem "N�mero factura"
    cboObsFactura.ItemData(cboObsFactura.NewIndex) = 1

    cboObsFactura.AddItem "Fecha integraci�n"
    cboObsFactura.ItemData(cboObsFactura.NewIndex) = 2

End Sub



' ---- [19/10/2009] [LAURA]: a�adir campo modo analitica
Private Sub CargarComboModoAnalitica()
'### Combo modo analitica
'Cargaremos el combo, o bien desde una tabla o con valores fijos o como
'se quiera, la cuestion es cargarlo
' El estilo del combo debe de ser 2 - Dropdown List
' Si queremos que este ordenado, o lo ordenamos por la sentencia sql
' o marcamos la opcion sorted del combo
'0-Trabajador, 1-Familia, 2-Proyecto

    Me.CboModAnalitica.Clear
    Me.CboModAnalitica.AddItem "Trabajador"
    CboModAnalitica.ItemData(CboModAnalitica.NewIndex) = 0
    
    CboModAnalitica.AddItem "Familia"
    CboModAnalitica.ItemData(CboModAnalitica.NewIndex) = 1
    
    CboModAnalitica.AddItem "Proyecto"
    CboModAnalitica.ItemData(CboModAnalitica.NewIndex) = 2
End Sub








Private Sub txtNumCopias_GotFocus(Index As Integer)
    ConseguirFoco txtNumCopias(Index), Modo
End Sub

Private Sub txtNumCopias_KeyPress(Index As Integer, KeyAscii As Integer)
  KEYpress KeyAscii
End Sub

Private Sub txtNumCopias_LostFocus(Index As Integer)
    txtNumCopias(Index).Text = Trim(txtNumCopias(Index).Text)
    If txtNumCopias(Index).Text = "" Then Exit Sub
    If Not PonerFormatoEntero(txtNumCopias(Index)) Then
        txtNumCopias(Index).Text = ""
    Else
        If Val(txtNumCopias(Index).Text) > 9 Then
            MsgBox "Numero maximo copias=9", vbExclamation
            txtNumCopias(Index).Text = ""
            PonerFoco txtNumCopias(Index)
        End If
    End If
    
End Sub


Private Sub PonerValoresNumerosDeCopia()
Dim C As String

    If Len(vParamAplic.NumeroCopias) <> Me.txtNumCopias.Count Then
        
        vParamAplic.NumeroCopias = Mid(vParamAplic.NumeroCopias & "1111111111", 1, txtNumCopias.Count)
    End If


    For NumRegElim = 1 To Len(vParamAplic.NumeroCopias)
        C = Mid(vParamAplic.NumeroCopias, NumRegElim, 1)
        If C = "" Then
            C = "1"
        Else
            If Not IsNumeric(C) Then C = "1"
        End If
        Me.txtNumCopias(NumRegElim - 1).Text = C
    Next NumRegElim
    
End Sub

Private Sub AsignarNumeroDeCopias()
Dim Aux As String
Dim i As Integer

    Aux = ""
    For i = 0 To Me.txtNumCopias.Count - 1
        If txtNumCopias(i).Text = "" Then
            Aux = Aux & "0"
        Else
            If Val(txtNumCopias(i).Text) = 0 Then
                Aux = Aux & "1"
            Else
                Aux = Aux & Mid(txtNumCopias(i).Text, 1, 1)
            End If
        End If
    Next i
    vParamAplic.NumeroCopias = Aux
End Sub
