VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmComHcoFacturSA 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Histórico de Facturas Proveedores"
   ClientHeight    =   8805
   ClientLeft      =   45
   ClientTop       =   4335
   ClientWidth     =   14400
   Icon            =   "frmComHcoFacturSA.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmComHcoFacturSA.frx":000C
   ScaleHeight     =   8805
   ScaleWidth      =   14400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   710
      Left            =   120
      TabIndex        =   101
      Top             =   385
      Width           =   13335
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   31
         Left            =   3960
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "Fecha Recepción|F|N|||scafpc|fecrecep|dd/mm/yyyy|N|"
         Top             =   315
         Width           =   1065
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   3
         Left            =   9090
         MaxLength       =   40
         TabIndex        =   5
         Tag             =   "Nombre Proveedor|T|N|||scafpc|nomprove||N|"
         Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwwwwwwwwww aq"
         Top             =   240
         Width           =   3990
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   2
         Left            =   8205
         MaxLength       =   6
         TabIndex        =   4
         Tag             =   "Cod. Proveedor|N|N|0|999999|scafpc|codprove|000000|S|"
         Text            =   "Text1"
         Top             =   240
         Width           =   780
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   1
         Left            =   2790
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Fecha Factura|F|N|||scafpc|fecfactu|dd/mm/yyyy|S|"
         Top             =   315
         Width           =   1065
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000013&
         Height          =   315
         Index           =   0
         Left            =   240
         MaxLength       =   20
         TabIndex        =   0
         Tag             =   "Nº Factura|T|N|||scafpc|numfactu||S|"
         Text            =   "Text1 7"
         Top             =   315
         Width           =   2205
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Contabilizado"
         Height          =   255
         Left            =   5640
         TabIndex        =   3
         Tag             =   "Contabilizado|N|N|0|1|scafpc|intconta||N|"
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "F. Recepción"
         Height          =   255
         Index           =   2
         Left            =   3960
         TabIndex        =   108
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Proveedor"
         Height          =   255
         Index           =   0
         Left            =   7080
         TabIndex        =   104
         Top             =   240
         Width           =   735
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   7935
         Picture         =   "frmComHcoFacturSA.frx":0A0E
         ToolTipText     =   "Buscar proveedor"
         Top             =   270
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "F. Factura"
         Height          =   255
         Index           =   29
         Left            =   2790
         TabIndex        =   103
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Factura"
         Height          =   255
         Index           =   28
         Left            =   240
         TabIndex        =   102
         Top             =   120
         Width           =   855
      End
   End
   Begin MSAdodcLib.Adodc Data2 
      Height          =   330
      Left            =   7320
      Top             =   3720
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
      Left            =   7560
      Top             =   3720
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
      Height          =   7200
      Left            =   120
      TabIndex        =   32
      Tag             =   "Fecha Oferta|F|N|||scapre|fecentre|dd/mm/yyyy|N|"
      Top             =   1095
      Width           =   14055
      _ExtentX        =   24791
      _ExtentY        =   12700
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
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
      TabPicture(0)   =   "frmComHcoFacturSA.frx":0B10
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "FrameObserva"
      Tab(0).Control(1)=   "FrameFactura"
      Tab(0).Control(2)=   "FrameCliente"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Albaranes"
      TabPicture(1)   =   "frmComHcoFacturSA.frx":0B2C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "imgBuscar(6)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label1(9)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "imgBuscar(5)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label1(21)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label1(6)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label1(18)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label1(8)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label1(13)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "imgBuscar(7)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label1(35)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label1(3)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Label1(46)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "imgBuscar2(11)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "imgBuscar2(10)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "imgBuscar2(9)"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Label1(22)"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Label1(23)"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "Label1(24)"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "Label1(34)"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "Label1(47)"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "Label1(40)"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "imgAmpliaci"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "DataGrid2"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "DataGrid1"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "txtAux(7)"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "txtAux(6)"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "txtAux(5)"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "txtAux(4)"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "Text3(1)"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "Text2(1)"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "Text3(0)"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "Text2(0)"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "Text3(2)"
      Tab(1).Control(32).Enabled=   0   'False
      Tab(1).Control(33)=   "Text3(3)"
      Tab(1).Control(33).Enabled=   0   'False
      Tab(1).Control(34)=   "txtAux(0)"
      Tab(1).Control(34).Enabled=   0   'False
      Tab(1).Control(35)=   "txtAux(1)"
      Tab(1).Control(35).Enabled=   0   'False
      Tab(1).Control(36)=   "txtAux(2)"
      Tab(1).Control(36).Enabled=   0   'False
      Tab(1).Control(37)=   "txtAux(3)"
      Tab(1).Control(37).Enabled=   0   'False
      Tab(1).Control(38)=   "txtAux3(0)"
      Tab(1).Control(38).Enabled=   0   'False
      Tab(1).Control(39)=   "txtAux3(1)"
      Tab(1).Control(39).Enabled=   0   'False
      Tab(1).Control(40)=   "txtAux(8)"
      Tab(1).Control(40).Enabled=   0   'False
      Tab(1).Control(41)=   "chkDocArchi"
      Tab(1).Control(41).Enabled=   0   'False
      Tab(1).Control(42)=   "Text3(9)"
      Tab(1).Control(42).Enabled=   0   'False
      Tab(1).Control(43)=   "Text2(2)"
      Tab(1).Control(43).Enabled=   0   'False
      Tab(1).Control(44)=   "Text3(10)"
      Tab(1).Control(44).Enabled=   0   'False
      Tab(1).Control(45)=   "Text2(16)"
      Tab(1).Control(45).Enabled=   0   'False
      Tab(1).Control(46)=   "Text2(17)"
      Tab(1).Control(46).Enabled=   0   'False
      Tab(1).Control(47)=   "txtAux2(8)"
      Tab(1).Control(47).Enabled=   0   'False
      Tab(1).Control(48)=   "txtDesc(11)"
      Tab(1).Control(48).Enabled=   0   'False
      Tab(1).Control(49)=   "txtDesc(10)"
      Tab(1).Control(49).Enabled=   0   'False
      Tab(1).Control(50)=   "txtDesc(9)"
      Tab(1).Control(50).Enabled=   0   'False
      Tab(1).Control(51)=   "txtAux(11)"
      Tab(1).Control(51).Enabled=   0   'False
      Tab(1).Control(52)=   "txtAux(10)"
      Tab(1).Control(52).Enabled=   0   'False
      Tab(1).Control(53)=   "txtAux(9)"
      Tab(1).Control(53).Enabled=   0   'False
      Tab(1).Control(54)=   "Text3(11)"
      Tab(1).Control(54).Enabled=   0   'False
      Tab(1).Control(55)=   "Text3(12)"
      Tab(1).Control(55).Enabled=   0   'False
      Tab(1).Control(56)=   "Text3(13)"
      Tab(1).Control(56).Enabled=   0   'False
      Tab(1).Control(57)=   "FrameEuler"
      Tab(1).Control(57).Enabled=   0   'False
      Tab(1).ControlCount=   58
      Begin VB.Frame FrameEuler 
         Height          =   1695
         Left            =   10680
         TabIndex        =   138
         Top             =   3240
         Width           =   3255
         Begin VB.TextBox txtAux 
            BackColor       =   &H80000013&
            Height          =   315
            Index           =   12
            Left            =   0
            MaxLength       =   3
            TabIndex        =   90
            Tag             =   "codtipom"
            Text            =   "Codtipom"
            Top             =   480
            Width           =   495
         End
         Begin VB.TextBox txtAux 
            BackColor       =   &H80000013&
            Height          =   315
            Index           =   13
            Left            =   480
            MaxLength       =   20
            TabIndex        =   91
            Tag             =   "numeroalb"
            Text            =   "numeroalb"
            Top             =   480
            Width           =   975
         End
         Begin VB.TextBox txtAux 
            BackColor       =   &H80000013&
            Height          =   315
            Index           =   14
            Left            =   1800
            MaxLength       =   20
            TabIndex        =   92
            Tag             =   "fechaalb"
            Text            =   "99/99/9999"
            Top             =   480
            Width           =   1095
         End
         Begin VB.TextBox txtDesc 
            BackColor       =   &H80000018&
            Height          =   675
            Index           =   0
            Left            =   0
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   139
            Text            =   "frmComHcoFacturSA.frx":0B48
            Top             =   840
            Width           =   3165
         End
         Begin VB.Image imgBuscar2 
            Height          =   240
            Index           =   12
            Left            =   2640
            Picture         =   "frmComHcoFacturSA.frx":0B58
            ToolTipText     =   "Buscar cliente"
            Top             =   240
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha alb"
            Height          =   195
            Index           =   44
            Left            =   1800
            TabIndex        =   141
            Top             =   240
            Width           =   840
         End
         Begin VB.Label Label1 
            Caption         =   "Albarán"
            Height          =   195
            Index           =   43
            Left            =   0
            TabIndex        =   140
            Top             =   240
            Width           =   840
         End
         Begin VB.Line Line3 
            X1              =   0
            X2              =   3000
            Y1              =   1560
            Y2              =   1560
         End
         Begin VB.Line Line2 
            X1              =   0
            X2              =   3240
            Y1              =   120
            Y2              =   120
         End
      End
      Begin VB.TextBox Text3 
         Height          =   280
         Index           =   13
         Left            =   10920
         MaxLength       =   80
         TabIndex        =   26
         Tag             =   "O|T|S|||scafpa|SReferencia||N|"
         Top             =   2160
         Width           =   2805
      End
      Begin VB.TextBox Text3 
         Height          =   280
         Index           =   12
         Left            =   8040
         MaxLength       =   80
         TabIndex        =   25
         Tag             =   "O|T|S|||scafpa|NReferencia||N|"
         Top             =   2160
         Width           =   2805
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   11
         Left            =   6720
         MaxLength       =   10
         TabIndex        =   24
         Tag             =   "Fec. entrega|F|S|||scafpa|fecentrega|dd/mm/yyyy||"
         Top             =   2160
         Width           =   1185
      End
      Begin VB.Frame FrameObserva 
         Enabled         =   0   'False
         ForeColor       =   &H00972E0B&
         Height          =   2055
         Left            =   -73680
         TabIndex        =   128
         Tag             =   "Observación 4|T|S|||scafac1|observa4||N|"
         Top             =   4920
         Width           =   11055
         Begin VB.TextBox Text3 
            Height          =   300
            Index           =   4
            Left            =   1080
            MaxLength       =   80
            TabIndex        =   133
            Tag             =   "Observación 1|T|S|||scafpa|observa1||N|"
            Text            =   "Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Te"
            Top             =   240
            Width           =   8940
         End
         Begin VB.TextBox Text3 
            Height          =   300
            Index           =   5
            Left            =   1080
            MaxLength       =   80
            TabIndex        =   132
            Tag             =   "Observación 2|T|S|||scafpa|observa2||N|"
            Text            =   "Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Te"
            Top             =   570
            Width           =   8940
         End
         Begin VB.TextBox Text3 
            Height          =   300
            Index           =   6
            Left            =   1080
            MaxLength       =   80
            TabIndex        =   131
            Tag             =   "Observación 3|T|S|||scafpa|observa3||N|"
            Text            =   "Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Te"
            Top             =   900
            Width           =   8940
         End
         Begin VB.TextBox Text3 
            Height          =   300
            Index           =   7
            Left            =   1080
            MaxLength       =   80
            TabIndex        =   130
            Tag             =   "Observación 4|T|S|||scafpa|observa4||N|"
            Text            =   "Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Te"
            Top             =   1230
            Width           =   8940
         End
         Begin VB.TextBox Text3 
            Height          =   300
            Index           =   8
            Left            =   1080
            MaxLength       =   80
            TabIndex        =   129
            Tag             =   "Observación 5|T|S|||scafpa|observa5||N|"
            Text            =   "Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Te"
            Top             =   1560
            Width           =   8940
         End
         Begin VB.Label Label1 
            Caption         =   "Observaciones"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   27
            Left            =   240
            TabIndex        =   134
            Top             =   0
            Width           =   1065
         End
      End
      Begin VB.TextBox txtAux 
         BackColor       =   &H80000013&
         Height          =   315
         Index           =   9
         Left            =   10680
         MaxLength       =   10
         TabIndex        =   89
         Tag             =   "cliente"
         Text            =   "cc"
         Top             =   2880
         Width           =   735
      End
      Begin VB.TextBox txtAux 
         BackColor       =   &H80000013&
         Height          =   315
         Index           =   10
         Left            =   10680
         MaxLength       =   10
         TabIndex        =   93
         Tag             =   "obra"
         Text            =   "cc"
         Top             =   3600
         Width           =   615
      End
      Begin VB.TextBox txtAux 
         BackColor       =   &H80000013&
         Height          =   315
         Index           =   11
         Left            =   10680
         MaxLength       =   20
         TabIndex        =   94
         Tag             =   "actuacion"
         Text            =   "cc"
         Top             =   4200
         Width           =   1215
      End
      Begin VB.TextBox txtDesc 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   9
         Left            =   11400
         Locked          =   -1  'True
         TabIndex        =   124
         Text            =   "ABCDKFJADKSFJAK"
         Top             =   2880
         Width           =   2445
      End
      Begin VB.TextBox txtDesc 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   10
         Left            =   11400
         Locked          =   -1  'True
         TabIndex        =   123
         Text            =   "ABCDKFJADKSFJAK"
         Top             =   3600
         Width           =   2445
      End
      Begin VB.TextBox txtDesc 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   11
         Left            =   12000
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   122
         Text            =   "ABCDKFJADKSFJAK"
         Top             =   4200
         Width           =   1845
      End
      Begin VB.TextBox txtAux2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   8
         Left            =   11280
         Locked          =   -1  'True
         MaxLength       =   60
         TabIndex        =   120
         Text            =   "nom ccoste"
         Top             =   5160
         Visible         =   0   'False
         Width           =   2565
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   17
         Left            =   11760
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   118
         Text            =   "ABCDKFJADKSFJAK"
         Top             =   4680
         Width           =   1725
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   1275
         Index           =   16
         Left            =   10680
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   96
         Text            =   "frmComHcoFacturSA.frx":10E2
         Top             =   5760
         Width           =   3165
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   10
         Left            =   6240
         MaxLength       =   30
         TabIndex        =   19
         Tag             =   "Forma de envio|N|S|0|9999|scafpa|codenvio|0000|N|"
         Text            =   "Text1"
         Top             =   1440
         Width           =   660
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   2
         Left            =   6900
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   115
         Text            =   "Text2"
         Top             =   1440
         Width           =   3765
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   9
         Left            =   12600
         MaxLength       =   10
         TabIndex        =   20
         Tag             =   "Fec. archiv|F|S|||scafpa|fecenvio|dd/mm/yyyy||"
         Top             =   480
         Width           =   1185
      End
      Begin VB.CheckBox chkDocArchi 
         Alignment       =   1  'Right Justify
         Caption         =   "Documento archivado"
         Height          =   330
         Left            =   11520
         TabIndex        =   21
         Top             =   960
         Width           =   2175
      End
      Begin VB.TextBox txtAux 
         BackColor       =   &H80000013&
         Height          =   315
         Index           =   8
         Left            =   10680
         MaxLength       =   4
         TabIndex        =   95
         Tag             =   "Centro coste|T|S|||slifac|codccost||N|"
         Text            =   "cc"
         Top             =   5160
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtAux3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   1
         Left            =   2040
         MaxLength       =   30
         TabIndex        =   100
         Tag             =   "Fecha Albaran|F|N|||scafpa|fechaalb|dd/mm/yyyy|N|"
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
         Index           =   0
         Left            =   1080
         MaxLength       =   15
         TabIndex        =   99
         Tag             =   "Nº Albaran|T|N|||scafpa|numalbar||N|"
         Text            =   "numalbar"
         Top             =   2160
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   3
         Left            =   3240
         MaxLength       =   12
         TabIndex        =   84
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
         Left            =   2160
         MaxLength       =   12
         TabIndex        =   83
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
         Left            =   1320
         MaxLength       =   12
         TabIndex        =   82
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
         Left            =   360
         MaxLength       =   12
         TabIndex        =   81
         Tag             =   "Almacen|N|N|0|999|slifac|codalmac|000|N|"
         Text            =   "almacen"
         Top             =   4320
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   3
         Left            =   4320
         TabIndex        =   22
         Tag             =   "Fecha Pedido|F|S|||scafpa|fecpedpr|dd/mm/yyyy|N|"
         Top             =   2160
         Width           =   1185
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   2
         Left            =   5640
         MaxLength       =   7
         TabIndex        =   23
         Tag             =   "Nº Pedido|N|S|||scafpa|numpedpr|0000000|N|"
         Text            =   "Text1 7"
         Top             =   2160
         Width           =   885
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   0
         Left            =   6900
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   78
         Text            =   "Text2"
         Top             =   480
         Width           =   3765
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   0
         Left            =   6240
         MaxLength       =   30
         TabIndex        =   17
         Tag             =   "Trabajador Albaran|N|S|0|9999|scafpa|codtrab2|0000|N|"
         Text            =   "Text1"
         Top             =   480
         Width           =   660
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   1
         Left            =   6900
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   77
         Text            =   "Text2"
         Top             =   960
         Width           =   3765
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   1
         Left            =   6240
         MaxLength       =   30
         TabIndex        =   18
         Tag             =   "Trabajador pedido|N|S|0|9999|scafpa|codtrab1|0000|N|"
         Text            =   "Text1"
         Top             =   960
         Width           =   660
      End
      Begin VB.Frame FrameFactura 
         Height          =   2220
         Left            =   -73680
         TabIndex        =   46
         Top             =   2400
         Width           =   11055
         Begin VB.Frame FrmRetencionSocios 
            Height          =   855
            Left            =   240
            TabIndex        =   109
            Top             =   1200
            Width           =   3615
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   33
               Left            =   1920
               MaxLength       =   15
               TabIndex        =   113
               Tag             =   "Imp Ret|N|S|||scafpc|impret|#,##0.00|N|"
               Text            =   "Text1 7"
               Top             =   360
               Width           =   1485
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   32
               Left            =   960
               MaxLength       =   15
               TabIndex        =   110
               Tag             =   "% Ret|N|S|||scafpc|porret|#0.00|N|"
               Text            =   "Text1 7"
               Top             =   360
               Width           =   525
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "%"
               Height          =   195
               Index           =   5
               Left            =   1560
               TabIndex        =   112
               Top             =   360
               Width           =   120
            End
            Begin VB.Label Label1 
               Caption         =   "Retención"
               Height          =   255
               Index           =   4
               Left            =   120
               TabIndex        =   111
               Top             =   360
               Width           =   855
            End
         End
         Begin VB.TextBox Text1 
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
            Index           =   30
            Left            =   9360
            MaxLength       =   15
            TabIndex        =   71
            Tag             =   "Total Factura|N|N|||scafpc|totalfac|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   1680
            Width           =   1605
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   29
            Left            =   7560
            MaxLength       =   15
            TabIndex        =   66
            Tag             =   "Importe IVA 3|N|S|||scafpc|impoiva3|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   1725
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   23
            Left            =   5040
            MaxLength       =   5
            TabIndex        =   65
            Tag             =   "% IVA 3|N|S|0|99.90|scafpc|porciva3|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   1725
            Width           =   525
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   20
            Left            =   4320
            MaxLength       =   3
            TabIndex        =   64
            Tag             =   "Cod. IVA 3|N|S|0|999|scafpc|tipoiva3|000|N|"
            Text            =   "Text1 7"
            Top             =   1725
            Width           =   525
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   26
            Left            =   5760
            MaxLength       =   15
            TabIndex        =   63
            Tag             =   "Base Imponible 3|N|S|||scafpc|baseiva3|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   1725
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   28
            Left            =   7560
            MaxLength       =   15
            TabIndex        =   62
            Tag             =   "Importe IVA 2|N|S|||scafpc|impoiva2|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   1395
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   22
            Left            =   5040
            MaxLength       =   5
            TabIndex        =   61
            Tag             =   "& IVA 2|N|S|0|99.90|scafpc|porciva2|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   1395
            Width           =   525
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   19
            Left            =   4320
            MaxLength       =   3
            TabIndex        =   60
            Tag             =   "Cod. IVA 2|N|S|0|999|scafpc|tipoiva2|000|N|"
            Text            =   "Text1 7"
            Top             =   1395
            Width           =   525
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   25
            Left            =   5760
            MaxLength       =   15
            TabIndex        =   59
            Tag             =   "Base Imponible 2 |N|S|||scafpc|baseiva2|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   1395
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   27
            Left            =   7560
            MaxLength       =   15
            TabIndex        =   58
            Tag             =   "Importe IVA 1|N|N|||scafpc|impoiva1|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   1080
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   21
            Left            =   5040
            MaxLength       =   5
            TabIndex        =   57
            Tag             =   "% IVA 1|N|S|0|99.90|scafpc|porciva1|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   1080
            Width           =   525
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   18
            Left            =   4320
            MaxLength       =   3
            TabIndex        =   56
            Tag             =   "Cod. IVA 1|N|S|0|999|scafpc|tipoiva1|000|N|"
            Text            =   "Text1 7"
            Top             =   1080
            Width           =   525
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   24
            Left            =   5760
            MaxLength       =   15
            TabIndex        =   55
            Tag             =   "Base Imponible 1|N|N|||scafpc|baseiva1|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   1080
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   17
            Left            =   5760
            MaxLength       =   15
            TabIndex        =   50
            Text            =   "Text1 7"
            Top             =   435
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   16
            Left            =   3960
            MaxLength       =   15
            TabIndex        =   49
            Tag             =   "Imp. Dto Gn|N|N|||scafpc|impgnral|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   435
            Width           =   1365
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   15
            Left            =   2160
            MaxLength       =   15
            TabIndex        =   48
            Tag             =   "Imp. Dto PP|N|N|||scafpc|impppago|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   435
            Width           =   1365
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   14
            Left            =   240
            MaxLength       =   15
            TabIndex        =   47
            Tag             =   "Imp.Bruto|N|N|||scafpc|brutofac|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   435
            Width           =   1485
         End
         Begin VB.Label Label1 
            Caption         =   "Cod. IVA"
            Height          =   255
            Index           =   42
            Left            =   4320
            TabIndex        =   98
            Top             =   840
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "% IVA"
            Height          =   255
            Index           =   41
            Left            =   5040
            TabIndex        =   97
            Top             =   870
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
            ForeColor       =   &H000000C0&
            Height          =   255
            Index           =   39
            Left            =   9330
            TabIndex        =   75
            Top             =   1440
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
            Left            =   9120
            TabIndex        =   74
            Top             =   1680
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
            TabIndex        =   73
            Top             =   2160
            Width           =   135
         End
         Begin VB.Line Line1 
            X1              =   4320
            X2              =   7320
            Y1              =   825
            Y2              =   825
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
            Index           =   37
            Left            =   7320
            TabIndex        =   72
            Top             =   960
            Width           =   135
         End
         Begin VB.Label Label1 
            Caption         =   "Imp. IVA"
            Height          =   255
            Index           =   33
            Left            =   7680
            TabIndex        =   70
            Top             =   870
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
            TabIndex        =   69
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
            Left            =   3720
            TabIndex        =   68
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
            Left            =   1920
            TabIndex        =   67
            Top             =   360
            Width           =   135
         End
         Begin VB.Label Label1 
            Caption         =   "Base Imponible"
            Height          =   255
            Index           =   14
            Left            =   5880
            TabIndex        =   54
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Imp. Dto Gn"
            Height          =   255
            Index           =   12
            Left            =   4080
            TabIndex        =   53
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Imp. Dto PP"
            Height          =   255
            Index           =   11
            Left            =   2280
            TabIndex        =   52
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Imp. Bruto"
            Height          =   255
            Index           =   10
            Left            =   240
            TabIndex        =   51
            Top             =   240
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
         Left            =   4080
         MaxLength       =   12
         TabIndex        =   85
         Tag             =   "Precio|N|N|0|999999.0000|slifac|precioar|###,##0.0000|N|"
         Text            =   "Precio"
         Top             =   4320
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   5
         Left            =   5760
         MaxLength       =   5
         TabIndex        =   86
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
         Index           =   6
         Left            =   6480
         MaxLength       =   30
         TabIndex        =   87
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
         Index           =   7
         Left            =   7080
         MaxLength       =   12
         TabIndex        =   88
         Tag             =   "Importe|N|N|0||slifac|importel|#,###,###,##0.00|N|"
         Text            =   "Importe"
         Top             =   4320
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Frame FrameCliente 
         Caption         =   "Datos Proveedor"
         ForeColor       =   &H00972E0B&
         Height          =   1875
         Left            =   -73680
         TabIndex        =   33
         Top             =   480
         Width           =   11055
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   13
            Left            =   7530
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   106
            Text            =   "Text2"
            Top             =   240
            Width           =   3285
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   13
            Left            =   6945
            MaxLength       =   4
            TabIndex        =   105
            Tag             =   "Trabajador|N|N|0|9999|scafpc|codtraba|0000|N|"
            Text            =   "Text1"
            Top             =   240
            Width           =   540
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   9
            Left            =   1125
            MaxLength       =   30
            TabIndex        =   11
            Tag             =   "Provincia|T|N|||scafpc|proprove||N|"
            Text            =   "Text1 Text1 Text1 Text1 Text22"
            Top             =   1350
            Width           =   2445
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   7
            Left            =   1125
            MaxLength       =   6
            TabIndex        =   9
            Tag             =   "CPostal|T|N|||scafpc|codpobla||N|"
            Text            =   "Text15"
            Top             =   990
            Width           =   630
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   8
            Left            =   1755
            MaxLength       =   30
            TabIndex        =   10
            Tag             =   "Población|T|N|||scafpc|pobprove||N|"
            Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwaq"
            Top             =   990
            Width           =   3405
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   5
            Left            =   3195
            MaxLength       =   20
            TabIndex        =   7
            Tag             =   "teléfono proveedor|T|S|||scafpc|telprove||N|"
            Text            =   "12345678911234567899"
            Top             =   285
            Width           =   1965
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   4
            Left            =   1125
            MaxLength       =   15
            TabIndex        =   6
            Tag             =   "NIF proveedor|T|N|||scafpc|nifprove||N|"
            Text            =   "123456789"
            Top             =   285
            Width           =   1110
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   10
            Left            =   6945
            MaxLength       =   3
            TabIndex        =   12
            Tag             =   "Forma de Pago|N|N|0|999|scafpc|codforpa|000|N|"
            Text            =   "Text1"
            Top             =   645
            Width           =   540
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   10
            Left            =   7530
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   35
            Text            =   "Text2"
            Top             =   645
            Width           =   3285
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   11
            Left            =   7560
            MaxLength       =   5
            TabIndex        =   13
            Tag             =   "Descuento P.Pago|N|N|0|99.90|scafpc|dtoppago|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   1350
            Width           =   525
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   12
            Left            =   8925
            MaxLength       =   5
            TabIndex        =   14
            Tag             =   "Descuento General|N|N|0|99.90|scafpc|dtognral|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   1350
            Width           =   525
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   6
            Left            =   1125
            MaxLength       =   35
            TabIndex        =   8
            Tag             =   "Domicilio|T|N|||scafpc|domprove||N|"
            Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwwwww aq"
            Top             =   645
            Width           =   4030
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   4
            Left            =   6660
            Picture         =   "frmComHcoFacturSA.frx":111F
            ToolTipText     =   "Buscar trabajador"
            Top             =   270
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Trabajador"
            Height          =   255
            Index           =   1
            Left            =   5730
            TabIndex        =   107
            Top             =   240
            Width           =   855
         End
         Begin VB.Image imgBuscar 
            Enabled         =   0   'False
            Height          =   240
            Index           =   2
            Left            =   855
            Picture         =   "frmComHcoFacturSA.frx":1221
            ToolTipText     =   "Buscar población"
            Top             =   1005
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Provincia"
            Height          =   255
            Index           =   17
            Left            =   120
            TabIndex        =   42
            Top             =   1350
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Población"
            Height          =   255
            Index           =   16
            Left            =   120
            TabIndex        =   41
            Top             =   990
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Teléfono"
            Height          =   255
            Index           =   19
            Left            =   2445
            TabIndex        =   40
            Top             =   285
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "NIF"
            Height          =   255
            Index           =   20
            Left            =   120
            TabIndex        =   39
            Top             =   285
            Width           =   615
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   1
            Left            =   855
            Picture         =   "frmComHcoFacturSA.frx":1323
            ToolTipText     =   "Buscar proveedor varios"
            Top             =   300
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Forma Pago"
            Height          =   255
            Index           =   15
            Left            =   5730
            TabIndex        =   38
            Top             =   645
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Dto. P.P"
            Height          =   255
            Index           =   25
            Left            =   6900
            TabIndex        =   37
            Top             =   1350
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Dto. Gral"
            Height          =   255
            Index           =   26
            Left            =   8235
            TabIndex        =   36
            Top             =   1350
            Width           =   735
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   3
            Left            =   6660
            Picture         =   "frmComHcoFacturSA.frx":1425
            ToolTipText     =   "Buscar forma de pago"
            Top             =   675
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Domicilio"
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   34
            Top             =   645
            Width           =   735
         End
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmComHcoFacturSA.frx":1527
         Height          =   4305
         Left            =   240
         TabIndex        =   45
         Top             =   2760
         Width           =   10335
         _ExtentX        =   18230
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
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "frmComHcoFacturSA.frx":153C
         Height          =   1995
         Left            =   240
         TabIndex        =   76
         Top             =   520
         Width           =   3600
         _ExtentX        =   6350
         _ExtentY        =   3519
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
      Begin VB.Image imgAmpliaci 
         Height          =   240
         Left            =   12000
         Picture         =   "frmComHcoFacturSA.frx":1551
         ToolTipText     =   "Buscar actuacion"
         Top             =   5520
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "S.referencia pedido"
         Height          =   255
         Index           =   40
         Left            =   10920
         TabIndex        =   137
         Top             =   1920
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "N.referencia pedido"
         Height          =   255
         Index           =   47
         Left            =   8040
         TabIndex        =   136
         Top             =   1920
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Fec. entrega"
         Height          =   255
         Index           =   34
         Left            =   6720
         TabIndex        =   135
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Obra"
         Height          =   255
         Index           =   24
         Left            =   10680
         TabIndex        =   127
         Top             =   3360
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente"
         Height          =   195
         Index           =   23
         Left            =   10680
         TabIndex        =   126
         Top             =   2640
         Width           =   480
      End
      Begin VB.Label Label1 
         Caption         =   "Actuacion"
         Height          =   255
         Index           =   22
         Left            =   10680
         TabIndex        =   125
         Top             =   3960
         Width           =   855
      End
      Begin VB.Image imgBuscar2 
         Height          =   240
         Index           =   9
         Left            =   11160
         Picture         =   "frmComHcoFacturSA.frx":1653
         ToolTipText     =   "Buscar cliente"
         Top             =   2640
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgBuscar2 
         Height          =   240
         Index           =   10
         Left            =   11160
         Picture         =   "frmComHcoFacturSA.frx":1755
         ToolTipText     =   "Buscar obra"
         Top             =   3360
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgBuscar2 
         Height          =   240
         Index           =   11
         Left            =   11520
         Picture         =   "frmComHcoFacturSA.frx":1857
         ToolTipText     =   "Buscar actuacion"
         Top             =   3960
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Centro coste"
         Height          =   255
         Index           =   46
         Left            =   10680
         TabIndex        =   121
         Top             =   4920
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Lote"
         Height          =   255
         Index           =   3
         Left            =   10680
         TabIndex        =   119
         Top             =   4680
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Ampliación Línea"
         Height          =   255
         Index           =   35
         Left            =   10680
         TabIndex        =   117
         Top             =   5520
         Width           =   1335
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   7
         Left            =   5880
         Picture         =   "frmComHcoFacturSA.frx":1959
         ToolTipText     =   "Buscar trabajador"
         Top             =   1455
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Forma de envio"
         Height          =   255
         Index           =   13
         Left            =   4320
         TabIndex        =   116
         Top             =   1440
         Width           =   1425
      End
      Begin VB.Label Label1 
         Caption         =   "Fec. archiv"
         Height          =   255
         Index           =   8
         Left            =   11520
         TabIndex        =   114
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Pedido"
         Height          =   255
         Index           =   18
         Left            =   4320
         TabIndex        =   80
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Pedido"
         Height          =   255
         Index           =   6
         Left            =   5760
         TabIndex        =   79
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Trabajador Albaran"
         Height          =   255
         Index           =   21
         Left            =   4320
         TabIndex        =   44
         Top             =   525
         Width           =   1455
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   5
         Left            =   5895
         Picture         =   "frmComHcoFacturSA.frx":1A5B
         ToolTipText     =   "Buscar trabajador"
         Top             =   480
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Trabajador Pedido"
         Height          =   255
         Index           =   9
         Left            =   4320
         TabIndex        =   43
         Top             =   975
         Width           =   1425
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   6
         Left            =   5880
         Picture         =   "frmComHcoFacturSA.frx":1B5D
         ToolTipText     =   "Buscar trabajador"
         Top             =   975
         Width           =   240
      End
   End
   Begin VB.Frame Frame1 
      Height          =   475
      Index           =   0
      Left            =   120
      TabIndex        =   28
      Top             =   8295
      Width           =   2175
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   240
         Left            =   240
         TabIndex        =   29
         Top             =   180
         Width           =   1755
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   13050
      TabIndex        =   16
      Top             =   8400
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   11880
      TabIndex        =   15
      Top             =   8400
      Width           =   1035
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   30
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
         NumButtons      =   18
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
            Object.Tag             =   "0"
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
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Imprimir Factura"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Último"
            ImageIndex      =   9
         EndProperty
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   8400
         TabIndex        =   31
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   13050
      TabIndex        =   27
      Top             =   8400
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
         Enabled         =   0   'False
         Shortcut        =   ^I
         Visible         =   0   'False
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
Attribute VB_Name = "frmComHcoFacturSA"
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
Public hcoFechaMovim As Date 'Fecha del Movim
Public hcoCodProve As Long 'Codigo de Proveedor    'DAVID.  Estaba integer

'========== VBLES PRIVADAS ====================
Private WithEvents frmB As frmBuscaGrid 'Form para busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmCP As frmCPostal 'Codigos Postales
Attribute frmCP.VB_VarHelpID = -1

Private WithEvents frmProv As frmComProveedores  'Form Mto Proveedores
Attribute frmProv.VB_VarHelpID = -1
Private WithEvents frmPV As frmComProveV  'Form Mto Proveedores Varios
Attribute frmPV.VB_VarHelpID = -1
Private WithEvents frmFP As frmFacFormasPago 'Form Mto Formas de Pago
Attribute frmFP.VB_VarHelpID = -1
Private WithEvents frmT As frmAdmTrabajadores  'Form Mto Trabajadores
Attribute frmT.VB_VarHelpID = -1
Private WithEvents frmFE As frmFacFormasEnvio
Attribute frmFE.VB_VarHelpID = -1
Private WithEvents frmCli As frmFacClientes
Attribute frmCli.VB_VarHelpID = -1
Private WithEvents frmAc As frmObraActua
Attribute frmAc.VB_VarHelpID = -1
Private WithEvents frmF As frmCal
Attribute frmF.VB_VarHelpID = -1



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

Dim TituloLinea As String 'Descripcion de la linea que estamos en Mantenimiento

Dim primeravez As Boolean

'Si el cliente mostrado es de Varios o No
Dim EsDeVarios As Boolean


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

Private BuscaChekc As String

Private Sub Check1_Click()
    If Modo = 1 Then CheckCadenaBusqueda Check1, BuscaChekc
End Sub

Private Sub Check1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub chkDocArchi_Click()
     If Modo = 1 Then CheckCadenaBusqueda chkDocArchi, BuscaChekc
End Sub
Private Sub chkDocArchi_KeyDown(KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub
Private Sub chkDocArchi_KeyPress(KeyAscii As Integer)
    If FrameObserva.visible Then
        PonerFoco Text3(4)
    Else
        PonerFocoBtn cmdAceptar
    End If
    
End Sub

Private Sub cmdAceptar_Click()

    Screen.MousePointer = vbHourglass
    On Error GoTo Error1

    Select Case Modo
        Case 1  'BUSQUEDA
            HacerBusqueda
            
        Case 4  'MODIFICAR
            If DatosOk Then
               If ModificarFactura Then
                    TerminaBloquear
'                    PosicionarData
               Else
                    '---- Laura 24/10/2006
                    'como no hemos modificado dejamos la fecha como estaba ya que ahora se puede modificar
                    Text1(1).Text = Me.Data1.Recordset!FecFactu
               End If
               PosicionarData
            End If
            
         Case 5 'InsertarModificar LINEAS
            'Actualizar el registro en la tabla de lineas 'slialb'
            If ModificaLineas = 1 Then 'INSERTAR lineas Albaran

            ElseIf ModificaLineas = 2 Then 'MODIFICAR lineas
                If ModificarLinea Then
                    TerminaBloquear
                    NumRegElim = Data2.Recordset.AbsolutePosition
                    
                    CargaGrid2 DataGrid1, Data2
                    ModificaLineas = 0
                    PonerBotonCabecera True
                    HabilitarLineas False
           
                    LLamaLineas Modo, 0, "DataGrid1"
                    PosicionarData
                    If (Not Data2.Recordset.EOF) And (Not Data2.Recordset.BOF) Then
                        SituarDataPosicion Data2, NumRegElim, ""
                    End If
                End If
                Me.DataGrid1.Enabled = True
                Me.DataGrid2.Enabled = True
            End If
    End Select
    Screen.MousePointer = vbDefault

Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
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
            HabilitarLineas False
            ModificaLineas = 0
            LLamaLineas Modo, 0, "DataGrid1"
            PonerForaGrid
            PonerBotonCabecera True
            Me.DataGrid1.Enabled = True
    End Select
End Sub

Private Sub HabilitarLineas(Habilitar As Boolean)
Dim J As Integer
Dim b As Boolean
            BloquearTxt Text2(16), Not Habilitar
            BloquearTxt Text2(17), Not Habilitar
            
            For J = 9 To 12
                BloquearTxt txtAux(J), Not Habilitar
                Me.imgBuscar2(J).visible = Habilitar
            Next
            If vEmpresa.TieneAnalitica Then
                b = Habilitar And Me.Check1.Value = 0
                BloquearTxt txtAux(8), Not b
            End If
            
            If InstalacionEsEulerTaxco Then
                b = CDate(Text1(31).Text) >= vEmpresa.FechaIni
                b = b And Habilitar
                For J = 12 To 14
                    BloquearTxt txtAux(J), Not Habilitar
                Next
            
            End If
End Sub
Private Sub BotonBuscar()
Dim anc As Single

    'Buscar
    If Modo <> 1 Then
        LimpiarCampos
        BuscaChekc = ""
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

    If chkVistaPrevia.Value = 1 Then
        MandaBusquedaPrevia ""
    Else
        LimpiarCampos

        LimpiarDataGrids
        CadenaConsulta = "Select scafpc.* "
        CadenaConsulta = CadenaConsulta & "from " & NombreTabla & Ordenacion
        

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
Dim DeVarios As Boolean

    'solo se puede modificar la factura si no esta contabilizada
    If FactContabilizada Then
        TerminaBloquear
        Exit Sub
    End If
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 4
    
    PonerFocoChk Me.Check1
        
    'Si es proveedor de Varios no se pueden modificar sus datos
    DeVarios = EsProveedorVarios(Text1(2).Text)
    BloquearDatosProve (DeVarios)
End Sub


Private Sub BotonModificarLinea()
'Modificar una linea
Dim vWhere As String
Dim anc As Single
Dim J As Byte
On Error GoTo EModificarLinea




    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Then
        TerminaBloquear
        Exit Sub '1= Insertar
    End If
    
    If Data2.Recordset.EOF Then
        TerminaBloquear
        Exit Sub
    End If
    
    vWhere = ObtenerWhereCP(False)
    vWhere = vWhere & " AND numalbar='" & data3.Recordset.Fields!Numalbar & "'"
    vWhere = vWhere & " and numlinea=" & Data2.Recordset!numlinea
    If Not BloqueaRegistro(NomTablaLineas, vWhere) Then
        TerminaBloquear
        Exit Sub
    End If

    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        J = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, J
        DataGrid1.Refresh
    End If
    
    anc = DataGrid1.Top
    If DataGrid1.Row < 0 Then
        anc = anc + 210
    Else
        anc = anc + DataGrid1.RowTop(DataGrid1.Row) + 20
    End If

    For J = 0 To 2
        txtAux(J).Text = DataGrid1.Columns(J + 5).Text
    Next J
    Text2(16).Text = DataGrid1.Columns(J + 5).Text
    For J = J + 1 To 8
        txtAux(J - 1).Text = DataGrid1.Columns(J + 5).Text
    Next J
    Text2(17).Text = DataGrid1.Columns(14).Text
    
    PonerForaGrid
    
    
    ModificaLineas = 2 'Modificar
    LLamaLineas ModificaLineas, anc, "DataGrid1"
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    Me.lblIndicador.Caption = "MODIFICAR LINEAS"
    PonerBotonCabecera False
    HabilitarLineas True
    
'    PonerFoco txtAux(4)
    'PonerFoco Text2(16)
    PonerFoco txtAux(9)
    Me.DataGrid1.Enabled = False

EModificarLinea:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub LLamaLineas(xModo As Byte, Optional alto As Single, Optional grid As String)
Dim jj As Integer
Dim b As Boolean

        If grid = "DataGrid2" Then
            DeseleccionaGrid Me.DataGrid2
            b = (xModo = 1)
             For jj = 0 To txtAux3.Count - 1
                txtAux3(jj).Height = DataGrid2.RowHeight
                txtAux3(jj).Top = alto
                txtAux3(jj).visible = b
            Next jj
        End If
'    End Select
End Sub


Private Sub BotonEliminar()
'Eliminar Registro de la Cabecera: Tabla de Facturas (scafac)
' y los registros correspondientes de las tablas cab. albaranes (scafac1)
' y las lineas de la factura (slifac)
Dim cad As String
'Dim NumPedElim As Long
On Error GoTo EEliminar

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
    'solo se puede modificar la factura si no esta contabilizada
    If FactContabilizada Then Exit Sub
    
    cad = "Cabecera de Facturas." & vbCrLf
    cad = cad & "-----------------------------------" & vbCrLf & vbCrLf
    cad = cad & "Va a eliminar la Factura:            "
    cad = cad & vbCrLf & "Proveedor:  " & Text1(2).Text & " - " & Text1(3).Text
    cad = cad & vbCrLf & "Nº Fact.:  " & Text1(0).Text
    cad = cad & vbCrLf & "Fecha:  " & Format(Text1(1).Text, "dd/mm/yyyy")

    cad = cad & vbCrLf & vbCrLf & " ¿Desea Eliminarla? "
    
    'Borramos
    If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        On Error GoTo EEliminar
        Screen.MousePointer = vbHourglass
        NumRegElim = Data1.Recordset.AbsolutePosition
       ' NumPedElim = Data1.Recordset.Fields(1).Value   MAAAAl. Ya que el nºfac prov es ALFANUMERICO
        
        If Not Eliminar Then
            Screen.MousePointer = vbDefault
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


'Private Sub cmdObserva_Click()
'    If Modo <> 2 And Modo <> 4 Then Exit Sub
'    If Me.FrameObserva.visible = False Then
'        Me.DataGrid1.visible = False
'        Me.FrameObserva.visible = True
'        Me.cmdObserva.Picture = frmPpal.imgListComun.ListImages(18).Picture
''        CargarICO Me.cmdObserva, "volver.ico"
'        Me.cmdObserva.ToolTipText = "volver lineas albaran"
'        BloqueaText3
'    Else
'        Me.DataGrid1.visible = True
'        Me.FrameObserva.visible = False
''        CargarICO Me.cmdObserva, "message.ico"
'        Me.cmdObserva.Picture = frmPpal.imgListComun.ListImages(41).Picture
'        Me.cmdObserva.ToolTipText = "ver observaciones albaran"
'    End If
'End Sub


Private Sub BloqueaText3()
Dim I As Byte
    'bloquear los Text3 que son las lineas de scafpa
    For I = 0 To 1
        BloquearTxt Text3(I), (Modo <> 4) And Modo <> 1
    Next I
    If Me.FrameObserva.visible Then
        For I = 4 To 8
            BloquearTxt Text3(I), (Modo <> 4)
        Next I
        
    End If
    'numpedpr, fecpedpr siempre bloqueados
    For I = 2 To 3
        'Feb 2011. Dejamos que busque por ellos
        BloquearTxt Text3(I), Modo <> 1
        'Feb 2015. Dejamos que busque por ellos
        BloquearTxt Text3(I + 10), Modo <> 1
    Next I
    BloquearTxt Text3(11), Modo <> 1
    Me.chkDocArchi.Enabled = Modo = 1 Or Modo = 4
    BloquearTxt Text3(9), Not chkDocArchi.Enabled
    BloquearTxt Text3(10), Not chkDocArchi.Enabled
End Sub


Private Sub cmdRegresar_Click()
'Este es el boton Cabecera
Dim cad As String

    'Quitar lineas y volver a la cabecera
    If Modo = 5 Then  'modo 5: Mantenimientos Lineas
        PonerModo 2
        DataGrid2.Enabled = True
        If Not Data1.Recordset.EOF Then _
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



    
    If Not Data2.Recordset.EOF Then
        If Not DGrid_CambiarFila(DataGrid1) Then Exit Sub
    End If
    
    If Not Data2.Recordset.EOF And ModificaLineas <> 1 Then PonerForaGrid
    

'
'
'    If Not Data2.Recordset.EOF Then
'        If ModificaLineas <> 1 Then
'            Text2(16).Text = DBLet(Data2.Recordset.Fields!Ampliaci)
'            Text2(17).Text = DBLet(Data2.Recordset.Fields!numlotes)
'        End If
'
'        '- centro de coste
'        ' ---- [20/10/2009] [LAURA]: añadir campo centro de coste familia
'        If vEmpresa.TieneAnalitica Then
'            Me.txtAux(8).Text = DBLet(Data2.Recordset!CodCCost, "T")
'            Me.txtAux2(8).Text = PonerNombreCCoste(Me.txtAux(8))
'        Else
'            txtAux2(8).Text = ""
'        End If
'
'
'    Else
'        Text2(16).Text = ""
'        Text2(17).Text = ""
'        txtAux2(8).Text = ""
'    End If
    Exit Sub

Error1:
    MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub



Private Sub DataGrid2_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim I As Byte

    If Not data3.Recordset.EOF Then
        Text3(0).Text = DBLet(data3.Recordset.Fields!codtrab2, "T")
        Text3_LostFocus (0)
        Text3(1).Text = DBLet(data3.Recordset.Fields!CodTrab1, "T")
        Text3_LostFocus (1)

        Text3(2).Text = DBLet(data3.Recordset.Fields!numpedpr, "N")
        If Text3(2).Text <> "0" Then
            FormateaCampo Text3(2)
        Else
            Text3(2).Text = ""
        End If
        Text3(3).Text = DBLet(data3.Recordset.Fields!fecpedpr, "F")
        
        'Observaciones
        Text3(4).Text = DBLet(data3.Recordset.Fields!observa1, "T")
        Text3(5).Text = DBLet(data3.Recordset.Fields!observa2, "T")
        Text3(6).Text = DBLet(data3.Recordset.Fields!observa3, "T")
        Text3(7).Text = DBLet(data3.Recordset.Fields!observa4, "T")
        Text3(8).Text = DBLet(data3.Recordset.Fields!observa5, "T")
        Text3(9).Text = DBLet(data3.Recordset.Fields!FecEnvio, "T")
        Text3(10).Text = DBLet(data3.Recordset.Fields!CodEnvio, "T")
        Text3_LostFocus (10)
        Me.chkDocArchi.Value = DBLet(data3.Recordset!docarchiv, "N")
                
        'Los campos nuevos NReferencia SReferencia fecentrega
        
        Text3(11).Text = DBLet(data3.Recordset.Fields!fecentrega, "F")
        Text3(12).Text = DBLet(data3.Recordset.Fields!NReferencia, "T")
        Text3(13).Text = DBLet(data3.Recordset.Fields!SReferencia, "T")
        
        'Datos de la tabla slipre
        CargaGrid DataGrid1, Data2, True
    Else
        For I = 0 To Text3.Count - 1
            Text3(I).Text = ""
        Next I
        Me.chkDocArchi.Value = 0
        Text2(0).Text = ""
        Text2(1).Text = ""
        'Datos de la tabla slipre
        CargaGrid DataGrid1, Data2, False
        
        
    End If
    PonerForaGrid
End Sub


Private Sub Form_activate()
    Screen.MousePointer = vbDefault
    
    'Viene de DblClick en frmAlmMovimArticulos y carga el form con los valores
    If hcoCodMovim <> "" And Not Data1.Recordset.EOF Then PonerCadenaBusqueda
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub Form_Load()
    'Icono del formulario
    Me.Icon = frmPpal.Icon

    ' ICONITOS DE LA BARRA
    btnPrimero = 15
    With Me.Toolbar1
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 1   'Botón Buscar
        .Buttons(2).Image = 2   'Botón Todos
        .Buttons(5).Image = 4   'Modificar
        .Buttons(6).Image = 5   'Borrar
        .Buttons(9).Image = 10 'Mto Lineas Ofertas
        .Buttons(10).Image = 16 'Imprimir
        .Buttons(12).Image = 15  'Salir
        .Buttons(btnPrimero).Image = 6  'Primero
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
        .Buttons(btnPrimero + 3).Image = 9 'Último
    End With
    Me.SSTab1.Tab = 0
      
      
    'El frame FrmRetencionSocios es el que tendra el % retencion aplicable a los socios
    'Solo sera visible SI (y solo si) en paretros.ctaretnecion <>""
    FrmRetencionSocios.visible = vParamAplic.CtaReten <> ""
      
    LimpiarCampos   'Limpia los campos TextBox
     
    'cargar icono de observaciones de los albaranes de factura
'    CargarICO Me.cmdObserva, "message.ico"
'    Me.cmdObserva.Picture = frmPpal.imgListComun.ListImages(41).Picture
'    Me.FrameObserva.visible = False
'    Me.cmdObserva.ToolTipText = "ver observaciones albaran"
    
    VieneDeBuscar = False
            
    '## A mano
    NombreTabla = "scafpc"
    NomTablaLineas = "slifpc" 'Tabla lineas de Facturacion
    Ordenacion = " ORDER BY scafpc.fecrecep desc ,scafpc.codprove, scafpc.numfactu "
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    'ASignamos un SQL al DATA1
    CadenaConsulta = "Select * from " & NombreTabla
    
    If hcoCodMovim <> "" Then
        'Se llama desde Dobleclick en frmAlmMovimArticulos
'        CadenaConsulta = CadenaConsulta & " WHERE numalbar='" & hcoCodMovim & "' AND fechaalb= """ & Format(hcoFechaMovim, "yyyy-mm-dd") & """"
'        CadenaConsulta = CadenaConsulta & " AND codprove=" & hcoCodProve
        CadenaConsulta = CadenaConsulta & ObtenerSelFactura
    Else
        CadenaConsulta = CadenaConsulta & " where FALSE "
    End If
    
    Data1.ConnectionString = conn
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
        
        
    Me.Label1(46).visible = (vEmpresa.TieneAnalitica)
    Me.txtAux2(8).visible = (vEmpresa.TieneAnalitica)
    txtAux(8).visible = (vEmpresa.TieneAnalitica)
        
        
    primeravez = InstalacionEsEulerTaxco
    FrameEuler.visible = primeravez
    FrameEuler.BorderStyle = 0
    Me.txtAux(10).visible = Not primeravez
    Me.txtAux(11).visible = Not primeravez
    Me.txtDesc(10).visible = Not primeravez
    Me.txtDesc(11).visible = Not primeravez
    Me.imgBuscar2(10).visible = Not primeravez
    Me.imgBuscar2(11).visible = Not primeravez
    Text2(17).visible = Not primeravez
        
    'Cargar el DataGrid de lineas de Revisiones inicialmente a nada DATA2
    primeravez = True
    
    If hcoCodMovim = "" Then
        If DatosADevolverBusqueda = "" Then
            PonerModo 0
        Else
            PonerModo 1
        End If
        'Poner los grid sin apuntar a nada
        LimpiarDataGrids
        primeravez = False
    Else
         PonerModo 0
    End If
End Sub


Private Sub LimpiarCampos()
On Error Resume Next
    limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador.Caption = ""
    'Aqui va el especifico de cada form es
    '### a mano
    Me.Check1.Value = 0
    Me.chkDocArchi.Value = 0
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
End Sub


Private Sub frmAc_DatoSeleccionado(CadenaSeleccion As String)
    BuscaChekc = CadenaSeleccion
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
Dim CadB As String
Dim Aux As String
      
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        CadB = ""
        Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
        CadB = Aux
        Aux = ValorDevueltoFormGrid(Text1(1), CadenaDevuelta, 2)
        CadB = CadB & " and " & Aux
        Aux = ValorDevueltoFormGrid(Text1(2), CadenaDevuelta, 3)
        CadB = CadB & " and " & Aux
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB
        CadenaConsulta = CadenaConsulta & " GROUP BY scafpc.codprove, scafpc.numfactu, scafpc.fecfactu "
        CadenaConsulta = CadenaConsulta & " " & Ordenacion
        PonerCadenaBusqueda
'        Text1(0).Text = RecuperaValor(CadenaDevuelta, 2)
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub frmCli_DatoSeleccionado(CadenaSeleccion As String)
    BuscaChekc = CadenaSeleccion
End Sub

Private Sub frmCP_DatoSeleccionado(CadenaSeleccion As String)
'Formulario Mantenimiento C. Postales
Dim Indice As Byte
Dim devuelve As String

        Indice = 7
        Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'CPostal
        Text1(Indice + 1).Text = ObtenerPoblacion(Text1(Indice).Text, devuelve)  'Poblacion
        'provincia
        Text1(Indice + 2).Text = devuelve
End Sub

Private Sub frmF_Selec(vFecha As Date)
    BuscaChekc = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmFE_DatoSeleccionado(CadenaSeleccion As String)
        Text3(10).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000") 'Cod env
        Text2(2).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom env
End Sub

Private Sub frmFP_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Formas de Pago
Dim Indice As Byte
    Indice = 10
    Text1(Indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000") 'Cod Forma Pago
    Text2(10).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Forma Pago
End Sub


Private Sub frmProv_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Proveedores
    Text1(2).Text = RecuperaValor(CadenaSeleccion, 1)  'Cod Prove
End Sub

Private Sub frmPV_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Proveedores Varios
Dim Indice As Byte

    Indice = 4
    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'NIF
    Text1(Indice - 1).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Clien
    PonerDatosProveVario (Text1(Indice).Text)
End Sub

Private Sub frmT_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Trabajadores
Dim Indice As Byte

    Indice = Val(Me.imgBuscar(4).Tag)
    If Indice = 4 Then
        Indice = Indice + 9
        Text1(Indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000") 'Cod Trabajador
        Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Trabajador
    Else
        Text3(Indice - 5).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000") 'Cod Trabajador
        Text2(Indice - 5).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Trabajador
    End If
End Sub


Private Sub imgAmpliaci_Click()
   Dim b As Boolean
   
    If Modo < 2 Then Exit Sub
    
    CadenaDesdeOtroForm = ""
    If Not Data2.Recordset.EOF Then CadenaDesdeOtroForm = DBLet(Data2.Recordset!Ampliaci, "T")
            
            
       b = False
    If Text2(16).Enabled Then
        If Not Text2(16).Locked Then b = True
    End If
    
    frmFacClienteObser.Modificar = b
    frmFacClienteObser.Text1 = CadenaDesdeOtroForm
    frmFacClienteObser.Show vbModal
    If b Then
        If RecuperaValor(CadenaDesdeOtroForm, 1) = "1" Then
            Text2(16).Text = Mid(CadenaDesdeOtroForm, 3)
        End If
    End If
    
End Sub

Private Sub imgBuscar_Click(index As Integer)
Dim Indice As Byte

    If Modo = 2 Or Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass

    Select Case index
        Case 0 'Cod. Proveedor
            PonerFoco Text1(2)
            Set frmProv = New frmComProveedores
            frmProv.DatosADevolverBusqueda = "0"
            frmProv.Show vbModal
            Set frmProv = Nothing
            Indice = 2
            PonerFoco Text1(Indice)
            
        Case 1 'NIF para proveedor de Varios
            Set frmPV = New frmComProveV
            frmPV.DatosADevolverBusqueda = "0"
            frmPV.Show vbModal
            Set frmPV = Nothing
            Indice = 7
            PonerFoco Text1(Indice)
            
        Case 2 'Cod. Postal
            Set frmCP = New frmCPostal
            frmCP.DatosADevolverBusqueda = "0"
            frmCP.Show vbModal
            Set frmCP = Nothing
            Indice = 7
            VieneDeBuscar = True
            PonerFoco Text1(Indice)
      
         Case 3 'Forma de Pago
            Indice = 10
            PonerFoco Text1(Indice)
            Set frmFP = New frmFacFormasPago
            frmFP.DatosADevolverBusqueda = "0"
            frmFP.Show vbModal
            Set frmFP = Nothing
            
        Case 4, 5, 6 'Realizada Por Trabajador (Pedido, Albaran, Preparador Material
            Me.imgBuscar(4).Tag = index
            Set frmT = New frmAdmTrabajadores
            frmT.DatosADevolverBusqueda = "0"
            frmT.Show vbModal
            Set frmT = Nothing
            If index = 4 Then
                PonerFoco Text1(13)
            Else
                PonerFoco Text3(index - 5)
            End If
        Case 7
            Set frmFE = New frmFacFormasEnvio
            frmFE.DatosADevolverBusqueda = "0|1|"
            frmFE.Show vbModal
            Set frmFE = Nothing
            
    End Select
    
    Screen.MousePointer = vbDefault
End Sub




Private Sub imgBuscar2_Click(index As Integer)
    If Modo <> 5 Then Exit Sub
    
    If index = 9 Then
            BuscaChekc = ""
            Set frmCli = New frmFacClientes
            frmCli.DatosADevolverBusqueda = "0"
            frmCli.Show vbModal
            Set frmCli = Nothing
            If BuscaChekc <> "" Then
                txtAux(9).Text = RecuperaValor(BuscaChekc, 1) 'Cod cliente
                Me.txtDesc(9).Text = RecuperaValor(BuscaChekc, 2) 'Nom clien
                BuscaChekc = ""
                PonerFoco txtAux(10)
            End If
            
    ElseIf index = 12 Then
        Set frmF = New frmCal
        BuscaChekc = ""
        frmF.Fecha = Now
        If Me.txtAux(14).Text <> "" Then frmF.Fecha = CDate(txtAux(14).Text)
        frmF.Show vbModal
        If BuscaChekc <> "" Then
             Me.txtAux(14).Text = BuscaChekc
             txtAux_LostFocus 14
        End If
    
    Else
        'Obra actuacion. Llamaraemos al mismo
        If Me.txtAux(9).Text = "" Then
            MsgBox "Indique el cliente", vbExclamation
            PonerFoco txtAux(9)
            
        Else
            BuscaChekc = ""
            Set frmAc = New frmObraActua
            frmAc.DatosADevolverBusqueda = txtAux(9).Text & "|" & txtAux(10).Text & "|"
            frmAc.Show vbModal
            Set frmAc = Nothing
            If BuscaChekc <> "" Then
                txtAux(11).Text = RecuperaValor(BuscaChekc, 3)
                txtDesc(11).Text = RecuperaValor(BuscaChekc, 4) & "  " & RecuperaValor(BuscaChekc, 5)
                
                If txtAux(10).Text = "" Then
                    txtAux(10).Text = RecuperaValor(BuscaChekc, 2)
                    PonerClieObraActuacion 10, False
                End If
                BuscaChekc = ""
            End If
        End If
    End If
    
    
End Sub

Private Sub mnBuscar_Click()
    Me.SSTab1.Tab = 0
    BotonBuscar
End Sub


Private Sub mnEliminar_Click()
    If Modo = 5 Then 'Eliminar lineas de Pedido
'         BotonEliminarLinea
    Else   'Eliminar Pedido
         BotonEliminar
    End If
End Sub


Private Sub mnImprimir_Click()
'Imprimir Factura
'    BotonImprimir (53) '53: Informe de Facturas
End Sub


Private Sub mnLineas_Click()
    If Data1.Recordset.EOF Then Exit Sub
    
    'Si son facturas de liquidacion de soccios NO dejamos modificarlas
    If Me.FrmRetencionSocios.visible Then
        If DBLet(Data1.Recordset!PorRet, "N") > 0 Then
            MsgBox "Factura liquidación socios. No puede modificarse", vbExclamation
            Exit Sub
        End If
    End If

    BotonMtoLineas 1, "Facturas"
End Sub


Private Sub mnModificar_Click()
    If Data1.Recordset.EOF Then Exit Sub
    
    'Si son facturas de liquidacion de soccios NO dejamos modificarlas
    If Me.FrmRetencionSocios.visible Then
        If DBLet(Data1.Recordset!PorRet, "N") > 0 Then
            MsgBox "Factura liquidación socios. No puede modificarse", vbExclamation
            Exit Sub
        End If
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
        'bloquea la tabla cabecera de factura: scafpc
        If BLOQUEADesdeFormulario(Me) Then
            'bloquear la tabla cabecera de albaranes de la factura: scafpa
            If BloqueaAlbxFac Then BotonModificar
        End If
    End If
End Sub


Private Function BloqueaAlbxFac() As Boolean
'bloquea todos los albaranes de la factura
Dim SQL As String
On Error GoTo EBloqueaAlb

    BloqueaAlbxFac = False
    'bloquear cabecera albaranes x factura
    SQL = "select * FROM scafpa "
    SQL = SQL & ObtenerWhereCP(True) & " FOR UPDATE"
    conn.Execute SQL, , adCmdText
    BloqueaAlbxFac = True

EBloqueaAlb:
    If Err.Number <> 0 Then BloqueaAlbxFac = False
End Function


Private Function BloqueaLineasFac() As Boolean
'bloquea TODAS las lineas de la factura
Dim SQL As String
    
    On Error GoTo EBloqueaLin

    BloqueaLineasFac = False
    
    'bloquear cabecera albaranes x factura
    SQL = "select * FROM slifpc "
    SQL = SQL & ObtenerWhereCP(True) & " FOR UPDATE"
    conn.Execute SQL, , adCmdText
    BloqueaLineasFac = True

EBloqueaLin:
    If Err.Number <> 0 Then BloqueaLineasFac = False
End Function


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


Private Sub Text1_Change(index As Integer)
    If index = 9 Then HaCambiadoCP = True 'Cod. Postal
End Sub

'### A mano
'Los metodos del text tendran que estar
'Los descomentamos cuando esten puestos ya los controles
Private Sub Text1_GotFocus(index As Integer)
    kCampo = index
    If index = 9 Then HaCambiadoCP = False 'CPostal
    ConseguirFoco Text1(index), Modo
End Sub


Private Sub Text1_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
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
        
    If Not PerderFocoGnral(Text1(index), Modo) Then Exit Sub
        
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
          
    'Si queremos hacer algo ..
    Select Case index
        Case 1, 31 'Fecha factura,fecha recepcion,doc archivado
            If Text1(index).Text <> "" Then PonerFormatoFecha Text1(index)
                
        Case 13 'Cod trabajador
            Text2(index).Text = PonerNombreDeCod(Text1(index), conAri, "straba", "nomtraba", "codtraba")
            If Modo > 2 Then
                If Text2(index).Text = "" Then
                    Text1(index).Text = ""
                    PonerFoco Text1(index)
                End If
            End If
        Case 2 'Cod. prove
            If Modo = 1 Then 'Modo=1 Busqueda
                Text1(index + 1).Text = PonerNombreDeCod(Text1(index), conAri, "sprove", "nomprove")
            Else
                PonerDatosProveedor (Text1(index).Text)
            End If
        
        Case 4 'NIF
            If Not EsDeVarios Then Exit Sub
            If Modo = 4 Then 'Modificar
                'si no se ha modificado el nif del cliente no hacer nada
                If Text1(4).Text = DBLet(Data1.Recordset!nifProve, "T") Then
                    Exit Sub
                End If
            End If
            PonerDatosProveVario (Text1(index).Text)
        
        Case 7 'Cod. Postal
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
        
        
        Case 10 'Forma de Pago
            If PonerFormatoEntero(Text1(index)) Then
                Text2(index).Text = PonerNombreDeCod(Text1(index), conAri, "sforpa", "nomforpa")
            Else
                Text2(index).Text = ""
            End If
                If Modo > 2 Then
                    If Text2(index).Text = "" Then
                        Text1(index).Text = ""
                        PonerFoco Text1(index)
                    End If
                End If
        Case 11, 12 'Descuentos
            If Modo = 4 Then 'comprobar que el dato a cambiado
                
                If Text1(index).Text = "" Then
                    Text1(index).Text = "0"
                Else
                    If Not PonerFormatoDecimal(Text1(index), 4) Then Text1(index).Text = "0"
                End If
            
                If index = 11 Then
                    
                    If CCur(Text1(index).Text) = CCur(Data1.Recordset!DtoPPago) Then Exit Sub
                ElseIf index = 12 Then
                    If CCur(Text1(index).Text) = CCur(Data1.Recordset!DtoGnral) Then Exit Sub
                End If
            End If
            
            If Modo = 3 Or Modo = 4 Then
                If Text1(index).Text <> "" Then PonerFormatoDecimal Text1(index), 4 'Tipo 4: Decimal(4,2)
                If Not ActualizarDatosFactura Then
                   If index = 11 Then Text1(index).Text = Data1.Recordset!DtoPPago
                   If index = 12 Then Text1(index).Text = Data1.Recordset!DtoGnral
                   FormateaCampo Text1(index)
                End If
            End If
    End Select
End Sub


Private Sub HacerBusqueda()
Dim CadB As String

    CadB = ObtenerBusqueda(Me, False, BuscaChekc)
    
    'De momento a mano
    If InStr(1, BuscaChekc, "chkDocArchi") > 0 Then
        'Ha clicado sobnre doarchiv
        If CadB <> "" Then CadB = CadB & " AND "
        CadB = CadB & " scafpa.docarchiv = " & Me.chkDocArchi.Value
    End If
    
    
    If chkVistaPrevia = 1 Then
        MandaBusquedaPrevia CadB
    ElseIf CadB <> "" Then
        'Se muestran en el mismo form
        CadenaConsulta = "select " & NombreTabla & ".* from " & NombreTabla & " LEFT OUTER JOIN scafpa ON " & NombreTabla & ".codprove=scafpa.codprove AND " & NombreTabla & ".numfactu=scafpa.numfactu AND " & NombreTabla & ".fecfactu=scafpa.fecfactu "
        CadenaConsulta = CadenaConsulta & " WHERE " & CadB
        CadenaConsulta = CadenaConsulta & " GROUP BY scafpc.codprove, scafpc.numfactu, scafpc.fecfactu "
        CadenaConsulta = CadenaConsulta & " " & Ordenacion
        
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
'        cad = cad & ParaGrid(Text1(1), 10, "Tipo Fac.")
        cad = cad & ParaGrid(Text1(0), 18, "Nº Factura")
        cad = cad & ParaGrid(Text1(1), 15, "Fecha Fac.")
        cad = cad & ParaGrid(Text1(2), 12, "Prov.")
        cad = cad & ParaGrid(Text1(3), 55, "Nombre Prov")
        tabla = NombreTabla & " LEFT OUTER JOIN scafpa ON " & NombreTabla & ".codprove=scafpa.codprove AND " & NombreTabla & ".numfactu=scafpa.numfactu AND " & NombreTabla & ".fecfactu=scafpa.fecfactu "
        Titulo = "Facturas"
        devuelve = "0|1|2|"
           
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
        frmB.vConexionGrid = conAri  'Conexión a BD: Ariges
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
        If Modo = 1 Then PonerFoco Text1(0)
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
'Carga el grid de los AlbaranesxFactura, es decir, la tabla scafpc de la factura seleccionada
On Error GoTo EPonerLineas

    Screen.MousePointer = vbHourglass
    
    'Datos de la tabla albaranes x factura: scafpa
    CargaGrid DataGrid2, data3, True
    PonerForaGrid
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
    
    'Para que no de error el chkarch
    PonerCamposForma Me, Data1
    
    'Poner la base imponible (impbruto - dtoppago - dtognral
    BrutoFac = CSng(Text1(14).Text) - CSng(Text1(15).Text) - CSng(Text1(16).Text)
    Text1(17).Text = Format(BrutoFac, FormatoImporte)
    
    'poner descripcion campos
    Text2(10).Text = PonerNombreDeCod(Text1(10), conAri, "sforpa", "nomforpa")
    Text2(13).Text = PonerNombreDeCod(Text1(13), conAri, "straba", "nomtraba", "codtraba")
    
    
    PonerCamposLineas 'Pone los datos de las tablas de lineas de Ofertas
    
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
    '## No tiene el boton modificar y no utiliza la funcion general
    ActualizarToolbar Modo, Kmodo
    
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
    'si estamos en modificar bloquea las compos que son clave primaria
    BloquearText1 Me, Modo
    '---- laura 24/10/2006: si ponemos las claves de la tabla con ON UPDATE CASCADE
    'podemos permitir modificar la fecha de la factura que es clave primaria
'    If Modo = 4 Then BloquearTxt Text1(1), False
    
    
    Me.Check1.Enabled = (Modo = 1 Or Modo = 3 Or Modo = 4)
    
    b = (Modo <> 1)
    'Campos Nº Factura bloqueado y en azul
    BloquearTxt Text1(0), b, True
    BloquearTxt Text1(3), b 'referencia
    
    FrmRetencionSocios.Enabled = Not b
    
    
    'Importes siempre bloqueados
    For I = 14 To 30
        BloquearTxt Text1(I), (Modo <> 1)
    Next I

    'Campo B.Imp y Imp. IVA siempre en azul
    Text1(17).BackColor = &HFFFFC0
    Text1(27).BackColor = &HFFFFC0
    Text1(28).BackColor = &HFFFFC0
    Text1(29).BackColor = &HFFFFC0
    Text1(30).BackColor = &HC0C0FF
    
    'bloquear los Text3 que son las lineas de scafac1
    BloqueaText3
    
    'Si no es modo lineas Boquear los TxtAux
    For I = 0 To txtAux.Count - 1
        BloquearTxt txtAux(I), True
    Next I
    
    
    'Si no es modo Busqueda Bloquear los TxtAux3 (son los txtaux de los albaranes de factura)
    For I = 0 To txtAux3.Count - 1
        BloquearTxt txtAux3(I), (Modo <> 1)
    Next I
    
    'ampliacion linea
    b = (Modo = 5) And Me.DataGrid1.visible
    'Modo Linea de Albaranes
    'Me.Label1(35).visible = b
    'Me.Label1(3).visible = b
    'Me.Text2(16).visible = b '
    'Me.Text2(17).visible = b
'    BloquearTxt Text2(16), (Modo <> 5) Or (Modo = 5 And ModificaLineas <> 1)
'    BloquearTxt Text2(17), (Modo <> 5) Or (Modo = 5 And ModificaLineas <> 1)


    ' ---- [20/10/2009] [LAURA] : añadir del centro de coste



    '---------------------------------------------
    b = (Modo <> 0 And Modo <> 2 And Modo <> 5)
    cmdCancelar.visible = b
    cmdAceptar.visible = b
    Me.chkDocArchi.Enabled = b
    
    For I = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(I).Enabled = b
    Next I
    Me.imgBuscar(0).Enabled = (Modo = 1)
    Me.imgBuscar(1).visible = False
                    
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
Dim b As Boolean
On Error GoTo EDatosOK

    DatosOk = False
    
    'Para que no den errores los 0's de los importes de dtos
    ComprobarDatosTotales
        
    'comprobamos datos OK de la tabla scafac
    b = CompForm(Me, 1) 'Comprobar formato datos ok de la cabecera: opcion=1
    If Not b Then Exit Function
       
    BuscaChekc = ""
    If Text3(0).Text = "" Or Text2(0).Text = "" Then BuscaChekc = BuscaChekc & "  -" & Label1(21).Caption & vbCrLf
    
    If Text3(1).Text = "" Xor Text2(1).Text = "" Then BuscaChekc = BuscaChekc & "  -" & Label1(9).Caption & vbCrLf
    If Text3(10).Text = "" Xor Text2(2).Text = "" Then BuscaChekc = BuscaChekc & "  -" & Label1(13).Caption & vbCrLf
    If BuscaChekc <> "" Then
        BuscaChekc = "Error en campos: " & vbCrLf & BuscaChekc
        MsgBox BuscaChekc, vbExclamation
        BuscaChekc = ""
        b = False
    End If
       
       
       
       
       
    DatosOk = b
    
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Function DatosOkLinea() As Boolean
Dim b As Boolean
Dim I As Byte

On Error GoTo EDatosOkLinea

    DatosOkLinea = False
    b = True

    For I = 0 To txtAux.Count - 1
        If I = 4 Or I = 5 Or I = 6 Then
            If txtAux(I).Text = "" Then
                MsgBox "El campo " & txtAux(I).Tag & " no puede ser nulo", vbExclamation
                b = False
                PonerFoco txtAux(I)
                Exit Function
            End If
        End If
    Next I
            
            
    
        'obra actuacion
    'B=true
    NumRegElim = 0
    
    BuscaChekc = "" 'para saber si ha puesto alguna de ellas
    For I = 9 To 11
       If txtAux(I).Text = "" Xor Me.txtDesc(I).Text = "" Then BuscaChekc = BuscaChekc & vbCrLf & txtAux(I).Tag
       If txtAux(I).Text <> "" Then NumRegElim = 1
       
    Next
    If BuscaChekc <> "" Then BuscaChekc = "Error en: " & vbCrLf & BuscaChekc
        
    
    'Si indica alguno, debe indicarlos todos
    If NumRegElim = 1 Then
        If BuscaChekc = "" Then
            'Ha puesto alguno de los campos(no deberia haber pasado)
            If txtAux(9).Text = "" Or txtAux(10).Text = "" Or txtAux(11).Text = "" Then
                BuscaChekc = "Faltan campos en la obra actuacion"
            Else
                'Compruebo que exista
                BuscaChekc = "codclien =" & txtAux(9).Text & " AND coddirec= " & txtAux(10).Text & " AND actuacion "
                BuscaChekc = DevuelveDesdeBDNew(conAri, "sactuaobra", "concat(fechaini,' ',if(observa is null,'',observa))", BuscaChekc, txtAux(11).Text, "T")
                If BuscaChekc = "" Then
                    BuscaChekc = "No existe la obra-actuacion"
                Else
                    BuscaChekc = ""
                End If
            End If
        End If
    End If
    If InstalacionEsEulerTaxco Then BuscaChekc = ""
    
    If BuscaChekc <> "" Then
        MsgBox BuscaChekc, vbExclamation
        PonerFoco txtAux(9)
        Exit Function
    End If
            
            
            
            
    'EULER y numero de albaran
    If InstalacionEsEulerTaxco Then
        
        BuscaChekc = "" 'para saber si ha puesto alguna de ellas
        For I = 12 To 14
            If txtAux(I).Text <> "" Then BuscaChekc = BuscaChekc & "1"
             
        Next
        
        If BuscaChekc <> "" Then
            If Len(BuscaChekc) <> 3 Then
                MsgBox "Falta identificar el albaran correctamente", vbExclamation
                Exit Function
            Else
                'LEN 3, vemaos si existe
                
                BuscaChekc = txtDesc(0).Text
                If BuscaChekc = "" Then BuscaChekc = "NO EXISTE"
                If BuscaChekc = "NO EXISTE" Then
                    BuscaChekc = "No existe el albaran indicado. ¿Continuar de igual modo?"
                    If MsgBox(BuscaChekc, vbQuestion + vbYesNo) = vbNo Then Exit Function
                End If
            End If
        End If
    End If
            
            
            
            
            
    DatosOkLinea = b
EDatosOkLinea:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Sub Text2_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
    If index = 16 And KeyCode = 40 Then 'campo Amliacion Linea y Flecha hacia abajo
        PonerFocoBtn Me.cmdAceptar
    Else
        KEYdown KeyCode
    End If
End Sub

Private Sub Text2_KeyPress(index As Integer, KeyAscii As Integer)
    If index = 17 And KeyAscii = 13 Then 'campo nº de lote y ENTER
        PonerFocoBtn Me.cmdAceptar
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub Text3_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub Text3_KeyPress(index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub Text3_LostFocus(index As Integer)

    If Modo = 1 Then
        If Not PerderFocoGnral(Text3(index), Modo) Then Exit Sub
    End If
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    If Modo = 4 Then
        'Modificar
        If index = 0 Or index = 1 Or index = 10 Then
            If Not PonerFormatoEntero(Text3(index)) Then Text3(index).Text = ""
        End If
    End If
    
    Select Case index
        Case 0, 1 'trabajador
            Text2(index).Text = PonerNombreDeCod(Text3(index), conAri, "straba", "nomtraba", "codtraba", "Cod. Trabajador", "N")
        Case 8 'observa 5
            PonerFocoBtn Me.cmdAceptar
            
        Case 9
            If Text3(index).Text <> "" Then PonerFormatoFecha Text3(index)
        Case 10
                    
            Text2(2).Text = PonerNombreDeCod(Text3(index), conAri, "senvio", "nomenvio", "codenvio", "Forma de envio", "N")

    End Select
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.index
        Case 1  'Buscar
            mnBuscar_Click
        Case 2  'Todos
            BotonVerTodos

        Case 5  'Modificar
            mnModificar_Click
        Case 6  'Borrar
            mnEliminar_Click
        Case 9  'Lineas
            mnLineas_Click
        Case 10 'Imprimir Albaran
            mnImprimir_Click
        Case 12    'Salir
            mnSalir_Click
        Case btnPrimero To btnPrimero + 3 'Flechas Desplazamiento
            Desplazamiento (Button.index - btnPrimero)
    End Select
End Sub


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim Cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, Cerrar
    If Cerrar Then Unload Me
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
    End If
End Sub
    
    
Private Function ModificarLinea() As Boolean
'Modifica un registro en la tabla de lineas de Albaran: slialb
Dim SQL As String
Dim vWhere As String
Dim b As Boolean

    On Error GoTo EModificarLinea

    ModificarLinea = False
    If Data2.Recordset.EOF Then Exit Function
    
    vWhere = ObtenerWhereCP(True)
    vWhere = vWhere & " AND numalbar='" & data3.Recordset.Fields!Numalbar & "'"
    vWhere = vWhere & " AND numlinea=" & Data2.Recordset.Fields!numlinea
    
    If DatosOkLinea() Then
        SQL = "UPDATE " & NomTablaLineas & " SET "
        SQL = SQL & " ampliaci=" & DBSet(Text2(16).Text, "T")
        'SQL = SQL & "precioar= " & DBSet(txtAux(4).Text, "N") & ", "
        'SQL = SQL & "dtoline1= " & DBSet(txtAux(5).Text, "N") & ", dtoline2= " & DBSet(txtAux(6).Text, "N") & ", "
        'SQL = SQL & "importel= " & DBSet(txtAux(7).Text, "N")
        SQL = SQL & ", numlotes=" & DBSet(Text2(17).Text, "T")
        SQL = SQL & ", codclien=" & DBSet(txtAux(9).Text, "N", "S")
        SQL = SQL & ", coddirec=" & DBSet(txtAux(10).Text, "N", "S")
        SQL = SQL & ", actuacion=" & DBSet(txtAux(11).Text, "T", "S")
        
        If vEmpresa.TieneAnalitica Then SQL = SQL & ", codccost=" & DBSet(txtAux(8).Text, "T", "S")
        
        'Julio 2015. Euler
        If InstalacionEsEulerTaxco Then
            'codtipomV numalbarV fechaalbV
            SQL = SQL & "," & "codtipomv=" & DBSet(txtAux(12).Text, "T", "S")
            SQL = SQL & "," & "numalbarV=" & DBSet(txtAux(13).Text, "N", "S")
            SQL = SQL & "," & "fechaalbV=" & DBSet(txtAux(14).Text, "F", "S")
        End If
        
        SQL = SQL & vWhere
    End If
    
    If SQL <> "" Then
        'actualizar la factura y vencimientos
        b = ModificarFactura(SQL)
        ModificarLinea = b
    End If
    
EModificarLinea:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Modificar Lineas Factura" & vbCrLf & Err.Description
        b = False
    End If
    ModificarLinea = b
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
    'Habilitar las opciones correctas del menu segun Modo
'    PonerModoOpcionesMenu (Modo)
'    PonerOpcionesMenu 'Habilitar las opciones correctas del menu segun Nivel de Acceso
    DataGrid2.Enabled = Not b
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub CargaGrid(ByRef vDataGrid As DataGrid, ByRef vData As Adodc, enlaza As Boolean)
Dim b As Boolean
Dim Opcion As Byte
Dim SQL As String

    On Error GoTo ECargaGrid

'    b = DataGrid1.Enabled

    If vDataGrid.Name = "DataGrid1" Then
        Opcion = 1
    Else
        Opcion = 2
    End If
    
    SQL = MontaSQLCarga(enlaza, Opcion)
    CargaGridGnral vDataGrid, vData, SQL, primeravez
    
    CargaGrid2 vDataGrid, vData
    vDataGrid.ScrollBars = dbgAutomatic
    
    b = (Modo = 5) And (ModificaLineas = 1 Or ModificaLineas = 2)
    vDataGrid.Enabled = Not b
    primeravez = False
    Exit Sub
    
ECargaGrid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub


Private Sub CargaGrid2(ByRef vDataGrid As DataGrid, ByRef vData As Adodc)
Dim tots As String

    On Error GoTo ECargaGrid
    
    vData.Refresh
    Select Case vDataGrid.Name
        Case "DataGrid1" 'Lineas de Albaran
            'SQL = "SELECT codtipom, numfactu, fecfactu, numalbar, numlinea,
            'codalmac, codartic, nomartic, ampliaci, cantidad, precioar, origpre, dtoline1, dtoline2, importel "
            tots = "N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;"
            tots = tots & "S|txtAux(0)|T|Alm.|440|;S|txtAux(1)|T|Artículo|1750|;S|txtAux(2)|T|Nombre Art.|3150|;"
            tots = tots & "N||||0|;S|txtAux(3)|T|Cantidad|950|;S|txtAux(4)|T|Precio|1100|;S|txtAux(5)|T|Dto.1|560|;S|txtAux(6)|T|Dto.2|560|;"
            tots = tots & "S|txtAux(7)|T|Importe|1150|;N||||0|;"
            'If vEmpresa.TieneAnalitica Then
            '
            '    tots = tots & "S|txtAux(8)|T|CCost|620|;"
            'Else
                tots = tots & "N||||0|;"
           ' End If
            
'            tots = tots & "N||||0|;"
'            tots = tots & "N||||0|;"
'            tots = tots & "N||||0|;"
                        
            For kCampo = 1 To 6
                tots = tots & "N||||0|;"
            Next
                        
            
            
            arregla tots, DataGrid1, Me
            DataGrid1.Columns(9).Alignment = dbgRight
            DataGrid1.Columns(10).Alignment = dbgRight
            DataGrid1.Columns(12).Alignment = dbgRight
            DataGrid1.Columns(13).Alignment = dbgRight
                       
         Case "DataGrid2" 'albaranes x articulo
            'SQL = "SELECT codtipom,numfactu,fecfactu,codtipoa,numalbar, fechaalb,"
            'numpedcl,fecpedcl,sementre,numofert,fecofert, referenc, codenvio,codtraba, codtrab1, codtrab2,observa1,observa2,observa3,observa4,observa5,fecnvio,docarchiv,  "
            tots = "N||||0|;N||||0|;N||||0|;"
            tots = tots & "S|txtAux3(0)|T|Albaran|1400|;S|txtAux3(1)|T|Fecha|1300|;"
            tots = tots & "N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;"
            tots = tots & "N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;"
            arregla tots, DataGrid2, Me
        
            DataGrid2_RowColChange 1, 1
    End Select
    
    vDataGrid.HoldFields
    Exit Sub
ECargaGrid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub


Private Sub txtAux_GotFocus(index As Integer)
    ConseguirFoco txtAux(index), Modo
End Sub

Private Sub TxtAux_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
End Sub


Private Sub txtAux_KeyPress(index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub txtAux_LostFocus(index As Integer)

    If Not PerderFocoGnralLineas(txtAux(index), ModificaLineas) Then Exit Sub
    
    Select Case index
        Case 4 'Precio
            If txtAux(index).Text <> "" Then
                PonerFormatoDecimal txtAux(index), 2 'Tipo 2: Decimal(10,4)
            End If
            
        Case 5, 6 'Descuentos
            PonerFormatoDecimal txtAux(index), 4 'Tipo 4: Decimal(4,2)
            If index = 6 Then PonerFoco Me.Text2(16)
            
        Case 7 'Importe Linea
            PonerFormatoDecimal txtAux(index), 1 'Tipo 3: Decimal(12,2)
        Case 8
            txtAux(index).Text = UCase(Trim(txtAux(index).Text))
            If txtAux(index).Text <> "" Then
                BuscaChekc = DevuelveDesdeBD(conConta, "nomccost", "cabccost", "codccost", txtAux(index).Text, "T")
                If BuscaChekc = "" Then
                    MsgBox "No existe el centro de coste: " & txtAux(index).Text, vbExclamation
                    PonerFoco txtAux(index)
                End If
            Else
                BuscaChekc = ""
            End If
            txtAux2(index).Text = BuscaChekc
            BuscaChekc = ""
        Case 9 To 11
            PonerClieObraActuacion index, False
            
            
        Case 12
            txtAux(index).Text = UCase(txtAux(index).Text)
            
        Case 13
            'NUmero
            If Not PonerFormatoEntero(txtAux(index)) Then txtAux(index).Text = ""
            
        Case 14
            'Fecha
            If txtAux(index).Text <> "" Then PonerFormatoFecha txtAux(index)
    End Select
    
    If (index = 3 Or index = 4 Or index = 6 Or index = 7) Then 'Cant., Precio, Dto1, Dto2
        If txtAux(1).Text = "" Then Exit Sub
        txtAux(7).Text = CalcularImporte(txtAux(3).Text, txtAux(4).Text, txtAux(5).Text, txtAux(6).Text, vParamAplic.TipoDtos)
        PonerFormatoDecimal txtAux(7), 1
    End If
    
    If index >= 12 And index <= 14 Then
        'Buscamos el albaran-factura
        PonerDatosAlbaranFacturaEuler
        
    End If

End Sub


Private Sub BotonMtoLineas(numTab As Integer, cad As String)
    Me.SSTab1.Tab = numTab
    
    If Me.DataGrid1.visible Then 'Lineas de Albaranes
        If Me.Data2.Recordset.RecordCount < 1 Then
            MsgBox "La factura no tiene lineas.", vbInformation
            Exit Sub
        End If
        TituloLinea = cad
        
        ModificaLineas = 0
        PonerModo 5
        PonerBotonCabecera True
    End If
End Sub


Private Function Eliminar() As Boolean
Dim SQL As String
Dim Cta As String
Dim b As Boolean

    On Error GoTo FinEliminar

        b = False
        Eliminar = False
        If Data1.Recordset.EOF Then Exit Function
        
        conn.BeginTrans
        ConnConta.BeginTrans
        
        'Eliminar en la tabla pagos de la Contabilidad: spagop
        '------------------------------------------------
        Cta = DevuelveDesdeBDNew(conAri, "sprove", "codmacta", "codprove", Text1(2).Text, "N")
        SQL = " ctaprove='" & Cta & "' AND numfactu='" & Data1.Recordset.Fields!Numfactu & "'"
        SQL = SQL & " AND fecfactu='" & Format(Data1.Recordset.Fields!FecFactu, FormatoFecha) & "'"
        ConnConta.Execute "Delete from spagop WHERE " & SQL
        b = True
        
        
        'Eliminar en tablas de factura de Ariges: scafpc, scafpa, slifpc
        '---------------------------------------------------------------
        If b Then
            SQL = " " & ObtenerWhereCP(True)
        
            'Lineas de facturas (slifpc)
            conn.Execute "Delete from " & NomTablaLineas & SQL
        
            'Lineas de cabeceras de albaranes de la factura
            conn.Execute "Delete from scafpa " & SQL
            
            'Cabecera de facturas (scafpc)
            conn.Execute "Delete from " & NombreTabla & SQL
        End If
        
        'Eliminar los movimientos generados por el albaran que genero la factura
        '-----------------------------------------------------------------------
        If b Then
        
        End If
        
'        b = True
        
FinEliminar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar Factura", Err.Description
        b = False
    End If
    If Not b Then
        conn.RollbackTrans
        ConnConta.RollbackTrans
    Else
        conn.CommitTrans
        ConnConta.CommitTrans
    End If
    Eliminar = b
End Function


Private Sub LimpiarDataGrids()
'Pone los Grids sin datos, apuntando a ningún registro
On Error Resume Next
    CargaGrid DataGrid2, data3, False
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
             If Modo <> 5 Then
                PonerModo 2
                PonerCampos
             End If
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
Dim SQL As String
On Error Resume Next
    SQL = "codprove= " & Text1(2).Text & " and numfactu= '" & Text1(0).Text & "' and fecfactu='" & Format(Text1(1).Text, FormatoFecha) & "' "
    If conWhere Then SQL = " WHERE " & SQL
    ObtenerWhereCP = SQL
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
Dim SQL As String
    
    If Opcion = 1 Then
        SQL = "SELECT codprove, numfactu, fecfactu, numalbar, numlinea, codalmac, codartic, nomartic, ampliaci, cantidad, precioar, dtoline1, dtoline2, importel,numlotes,codccost "
        'Sept 2012
        SQL = SQL & ",codclien,coddirec,actuacion"
        
        'Julio 2015
        SQL = SQL & ",codtipomV,numalbarV,fechaalbV"
        
        SQL = SQL & " FROM slifpc " 'lineas de factura
    ElseIf Opcion = 2 Then
        SQL = "SELECT codprove,numfactu,fecfactu,numalbar, fechaalb, numpedpr,fecpedpr,codtrab1, codtrab2,observa1,observa2,observa3,observa4,observa5,fecenvio,docarchiv,codenvio,NReferencia ,SReferencia ,fecentrega  "
        SQL = SQL & " FROM scafpa " 'cabeceras albaranes de la factura
    End If
    
    If enlaza Then
        SQL = SQL & " " & ObtenerWhereCP(True)
        'lineas factura proveedor
        If Opcion = 1 Then SQL = SQL & " AND numalbar=" & DBSet(data3.Recordset.Fields!Numalbar, "T")
    Else
        SQL = SQL & " WHERE numfactu = -1"
    End If
    SQL = SQL & " ORDER BY codprove, numfactu, fecfactu,numalbar "
    If Opcion = 1 Then SQL = SQL & ", numlinea "
    MontaSQLCarga = SQL
End Function


Private Sub PonerModoOpcionesMenu(Modo As Byte)
'Activas unas Opciones de Menu y Toolbar según el modo en que estemos
Dim b As Boolean

        b = (Modo = 2) Or (Modo = 5 And ModificaLineas = 0)
        'Modificar
        Toolbar1.Buttons(5).Enabled = b
        Me.mnModificar.Enabled = b
        'eliminar
        Toolbar1.Buttons(6).Enabled = (Modo = 2)
        Me.mnEliminar.Enabled = (Modo = 2)
            
        b = (Modo = 2)
        'Mantenimiento lineas
        Toolbar1.Buttons(9).Enabled = b
        Me.mnLineas.Enabled = b
        'Imprimir
'        Toolbar1.Buttons(10).Enabled = b
'        Me.mnImprimir.Enabled = b
        
        b = (Modo >= 3) Or Modo = 1
        'Buscar
        Toolbar1.Buttons(1).Enabled = Not b
        Me.mnBuscar.Enabled = Not b
        'Ver Todos
        Toolbar1.Buttons(2).Enabled = Not b
        Me.mnVerTodos.Enabled = Not b
End Sub


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
           
            EsDeVarios = vProve.DeVarios
            BloquearDatosProve (EsDeVarios)
        
            If Modo = 4 And EsDeVarios Then 'Modificar
                'si no se ha modificado el proveedor no hacer nada
                If CLng(Text1(2).Text) = CLng(Data1.Recordset!Codprove) Then
                    Set vProve = Nothing
                    Exit Sub
                End If
            End If
        
            Text1(2).Text = vProve.Codigo
            FormateaCampo Text1(2)
            
            If (Modo = 3) Or (Modo = 4) Then
                Text1(3).Text = vProve.Nombre  'Nom prove
                Text1(6).Text = vProve.Domicilio
                Text1(7).Text = vProve.CPostal
                Text1(8).Text = vProve.Poblacion
                Text1(9).Text = vProve.Provincia
                Text1(4).Text = vProve.NIF
                Text1(5).Text = DBLet(vProve.TfnoAdmon, "T")
            End If
            
            Observaciones = DBLet(vProve.Observaciones)
            If Observaciones <> "" Then
                MsgBox Observaciones, vbInformation, "Observaciones del proveedor"
            End If
        End If
    Else
        LimpiarDatosProve
    End If
    Set vProve = Nothing

EPonerDatos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poner Datos Proveedor", Err.Description
End Sub


Private Sub PonerDatosProveVario(nifProve As String)
'Poner el los campos Text el valor del proveedor
Dim vProve As CProveedor
Dim b As Boolean
   
    If nifProve = "" Then Exit Sub
   
    Set vProve = New CProveedor
    b = vProve.LeerDatosProveVario(nifProve)
    If b Then
        Text1(3).Text = vProve.Nombre   'Nom proveedor
        Text1(6).Text = vProve.Domicilio
        Text1(7).Text = vProve.CPostal
        Text1(8).Text = vProve.Poblacion
        Text1(9).Text = vProve.Provincia
        Text1(5).Text = DBLet(vProve.TfnoAdmon, "T")
    End If
    Set vProve = Nothing
End Sub


Private Sub LimpiarDatosProve()
Dim I As Byte

    For I = 3 To 9
        Text1(I).Text = ""
    Next I
End Sub
   

Private Function ModificaAlbxFac() As Boolean
Dim SQL As String
Dim b As Boolean
On Error GoTo EModificaAlb
    
    ModificaAlbxFac = False
    If data3.Recordset.EOF Then Exit Function
    
    'comprobar datos OK de la scafac1
    b = CompForm(Me, 2) 'Comprobar formato datos ok de la cabecera alb: opcion=2
    If Not b Then Exit Function
    
    'Comprobaremos de albaranes factura
    BuscaChekc = ""
    If Text3(0).Text = "" Or Text2(0).Text = "" Then BuscaChekc = BuscaChekc & "  -" & Label1(21).Caption & vbCrLf
    If Text3(1).Text = "" Xor Text2(1).Text = "" Then BuscaChekc = BuscaChekc & "  -" & Label1(9).Caption & vbCrLf
    If Text3(10).Text = "" Xor Text2(2).Text = "" Then BuscaChekc = BuscaChekc & "  -" & Label1(13).Caption & vbCrLf
    If BuscaChekc <> "" Then
        BuscaChekc = "Error en campos: " & vbCrLf & BuscaChekc
        Err.Raise 513, , BuscaChekc
        
        'MsgBox BuscaChekc, vbExclamation
        'BuscaChekc = ""
        'Exit Function
    End If
    
    
    
    
    
    
    
    SQL = "UPDATE scafpa SET codtrab2=" & DBSet(Text3(0).Text, "N", "S") & ", "
    SQL = SQL & "codtrab1=" & DBSet(Text3(1).Text, "N", "S")
    If Me.FrameObserva.visible Then
        SQL = SQL & ", observa1=" & DBSet(Text3(4).Text, "T")
        SQL = SQL & ", observa2=" & DBSet(Text3(5).Text, "T")
        SQL = SQL & ", observa3=" & DBSet(Text3(6).Text, "T")
        SQL = SQL & ", observa4=" & DBSet(Text3(7).Text, "T")
        SQL = SQL & ", observa5=" & DBSet(Text3(8).Text, "T")
    End If
    SQL = SQL & ", fecenvio=" & DBSet(Text3(9).Text, "F", "S")
    SQL = SQL & ", docarchiv=" & DBSet(Me.chkDocArchi.Value, "N")
    SQL = SQL & ", codenvio=" & DBSet(Text3(10).Text, "N", "S")
    SQL = SQL & ObtenerWhereCP(True)
    'Antes oct 2011
    'SQL = SQL & " AND numalbar=" & Data3.Recordset.Fields!NumAlbar
    SQL = SQL & " AND numalbar=" & DBSet(data3.Recordset.Fields!Numalbar, "T")
    conn.Execute SQL
    ModificaAlbxFac = True
    
EModificaAlb:
If Err.Number <> 0 Then MuestraError Err.Number, "Modificar Albaranes de factura", Err.Description
End Function



Private Function ModificarFactura(Optional sqlLineas As String) As Boolean
'si se ha modificado la linea de slifac, añadir a la transaccion la modificación de la linea y recalcular
Dim bol As Boolean
Dim MenError As String
Dim SQL As String
Dim vFactu As CFacturaCom
Dim TocarEnTesoreria As Boolean
Dim Impag As String
On Error GoTo EModFact

    bol = False
    conn.BeginTrans
    ConnConta.BeginTrans
    
    If sqlLineas <> "" Then
        'actualizar el importe de la linea modificada
        MenError = "Modificando lineas de Factura."
        conn.Execute sqlLineas
    End If
    
    'recalcular las bases imponibles x IVA
    MenError = "Recalcular importes IVA"
    bol = ActualizarDatosFactura
    
    If bol Then
        'modificamos la scafpc
        MenError = "Modificando cabecera de factura"
        bol = ModificaDesdeFormulario(Me, 1)
        
        If bol Then
            'Si es proveedor de varios actualizar datos proveedor en tabla:sprvar
            MenError = "Modificando datos proveedor varios"
            bol = ActualizarProveVarios(Text1(2).Text, Text1(4).Text)
        End If
        
        If bol Then
            MenError = "Modificando albaranes de factura"
            'modificar la tabla: scafpa
            bol = ModificaAlbxFac
            
            If bol Then 'si se ha modificado la factura
                MenError = "Actualizando en Tesoreria"
                'y eliminar de tesoreria conta.spagop los registros de la factura
                
                'antes de Eliminar en las tablas de la Contabilidad
                Set vFactu = New CFacturaCom
                bol = vFactu.LeerDatos3(Text1(2).Text, Text1(0).Text, Text1(1).Text)
                
                If bol Then
                    'Eliminar de la spagop
                    TocarEnTesoreria = False
                    If vParamAplic.ContabilidadNueva Then
                        'Si no esta el pago, o ya esta pagado, NO tocamos nada de la tesoreria
                        If Modo = 5 Then
                            'NO TOCAREMOS EL VTO, estamos asignadndo albaranes en las linea
                            SQL = ""
                        Else
                            SQL = " codmacta='" & vFactu.CtaProve & "' AND numfactu='" & Data1.Recordset.Fields!Numfactu & "'"
                            SQL = SQL & " AND fecfactu='" & Format(Data1.Recordset.Fields!FecFactu, FormatoFecha) & "'"
                            SQL = DevuelveDesdeBD(conConta, "concat(impefect-coalesce(imppagad,0),'|',coalesce(imppagad,0),'|')", "pagos", SQL & " AND 1", "1")
                            Impag = RecuperaValor(SQL, 2)
                            SQL = RecuperaValor(SQL, 1)
                        End If
                        If SQL <> "" Then
                            If Impag = "" Then Impag = "0.00"
                           
                            
                            If vFactu.TotalFac <> CCur(SQL) Then
                                If Impag <> "0.00" Then
                                   ' MsgBox "Revise pago en tesoreria. Tiene pago realizado", vbExclamation
                                Else
                                    TocarEnTesoreria = True
                                End If
                                
                            Else
                                'Mismo importe.
                                If Impag <> "0.00" Then
                                    MsgBox "Revise pago en tesoreria. Tiene pago realizado", vbExclamation
                                Else
                                    TocarEnTesoreria = True
                                End If
                            End If
                        End If
                        SQL = " codmacta='" & vFactu.CtaProve & "' AND numfactu='" & Data1.Recordset.Fields!Numfactu & "'"
                        SQL = SQL & " AND fecfactu='" & Format(Data1.Recordset.Fields!FecFactu, FormatoFecha) & "'"
                        SQL = "Delete from pagos WHERE " & SQL
                    Else
                        TocarEnTesoreria = True
                        SQL = " ctaprove='" & vFactu.CtaProve & "' AND numfactu='" & Data1.Recordset.Fields!Numfactu & "'"
                        SQL = SQL & " AND fecfactu='" & Format(Data1.Recordset.Fields!FecFactu, FormatoFecha) & "'"
                        SQL = "Delete from spagop WHERE " & SQL
                    End If
                    
                    
                    'Volvemos a grabar en TESORERIA. Tabla de Contabilidad: sconta.spagop
                    If TocarEnTesoreria Then
                        ConnConta.Execute SQL
                        If bol Then
                            bol = vFactu.InsertarEnTesoreria(MenError)
                        End If
                    End If
                End If
                Set vFactu = Nothing
            End If
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
        MenError = "Actualizando Factura." & vbCrLf & "----------------------------" & vbCrLf & MenError & vbCrLf & Err.Description
        MsgBox MenError, vbExclamation
    End If
End Function



Private Function FactContabilizada() As Boolean
Dim Cta As String, numasien As String
On Error GoTo EContab

    'comprabar que se puede modificar/eliminar la factura
    If Me.Check1.Value = 1 Then 'si esta contabilizada
        'comprobar en la contabilidad si esta contabilizada
        

        Cta = DevuelveDesdeBDNew(conAri, "sprove", "codmacta", "codprove", Text1(2).Text, "N")
 
        If Cta <> "" Then
            If vParamAplic.ContabilidadNueva Then
                numasien = "NO"
            Else
                numasien = DevuelveDesdeBDNew(conConta, "cabfactprov", "numasien", "codmacta", Cta, "T", , "numfacpr", Text1(0).Text, "T", "fecfacpr", Text1(1).Text, "F")
            End If
            If numasien <> "" Then
                FactContabilizada = True
                MsgBox "La factura esta contabilizada y no se puede modificar ni eliminar.", vbInformation
                Exit Function
            Else
                FactContabilizada = False
            End If
        Else
            FactContabilizada = True
            Exit Function
        End If
    Else
        FactContabilizada = False
    End If
EContab:
    If Err.Number <> 0 Then MuestraError Err.Number, "Comprobar Factura contabilizada", Err.Description
End Function


Private Sub TxtAux3_GotFocus(index As Integer)
    ConseguirFoco txtAux3(index), Modo
End Sub

Private Sub TxtAux3_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
    If index <> 0 And KeyCode <> 38 Then KEYdown KeyCode
End Sub

Private Sub TxtAux3_KeyPress(index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub TxtAux3_LostFocus(index As Integer)
    If Not PerderFocoGnral(txtAux3(index), Modo) Then Exit Sub
End Sub


Private Sub BloquearDatosProve(bol As Boolean)
Dim I As Byte

    'bloquear/desbloquear campos de datos segun sea de varios o no
    If Modo <> 5 Then
        Me.imgBuscar(1).visible = bol 'NIF
        Me.imgBuscar(1).Enabled = bol 'NIF
        Me.imgBuscar(2).Enabled = bol 'poblacion
        
        For I = 3 To 9 'si no es de varios no se pueden modificar los datos
            BloquearTxt Text1(I), Not bol
        Next I
    End If
End Sub


Private Function ActualizarProveVarios(Prove As String, NIF As String) As Boolean
'Modifica los datos de la tabla de Proveedores Varios
Dim vProve As CProveedor
On Error GoTo EActualizarCV

    ActualizarProveVarios = False
    
    Set vProve = New CProveedor
    If EsProveedorVarios(Prove) Then
        vProve.NIF = NIF
        vProve.Nombre = Text1(3).Text
        vProve.Domicilio = Text1(6).Text
        vProve.CPostal = Text1(7).Text
        vProve.Poblacion = Text1(8).Text
        vProve.Provincia = Text1(9).Text
        vProve.TfnoAdmon = Text1(5).Text
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


Private Function ObtenerSelFactura() As String
'Cuando venimos desde dobleClick en Movimientos de Articulos para Albaranes ya
'Facturados, abrimos este form pero cargando los datos de la factura
'correspendiente al albaran que se selecciono
Dim cad As String
Dim RS As ADODB.Recordset
On Error Resume Next

    cad = "SELECT codprove,numfactu,fecfactu FROM scafpa "
    cad = cad & " WHERE codprove=" & DBSet(hcoCodProve, "N") & " AND numalbar=" & DBSet(hcoCodMovim, "T")
    cad = cad & " AND fechaalb=" & DBSet(hcoFechaMovim, "F")

    Set RS = New ADODB.Recordset
    RS.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then 'where para la factura
        cad = " WHERE codprove=" & RS!Codprove & " AND numfactu= '" & RS!Numfactu & "' AND fecfactu=" & DBSet(RS!FecFactu, "F")
    Else
        cad = " where numfactu=-1"
    End If
    RS.Close
    Set RS = Nothing

    ObtenerSelFactura = cad
End Function



Private Function ActualizarDatosFactura() As Boolean
Dim vFactu As CFacturaCom
Dim cadSel As String

    Set vFactu = New CFacturaCom
    cadSel = ObtenerWhereCP(False)
    cadSel = "slifpc." & cadSel
    vFactu.DtoPPago = CCur(Text1(11).Text)
    vFactu.DtoGnral = CCur(Text1(12).Text)
    
    'Si tiene RETENCION
    If Me.FrmRetencionSocios.visible Then
        vFactu.PorRet = ImporteFormateado(Text1(32).Text)
        vFactu.ImpRet2 = ImporteFormateado(Text1(33).Text)
    End If

    
    
    If vFactu.CalcularDatosFactura2(cadSel, "scafpa", "slifpc", CDate(Text1(1).Text), False) Then
        Text1(14).Text = vFactu.BrutoFac
        Text1(15).Text = vFactu.ImpPPago
        Text1(16).Text = vFactu.ImpGnral
        Text1(17).Text = vFactu.BaseImp
        Text1(18).Text = vFactu.TipoIVA1
        Text1(19).Text = vFactu.TipoIVA2
        Text1(20).Text = vFactu.TipoIVA3
        Text1(21).Text = vFactu.PorceIVA1
        Text1(22).Text = vFactu.PorceIVA2
        Text1(23).Text = vFactu.PorceIVA3
        Text1(24).Text = vFactu.BaseIVA1
        Text1(25).Text = vFactu.BaseIVA2
        Text1(26).Text = vFactu.BaseIVA3
        Text1(27).Text = vFactu.ImpIVA1
        Text1(28).Text = vFactu.ImpIVA2
        Text1(29).Text = vFactu.ImpIVA3
        Text1(30).Text = vFactu.TotalFac
        If Me.FrmRetencionSocios.visible Then
            Text1(32).Text = vFactu.PorRet
            Text1(33).Text = vFactu.ImpRet2
        End If
        
        FormatoDatosTotales
        
        ActualizarDatosFactura = True
    Else
        ActualizarDatosFactura = False
        MuestraError Err.Number, "Recalculando Factura", Err.Description
    End If
    Set vFactu = Nothing
End Function


Private Sub FormatoDatosTotales()
Dim I As Byte

    For I = 14 To 17
'        Text1(I).Text = QuitarCero(Text1(I).Text)
        FormateaCampo Text1(I)
    Next I
    
    For I = 24 To 26
        If Text1(I).Text <> "" Then
            'Si la Base Imp. es 0
            If CSng(Text1(I).Text) = 0 Then
                Text1(I).Text = QuitarCero(Text1(I).Text)
                Text1(I - 3).Text = QuitarCero(Text1(I - 3).Text)
                Text1(I - 6).Text = QuitarCero(Text1(I - 6).Text)
                Text1(I + 3).Text = QuitarCero(Text1(I + 3).Text)
            Else
                FormateaCampo Text1(I)
                FormateaCampo Text1(I - 3)
                FormateaCampo Text1(I - 6)
                FormateaCampo Text1(I + 3)
            End If
        Else 'No hay Base Imponible
            Text1(I - 3).Text = QuitarCero(Text1(I - 3).Text)
            Text1(I - 6).Text = QuitarCero(Text1(I - 6).Text)
            Text1(I + 3).Text = ""
        End If
    Next I
    
    If Me.FrmRetencionSocios.visible Then
        FormateaCampo Text1(32)
        FormateaCampo Text1(33)
    End If
End Sub



Private Sub ComprobarDatosTotales()
Dim I As Byte

    For I = 14 To 17
        Text1(I).Text = ComprobarCero(Text1(I).Text)
    Next I
End Sub







Private Sub PonerForaGrid()
    'Dim RS As ADODB.Recordset
    'Dim SQL As String
    Dim Borrar As Boolean
    Dim J As Integer
    Dim C As String
    
On Error GoTo Error1
  
        Borrar = False

        If Not Data2.Recordset.EOF Then
            
            
            If vEmpresa.TieneAnalitica Then
                '- centro de coste
                ' ---- [19/10/2009] [LAURA]: añadir campo centro de coste familia
                Me.txtAux(8).Text = DBLet(Data2.Recordset!CodCCost, "T")
                Me.txtAux2(8).Text = PonerNombreCCoste(Me.txtAux(8))
            End If
              
            For J = 9 To 11
                Me.txtAux(J).Text = DBLet(Data2.Recordset.Fields(J + 7), "T")
                PonerClieObraActuacion J, True
            Next
            
            'Ampliacon
            Me.Text2(17).Text = DBLet(Data2.Recordset!numlotes, "T")
            Me.Text2(16).Text = DBLet(Data2.Recordset!Ampliaci, "T")
            
             If InstalacionEsEulerTaxco Then
                For J = 12 To 14
                        If IsNull(Data2.Recordset.Fields(J + 7)) Then
                            Me.txtAux(J).Text = ""
                        Else
                            If J = 14 Then
                                Me.txtAux(J).Text = DBLet(Me.Data2.Recordset.Fields(J + 7), "F")
                            Else
                                Me.txtAux(J).Text = DBLet(Me.Data2.Recordset.Fields(J + 7), "T")
                            End If
                        End If
                Next J
                PonerDatosAlbaranFacturaEuler
            End If
            
            
      Else
        'EOF
        Borrar = True
        
      End If   'De EOF
        
    

    
    
    

    
Error1:
    If Err.Number <> 0 Then
        MsgBox Err.Number & ": " & Err.Description, vbExclamation
        Borrar = True
    End If
    
    If Borrar Then
        
        For J = 8 To 11
            If J > 8 Then txtDesc(J).Text = ""
            txtAux(J).Text = ""
        Next
        Me.txtAux2(8).Text = ""
        Me.Text2(17).Text = "": Me.Text2(16).Text = ""
    End If

End Sub





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
        If txtAux(10).Text = "" Or txtAux(9).Text = "" Then
            If Not DesdePonerCampos Then
                MsgBox "Ponga el cliente/obra", vbExclamation
                txtAux(Cual).Text = ""
                If txtAux(9).Text = "" Then
                    PonerFoco txtAux(9)
                Else
                    PonerFoco txtAux(10)
                End If
                
            End If
            D = ""
            Exit Sub
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
            cad = "ALBARAN: " & cad
        End If
        txtDesc(0).Text = cad
    End If
        
End Sub
