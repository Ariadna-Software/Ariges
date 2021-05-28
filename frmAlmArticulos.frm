VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAlmArticulos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Artículos"
   ClientHeight    =   8430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11775
   Icon            =   "frmAlmArticulos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8430
   ScaleWidth      =   11775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtReser 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Height          =   315
      Left            =   6960
      Locked          =   -1  'True
      MaxLength       =   13
      TabIndex        =   194
      Text            =   "Text1"
      Top             =   7920
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdEuler 
      Caption         =   "Copiar de artículo"
      Height          =   375
      Left            =   4080
      TabIndex        =   189
      Top             =   7920
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton cmdGenerar 
      Caption         =   "Generar denominación"
      Height          =   375
      Left            =   6480
      TabIndex        =   115
      Top             =   7920
      Visible         =   0   'False
      Width           =   2175
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6540
      Left            =   240
      TabIndex        =   56
      Top             =   1080
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   11536
      _Version        =   393216
      Style           =   1
      Tabs            =   8
      Tab             =   6
      TabsPerRow      =   8
      TabHeight       =   520
      TabCaption(0)   =   "Datos básicos   "
      TabPicture(0)   =   "frmAlmArticulos.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "imgCuentas(3)"
      Tab(0).Control(1)=   "Label1(17)"
      Tab(0).Control(2)=   "Label1(7)"
      Tab(0).Control(3)=   "Label1(8)"
      Tab(0).Control(4)=   "Label1(6)"
      Tab(0).Control(5)=   "Label1(5)"
      Tab(0).Control(6)=   "imgCuentas(0)"
      Tab(0).Control(7)=   "imgCuentas(2)"
      Tab(0).Control(8)=   "imgCuentas(1)"
      Tab(0).Control(9)=   "imgCuentas(5)"
      Tab(0).Control(10)=   "imgCuentas(4)"
      Tab(0).Control(11)=   "Label1(9)"
      Tab(0).Control(12)=   "Label1(20)"
      Tab(0).Control(13)=   "Label1(19)"
      Tab(0).Control(14)=   "Label1(3)"
      Tab(0).Control(15)=   "Label1(4)"
      Tab(0).Control(16)=   "Label1(16)"
      Tab(0).Control(17)=   "imgFecha(0)"
      Tab(0).Control(18)=   "lblSumaStocks"
      Tab(0).Control(19)=   "Label1(37)"
      Tab(0).Control(20)=   "Label1(38)"
      Tab(0).Control(21)=   "Label1(24)"
      Tab(0).Control(22)=   "Label1(14)"
      Tab(0).Control(23)=   "Label1(42)"
      Tab(0).Control(24)=   "Line7(0)"
      Tab(0).Control(25)=   "Line7(1)"
      Tab(0).Control(26)=   "Label1(44)"
      Tab(0).Control(27)=   "FrameLitrosUd"
      Tab(0).Control(28)=   "FrameDatosAlmacen2"
      Tab(0).Control(29)=   "chkSeries"
      Tab(0).Control(30)=   "chkConjunto"
      Tab(0).Control(31)=   "Text2(3)"
      Tab(0).Control(32)=   "Text1(5)"
      Tab(0).Control(33)=   "Text1(2)"
      Tab(0).Control(34)=   "Text1(3)"
      Tab(0).Control(35)=   "Text1(7)"
      Tab(0).Control(36)=   "Text1(4)"
      Tab(0).Control(37)=   "Text2(2)"
      Tab(0).Control(38)=   "Text2(5)"
      Tab(0).Control(39)=   "Text2(1)"
      Tab(0).Control(40)=   "Text2(0)"
      Tab(0).Control(41)=   "Text2(4)"
      Tab(0).Control(42)=   "Text1(6)"
      Tab(0).Control(43)=   "Text1(12)"
      Tab(0).Control(44)=   "Text1(11)"
      Tab(0).Control(45)=   "Text1(9)"
      Tab(0).Control(46)=   "cboStatus"
      Tab(0).Control(47)=   "Text1(10)"
      Tab(0).Control(48)=   "txtSumaStock"
      Tab(0).Control(49)=   "chkCtrStock"
      Tab(0).Control(50)=   "chkMateriaPrima"
      Tab(0).Control(51)=   "Text1(8)"
      Tab(0).Control(52)=   "Text1(31)"
      Tab(0).Control(53)=   "chkRotacion"
      Tab(0).Control(54)=   "txtPVPIVA"
      Tab(0).Control(55)=   "Text1(17)"
      Tab(0).Control(56)=   "Text1(34)"
      Tab(0).Control(57)=   "cboTipoComiArtVario"
      Tab(0).Control(58)=   "cboUnidadCompra"
      Tab(0).ControlCount=   59
      TabCaption(1)   =   "Otros Datos"
      TabPicture(1)   =   "frmAlmArticulos.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Text1(33)"
      Tab(1).Control(1)=   "Text2(9)"
      Tab(1).Control(2)=   "chkWeb"
      Tab(1).Control(3)=   "framePortes"
      Tab(1).Control(4)=   "Text1(28)"
      Tab(1).Control(5)=   "Text1(19)"
      Tab(1).Control(6)=   "Text1(20)"
      Tab(1).Control(7)=   "Text1(21)"
      Tab(1).Control(8)=   "Label1(45)"
      Tab(1).Control(9)=   "imgCuentas(10)"
      Tab(1).Control(10)=   "Label2(1)"
      Tab(1).Control(11)=   "Label2(11)"
      Tab(1).Control(12)=   "Label2(2)"
      Tab(1).Control(13)=   "Label2(3)"
      Tab(1).ControlCount=   14
      TabCaption(2)   =   "Componentes"
      TabPicture(2)   =   "frmAlmArticulos.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Line4"
      Tab(2).Control(1)=   "Label5(0)"
      Tab(2).Control(2)=   "Label5(1)"
      Tab(2).Control(3)=   "Label5(2)"
      Tab(2).Control(4)=   "Label5(3)"
      Tab(2).Control(5)=   "Label5(4)"
      Tab(2).Control(6)=   "Label5(5)"
      Tab(2).Control(7)=   "Line5"
      Tab(2).Control(8)=   "DataGrid1"
      Tab(2).Control(9)=   "cmdAux"
      Tab(2).Control(10)=   "txtAux(0)"
      Tab(2).Control(11)=   "txtAux(1)"
      Tab(2).Control(12)=   "txtAux2"
      Tab(2).Control(13)=   "txtAux(3)"
      Tab(2).Control(14)=   "txtAux(4)"
      Tab(2).Control(15)=   "txtAux(5)"
      Tab(2).Control(16)=   "txtConjunto(0)"
      Tab(2).Control(17)=   "txtConjunto(1)"
      Tab(2).Control(18)=   "txtConjunto(2)"
      Tab(2).Control(19)=   "txtConjunto(3)"
      Tab(2).Control(20)=   "txtConjunto(4)"
      Tab(2).Control(21)=   "txtConjunto(5)"
      Tab(2).Control(22)=   "cmdActualizarImportes1(0)"
      Tab(2).Control(23)=   "cmdActualizarImportes1(1)"
      Tab(2).Control(24)=   "txtAux(6)"
      Tab(2).Control(25)=   "Data2"
      Tab(2).Control(26)=   "txtAux(7)"
      Tab(2).ControlCount=   27
      TabCaption(3)   =   "Control instalación / producción"
      TabPicture(3)   =   "frmAlmArticulos.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "txtAux(11)"
      Tab(3).Control(1)=   "txtAux(10)"
      Tab(3).Control(2)=   "cboCalidad"
      Tab(3).Control(3)=   "txtAux(9)"
      Tab(3).Control(4)=   "txtAux(2)"
      Tab(3).Control(5)=   "DataGrid2"
      Tab(3).Control(6)=   "Data3"
      Tab(3).ControlCount=   7
      TabCaption(4)   =   "Stocks"
      TabPicture(4)   =   "frmAlmArticulos.frx":007C
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "cmdAlma"
      Tab(4).Control(1)=   "Text3(0)"
      Tab(4).Control(2)=   "Text2(8)"
      Tab(4).Control(3)=   "Text3(2)"
      Tab(4).Control(4)=   "FrameArtxAlmac"
      Tab(4).Control(5)=   "DataGrid3"
      Tab(4).ControlCount=   6
      TabCaption(5)   =   "  EAN  / Equivalencias"
      TabPicture(5)   =   "frmAlmArticulos.frx":0098
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "cmdEquiv"
      Tab(5).Control(1)=   "Text6(0)"
      Tab(5).Control(2)=   "Text6(1)"
      Tab(5).Control(3)=   "txtAux(8)"
      Tab(5).Control(4)=   "DataGrid4"
      Tab(5).Control(5)=   "Data5"
      Tab(5).Control(6)=   "DataGrid6"
      Tab(5).Control(7)=   "Data7"
      Tab(5).Control(8)=   "Label2(6)"
      Tab(5).Control(9)=   "Label2(4)"
      Tab(5).ControlCount=   10
      TabCaption(6)   =   "Documentos"
      TabPicture(6)   =   "frmAlmArticulos.frx":00B4
      Tab(6).ControlEnabled=   -1  'True
      Tab(6).Control(0)=   "Label2(0)"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).Control(1)=   "lw1"
      Tab(6).Control(1).Enabled=   0   'False
      Tab(6).Control(2)=   "FrameDisponible"
      Tab(6).Control(2).Enabled=   0   'False
      Tab(6).Control(3)=   "Frame4"
      Tab(6).Control(3).Enabled=   0   'False
      Tab(6).ControlCount=   4
      TabCaption(7)   =   "Fitosanitarios"
      TabPicture(7)   =   "frmAlmArticulos.frx":00D0
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Label2(5)"
      Tab(7).Control(1)=   "Label1(41)"
      Tab(7).Control(2)=   "data6"
      Tab(7).Control(3)=   "DataGrid5"
      Tab(7).Control(4)=   "FrameServicios"
      Tab(7).Control(5)=   "cmdMatAux"
      Tab(7).Control(6)=   "Text5(0)"
      Tab(7).Control(7)=   "Text5(1)"
      Tab(7).Control(8)=   "cboADV"
      Tab(7).ControlCount=   9
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   33
         Left            =   -74760
         MaxLength       =   16
         TabIndex        =   44
         Tag             =   "SIGAUS|T|S|||sartic|artsigaus||N|"
         Text            =   "T"
         Top             =   5640
         Width           =   1485
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   9
         Left            =   -73200
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   196
         Text            =   "Text2"
         Top             =   5640
         Width           =   4485
      End
      Begin VB.ComboBox cboUnidadCompra 
         Height          =   315
         Left            =   -65520
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Tag             =   "Unidad compra|N|S|||sartic|unidadesCompra||N|"
         Top             =   3120
         Width           =   1335
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   11
         Left            =   -67080
         MaxLength       =   60
         TabIndex        =   193
         Text            =   "calid max"
         Top             =   5640
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   10
         Left            =   -68040
         MaxLength       =   60
         TabIndex        =   192
         Text            =   "min calid"
         Top             =   5640
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.ComboBox cboCalidad 
         Height          =   315
         Left            =   -73440
         Style           =   2  'Dropdown List
         TabIndex        =   190
         Top             =   5640
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   9
         Left            =   -71040
         MaxLength       =   60
         TabIndex        =   191
         Text            =   "Especfi calidad"
         Top             =   5640
         Visible         =   0   'False
         Width           =   2955
      End
      Begin VB.ComboBox cboTipoComiArtVario 
         Height          =   315
         Left            =   -66480
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Tag             =   "Tipo comision|N|S|||sartic|TipoComiArtVario||N|"
         Top             =   4080
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   34
         Left            =   -66480
         MaxLength       =   8
         TabIndex        =   28
         Tag             =   "Embalaje grande|N|S|||sartic|unicajas2||N|"
         Text            =   "Text1"
         Top             =   3600
         Width           =   615
      End
      Begin VB.ComboBox cboADV 
         Height          =   315
         ItemData        =   "frmAlmArticulos.frx":00EC
         Left            =   -74640
         List            =   "frmAlmArticulos.frx":00EE
         Style           =   2  'Dropdown List
         TabIndex        =   182
         Tag             =   "Partes trabajo|N|N|||sartic|partesADV|||"
         Top             =   960
         Width           =   1935
      End
      Begin VB.CheckBox chkWeb 
         Caption         =   "Se muestra en la web"
         Height          =   315
         Left            =   -71760
         TabIndex        =   181
         Tag             =   "Se muestra en la web|N|N|0|1|sartic|oftweb||N|"
         Top             =   6120
         Width           =   2415
      End
      Begin VB.TextBox txtAux 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   7
         Left            =   -65520
         TabIndex        =   180
         Tag             =   "C|N|S|||||0||"
         Text            =   "Dato2"
         ToolTipText     =   "Materia prima"
         Top             =   3240
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   17
         Left            =   -73200
         MaxLength       =   12
         TabIndex        =   18
         Tag             =   "Precio Venta al público|N|N|0|999999.0000|sartic|preciove|###,##0.0000|N|"
         Text            =   "Text1"
         Top             =   6000
         Width           =   1215
      End
      Begin VB.TextBox txtPVPIVA 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   -70080
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   37
         Text            =   "Text1"
         Top             =   6000
         Width           =   1215
      End
      Begin VB.CommandButton cmdEquiv 
         Caption         =   "+"
         Height          =   255
         Left            =   -69600
         TabIndex        =   176
         ToolTipText     =   "Materias activas"
         Top             =   5280
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox Text6 
         Height          =   315
         Index           =   0
         Left            =   -70920
         MaxLength       =   16
         TabIndex        =   38
         Text            =   "Text3"
         Top             =   5280
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox Text6 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   1
         Left            =   -69360
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   177
         Text            =   "Text2"
         Top             =   5280
         Visible         =   0   'False
         Width           =   3045
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   1
         Left            =   -68040
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   173
         Text            =   "Text2"
         Top             =   1560
         Visible         =   0   'False
         Width           =   3045
      End
      Begin VB.TextBox Text5 
         Height          =   315
         Index           =   0
         Left            =   -69000
         MaxLength       =   8
         TabIndex        =   171
         Text            =   "Text3"
         Top             =   1560
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.CommandButton cmdMatAux 
         Caption         =   "+"
         Height          =   255
         Left            =   -68280
         TabIndex        =   172
         ToolTipText     =   "Materias activas"
         Top             =   1560
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Frame FrameServicios 
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   3735
         Left            =   -74760
         TabIndex        =   157
         Top             =   1440
         Width           =   4815
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   32
            Left            =   120
            MaxLength       =   4
            TabIndex        =   167
            Tag             =   "Cod. Categoría|T|S|||sartic|numadr||N|"
            Text            =   "Tex"
            Top             =   2640
            Width           =   765
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   7
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   166
            Text            =   "Text2"
            Top             =   2640
            Width           =   3645
         End
         Begin VB.Frame Frame3 
            Caption         =   "Registro fitosanitarios"
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
            Height          =   855
            Left            =   120
            TabIndex        =   160
            Top             =   1080
            Width           =   4060
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Height          =   315
               Index           =   24
               Left            =   2400
               MaxLength       =   10
               TabIndex        =   162
               Tag             =   "Fecha vigencia|F|S|||sartic|fecvigen||N|"
               Text            =   "Tex"
               Top             =   430
               Width           =   1400
            End
            Begin VB.TextBox Text1 
               Height          =   315
               Index           =   23
               Left            =   240
               MaxLength       =   15
               TabIndex        =   161
               Tag             =   "Nº serie|T|S|||sartic|numserie||N|"
               Text            =   "Tex"
               Top             =   430
               Width           =   1965
            End
            Begin VB.Label Label1 
               Caption         =   "Fecha vigencia"
               Height          =   255
               Index           =   32
               Left            =   2400
               TabIndex        =   164
               Top             =   230
               Width           =   1215
            End
            Begin VB.Label Label1 
               Caption         =   "Nº"
               Height          =   255
               Index           =   31
               Left            =   240
               TabIndex        =   163
               Top             =   230
               Width           =   1215
            End
            Begin VB.Image imgFecha 
               Height          =   240
               Index           =   3
               Left            =   3600
               Picture         =   "frmAlmArticulos.frx":00F0
               ToolTipText     =   "Buscar fecha"
               Top             =   180
               Width           =   240
            End
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   22
            Left            =   120
            MaxLength       =   3
            TabIndex        =   159
            Tag             =   "Cod. Categoría|T|S|||sartic|codcateg||N|"
            Text            =   "Tex"
            Top             =   480
            Width           =   645
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   22
            Left            =   840
            Locked          =   -1  'True
            TabIndex        =   158
            Text            =   "Text2"
            Top             =   480
            Width           =   3645
         End
         Begin VB.Label Label1 
            Caption         =   "NºADR"
            Height          =   255
            Index           =   39
            Left            =   120
            TabIndex        =   168
            Top             =   2400
            Width           =   645
         End
         Begin VB.Image imgCuentas 
            Height          =   240
            Index           =   9
            Left            =   840
            Picture         =   "frmAlmArticulos.frx":067A
            ToolTipText     =   "Buscar familia"
            Top             =   2400
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Cod. Categoría"
            Height          =   255
            Index           =   21
            Left            =   120
            TabIndex        =   165
            Top             =   240
            Width           =   1125
         End
         Begin VB.Image imgCuentas 
            Height          =   240
            Index           =   8
            Left            =   1320
            Picture         =   "frmAlmArticulos.frx":077C
            ToolTipText     =   "Buscar familia"
            Top             =   240
            Width           =   240
         End
      End
      Begin VB.CheckBox chkRotacion 
         Caption         =   "Rotación"
         Height          =   315
         Left            =   -64920
         TabIndex        =   34
         Tag             =   "Rotacion|N|N|0|1|sartic|rotacion||N|"
         Top             =   5280
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   31
         Left            =   -66480
         MaxLength       =   18
         TabIndex        =   19
         Tag             =   "Refprov|T|S|||sartic|referprov|||"
         Text            =   "Text1"
         Top             =   480
         Width           =   1830
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   8
         Left            =   -74400
         MaxLength       =   60
         TabIndex        =   153
         Text            =   "Dat"
         Top             =   4440
         Visible         =   0   'False
         Width           =   2595
      End
      Begin MSAdodcLib.Adodc Data2 
         Height          =   330
         Left            =   -66240
         Top             =   840
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
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
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         Height          =   3735
         Left            =   120
         TabIndex        =   151
         Top             =   600
         Width           =   855
         Begin MSComctlLib.Toolbar Toolbar2 
            Height          =   3690
            Left            =   120
            TabIndex        =   152
            Top             =   0
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   6509
            ButtonWidth     =   609
            ButtonHeight    =   582
            AllowCustomize  =   0   'False
            Appearance      =   1
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   13
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Tarifas"
                  Object.Tag             =   "0"
                  Style           =   2
                  Value           =   1
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Precios especiales"
                  Object.Tag             =   "1"
                  Style           =   2
               EndProperty
               BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Promociones"
                  Object.Tag             =   "2"
                  Style           =   2
               EndProperty
               BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Pedidos"
                  Object.Tag             =   "3"
                  Style           =   2
               EndProperty
               BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.Visible         =   0   'False
                  Object.ToolTipText     =   "Precios especiales"
                  Style           =   2
               EndProperty
               BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.Visible         =   0   'False
                  Style           =   3
               EndProperty
               BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Movimientos"
                  Object.Tag             =   "4"
                  Style           =   2
               EndProperty
               BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Precios proveedor"
                  Object.Tag             =   "5"
                  Style           =   2
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame FrameDisponible 
         Height          =   2295
         Left            =   8160
         TabIndex        =   141
         Top             =   3720
         Width           =   2895
         Begin VB.TextBox Text4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   0
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   145
            Text            =   "Text4"
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox Text4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   1
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   144
            Text            =   "Text4"
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox Text4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   2
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   143
            Text            =   "Text4"
            Top             =   1200
            Width           =   1215
         End
         Begin VB.TextBox Text4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   3
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   142
            Text            =   "Text4"
            Top             =   1800
            Width           =   1215
         End
         Begin VB.Image imgAyuda 
            Height          =   240
            Index           =   0
            Left            =   960
            Picture         =   "frmAlmArticulos.frx":087E
            ToolTipText     =   "Cantidad disponible"
            Top             =   240
            Width           =   240
         End
         Begin VB.Label Label4 
            Caption         =   "Reservas"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   149
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label4 
            Caption         =   "Pedido prov"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   148
            Top             =   1200
            Width           =   855
         End
         Begin VB.Label Label4 
            Caption         =   "Stock"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   147
            Top             =   240
            Width           =   855
         End
         Begin VB.Line Line1 
            X1              =   120
            X2              =   2520
            Y1              =   1560
            Y2              =   1560
         End
         Begin VB.Label Label4 
            Caption         =   "Disponible"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   146
            Top             =   1860
            Width           =   855
         End
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   8
         Left            =   -66480
         MaxLength       =   10
         TabIndex        =   22
         Tag             =   "Num. orden|N|S|||sartic|numorden|||"
         Text            =   "Text1"
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox txtAux 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   6
         Left            =   -66480
         TabIndex        =   137
         Tag             =   "C|T|S|||||||"
         Text            =   "Dato2"
         ToolTipText     =   "Materia prima"
         Top             =   2880
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.CheckBox chkMateriaPrima 
         Caption         =   "Materia prima"
         Height          =   315
         Left            =   -66360
         TabIndex        =   33
         Tag             =   "Materia prima|N|N|0|1|sartic|mateprima||N|"
         Top             =   5280
         Width           =   1815
      End
      Begin VB.Frame framePortes 
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
         ForeColor       =   &H00972E0B&
         Height          =   855
         Left            =   -68040
         TabIndex        =   135
         Top             =   5400
         Width           =   3975
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   30
            Left            =   2640
            MaxLength       =   15
            TabIndex        =   45
            Tag             =   "Kilos|N|S|||sartic|pesoarti|#,##0.00||"
            Text            =   "Tex"
            Top             =   360
            Width           =   1125
         End
         Begin VB.Label Label1 
            Caption         =   "Kilos"
            Height          =   255
            Index           =   36
            Left            =   1920
            TabIndex        =   136
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.CommandButton cmdActualizarImportes1 
         Height          =   375
         Index           =   1
         Left            =   -65640
         Picture         =   "frmAlmArticulos.frx":13B8
         Style           =   1  'Graphical
         TabIndex        =   132
         ToolTipText     =   "Modificar componente"
         Top             =   5640
         Width           =   375
      End
      Begin VB.CommandButton cmdActualizarImportes1 
         Height          =   375
         Index           =   0
         Left            =   -65040
         Picture         =   "frmAlmArticulos.frx":1DBA
         Style           =   1  'Graphical
         TabIndex        =   131
         ToolTipText     =   "Actualizar importes"
         Top             =   5640
         Width           =   375
      End
      Begin VB.TextBox txtConjunto 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   -67680
         TabIndex        =   129
         Text            =   "Text5"
         Top             =   5760
         Width           =   1215
      End
      Begin VB.TextBox txtConjunto 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   -69000
         TabIndex        =   127
         Text            =   "Text5"
         Top             =   5760
         Width           =   1215
      End
      Begin VB.TextBox txtConjunto 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   -70320
         TabIndex        =   125
         Text            =   "Text5"
         Top             =   5760
         Width           =   1215
      End
      Begin VB.TextBox txtConjunto 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   -72120
         TabIndex        =   123
         Text            =   "Text5"
         Top             =   5760
         Width           =   1215
      End
      Begin VB.TextBox txtConjunto 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   -73440
         TabIndex        =   121
         Text            =   "Text5"
         Top             =   5760
         Width           =   1215
      End
      Begin VB.TextBox txtConjunto 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   -74760
         TabIndex        =   119
         Text            =   "Text5"
         Top             =   5760
         Width           =   1215
      End
      Begin VB.TextBox txtAux 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   5
         Left            =   -67200
         TabIndex        =   118
         Tag             =   "C|N|S|||||###,##0.0000||"
         Text            =   "Dato2"
         Top             =   3120
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.TextBox txtAux 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   4
         Left            =   -65880
         TabIndex        =   117
         Tag             =   "C|N|S|||||###,##0.0000||"
         Text            =   "Dato2"
         Top             =   3120
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.TextBox txtAux 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   3
         Left            =   -68640
         TabIndex        =   116
         Tag             =   "C|N|S|||||###,##0.0000||"
         Text            =   "Dato2"
         Top             =   3120
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   28
         Left            =   -74760
         MaxLength       =   60
         ScrollBars      =   2  'Vertical
         TabIndex        =   43
         Tag             =   "Taux|T|S|||sartic|txtauxdocumento|||"
         Top             =   4920
         Width           =   6015
      End
      Begin VB.CommandButton cmdAlma 
         Caption         =   "+"
         Height          =   255
         Left            =   -74040
         TabIndex        =   112
         Top             =   3600
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox Text3 
         Height          =   315
         Index           =   0
         Left            =   -74760
         MaxLength       =   8
         TabIndex        =   92
         Tag             =   "Código Almacen|N|N|||salmac|codalmac|0|S|"
         Text            =   "Text3"
         Top             =   3600
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   8
         Left            =   -73920
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   110
         Text            =   "Text2"
         Top             =   3600
         Visible         =   0   'False
         Width           =   3045
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   2
         Left            =   -70800
         MaxLength       =   16
         TabIndex        =   93
         Tag             =   "Cantidad Stock|N|N|||salmac|canstock|#,###,###,##0.00|N|"
         Text            =   "Text3"
         Top             =   3600
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.Frame FrameArtxAlmac 
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
         Height          =   4455
         Left            =   -68520
         TabIndex        =   89
         Top             =   840
         Width           =   4455
         Begin VB.TextBox Text3 
            Height          =   315
            Index           =   1
            Left            =   240
            MaxLength       =   15
            TabIndex        =   94
            Tag             =   "Ubicación|T|N|||salmac|ubialmac||N|"
            Text            =   "Text3"
            Top             =   480
            Width           =   990
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   3
            Left            =   2760
            MaxLength       =   16
            TabIndex        =   96
            Tag             =   "Stock Mínimo|N|S|||salmac|stockmin|#,###,###,##0.00|N|"
            Text            =   "Text3"
            Top             =   960
            Width           =   1485
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   4
            Left            =   2760
            MaxLength       =   16
            TabIndex        =   97
            Tag             =   "Punto de Pedido|N|S|||salmac|puntoped|#,###,###,##0.00|N|"
            Text            =   "Text3"
            Top             =   1440
            Width           =   1485
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   6
            Left            =   2880
            MaxLength       =   16
            TabIndex        =   95
            Tag             =   "Stock inventario|N|S|||salmac|stockinv|#,###,###,##0.00|N|"
            Text            =   "Text3"
            Top             =   3480
            Width           =   1485
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   7
            Left            =   240
            MaxLength       =   10
            TabIndex        =   99
            Tag             =   "Fecha inventario|F|S|||salmac|fechainv|dd/mm/yyyy|N|"
            Text            =   "Text3"
            Top             =   3480
            Width           =   1125
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   8
            Left            =   1440
            MaxLength       =   8
            TabIndex        =   100
            Tag             =   "Hora Inventario|H|S|||salmac|horainve|hh:mm:ss|N|"
            Text            =   "Text3"
            Top             =   3480
            Width           =   1125
         End
         Begin VB.CheckBox chkInventario 
            Height          =   195
            Left            =   240
            TabIndex        =   102
            Top             =   4080
            Width           =   255
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   5
            Left            =   2760
            MaxLength       =   16
            TabIndex        =   98
            Tag             =   "Stock Máximo|N|S|||salmac|stockmax|#,###,###,##0.00|N|"
            Text            =   "Text3"
            Top             =   1920
            Width           =   1485
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   6
            Left            =   1320
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   91
            Text            =   "Text2"
            Top             =   480
            Width           =   2925
         End
         Begin VB.Label Label1 
            Caption         =   "INVENTARIO"
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
            Index           =   2
            Left            =   240
            TabIndex        =   155
            Top             =   2880
            Width           =   1215
         End
         Begin VB.Line Line6 
            X1              =   120
            X2              =   4320
            Y1              =   2640
            Y2              =   2640
         End
         Begin VB.Label Label3 
            Caption         =   "Realizando Inventario"
            Height          =   255
            Left            =   600
            TabIndex        =   101
            Top             =   4080
            Width           =   1815
         End
         Begin VB.Label Label1 
            Caption         =   "Ubicación"
            Height          =   255
            Index           =   23
            Left            =   240
            TabIndex        =   109
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Stock Mínimo"
            Height          =   255
            Index           =   25
            Left            =   240
            TabIndex        =   108
            Top             =   1080
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Punto de Pedido"
            Height          =   255
            Index           =   26
            Left            =   240
            TabIndex        =   107
            Top             =   1500
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Stock "
            Height          =   255
            Index           =   28
            Left            =   3600
            TabIndex        =   106
            Top             =   3240
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "Hora "
            Height          =   255
            Index           =   30
            Left            =   1440
            TabIndex        =   104
            Top             =   3240
            Width           =   495
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   2
            Left            =   840
            Picture         =   "frmAlmArticulos.frx":2344
            ToolTipText     =   "Buscar fecha"
            Top             =   3240
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Stock Máximo"
            Height          =   255
            Index           =   27
            Left            =   240
            TabIndex        =   103
            Top             =   1920
            Width           =   1215
         End
         Begin VB.Image imgCuentas 
            Height          =   240
            Index           =   6
            Left            =   4080
            Picture         =   "frmAlmArticulos.frx":28CE
            ToolTipText     =   "Buscar almacen"
            Top             =   4080
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image imgCuentas 
            Height          =   240
            Index           =   7
            Left            =   1080
            Picture         =   "frmAlmArticulos.frx":29D0
            ToolTipText     =   "Buscar ubicación"
            Top             =   240
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha "
            Height          =   255
            Index           =   29
            Left            =   240
            TabIndex        =   105
            Top             =   3240
            Width           =   735
         End
      End
      Begin VB.TextBox Text1 
         Height          =   975
         Index           =   19
         Left            =   -74760
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   40
         Tag             =   "Texto para Ventas|T|S|||sartic|textoven|||"
         Top             =   840
         Width           =   6015
      End
      Begin VB.TextBox Text1 
         Height          =   855
         Index           =   20
         Left            =   -74760
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   41
         Tag             =   "Texto para compras|T|S|||sartic|textocom|||"
         Top             =   2400
         Width           =   6015
      End
      Begin VB.TextBox Text1 
         Height          =   855
         Index           =   21
         Left            =   -74760
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   42
         Tag             =   "Control de instalación|T|S|||sartic|controli|||"
         Top             =   3705
         Width           =   6015
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   2
         Left            =   -74040
         MaxLength       =   60
         TabIndex        =   78
         Text            =   "Dat"
         Top             =   2880
         Visible         =   0   'False
         Width           =   7035
      End
      Begin VB.TextBox txtAux2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   290
         Left            =   -73200
         TabIndex        =   76
         Text            =   "Dato2"
         Top             =   3180
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.TextBox txtAux 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   1
         Left            =   -70320
         TabIndex        =   75
         Tag             =   "C|N|N|||||###,##0.0000||"
         Text            =   "Dato2"
         Top             =   3180
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   0
         Left            =   -74280
         TabIndex        =   74
         Text            =   "Dat"
         Top             =   3180
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.CommandButton cmdAux 
         Appearance      =   0  'Flat
         Caption         =   "+"
         Height          =   315
         Left            =   -73440
         TabIndex        =   73
         Top             =   3180
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.CheckBox chkCtrStock 
         Caption         =   "¿Control de stock?"
         Height          =   315
         Left            =   -66360
         TabIndex        =   31
         Tag             =   "Control de stock|N|N|0|1|sartic|ctrstock||N|"
         Top             =   4800
         Width           =   1815
      End
      Begin VB.TextBox txtSumaStock 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0FF&
         Height          =   315
         Left            =   -66000
         Locked          =   -1  'True
         MaxLength       =   13
         TabIndex        =   68
         Text            =   "Text1"
         Top             =   6000
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   10
         Left            =   -66480
         MaxLength       =   10
         TabIndex        =   21
         Tag             =   "Fecha de Alta|F|N|||sartic|fecaltas|dd/mm/yyyy|N|"
         Text            =   "Text1"
         Top             =   1335
         Width           =   1335
      End
      Begin VB.ComboBox cboStatus 
         Height          =   315
         ItemData        =   "frmAlmArticulos.frx":2AD2
         Left            =   -66480
         List            =   "frmAlmArticulos.frx":2AD4
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Tag             =   "Situación Artículo|N|N|||sartic|codstatu||N|"
         Top             =   2223
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   9
         Left            =   -66480
         MaxLength       =   18
         TabIndex        =   20
         Tag             =   "Código Asociación|T|S|||sartic|codtelem||N|"
         Text            =   "Text1"
         Top             =   927
         Width           =   1830
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   11
         Left            =   -66480
         MaxLength       =   8
         TabIndex        =   25
         Tag             =   "Días de garantia|N|N|0|99999|sartic|garantia||N|"
         Text            =   "Text1"
         Top             =   2655
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   12
         Left            =   -66480
         MaxLength       =   8
         TabIndex        =   26
         Tag             =   "Unidades por caja|N|N|||sartic|unicajas||N|"
         Text            =   "Text1"
         Top             =   3120
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   6
         Left            =   -73200
         MaxLength       =   8
         TabIndex        =   7
         Tag             =   "Cod. Tipo Artículo|T|N|||sartic|codtipar||N|"
         Text            =   "Te"
         Top             =   2223
         Width           =   885
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   4
         Left            =   -72240
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   53
         Text            =   "Text2"
         Top             =   2223
         Width           =   3405
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   0
         Left            =   -72240
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   46
         Text            =   "Text2"
         Top             =   495
         Width           =   3405
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   1
         Left            =   -72240
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   47
         Text            =   "Text2"
         Top             =   927
         Width           =   3405
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   5
         Left            =   -72240
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   55
         Text            =   "Text2"
         Top             =   2655
         Width           =   3405
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   2
         Left            =   -72240
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   48
         Text            =   "Text2"
         Top             =   1359
         Width           =   3405
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   4
         Left            =   -73200
         MaxLength       =   4
         TabIndex        =   5
         Tag             =   "Cod. Marca|N|N|0|9999|sartic|codmarca|0000|N|"
         Text            =   "Text1"
         Top             =   1359
         Width           =   885
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   7
         Left            =   -73200
         MaxLength       =   3
         TabIndex        =   8
         Tag             =   "Tipo de IVA|N|N|0||sartic|codigiva||N|"
         Text            =   "T"
         Top             =   2655
         Width           =   885
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   3
         Left            =   -73200
         TabIndex        =   4
         Tag             =   "Cod. Familia|N|N|0|32000|sartic|codfamia|0000|N|"
         Text            =   "Text1"
         Top             =   927
         Width           =   885
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   2
         Left            =   -73200
         MaxLength       =   6
         TabIndex        =   3
         Tag             =   "Cod. Proveedor|N|N|0|999999|sartic|codprove|000000|N|"
         Text            =   "Text1"
         Top             =   495
         Width           =   885
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   5
         Left            =   -73200
         MaxLength       =   2
         TabIndex        =   6
         Tag             =   "Cod. Tipo Unidad|N|N|0|99|sartic|codunida|00|N|"
         Text            =   "Text1"
         Top             =   1791
         Width           =   885
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   3
         Left            =   -72240
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   51
         Text            =   "Text2"
         Top             =   1791
         Width           =   3405
      End
      Begin VB.CheckBox chkConjunto 
         Caption         =   "Tiene componentes"
         Height          =   315
         Left            =   -68400
         TabIndex        =   32
         Tag             =   "¿Es conjunto?|N|N|0|1|sartic|conjunto||N|"
         Top             =   5280
         Width           =   1935
      End
      Begin VB.CheckBox chkSeries 
         Caption         =   "¿Control Nº Serie?"
         Height          =   315
         Left            =   -68400
         TabIndex        =   30
         Tag             =   "¿Control nº serie?|N|N|0|1|sartic|nseriesn||N|"
         Top             =   4800
         Width           =   1815
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Height          =   5250
         Left            =   -73560
         TabIndex        =   79
         Top             =   720
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   9260
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
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
      Begin MSAdodcLib.Adodc Data3 
         Height          =   330
         Left            =   -66360
         Top             =   720
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
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
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   3840
         Left            =   -74760
         TabIndex        =   77
         Top             =   960
         Width           =   10725
         _ExtentX        =   18918
         _ExtentY        =   6773
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
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
      Begin VB.Frame FrameDatosAlmacen2 
         BorderStyle     =   0  'None
         Caption         =   "Datos Relacionados con Almacen"
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
         Height          =   2535
         Left            =   -74880
         TabIndex        =   84
         Top             =   3120
         Width           =   6135
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   35
            Left            =   1680
            MaxLength       =   12
            TabIndex        =   17
            Tag             =   "Precio Medio Acumulado|N|S|0|999999.0000|sartic|preciominvta|###,##0.0000|N|"
            Text            =   "Text1"
            Top             =   2160
            Width           =   1215
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   16
            Left            =   1680
            MaxLength       =   12
            TabIndex        =   13
            Tag             =   "Precio Standard|N|S|0|999999.0000|sartic|preciost|###,##0.0000|N|"
            Text            =   "Text1"
            Top             =   1080
            Width           =   1215
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   25
            Left            =   4920
            MaxLength       =   6
            TabIndex        =   16
            Tag             =   "Margen comercial|N|S|0|999.00|sartic|margecom|##0.00|N|"
            Text            =   "Text1"
            Top             =   1560
            Width           =   1080
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   27
            Left            =   4920
            MaxLength       =   10
            TabIndex        =   12
            Tag             =   "Fecha último cambio P.V.P.|F|S|||sartic|ultfecpvp|dd/mm/yyyy|N|"
            Text            =   "Text1"
            Top             =   570
            Width           =   1095
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   26
            Left            =   4920
            MaxLength       =   12
            TabIndex        =   14
            Tag             =   "Precio anual matenimiento|N|S|0|999999.00|sartic|preanuman|###,##0.00|N|"
            Text            =   "Text1"
            Top             =   1080
            Width           =   1095
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   18
            Left            =   4920
            MaxLength       =   10
            TabIndex        =   10
            Tag             =   "Fecha última compra|F|S|||sartic|ultfecco|dd/mm/yyyy|N|"
            Text            =   "Text1"
            Top             =   120
            Width           =   1095
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   15
            Left            =   1680
            MaxLength       =   12
            TabIndex        =   9
            Tag             =   "Precio Ultima Compra|N|S|0|999999.0000|sartic|preciouc|###,##0.0000|N|"
            Text            =   "Text1"
            Top             =   120
            Width           =   1215
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   14
            Left            =   1680
            MaxLength       =   12
            TabIndex        =   15
            Tag             =   "Precio Medio Acumulado|N|S|0|999999.0000|sartic|precioma|###,##0.0000|N|"
            Text            =   "Text1"
            Top             =   1560
            Width           =   1215
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   13
            Left            =   1680
            MaxLength       =   12
            TabIndex        =   11
            Tag             =   "Precio Medio Ponderado|N|S|0|999999.0000|sartic|preciomp|###,##0.0000|N|"
            Text            =   "Text1"
            Top             =   570
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Pr. mínimo venta"
            Height          =   255
            Index           =   43
            Left            =   120
            TabIndex        =   187
            Top             =   2190
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Pr.Standard"
            Height          =   255
            Index           =   13
            Left            =   120
            TabIndex        =   185
            Top             =   1110
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Margen Comercial"
            Height          =   255
            Index           =   33
            Left            =   3120
            TabIndex        =   184
            Top             =   1590
            Width           =   1335
         End
         Begin VB.Line Line3 
            BorderColor     =   &H00C0C0C0&
            X1              =   3120
            X2              =   6000
            Y1              =   1440
            Y2              =   1440
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00C0C0C0&
            X1              =   6960
            X2              =   8280
            Y1              =   1440
            Y2              =   1440
         End
         Begin VB.Label Label1 
            Caption         =   "Últ. cambio P.V.P."
            Height          =   255
            Index           =   22
            Left            =   3120
            TabIndex        =   113
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Pre. anual mantenim."
            Height          =   255
            Index           =   34
            Left            =   3120
            TabIndex        =   90
            Top             =   1080
            Width           =   1815
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   1
            Left            =   4440
            Picture         =   "frmAlmArticulos.frx":2AD6
            ToolTipText     =   "Buscar fecha"
            Top             =   120
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Últ. fec. compra"
            Height          =   255
            Index           =   15
            Left            =   3120
            TabIndex        =   88
            Top             =   120
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Precio última compra"
            Height          =   255
            Index           =   12
            Left            =   120
            TabIndex        =   87
            Top             =   120
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Pr. Med  Acumulado"
            Height          =   255
            Index           =   11
            Left            =   120
            TabIndex        =   86
            Top             =   1590
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Pr Med. Ponderado"
            Height          =   255
            Index           =   10
            Left            =   120
            TabIndex        =   85
            Top             =   600
            Width           =   1575
         End
      End
      Begin MSDataGridLib.DataGrid DataGrid3 
         Height          =   4455
         Left            =   -74760
         TabIndex        =   111
         Top             =   960
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   7858
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
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
      Begin VB.Frame FrameLitrosUd 
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
         ForeColor       =   &H00972E0B&
         Height          =   735
         Left            =   -64920
         TabIndex        =   133
         Top             =   1560
         Width           =   1095
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   29
            Left            =   0
            MaxLength       =   15
            TabIndex        =   23
            Text            =   "Tex"
            Top             =   240
            Width           =   765
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Lt  x  ud"
            Height          =   195
            Index           =   35
            Left            =   0
            TabIndex        =   134
            Top             =   0
            Width           =   570
         End
      End
      Begin MSComctlLib.ListView lw1 
         Height          =   5535
         Left            =   960
         TabIndex        =   140
         Top             =   600
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   9763
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSDataGridLib.DataGrid DataGrid4 
         Height          =   4440
         Left            =   -74520
         TabIndex        =   150
         Top             =   1200
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   7832
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
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
      Begin MSAdodcLib.Adodc Data5 
         Height          =   330
         Left            =   -74880
         Top             =   5400
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
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
      Begin MSDataGridLib.DataGrid DataGrid5 
         Height          =   5055
         Left            =   -69600
         TabIndex        =   169
         Top             =   1080
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   8916
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
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
      Begin MSAdodcLib.Adodc data6 
         Height          =   330
         Left            =   -65760
         Top             =   480
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
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
      Begin MSDataGridLib.DataGrid DataGrid6 
         Height          =   4440
         Left            =   -70920
         TabIndex        =   175
         Top             =   1200
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   7832
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
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
      Begin MSAdodcLib.Adodc Data7 
         Height          =   330
         Left            =   -70920
         Top             =   5400
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
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
         Caption         =   "data6"
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
      Begin VB.Label Label1 
         Caption         =   "Art. SIGAUS/Tasas"
         Height          =   195
         Index           =   45
         Left            =   -74760
         TabIndex        =   197
         Top             =   5400
         Width           =   1395
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Index           =   10
         Left            =   -73320
         Picture         =   "frmAlmArticulos.frx":3060
         ToolTipText     =   "Buscar tipo IVA"
         Top             =   5400
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo comision Varios"
         Height          =   255
         Index           =   44
         Left            =   -68400
         TabIndex        =   188
         Top             =   4080
         Width           =   1695
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00808080&
         Index           =   1
         X1              =   -68400
         X2              =   -63840
         Y1              =   4560
         Y2              =   4560
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00808080&
         Index           =   0
         X1              =   -74880
         X2              =   -63840
         Y1              =   5760
         Y2              =   5760
      End
      Begin VB.Label Label1 
         Caption         =   "Ud.g"
         Height          =   255
         Index           =   42
         Left            =   -68400
         TabIndex        =   186
         Top             =   3600
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "En partes trabajo"
         Height          =   255
         Index           =   41
         Left            =   -74640
         TabIndex        =   183
         Top             =   720
         Width           =   2085
      End
      Begin VB.Label Label1 
         Caption         =   "P.V.P."
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
         Index           =   14
         Left            =   -74760
         TabIndex        =   179
         Top             =   6060
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "P.V.P. + IVA"
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
         Index           =   24
         Left            =   -71760
         TabIndex        =   178
         Top             =   6060
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Equivalencias"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   300
         Index           =   6
         Left            =   -70920
         TabIndex        =   174
         Top             =   720
         Width           =   2865
      End
      Begin VB.Label Label2 
         Caption         =   "Materias activas"
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
         Left            =   -69600
         TabIndex        =   170
         Top             =   720
         Width           =   2865
      End
      Begin VB.Label Label1 
         Caption         =   "Referencia prove."
         Height          =   255
         Index           =   38
         Left            =   -68400
         TabIndex        =   156
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Códigos de Barras"
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
         Left            =   -74520
         TabIndex        =   154
         Top             =   720
         Width           =   2865
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
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
         Height          =   300
         Index           =   0
         Left            =   8040
         TabIndex        =   139
         Top             =   480
         Width           =   2865
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Orden"
         Height          =   255
         Index           =   37
         Left            =   -68400
         TabIndex        =   138
         Top             =   1830
         Width           =   1095
      End
      Begin VB.Label lblSumaStocks 
         Caption         =   "Suma Stock Almacenes"
         Height          =   195
         Left            =   -68040
         TabIndex        =   69
         Top             =   6060
         Width           =   1695
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         X1              =   -70320
         X2              =   -66480
         Y1              =   5400
         Y2              =   5400
      End
      Begin VB.Label Label5 
         Caption         =   "Diferencia"
         Height          =   255
         Index           =   5
         Left            =   -67680
         TabIndex        =   130
         Top             =   5520
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "PVP real"
         Height          =   255
         Index           =   4
         Left            =   -69000
         TabIndex        =   128
         Top             =   5520
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "PVP articulo"
         Height          =   255
         Index           =   3
         Left            =   -70320
         TabIndex        =   126
         Top             =   5520
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Diferencia"
         Height          =   255
         Index           =   2
         Left            =   -72120
         TabIndex        =   124
         Top             =   5520
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Coste real"
         Height          =   255
         Index           =   1
         Left            =   -73440
         TabIndex        =   122
         Top             =   5520
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Coste articulo"
         Height          =   255
         Index           =   0
         Left            =   -74760
         TabIndex        =   120
         Top             =   5520
         Width           =   1095
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00000080&
         BorderWidth     =   3
         X1              =   -74760
         X2              =   -70920
         Y1              =   5400
         Y2              =   5400
      End
      Begin VB.Label Label2 
         Caption         =   "Texto auxiliar documentos"
         Height          =   240
         Index           =   1
         Left            =   -74760
         TabIndex        =   114
         Top             =   4680
         Width           =   2655
      End
      Begin VB.Label Label2 
         Caption         =   "Texto para Ventas"
         Height          =   240
         Index           =   11
         Left            =   -74760
         TabIndex        =   83
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Texto para Compras"
         Height          =   240
         Index           =   2
         Left            =   -74760
         TabIndex        =   82
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Control de Instalación"
         Height          =   240
         Index           =   3
         Left            =   -74760
         TabIndex        =   81
         Top             =   3480
         Width           =   1575
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   -67080
         Picture         =   "frmAlmArticulos.frx":3162
         ToolTipText     =   "Buscar fecha"
         Top             =   1335
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha de Alta"
         Height          =   255
         Index           =   16
         Left            =   -68400
         TabIndex        =   67
         Top             =   1365
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Situación Artículo"
         Height          =   255
         Index           =   4
         Left            =   -68400
         TabIndex        =   66
         Top             =   2235
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo Asociación"
         Height          =   255
         Index           =   3
         Left            =   -68400
         TabIndex        =   65
         Top             =   945
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Dias de Garantia"
         Height          =   255
         Index           =   19
         Left            =   -68400
         TabIndex        =   64
         Top             =   2700
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "U p"
         Height          =   255
         Index           =   20
         Left            =   -68400
         TabIndex        =   63
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Artículo"
         Height          =   195
         Index           =   9
         Left            =   -74760
         TabIndex        =   62
         Top             =   2265
         Width           =   915
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Index           =   4
         Left            =   -73500
         Picture         =   "frmAlmArticulos.frx":36EC
         ToolTipText     =   "Buscar tipo artículo"
         Top             =   2265
         Width           =   240
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Index           =   5
         Left            =   -73500
         Picture         =   "frmAlmArticulos.frx":37EE
         ToolTipText     =   "Buscar tipo IVA"
         Top             =   2700
         Width           =   240
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Index           =   1
         Left            =   -73500
         Picture         =   "frmAlmArticulos.frx":38F0
         ToolTipText     =   "Buscar familia"
         Top             =   930
         Width           =   240
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Index           =   2
         Left            =   -73500
         Picture         =   "frmAlmArticulos.frx":39F2
         ToolTipText     =   "Buscar marca"
         Top             =   1380
         Width           =   240
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Index           =   0
         Left            =   -73500
         Picture         =   "frmAlmArticulos.frx":3AF4
         Tag             =   "-1"
         ToolTipText     =   "Buscar proveedor"
         Top             =   495
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Cod.  Proveedor"
         Height          =   255
         Index           =   5
         Left            =   -74760
         TabIndex        =   61
         Top             =   510
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Cod. Familia"
         Height          =   255
         Index           =   6
         Left            =   -74760
         TabIndex        =   60
         Top             =   945
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo de I.V.A."
         Height          =   255
         Index           =   8
         Left            =   -74760
         TabIndex        =   59
         Top             =   2700
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Cod. Marca"
         Height          =   255
         Index           =   7
         Left            =   -74760
         TabIndex        =   58
         Top             =   1380
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Unidad"
         Height          =   195
         Index           =   17
         Left            =   -74760
         TabIndex        =   57
         Top             =   1830
         Width           =   870
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Index           =   3
         Left            =   -73500
         Picture         =   "frmAlmArticulos.frx":3BF6
         ToolTipText     =   "Buscar tipo unidad"
         Top             =   1815
         Width           =   240
      End
   End
   Begin VB.Frame Frame2 
      Height          =   620
      Left            =   240
      TabIndex        =   70
      Top             =   410
      Width           =   11055
      Begin VB.ComboBox cboArticuloVarios 
         Height          =   315
         Left            =   9600
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Tag             =   "Artículo de Varios|N|N|||sartic|artvario||N|"
         Top             =   210
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   1
         Left            =   4040
         MaxLength       =   40
         TabIndex        =   1
         Tag             =   "Denominación Artículo|T|N|||sartic|nomartic||N|"
         Text            =   "Text1"
         Top             =   210
         Width           =   4245
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   0
         Left            =   1040
         MaxLength       =   16
         TabIndex        =   0
         Tag             =   "Código Artículo|T1|N|||sartic|codartic||S|"
         Text            =   "Text1"
         Top             =   240
         Width           =   1710
      End
      Begin VB.Label Label1 
         Caption         =   "Artículo Varios"
         Height          =   255
         Index           =   18
         Left            =   8490
         TabIndex        =   80
         Top             =   220
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Denominación"
         Height          =   255
         Index           =   1
         Left            =   2950
         TabIndex        =   72
         Top             =   220
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Código Art."
         Height          =   255
         Index           =   0
         Left            =   200
         TabIndex        =   71
         Top             =   225
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Height          =   475
      Index           =   0
      Left            =   240
      TabIndex        =   49
      Top             =   7800
      Width           =   3615
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   240
         Left            =   120
         TabIndex        =   50
         Top             =   180
         Width           =   3435
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   10560
      TabIndex        =   36
      Top             =   7920
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   9360
      TabIndex        =   35
      Top             =   7920
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   360
      Top             =   7680
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
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
      TabIndex        =   52
      Top             =   0
      Width           =   11775
      _ExtentX        =   20770
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
            ImageIndex      =   1
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
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Nuevo"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Stocks Almacenes"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Componentes"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Instalaciones"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cod. EAN"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Materias activas"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Equivalencias articulos"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
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
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   7320
         TabIndex        =   54
         Top             =   120
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc Data4 
      Height          =   330
      Left            =   1800
      Top             =   7800
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
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
      Left            =   10560
      TabIndex        =   39
      Top             =   7920
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Reservas"
      Height          =   195
      Left            =   4920
      TabIndex        =   195
      Top             =   7980
      Visible         =   0   'False
      Width           =   1695
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
   Begin VB.Menu mnMantenimientos 
      Caption         =   "&Mantenimientos"
      Begin VB.Menu mnMtoStocksAlm 
         Caption         =   "&Stocks Almacenes"
         Shortcut        =   ^K
      End
      Begin VB.Menu mnMtoConjuntos 
         Caption         =   "Conjuntos"
         Shortcut        =   ^J
      End
      Begin VB.Menu mnMtoInstalaciones 
         Caption         =   "&Instalaciones"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnCodEAN 
         Caption         =   "Codigos &EAN"
      End
      Begin VB.Menu mnMateriasActivas 
         Caption         =   "Materias &activas"
      End
      Begin VB.Menu mnEquivalencias 
         Caption         =   "Equivalencias de artículos"
      End
   End
End
Attribute VB_Name = "frmAlmArticulos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
' ==== Modificaciones:  =====
' ---- [14/09/2009] (LAURA)  --> Modificar funcion "InsertarPreciosPorTarifa2" para crear en función del parámetro
                                 '"creatarifart" solo tarifa generar o todas las tarifas para el articulo
                                 
'---- [23/09/2009] LAURA  --> Añadir lineas de Cod. EAN
'---- [02/11/2009] LAURA  --> abrir el form y situarse en solapa Documentos|Pedidos
' ===========================


Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
    'Si empieza por ::   - QUIERO VER el articulo que va despues de los dos puntos
    '               ··    -Quiero CREAR un articulo desde TELEMATEL
    '                            codprove|nomprove|refprove|precio|nomartic|ean|codtelem|



Public DeConsulta As Boolean 'Muestra Form para consulta, solo buscar y ver todos activos

'---- [02/11/2009] LAURA  --> abrir el form y situarse en solapa Documentos|Pedidos
Public parNumTAb As Byte 'nº de tab en el q queremos q se situe al abrir el form
'----


Public Event DatoSeleccionado(CadenaSeleccion As String)


Private WithEvents frmB As frmBuscaGrid 'Form para busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmB2 As frmBuscaGrid 'Form para busquedas
Attribute frmB2.VB_VarHelpID = -1
Private WithEvents frmF As frmCal 'Form Calendario Fecha
Attribute frmF.VB_VarHelpID = -1

Private WithEvents frmP As frmBasico2 '%=%=frmComProveedores
Attribute frmP.VB_VarHelpID = -1
Private WithEvents frmM As frmAlmMarcas 'Marcas de Artículos
Attribute frmM.VB_VarHelpID = -1
Private WithEvents frmTU As frmAlmTipoUnidad
Attribute frmTU.VB_VarHelpID = -1
Private WithEvents frmTA As frmAlmTipoArticulo
Attribute frmTA.VB_VarHelpID = -1
Private WithEvents frmFA As frmBasico2 'frmAlmFamiliaArticulo
Attribute frmFA.VB_VarHelpID = -1
Private WithEvents frmA As frmAlmAlPropios 'Almacenes Propios
Attribute frmA.VB_VarHelpID = -1
Private WithEvents frmUbic As frmAlmUbicaciones 'ubicaciones de almacen
Attribute frmUbic.VB_VarHelpID = -1
Private WithEvents frmCat As frmAlmCategorias 'categorias articulo (control de lotes(S/N))
Attribute frmCat.VB_VarHelpID = -1
Private WithEvents frmMAct As frmAlmMatAct
Attribute frmMAct.VB_VarHelpID = -1
Private WithEvents frmADR As frmAlmADR
Attribute frmADR.VB_VarHelpID = -1

Private Modo As Byte
'-----------------------------
'Se distinguen varios modos
'   0.-  Formulario limpio sin nungun campo rellenado
'   1.-  Preparando para hacer la busqueda
'   2.-  Ya tenemos registros y los vamos a recorrer
'        y podemos editarlos Edicion del campo
'   3.-  Insercion de nuevo registro
'   4.-  Modificar un registro
'   5.-  Mantenimiento Lineas de Articulos x Almacen
'   6.-  Mantenimiento Lineas de Componentes de Conjuntos
'   7.-  Mantenimiento Lineas de Control de Instalaciones
'   8.-  Mantenimiento Lineas de EAN
'   9.-  Mantenimiento Lineas de Materias activas
'   10.- Mantenimiento Lineas de EQUIVALENICAS
'-------------------------------------------------------------------------
'-------------------------------------------------------------------------
Private ModoAnterior As Byte

Private ModoFrame As Byte
'ModoFrame: 0.-Inicio, 3.-Insertar, 4.-Modificar

Private CadenaConsulta As String
'SQL de la tabla principal del formulario

Private Ordenacion As String
Private NombreTabla As String  'Nombre de la tabla o de la
Private kCampo As Integer

Dim btnAnyadir As Byte
'Variable que indica el número del Boton  Anyadir en la Toolbar1

Dim btnPrimero As Byte
'Variable que indica el número del Boton  PrimerRegistro en la Toolbar1

'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean

Dim ModificaLineas As Byte
'1.- Añadir,  2.- Modificar,  3.- Borrar,  0.-Pasar control a Lineas

Dim IntercalaComponente As Boolean

Dim PrimeraVez As Boolean

Private TagText3 As String

'NUEVO: JULIO 2007. PARA BUSCAR POR CHECKS TB
'------------------------------------------------
Private BuscaChekc As String

'NUevo: Nov 2008
' Hay un campo para pintar el PVP con el IVA
' Guardaremos el tipo de iva y el % (para no tener que recaluclarlo cada ve
Private mPorIva As String

Private PriVezForm As Boolean

'Cunado esta metiendo componentes, si es materia prima, y va por porcentajes
Private MateriaPrima As Boolean



Private Sub cboArticuloVarios_KeyPress(KeyAscii As Integer)
    If InstalacionEsEulerTaxco Then
        If KeyAscii = 13 Then
            PonerFoco Text1(3)
            KeyAscii = 0
        End If
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub cboCalidad_Click()
    cboCalidad_LostFocus
    If cboCalidad.ListIndex >= 0 Then PonerFoco txtAux(9)
End Sub

Private Sub cboCalidad_LostFocus()

    If Modo = 7 Then
        If ModificaLineas = 1 Then
            If cboCalidad.ListIndex > 0 Then txtAux(9).Text = DevuelveDesdeBD(conAri, "especificaciones", "scalidad", "codigo", cboCalidad.ItemData(cboCalidad.ListIndex))
        End If
        
    End If
        
End Sub

Private Sub cboStatus_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub cboTipoComiArtVario_KeyPress(KeyAscii As Integer)
     KEYpress KeyAscii
End Sub



Private Sub cboUnidadCompra_KeyPress(KeyAscii As Integer)
     KEYpress KeyAscii
End Sub

Private Sub chkConjunto_Click()
 If Modo = 1 Then CheckCadenaBusqueda chkConjunto, BuscaChekc
End Sub

Private Sub chkConjunto_GotFocus()
     ConseguirfocoChk Modo
End Sub

Private Sub chkConjunto_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkCtrStock_Click()
 If Modo = 1 Then CheckCadenaBusqueda chkctrstock, BuscaChekc

End Sub

Private Sub chkctrstock_GotFocus()
     ConseguirfocoChk Modo
End Sub

Private Sub chkctrstock_KeyPress(KeyAscii As Integer)
     KEYpress KeyAscii
End Sub

Private Sub chkInventario_Click()
    If Modo = 1 Then CheckCadenaBusqueda chkInventario, BuscaChekc
End Sub

Private Sub chkInventario_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkInventario_LostFocus()
    PonerFocoBtn Me.cmdAceptar
End Sub



Private Sub chkMateriaPrima_Click()
    If Modo = 1 Then CheckCadenaBusqueda chkMateriaPrima, BuscaChekc
End Sub

Private Sub chkMateriaPrima_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub chkMateriaPrima_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkRotacion_Click()
    If Modo = 1 Then CheckCadenaBusqueda chkRotacion, BuscaChekc
End Sub

Private Sub chkRotacion_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub chkRotacion_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkSeries_Click()
 If Modo = 1 Then CheckCadenaBusqueda chkSeries, BuscaChekc
End Sub

Private Sub chkSeries_GotFocus()
     ConseguirfocoChk Modo
End Sub

Private Sub chkSeries_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub



Private Sub chkWeb_Click()
    If Modo = 1 Then CheckCadenaBusqueda chkWeb, BuscaChekc
End Sub

Private Sub chkWeb_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub chkWeb_KeyPress(KeyAscii As Integer)
 KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()
Dim cad As String, Indicador As String
Dim bol As Boolean

    Screen.MousePointer = vbHourglass
    On Error GoTo Error1
    
    Select Case Modo
        Case 1  'BUSQUEDA
            HacerBusqueda
            
        Case 3 'INSERTAR
            If DatosOk Then
'              If InsertarDesdeForm(Me) Then
'                InsetarArticulosPorAlmacen
'                InsertarPreciosPorTarifa
                If InsertarArticulo Then
                       
                    'Si viene de telematel
                    '   1.- Inserto el an si lleva
                    '   2.- Salgo cagando leches
                    '   JUNIO 2014. No salgo. Hago lo del precio y despues SALGO
                    If Mid(Me.DatosADevolverBusqueda, 1, 2) = "··" Then ActualizarEAN
                        
                    '    ActualizarEAN
                    '    Unload Me
                    '    Exit Sub
                    'End If
                    
                    
                    
                    
                    'Para que salga por lo menos a Herbelca y EULER
                    If vParamAplic.NumeroInstalacion = vbHerbelca Or InstalacionEsEulerTaxco Then
                        'Si no es de varios
                        If cboArticuloVarios.ListIndex = 0 Then
                            'le pasamos codartic, nomartic codprove nomprove
                            frmComPreciosProv2.NuevoDato = Text1(0).Text & "|" & Text1(1).Text & "|" & Text1(2).Text & "|" & Text2(0).Text & "|"  'Para que no se poing en modo insercion
                            frmComPreciosProv2.Show vbModal
                        End If
                    End If
                        
                    'If InstalacionEsEulerTaxco Then
                    If vParamAplic.NumeroInstalacion = vbEuler Then
                        'En EULER quieren ver tambien los precios de venta
                        frmFacTarifasPrecios.codArtic = Text1(0).Text
                        frmFacTarifasPrecios.Show vbModal
                    End If
                        
                    If Mid(Me.DatosADevolverBusqueda, 1, 2) = "··" Then
                        'AHORA si que salgo
                        Unload Me
                        Exit Sub
                    End If
                    
                    PosicionarData
                    
                    
                    
                    
                    
                    
                    
                End If
'              End If
            End If
            
        Case 4  'MODIFICAR
            If DatosOk Then
                If ModificaDesdeFormulario(Me, 1) Then
                    TerminaBloquear
                    'comprobar si se ha modificado el precio ult. compra
                    'y preguntar si modificar PVP y Tarifas
                    If CCur(DBLet(Data1.Recordset!precioUC, "N")) <> ImporteFormateado(Text1(15).Text) Then
                        ActualizarPreciosVenta
                    ElseIf CCur(DBLet(Data1.Recordset!PrecioVe, "N")) <> ImporteFormateado(Text1(17).Text) Then
                        'Comprobar si se ha modificado el precio de venta PVP y preguntar
                        'si se quieren actualizar las tarifas de precios
                        ActualizarPreciosPorTarifa
                    ElseIf CCur(DBLet(Data1.Recordset!margecom, "N")) <> ImporteFormateado(Text1(25).Text) Then
                        'comprobar si se ha modificado el margen comercial
                        'y preguntar si modificar PVP y Tarifas
                         ActualizarPreciosVenta
                    End If
                    TerminaBloquear
'                    DesBloqueaRegistroForm Text1(0)
                    PosicionarData
                End If
            End If
                
         Case 5 'InsertarModificar linea  '----------------
         
            'Actualizar el registro en la tabla de lineas 'salmac' (Artículos x Almacen)
            If InsertarModificarLinea Then
'                DesBloqueaRegistroForm Text1(0)
      
                NumRegElim = data4.Recordset.AbsolutePosition
                TerminaBloquear
                LLamaLineas2 0, 0, 4
                DataGrid3.AllowAddNew = False
                CargaGrid Me.DataGrid3, Me.data4, True
                SituarDataPosicion data4, NumRegElim, Indicador
                
                lblIndicador.Caption = Indicador
                PonerModoFrame 0
                PonerSumaStocks
                
               
                
            End If
            
          '---- [23/09/2009] LAURA : Añadir lineas de Cod. EAN (se añade modo 8)
          Case 6, 7, 8, 9, 10
                '6: InsertarModificar Conjuntos
                '7: InsertarModificar Instalaciones
                '8: InsertarModificar cod. EAN
                '9: InsertarModificar mATERIAS ACTIVAS
                '10: InsertarModificar equivalencias
             If Modo = 6 Then bol = InsertarModificarConjunto
             If Modo = 7 Then bol = InsertarModificarInstalacion
             If Modo = 8 Then bol = InsertarModificarCodigosEAN
             If Modo = 9 Then bol = InsertarModificarMATACT
             If Modo = 10 Then bol = InsertarModificarEQUIV
             
             If bol Then
                TerminaBloquear
                If Modo = 6 Then 'Conjunto
                  txtAux(0).visible = False
                  txtAux(1).visible = False
                  txtAux2.visible = False
                  cmdAux.visible = False
                  CargaGrid Me.DataGrid1, Me.Data2, True
                ElseIf Modo = 7 Then 'Instalacion
                    txtAux(2).visible = False
                    CargaGrid Me.DataGrid2, Me.data3, True
                '---- [23/09/2009] LAURA : Añadir lineas de Cod. EAN (se añade modo 8)
                ElseIf Modo = 8 Then 'codigos EAN
                    txtAux(8).visible = False
                    CargaGrid Me.DataGrid4, Me.data5, True
                    
                ElseIf Modo = 9 Then 'materias activas
                  
                    CargaGrid Me.DataGrid5, Me.data6, True
                ElseIf Modo = 10 Then 'equivalencias
                  
                    CargaGrid Me.DataGrid6, Me.data7, True
                '----
                End If
                
                If ModificaLineas = 2 Then 'Modificar
                    DesBloqueaRegistroForm Text1(0)
                    If Modo = 6 Then
                        Data2.Recordset.Find (Data2.Recordset.Fields(1).Name & " =" & CInt(Me.cmdAceptar.Tag))
                    ElseIf Modo = 7 Then
                        data3.Recordset.Find (data3.Recordset.Fields(1).Name & " =" & CInt(Me.cmdAceptar.Tag))
                    '---- [23/09/2009] LAURA : Añadir lineas de Cod. EAN (se añade modo 8)
                    ElseIf Modo = 8 Then
                        data5.Recordset.Find (data5.Recordset.Fields(1).Name & " =" & CInt(Me.cmdAceptar.Tag))
                    '----
                    End If
                    PonerBotonCabecera True
'                    Me.lblIndicador.Caption = ""
                    PonerFocoBtn Me.cmdAceptar
                    ModificaLineas = 0
                ElseIf ModificaLineas = 1 Then 'Insertar
                    IntercalaComponente = False
                    BotonAnyadirConjunto2
                End If
            End If
    End Select
    Screen.MousePointer = vbDefault
 
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub




Private Sub cmdActualizarImportes1_Click(Index As Integer)
Dim frmAr As frmAlmArticulos

    If Modo <> 2 Then Exit Sub
    MsgBox "Aqui"
    If ModificaLineas <> 0 Then
        MsgBox "Esta cambiando datos", vbExclamation
        Exit Sub
    End If
    
    If Index = 0 Then
        If txtConjunto(1).Text = "" Or txtConjunto(1).Text = "" Then
            MsgBox "Falta importes calculados", vbExclamation
            Exit Sub
        End If
        BuscaChekc = "¿Desea cambiar los importes PVP y UPC del árticulo principal?"
        If MsgBox(BuscaChekc, vbQuestion + vbYesNo) = vbNo Then Exit Sub
    End If
    
    
    If Index = 0 Then
        'ACtualizar importes
    
        'Haremos lo siguiente
        If BLOQUEADesdeFormulario(Me) Then
            'Fijaremos los nuevos importes
             
             If ModificarImportesDesdeConjuntos Then
                    TerminaBloquear
                    Text1(15).Text = Me.txtConjunto(1).Text
                    Text1(17).Text = Me.txtConjunto(4).Text
                    'comprobar si se ha modificado el precio ult. compra
                    'y preguntar si modificar PVP y Tarifas
                    If CCur(DBLet(Data1.Recordset!precioUC, "N")) <> ImporteFormateado(Text1(15).Text) Then ActualizarPreciosVenta
                    'Comprobar si se ha modificado el precio de venta PVP y preguntar
                    'si se quieren actualizar las tarifas de precios
                    If CCur(DBLet(Data1.Recordset!PrecioVe, "N")) <> ImporteFormateado(Text1(17).Text) Then ActualizarPreciosPorTarifa
                    

                    PosicionarData
            End If
        End If
    Else
        'VER ARTICULO LINEA
        Set frmAr = New frmAlmArticulos
        frmAr.DeConsulta = True
        frmAr.DatosADevolverBusqueda = "::" & DevNombreSQL(Data2.Recordset!codarti1)
        frmAr.Show vbModal
        Set frmAr = Nothing
        
        'Por si acaso ha cambiado
        'recargo el grid
        '--------------------------------------------------------------------------------------
        NumRegElim = Data2.Recordset.AbsolutePosition - 1
        
        CancelaADODC Me.Data2
        CargaGrid Me.DataGrid1, Me.Data2, True
        CancelaADODC Me.Data2
        ponerDatosConjuntos
        If NumRegElim > 0 Then Data2.Recordset.Move NumRegElim, 1
        
    End If
    BuscaChekc = ""
End Sub

Private Function ModificarImportesDesdeConjuntos() As Boolean
    On Error GoTo EM
    ModificarImportesDesdeConjuntos = False
    BuscaChekc = "UPDATE sartic set precioUC = " & TransformaComasPuntos(CStr(ImporteFormateado(Me.txtConjunto(1).Text)))
    BuscaChekc = BuscaChekc & " , preciove =" & TransformaComasPuntos(CStr(ImporteFormateado(Me.txtConjunto(4).Text)))
    BuscaChekc = BuscaChekc & " WHERE codartic = '" & DevNombreSQL(Data1.Recordset!codArtic) & "'"
    conn.Execute BuscaChekc
    ModificarImportesDesdeConjuntos = True
    Exit Function
EM:
    MuestraError Err.Number, "", Err.Description
End Function


Private Sub cmdAlma_Click()
    imgCuentas_Click 6
End Sub

Private Sub cmdAux_Click()
    MandaBusquedaPrevia " conjunto=0 "
    PonerFoco txtAux(1)
End Sub

Private Sub cmdCancelar_Click()
On Error Resume Next
    Select Case Modo
        Case 1 'Busqueda
            LimpiarCampos
            PonerModo 0
            PonerFoco Text1(0)
            
        Case 3 'Insertar
            If ModoAnterior = 0 Then
                LimpiarCampos
                PonerModo 0
            Else
                LimpiarCampos
                PonerModo 2
                PonerCampos
            End If
                
        Case 4  'Modificar
            TerminaBloquear
'            DesBloqueaRegistroForm Text1(0)
            PonerModo 2
            PonerCampos
            PonerFoco Text1(0)
        
        
        'QUITAR### TODO ESTE COMENTARIO ELIMINADO LAS LINEAS
'        Case 5 'Lineas Detalle
''            DesBloqueoManual NombreTabla
'            TerminaBloquear
'            PonerModoFrame 0
'            PonerCamposAlmacenes2
'            ModificaLineas = 0
'            PonerFoco Text3(1)
        
        '---- [23/09/2009] LAURA : Añadir lineas de Cod. EAN (se añade modo 8)
        Case 5, 6, 7, 8, 9, 10 'Lineas Conjuntos, Lineas Instalaciones
            ModificaLineas = 0
'            DesBloqueoManual NombreTabla
            TerminaBloquear
            Select Case Modo
            Case 5
                DataGrid3.AllowAddNew = False
                DataGrid2.Enabled = False
                PonerModoFrame 0
                LLamaLineas2 0, 0, 4
                NumRegElim = data4.Recordset.AbsolutePosition
                CargaGrid DataGrid3, data4, True
                SituarDataPosicion data4, NumRegElim, Me.lblIndicador.Caption
                If Not data4.Recordset.EOF Then PonerCamposAlmacenes2
                
            Case 6
                txtAux(0).visible = False
                txtAux(1).visible = False
                txtAux2.visible = False
                cmdAux.visible = False
                DataGrid1.AllowAddNew = False
                If Not (ModificaLineas = 2) Then 'Modificar
                    If Not Data2.Recordset.EOF Then Data2.Recordset.MoveFirst
                End If
                DataGrid1.Enabled = True
                txtAux2.BackColor = &H80000005
            Case 7
                txtAux(2).visible = False
                txtAux(9).visible = False
                txtAux(10).visible = False
                txtAux(11).visible = False
                Me.cboCalidad.visible = False
                DataGrid2.AllowAddNew = False
                If Not (ModificaLineas = 2) Then 'Modificar
                    If Not data3.Recordset.EOF Then data3.Recordset.MoveFirst
                End If
                DataGrid2.Enabled = True
            
                 
            '---- [23/09/2009] LAURA : Añadir lineas de Cod. EAN (se añade modo 8)
            Case 8 'Lineas codigos EAN
                txtAux(8).visible = False
                DataGrid4.AllowAddNew = False
                If Not (ModificaLineas = 2) Then 'Modificar
                    If Not data5.Recordset.EOF Then data5.Recordset.MoveFirst
                End If
                DataGrid4.Enabled = True
                
            Case 9
                'Lineas materias activas
                LLamaLineas2 0, 0, 6
                DataGrid5.AllowAddNew = False
                If Not (ModificaLineas = 2) Then 'Modificar
                    If Not data6.Recordset.EOF Then data6.Recordset.MoveFirst
                End If
                DataGrid5.Enabled = True
            Case 10
                'Lineas materias activas
                LLamaLineas2 0, 0, 7
                DataGrid6.AllowAddNew = False
                If Not (ModificaLineas = 2) Then 'Modificar
                    If Not data7.Recordset.EOF Then data7.Recordset.MoveFirst
                End If
                DataGrid6.Enabled = True
            '----
            End Select
            IntercalaComponente = False
            PonerBotonCabecera True
            PonerFocoBtn Me.cmdRegresar
    End Select
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub BotonAnyadir()

     If vUsu.Nivel2 = 2 Then Exit Sub


    LimpiarCampos 'Vacía los TextBox
    ModoAnterior = Modo 'Para el botón Cancelar en Modo Insertar
    PonerModo 3
    Me.SSTab1.Tab = 0
    
    'Poner valores por defecto
    Me.chkctrstock.Value = 1 'por defecto hay control de stock
    Me.Text1(10).Text = Format(Now, "dd/mm/yyyy") 'fecha alta
    Me.cboArticuloVarios.ListIndex = 0
    Me.cboStatus.ListIndex = 0
    cboADV.ListIndex = 0
    Me.Text1(11).Text = "0"
    Me.Text1(12).Text = "1"
    Me.chkMateriaPrima.Value = 0
    If vParamAplic.NumeroInstalacion = 2 Then cboTipoComiArtVario.ListIndex = 0
    cboUnidadCompra.ListIndex = 0
    
    If InstalacionEsEulerTaxco Then
        'EULER
        BuscaChekc = DevuelveDesdeBD(conAri, "concat(codprove,'|',nomprove,'|')", "sprove", "codprove in (select proveedorsartic from  eulerparam) AND 1", "1")
        If BuscaChekc <> "" Then
            Text1(2).Text = RecuperaValor(BuscaChekc, 1)
            Text2(0).Text = RecuperaValor(BuscaChekc, 2)
            BuscaChekc = ""
        End If
    
        PonerFoco Text1(1)
        
    Else
        PonerFoco Text1(0)
    End If
    
End Sub


Private Sub BotonAnyadirLinea()
Dim vWhere As String
    
    'Si no estaba modificando lineas salimos
    ' Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 2 Then Exit Sub
    Me.SSTab1.Tab = 4
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModoFrame 3    '3: Insertar
    ModificaLineas = 1 'Insertar

    'Obtenemos la siguiente numero de Artículo
    vWhere = "codartic=" & DBSet(Text1(0).Text, "T")
    Text3(0).Text = SugerirCodigoSiguienteStr("salmac", "codalmac", vWhere)
    lblIndicador.Caption = "INSERTAR ALMACEN"
    PonerFoco Text3(0)
End Sub





Private Sub BotonAnyadirConjunto2()
Dim NumF As String
Dim vWhere As String
Dim anc As Single
    
    'Si no estaba modificando lineas salimos
    ' Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 2 Then Exit Sub
        
        
        
        
    'Noviembre 2016
    'Quitamos esta prohibicion para que en los albaranes salgan varios plazos de seguriradad
    'If vParamAplic.NumeroInstalacion = 1 Then
    '    'En Alzira, en el rpt de facturas y albaranes, esta enlazando los ltoes con esta tabla,
    '    'si hubieran mas de una mataeria actuiva por cada articulo, los informes saldrian mal
    '    'al unir una linea de slifaclotes/Slialblotes con mas de una mtaria activa
    '    vWhere = DevuelveDesdeBD(conAri, "count(*)", "sarti5", "codartic", Text1(0).Text, "T")
    '    If Val(vWhere) > 0 Then
    '        MsgBox "No es posible asignar mas de una materia activa a un producto fitosanitario. Consulte soporte tecnico", vbCritical
    '        Exit Sub
    '    End If
    'End If
        
        
    
    ModificaLineas = 1
    PonerBotonCabecera False
    
    vWhere = "codartic=" & DBSet(Text1(0).Text, "T")
'    ancIni = 200
    Select Case Modo
    Case 5 'Lineas STOCK
        Me.SSTab1.Tab = 4
        lblIndicador.Caption = "INSERTAR STOCK"
        NumF = 1
        
    Case 6
        If IntercalaComponente Then
            NumF = Data2.Recordset!numlinea
        Else
            NumF = SugerirCodigoSiguienteStr("sarti1", "numlinea", vWhere)
        End If
        Me.SSTab1.Tab = 2
        lblIndicador.Caption = "INSERTAR CONJUNTO"
        
    Case 7 'Lineas Instalaciones
        NumF = SugerirCodigoSiguienteStr("sarti2", "numlinea", vWhere)
        Me.SSTab1.Tab = 3
        lblIndicador.Caption = "INSERTAR INSTALACIÓN"
        If vParamAplic.NumeroInstalacion = vbFontenas Then
            txtAux(9).Text = ""
            txtAux(10).Text = ""
            txtAux(11).Text = ""
        End If
    '---- [23/09/2009] LAURA : Añadir lineas de Cod. EAN
    Case 8 'Lineas cod. EAN
        NumF = SugerirCodigoSiguienteStr("sarti3", "numlinea", vWhere)
        Me.SSTab1.Tab = 5
        lblIndicador.Caption = "INSERTAR COD. EAN"
        
    Case 9 'Materias activas
        
        Me.SSTab1.Tab = 7
        lblIndicador.Caption = "INSERTAR MAT. ACTIVAS"
        
    Case 10 'Equivalencias
        
        Me.SSTab1.Tab = 5
        lblIndicador.Caption = "INSERTAR ART. EQUIVALENCIAS"
    '----
    End Select
    cmdAceptar.Tag = NumF
    
    Select Case Modo
    'If Modo = 6 Then 'Conjuntos
    Case 5 'Lineas STOCK
        PonerDatosForaGrid True
        PonerModoFrame 3
        AnyadirLinea DataGrid3, data4
        anc = ObtenerAlto(DataGrid3, 20)
        LLamaLineas2 anc, 1, 4
        PonerFoco Text3(0)
        BloquearTxt Text3(0), False
        
    Case 6

        txtAux(0).Text = ""
        txtAux2.Text = ""
        txtAux(1).Text = ""
        'Situamos el grid al final
        AnyadirLinea DataGrid1, Data2

        anc = ObtenerAlto(DataGrid1, 20)
        LLamaLineas2 anc, 1, 2
        
        BloquearTxt txtAux(0), False
        Me.cmdAux.Enabled = True
        PonerFoco txtAux(0)
        
        If IntercalaComponente Then
            lblIndicador.Caption = "I N T E R C A L A R"
            txtAux2.BackColor = &HC0E0FF
        Else
            txtAux2.BackColor = &H80000005
        End If
        
        
    Case 7 'Lineas INSTALACIONES
        Me.txtAux(2).Text = ""
        AnyadirLinea DataGrid2, data3
        anc = ObtenerAlto(DataGrid2, 20)
        LLamaLineas2 anc, 1, 3
        anc = 1
        If vParamAplic.NumeroInstalacion = vbFontenas Then
            If ModificaLineas = 1 Then anc = 1
        End If
        If anc = 1 Then
            PonerFocoCbo Me.cboCalidad
        Else
            PonerFoco txtAux(2)
        End If

    '---- [23/09/2009] LAURA : Añadir lineas de Cod. EAN
    Case 8 'Lineas cod. EAN
        Me.txtAux(8).Text = ""
        AnyadirLinea DataGrid4, data5
        anc = ObtenerAlto(DataGrid4, 20)
        LLamaLineas2 anc, 1, 5
        PonerFoco txtAux(8)

    Case 9
        ' 19/12/2011
        'materias activas
        Me.Text5(0).Text = ""
        Text5(1).Text = ""
        AnyadirLinea DataGrid5, data6
        anc = ObtenerAlto(DataGrid5, 20)
        LLamaLineas2 anc, 1, 6
        PonerFoco Text5(0)
        PonerFoco Text5(0)
        
    Case 10
        ' 23/feb/2012
        'Equivalenicas
        Me.Text6(0).Text = ""
        Text6(1).Text = ""
        AnyadirLinea DataGrid6, Me.data7
        anc = ObtenerAlto(DataGrid6, 20)
        LLamaLineas2 anc, 1, 7
        PonerFocoBtn Me.cmdEquiv
        PonerFoco Text6(0)
       
    '----
    End Select
End Sub


Private Sub BotonBuscar()
'Buscar
    LimpiarCampos
    If Modo <> 1 Then 'Modo 1: Busqueda
        BuscaChekc = ""
        PonerModo 1
        'Si pasamos el control aqui lo ponemos en amarillo
        PonerFoco Text1(0)
        Text1(0).BackColor = vbYellow
        'Si es de buscqueda , buscamos solo activos
        If DeConsulta Then Me.cboStatus.ListIndex = 0
    Else
        If DeConsulta Then
            If cboStatus.ListIndex < 0 Then cboStatus.ListIndex = 0
        End If
        HacerBusqueda
        If Data1.Recordset.EOF Then
            Text1(kCampo).Text = ""
            Text1(kCampo).BackColor = vbYellow
            PonerFoco Text1(kCampo)
        End If
    End If
End Sub


Private Sub BotonVerTodos()
Dim C As String
  
    LimpiarCampos
    If chkVistaPrevia.Value = 1 Then
        C = ""
        If DeConsulta Then C = "codstatu = 0"
    
        MandaBusquedaPrevia C
    Else
        C = "Select * from " & NombreTabla
        If DeConsulta Then C = C & " WHERE codstatu = 0 "
        CadenaConsulta = C & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Sub Desplazamiento(Index As Integer)
'Botones de Desplazamiento de la Toolbar
    Select Case Modo
        Case 5 'Modo Mantenimiento de Almacenes (Lineas)
            If data4.Recordset.EOF Then Exit Sub
            DesplazamientoData data4, Index
            PonerCamposAlmacenes2
        Case Else 'Datos de Cabecera
            If Data1.Recordset.EOF Then Exit Sub
            DesplazamientoData Data1, Index
            PonerCampos
            PonerModoOpcionesMenu (Modo) 'Poner opciones de menu según modo
            PonerOpcionesMenu   'Activar opciones de menu según nivel
                                'de permisos del usuario
    End Select
End Sub


Private Sub BotonModificar()
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 4
    PonerFoco Text1(1)
End Sub



Private Sub BotonModificarConjunto(ByRef vDataGrid As DataGrid, ByRef vData As Adodc)
Dim anc As Single
Dim i As Integer

    If vData.Recordset.EOF Then Exit Sub
    If vData.Recordset.RecordCount < 1 Then Exit Sub
   
    'Si no estaba modificando lineas salimos
    ' Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Then Exit Sub

    Screen.MousePointer = vbHourglass
    PonerBotonCabecera False
         
    If vDataGrid.Bookmark < vDataGrid.FirstRow Or vDataGrid.Bookmark > (vDataGrid.FirstRow + vDataGrid.VisibleRows - 1) Then
        i = vDataGrid.Bookmark - vDataGrid.FirstRow
        vDataGrid.Scroll 0, i
        vDataGrid.Refresh
    End If
    PonerFocoBtn Me.cmdAceptar
    vDataGrid.Enabled = False

    anc = ObtenerAlto(vDataGrid, 20)
    
    If Modo = 5 Then
        cmdAceptar.Tag = vData.Recordset!codAlmac
    Else
        
        cmdAceptar.Tag = vData.Recordset!numlinea
    End If
    
    Select Case Modo
    Case 5
        PonerModoFrame 4 'ModoFrame=4 -> Modificar
        Me.lblIndicador.Caption = "MODIFICAR ALMACEN"
        LLamaLineas2 anc, 2, 4
        BloquearTxt Text3(0), True
        Text3(0).Text = data4.Recordset!codAlmac
        Text3(2).Text = data4.Recordset!CanStock
        Text2(8).Text = data4.Recordset!nomalmac
        PonerFoco Text3(1)

    Case 6
        MateriaPrima = CStr(DBLet(vData.Recordset!MateriaPrima, "T")) = "*"
    ' If Modo = 6 Then 'Componentes de Conjunto
        Me.lblIndicador.Caption = "MODIFICAR CONJUNTO"
        Me.SSTab1.Tab = 2
         'Llamamos al form
        txtAux(0).Text = DataGrid1.Columns(2).Text
        'Feb 2011.   No bloqueamos el codartic
        'BloquearTxt txtAux(0), True
        Me.txtAux2.Text = DataGrid1.Columns(3).Text
        txtAux(1).Text = DataGrid1.Columns(4).Text
        LLamaLineas2 anc, 2, 2
        PonerFoco txtAux(1)
        'Feb 2011.   No bloqueamos el codartic
        'If ModificaLineas = 2 Then cmdAux.Enabled = False
    'Poner el foco
    'ElseIf Modo = 7 Then
    Case 7
        Me.lblIndicador.Caption = "MODIFICAR INSTALACIÓN"
        Me.SSTab1.Tab = 3
        If vParamAplic.NumeroInstalacion = vbFontenas Then
            txtAux(9).Text = DataGrid2.Columns(2).Text
            txtAux(10).Text = DataGrid2.Columns(3).Text
            txtAux(11).Text = DataGrid2.Columns(4).Text
        Else
            txtAux(2).Text = DataGrid2.Columns(2).Text
        End If
        LLamaLineas2 anc, 2, 3
        
        PonerFoco txtAux(IIf(vParamAplic.NumeroInstalacion = vbFontenas, 9, 2))
        
    '---- [23/09/2009] LAURA : Añadir lineas de Cod. EAN (se añade modo 8)
    Case 8 'Lineas cod. EAN
        Me.lblIndicador.Caption = "MODIFICAR COD. EAN"
        Me.SSTab1.Tab = 5
        txtAux(8).Text = DataGrid4.Columns(2).Text
        LLamaLineas2 anc, 2, 5
        PonerFoco txtAux(8)
    End Select
    Screen.MousePointer = vbDefault
End Sub


Private Sub BotonEliminar()
Dim SQL As String



     If vUsu.Nivel2 = 2 Then Exit Sub

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
    'No esta bloqueado
    If Val(Data1.Recordset!codstatu) = 1 Then
        MsgBox "Articulo bloqueado", vbExclamation
        Exit Sub
    End If
    
    
    'Tiene stock
    If ImporteFormateado(txtSumaStock.Text) <> 0 Then
        MsgBox "El articulo tiene stock", vbExclamation
        Exit Sub
    End If
    

    
    BuscaChekc = lblIndicador.Caption
    SQL = SePuedeEliminarArticulo(CStr(Data1.Recordset!codArtic), lblIndicador)
    lblIndicador.Caption = BuscaChekc
    BuscaChekc = ""
    If SQL <> "" Then
        SQL = "No se puede eliminar el articulo: " & Data1.Recordset!codArtic & vbCrLf & vbCrLf & SQL
        MsgBox SQL, vbExclamation
        Exit Sub
    End If
    SQL = "Cabecera de Artículos." & vbCrLf
    SQL = SQL & "---------------------------        " & vbCrLf & vbCrLf
    SQL = SQL & "Va a eliminar el Artículo:"
    SQL = SQL & vbCrLf & "Cod. Artic. :   " & Data1.Recordset.Fields(0)
    SQL = SQL & vbCrLf & "Descripción :   " & Data1.Recordset.Fields(1)
    SQL = SQL & vbCrLf & vbCrLf & " ¿Desea Eliminarlo? "
    
    If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        On Error GoTo Error2
        TerminaBloquear
        Screen.MousePointer = vbHourglass
        NumRegElim = Data1.Recordset.AbsolutePosition
        If Not Eliminar Then
            Screen.MousePointer = vbDefault
            Exit Sub
        ElseIf SituarDataTrasEliminar(Data1, NumRegElim) Then
            PonerModo 2
            PonerCampos
        Else
            LimpiarCampos
            PonerModo 0
        End If
    End If
    Screen.MousePointer = vbDefault
    
Error2:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Articulo", Err.Description
End Sub


Private Sub BotonEliminarLinea()
'Eliminar una linea De ArticulosxAlmacen
Dim cad As String

     On Error GoTo Error2

    If data4.Recordset.EOF Then Exit Sub
    If data4.Recordset.RecordCount < 1 Then Exit Sub
    If vUsu.Nivel > 1 Then Exit Sub
    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Or ModificaLineas = 2 Then Exit Sub '1= Insertar, 2=Modificar
    ModificaLineas = 3 'Eliminar
    
    '### a mano
    cad = "Seguro que desea eliminar de la BD el registro:"
    cad = cad & vbCrLf & "Cod. Artículo: " & Data1.Recordset.Fields(0)
    cad = cad & vbCrLf & "Cod. Almacen: " & data4.Recordset.Fields(1)

    'Borramos
    If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
       
        Screen.MousePointer = vbHourglass
        NumRegElim = data4.Recordset.AbsolutePosition
        
        cad = "DELETE FROM salmac where codartic = '" & DevNombreSQL(Data1.Recordset.Fields(0)) & "' AND codalmac = " & data4.Recordset!codAlmac
        conn.Execute cad
        
        CargaGrid Me.DataGrid3, Me.data4, True
        If data4.Recordset.EOF Then
            'Solo habia un registro
            LimpiarCamposAlmacenes
            PonerModoFrame 0
        Else
            SituarDataPosicion Me.data4, NumRegElim, cad
            PonerCamposAlmacenes2
        End If
        ModificaLineas = 0
    End If
    Screen.MousePointer = vbDefault
    
Error2:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then
        data4.Recordset.CancelUpdate
        MuestraError Err.Number, "Eliminar Linea de Articulo", Err.Description
    End If
End Sub


Private Sub BotonEliminarConjunto()
Dim SQL As String
    On Error GoTo Error2
    
    'Ciertas comprobaciones
    If Data2.Recordset.EOF Then Exit Sub
    
    SQL = "Seguro que desea eliminar el Componente de Conjunto:"
    SQL = SQL & vbCrLf & "Código: " & Data2.Recordset!codarti1
    SQL = SQL & vbCrLf & "Descripción: " & Data2.Recordset.Fields(3)
    
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then
        'Hay que eliminar
        SQL = "Delete from sarti1 where codartic=" & DBSet(Data2.Recordset!codArtic, "T")
        SQL = SQL & " and numlinea=" & Data2.Recordset!numlinea
        SQL = SQL & " and codarti1=" & DBSet(Data2.Recordset!codarti1, "T")
        conn.Execute SQL
        CancelaADODC Me.Data2
        CargaGrid Me.DataGrid1, Me.Data2, True
        CancelaADODC Me.Data2
    End If
    
Error2:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Componente de Conjunto", Err.Description
End Sub


Private Sub BotonEliminarInstalacion()
Dim SQL As String
    On Error GoTo Error2

    'Ciertas comprobaciones
    If data3.Recordset.EOF Then Exit Sub
    
    If vParamAplic.NumeroInstalacion = vbFontenas Then
        SQL = "Seguro que desea eliminar el registro:"
        SQL = SQL & vbCrLf & "Ensayo: " & data3.Recordset!ensayo
        SQL = SQL & vbCrLf & "Especificación: " & data3.Recordset!especificaciones
    Else
        SQL = "Seguro que desea eliminar el control de instalación:"
        SQL = SQL & vbCrLf & "Linea: " & data3.Recordset!numlinea
        SQL = SQL & vbCrLf & "Descripción: " & data3.Recordset!licontro
    End If
    If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        If vParamAplic.NumeroInstalacion = vbFontenas Then
            SQL = "Delete from sarti7 where codartic=" & DBSet(Data1.Recordset!codArtic, "T")
            SQL = SQL & " and codigoensayo=" & data3.Recordset!numlinea
        Else
            SQL = "Delete from sarti2 where codartic=" & DBSet(data3.Recordset!codArtic, "T")
            SQL = SQL & " and numlinea=" & data3.Recordset!numlinea
        End If
        conn.Execute SQL
        CancelaADODC Me.data3
        CargaGrid Me.DataGrid2, Me.data3, True
        CancelaADODC Me.data3
    End If
    
Error2:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Control de Instalaciones", Err.Description
End Sub



Private Sub BotonEliminarMateriaActiva()
Dim SQL As String
    On Error GoTo Error2

    'Ciertas comprobaciones
    If data6.Recordset.EOF Then Exit Sub
    
    SQL = "Seguro que desea eliminar la materia activa:"
    SQL = SQL & vbCrLf & "Linea: " & data6.Recordset!codigoma
    SQL = SQL & vbCrLf & "Descripción: " & data6.Recordset!nombrema
    
    If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        SQL = "Delete from sarti5 where codartic=" & DBSet(Data1.Recordset!codArtic, "T")
        SQL = SQL & " and Codigoma=" & data6.Recordset!codigoma
        conn.Execute SQL
        CancelaADODC Me.data5
        CargaGrid Me.DataGrid5, Me.data6, True
        'DataGrid5, data6, enlaza
        CancelaADODC Me.data6
    End If
    
Error2:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar materia activa", Err.Description
End Sub



'---- [23/09/2009] LAURA : Añadir lineas de Cod. EAN
Private Sub BotonEliminarCodigosEAN()
Dim SQL As String
    On Error GoTo ErrElimEAN

    'Ciertas comprobaciones
    If data5.Recordset.EOF Then Exit Sub
    
    SQL = "Seguro que desea eliminar el codigo EAN:"
    SQL = SQL & vbCrLf & "Linea: " & data5.Recordset!numlinea
    SQL = SQL & vbCrLf & "Cod. EAN: " & data5.Recordset!codigoea
    
    If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        SQL = "Delete from sarti3 where codartic=" & DBSet(data5.Recordset!codArtic, "T")
        SQL = SQL & " and numlinea=" & data5.Recordset!numlinea
        conn.Execute SQL
        CancelaADODC Me.data5
        CargaGrid Me.DataGrid4, Me.data5, True
        CancelaADODC Me.data5
    End If
    Exit Sub
    
ErrElimEAN:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Eliminar Codigos EAN", Err.Description
End Sub
'----

Private Sub BotonEliminarEquivalencia()
Dim SQL As String
    On Error GoTo Error2

    'Ciertas comprobaciones
    If data7.Recordset.EOF Then Exit Sub
    
    SQL = "Seguro que desea eliminar la equivalencia con el"
    SQL = SQL & vbCrLf & "Articulo: " & data7.Recordset!codarti1
    SQL = SQL & vbCrLf & "Descripción: " & Trim(data7.Recordset!NomArtic)
    
    If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        SQL = "Delete from sarti6 where codartic=" & DBSet(Data1.Recordset!codArtic, "T")
        SQL = SQL & " and codArti1=" & DBSet(data7.Recordset!codarti1, "T")
        conn.Execute SQL
        
        'Borramos la "viceversa"
        SQL = "Delete from sarti6 where codartic=" & DBSet(data7.Recordset!codarti1, "T")
        SQL = SQL & " and codArti1=" & DBSet(Data1.Recordset!codArtic, "T")
        conn.Execute SQL
        
        CancelaADODC Me.data7
        CargaGrid Me.DataGrid6, Me.data7, True
        'DataGrid5, data6, enlaza
        CancelaADODC Me.data7
    End If
    
Error2:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar equivalencia", Err.Description
End Sub



Private Sub BotonArticulosxAlmac()
    Screen.MousePointer = vbHourglass
    On Error GoTo ErrorArticAlmac
    
    Screen.MousePointer = vbHourglass
    'RESTAURO LOS tag's
    AccionesSobreTagText3_ False, False

    Me.SSTab1.Tab = 4
    PonerModo (5)
    PonerBotonCabecera True
    PonerFocoBtn Me.cmdRegresar
    Screen.MousePointer = vbDefault
    
    
    
'ANTEs ------------------------------------------------------
'
'    'Crear las lineas de Articulos x Almacen para el artículo
'    Me.SSTab1.Tab = 0
'
'    'ASignamos un SQL al DATA4
''    Me.Data4.ConnectionString = Conn
''    Data4.RecordSource = "Select * from salmac where codartic = '" & Text1(0).Text & "';"
''    Data4.Refresh
'
'    If Data4.Recordset.RecordCount <= 0 Then
'        MsgBox "No hay ningún registro en la tabla salmac", vbInformation
'        Screen.MousePointer = vbDefault
'        Exit Sub
'    Else
'        'Poner el modo en el formulario
'        PonerModo (5) 'Modo 5: Modificar lineas
'        PonerModoFrame 0 'TextBox Bloqueados inicialmente
'
'        'Data4.Recordset.MoveFirst
'        'PonerCamposAlmacenes
'        'PonerFocoBtn Me.cmdRegresar
'        Screen.MousePointer = vbDefault
'    End If
    Exit Sub
    
ErrorArticAlmac:
    MuestraError Err.Number, "PonerCadenaBusqueda", Err.Description
    PonerModo 0
    Screen.MousePointer = vbDefault
End Sub


Private Sub BotonConjuntos()
    On Error GoTo ErrorConjuntos
    
    Screen.MousePointer = vbHourglass
    
    If DBLet(Data1.Recordset!Conjunto, "N") = 1 Then
        
        Me.SSTab1.Tab = 2
        
        
        
        PonerModo (6)
        PonerBotonCabecera True
        
        DataGrid1.Enabled = True
        Me.DataGrid1.SetFocus
    
    Else
        MsgBox "El artículo no tiene la marca de """ & Me.chkConjunto.Caption & """", vbExclamation
    End If
    
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrorConjuntos:
    MuestraError Err.Number, "Conjuntos", Err.Description
    Screen.MousePointer = vbDefault
End Sub


Private Sub BotonInstalaciones()
    On Error GoTo ErrorInstala

    Screen.MousePointer = vbHourglass
    Me.SSTab1.Tab = 3
    PonerModo (7)
    PonerBotonCabecera True
    PonerFocoBtn Me.cmdRegresar
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrorInstala:
    MuestraError Err.Number, "Instalaciones", Err.Description
    Screen.MousePointer = vbDefault
End Sub



Private Sub BotonCodigosEAN()
    On Error GoTo ErrEAN
    Screen.MousePointer = vbHourglass
    
    Me.SSTab1.Tab = 5
    PonerModo (8)
    PonerBotonCabecera True
    PonerFocoBtn Me.cmdRegresar
    
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrEAN:
    MuestraError Err.Number, "Codigos EAN", Err.Description
    Screen.MousePointer = vbDefault
End Sub


Private Sub BotonMateriasActivas()
    On Error GoTo ErrorConjuntos
    
    If vParamAplic.Ariagro = "" Then Exit Sub 'por si acaso
    
    Screen.MousePointer = vbHourglass
    Me.SSTab1.Tab = 7
    
    PonerModo (9)
    PonerBotonCabecera True
    
    DataGrid5.Enabled = True
    Me.DataGrid5.SetFocus
    Screen.MousePointer = vbDefault
    
    
    Exit Sub
    
ErrorConjuntos:
    MuestraError Err.Number, "Conjuntos", Err.Description
    Screen.MousePointer = vbDefault
End Sub



Private Sub BotonEquivalencias()
    On Error GoTo ErrorConjuntos
    
    If Modo <> 2 Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    Me.SSTab1.Tab = 5
    
    PonerModo (10)
    PonerBotonCabecera True
    
    DataGrid6.Enabled = True
    Me.DataGrid6.SetFocus
    Screen.MousePointer = vbDefault
    
    
    Exit Sub
    
ErrorConjuntos:
    MuestraError Err.Number, "Equivalencias", Err.Description
    Screen.MousePointer = vbDefault
End Sub





Private Sub cmdEquiv_Click()

    AbreFrmBuscaGrid_El2 1
    
End Sub


Private Sub AbreFrmBuscaGrid_El2(DesdeEquivalencias As Byte)

    BuscaChekc = ParaGrid(Text1(0), 23, "Código")
    BuscaChekc = BuscaChekc & ParaGrid(Text1(1), 58, "Denominación")
    BuscaChekc = BuscaChekc & ParaGrid(Text1(9), 19, "Cod. asoc.")
    
    Screen.MousePointer = vbHourglass
    Set frmB2 = New frmBuscaGrid
    frmB2.vCampos = BuscaChekc
    frmB2.vTabla = "sartic"
    frmB2.vSQL = ""
    
    '###A mano
    frmB2.vDevuelve = "0|1|"
    frmB2.vTitulo = "Artículos"
    frmB2.vselElem = 1
    frmB2.vConexionGrid = conAri

    frmB2.vCargaFrame = False
    BuscaChekc = ""
    frmB2.Show vbModal
    Set frmB2 = Nothing
    If BuscaChekc <> "" Then
    
        If DesdeEquivalencias = 1 Then
            Text6(0).Text = RecuperaValor(BuscaChekc, 1)
            Text6(1).Text = RecuperaValor(BuscaChekc, 2)
            
            
        ElseIf DesdeEquivalencias = 2 Then
            Text1(33).Text = RecuperaValor(BuscaChekc, 1)
            Text2(9).Text = RecuperaValor(BuscaChekc, 2)
            
        Else
            DoEvents
            Screen.MousePointer = vbHourglass
            pPdfRpt = "Select codprove,codfamia,codmarca,codunida,codtipar,codigiva,margecom,preciove,"
            pPdfRpt = pPdfRpt & "referprov, unicajas, unicajas2, nomartic"
            pPdfRpt = pPdfRpt & ", CtrStock"
            pPdfRpt = pPdfRpt & " FROM sartic where codartic=" & DBSet(RecuperaValor(BuscaChekc, 1), "T")
            Set miRsAux = New ADODB.Recordset
            miRsAux.Open pPdfRpt, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            'NO PUEDE SER EOF
            BuscaChekc = "2|3|4|5|6|7|25|17|31|12|34|1|"
            For pRptvMultiInforme = 1 To miRsAux.Fields.Count - 1
                 
                pPdfRpt = RecuperaValor(BuscaChekc, pRptvMultiInforme)
                If IsNull(miRsAux.Fields(pRptvMultiInforme - 1)) Then
                    Text1(CInt(pPdfRpt)).Text = ""
                Else
                    Text1(CInt(pPdfRpt)).Text = miRsAux.Fields(pRptvMultiInforme - 1)
                End If
                '   los seis primeros
                If pRptvMultiInforme < 8 Then Text1_LostFocus CInt(pRptvMultiInforme)
            Next
            'Controstock. Campo 12 de rs
            chkctrstock.Value = miRsAux!CtrStock
            
            PonerFoco Text1(1)
            Text1(1).SelStart = Len(Text1(1).Text)
            
            miRsAux.Close
            Set miRsAux = Nothing
            pRptvMultiInforme = 0
            pPdfRpt = ""
            
            Screen.MousePointer = vbDefault
            
        End If
        BuscaChekc = ""
    End If
    
End Sub


Private Sub cmdEuler_Click()
    AbreFrmBuscaGrid_El2 0
End Sub

Private Sub cmdGenerar_Click()
    Dim Aux As String
    Aux = Text2(2) & " " & Text2(1) & " " & Text2(4) & " " & Text2(3)
    Text1(1).Text = Replace(Left(Aux, 40), "*", "")
    Text1(0).Text = SugerirCodAutomatico(Text1(4), Text1(3), Text1(6), Text1(5))
End Sub

Private Sub cmdMatAux_Click()
    'Materia activa
    Set frmMAct = New frmAlmMatAct
    frmMAct.DatosADevolverBusqueda = "0"
    frmMAct.Show vbModal
    Set frmMAct = Nothing
End Sub

Private Sub cmdRegresar_Click()
'Este es el boton Cabecera
Dim cad As String

    'Quitar lineas y volver a la cabecera
    '---- [23/09/2009] LAURA : Añadir lineas de Cod. EAN (se añade modo 8)
    'If Modo = 5 Or Modo = 6 Or Modo = 7 Or Modo = 8 Then
    If Modo >= 5 And Modo <= 10 Then
        If Modo = 6 Then
            'Componentes
            If vParamAplic.ComponentePorcentaje Then
                'Son porcentajes. Compruebo que la suma es 100
                If Not ComprobarPorcentajesCorrectos Then Exit Sub
            End If
        End If
        'modo 5: Lineas Articulos x Almacen
        'modo 6: Lineas Conjuntos
        'modo 7: Lineas Instalaciones
        'modo 8: Lineas cod. EAN
        Me.lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
        If DataGrid1.Row >= 0 Then
            DeseleccionaGrid Me.DataGrid1
            DataGrid1.Bookmark = 1
        End If
        If DataGrid2.Row >= 0 Then
            DeseleccionaGrid Me.DataGrid2
            DataGrid2.Bookmark = 1
        End If
        PonerModo 2

    Else 'Se llamo desde un botón de Prismático
        If Data1.Recordset.EOF Then
            MsgBox "Ningún registro devuelto.", vbExclamation
            Exit Sub
        End If
        If DeConsulta Then
            If cboStatus.ListIndex > 0 Then
                MsgBox "Articulo " & cboStatus.Text, vbExclamation
                Exit Sub
            End If
        End If
            
        cad = Data1.Recordset.Fields(0) & "|"
        cad = cad & Data1.Recordset.Fields(1) & "|"
        cad = cad & Data1.Recordset.Fields(8).Value & "|"
        cad = cad & Text2(4).Text & "|"
        RaiseEvent DatoSeleccionado(cad)
        Unload Me
    End If
End Sub




Private Sub Data4_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    If Modo = 5 And ModificaLineas > 0 Then Exit Sub
    If Not data4.Recordset.EOF Then
        If Not PrimeraVez Then PonerCamposAlmacenes2
    End If
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        If ModificaLineas = 0 Then
            PonerFocoBtn Me.cmdRegresar
        Else
            SendKeys "{tab}"
        End If
    End If
End Sub


Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    If Modo = 1 Then PonerFoco Text1(1)
    
    If PriVezForm Then
        PriVezForm = False
        
        

        
        'He abierto el form queriendo cargar un articulo
        If Mid(DatosADevolverBusqueda, 1, 2) = "::" Then
            DatosADevolverBusqueda = Mid(DatosADevolverBusqueda, 3)
            CadenaConsulta = "Select * from " & NombreTabla & " where codartic='" & DatosADevolverBusqueda & "'"
            PonerCadenaBusqueda
            
            If Me.chkConjunto.Value > 0 And vUsu.Nivel <= 1 Then
                Toolbar1.Buttons(11).Enabled = True
                Me.mnMtoConjuntos.Enabled = True
                If Me.parNumTAb = 6 Then parNumTAb = 1
            End If
            
            If Me.parNumTAb = 6 Then
                Toolbar2.Buttons(7).Value = tbrPressed
                Toolbar2_ButtonClick Toolbar2.Buttons(7)
            End If
            
            cmdRegresar.Cancel = True
         Else
            If Mid(DatosADevolverBusqueda, 1, 2) = "··" Then
                BotonAnyadir
                
                'Acabo de cargar los datos
                BuscaChekc = Mid(DatosADevolverBusqueda, 3)
                
                'Llevaremos codprove|nomprove|refprove|precio|ean|
                Text1(2).Text = RecuperaValor(BuscaChekc, 1)
                Text2(0).Text = RecuperaValor(BuscaChekc, 2)
                Text1(31).Text = RecuperaValor(BuscaChekc, 3)
                Text1(1).Text = RecuperaValor(BuscaChekc, 5)
                Text1(15).Text = RecuperaValor(BuscaChekc, 4)
                Text1(9).Text = RecuperaValor(BuscaChekc, 7)
                
                Text1(18).Text = Format(Now, "dd/mm/yyyy")
                Text1(25).Text = "0.00"
                Text1(17).Text = Text1(15).Text
                PonerFoco Text1(0)
                Text1(0).Text = Text1(31).Text
                
                BuscaChekc = ""
            End If
         End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    KEYpress KeyAscii
End Sub


Private Sub Form_Load()
    PriVezForm = True
    'Icono del formulario
    Me.Icon = frmPpal.Icon
    
    ' ICONITOS DE LA BARRA
    btnAnyadir = 6
    btnPrimero = 21 'Posicion del Boton Primero en la toolbar (+ 3 siguientes)
    With Me.Toolbar1
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 1   'Botón Buscar
        .Buttons(2).Image = 2   'Botón Todos
        .Buttons(6).Image = 3   'Insertar Nuevo
        .Buttons(7).Image = 4   'Modificar
        .Buttons(8).Image = 5   'Borrar
        .Buttons(10).Image = 10 'Stocks Almacenes
        .Buttons(11).Image = 11 'Conjuntos
        .Buttons(12).Image = 36 'Instalaciones
        .Buttons(13).Image = 23 'Cod. EAN '---- [23/09/2009] LAURA : Añadir lineas de Cod. EAN
        .Buttons(14).Image = 25 'Materias activas
        .Buttons(15).Image = 47 'Equivalencias
        
        .Buttons(18).Image = 16  'Imprimir
        .Buttons(19).Image = 15  'Salir
        .Buttons(btnPrimero).Image = 6  'Primero
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
        .Buttons(btnPrimero + 3).Image = 9 'Último
    End With
    
    
    'Como en un futuro se parametrizaran el numero de decimales...
    Text1(29).Tag = "Litros x Ud|N|S|||sartic|LitrosUnidad|" & FormatoCantidad & "|N|"
    
    
    If Me.parNumTAb > 0 Then
        Me.SSTab1.Tab = Me.parNumTAb
    Else
        Me.SSTab1.Tab = 0
    End If
    Me.SSTab1.TabVisible(2) = False
    Me.SSTab1.TabVisible(3) = False
    
    
    If vParamAplic.NumeroInstalacion = vbFontenas Then
        Me.SSTab1.TabCaption(3) = "Control calidad"
        Me.Toolbar1.Buttons(12).ToolTipText = "Calidad"
        Me.DataGrid2.Width = 9375
    End If
    
    chkWeb.visible = vParamAplic.NumeroInstalacion = 1
    
    SSTab1.TabVisible(7) = vParamAplic.Ariagro <> ""
    Me.mnMateriasActivas.visible = vParamAplic.Ariagro <> ""
    Me.Toolbar1.Buttons(14).visible = vParamAplic.Ariagro <> ""

    cboADV.visible = vParamAplic.NumeroInstalacion = 1
    Label1(41).visible = vParamAplic.NumeroInstalacion = 1
    'Documentos articulo. Cantidad reservada. En los que tienen produccion
    If vParamAplic.Produccion Then Label4(0).Caption = "En produccion"
    
    If vParamAplic.NumeroInstalacion = 2 Then
        'HERBELCA
        cboTipoComiArtVario.visible = True
        Label1(44).visible = True
        CargarComboComisionArticulosVarios
        
        Label1(20).Caption = "Ud. embalaje grande"
        Label1(42).Caption = "Ud. embalaje pequeña"
        
    Else
        'Resto
        Label1(20).Caption = "Unidades caja"
        Label1(42).Caption = "Ud embalaje"
        cboTipoComiArtVario.visible = False
        Label1(44).visible = False
    End If
    CargarComboTipoUnidadesCompra
    cboUnidadCompra.visible = True

    
    LimpiarCampos   'Limpia los campos TextBox
    PrimeraVez = True
    
        
    'Marzo 2015
    'FrameServicios.visible = vParamAplic.Servicios
    FrameServicios.visible = vParamAplic.Ariagro <> ""
    
    
    If InstalacionEsEulerTaxco Then
        'En EULER, ni codprove, ni refereprov SE VEN
        'Pero se insertan etc etc, por lo tanto los pongo "lejos" y en el zorder los paso al final
        Label1(5).visible = False
        Label1(38).visible = False
        imgCuentas(0).visible = False
        Text2(0).visible = False
        'Los txt no puedo ocultarlos

        Text1(2).Left = 13000
        Text1(31).Left = 13000
        Text1(2).TabIndex = 300
        Text1(31).TabIndex = 301
    End If
    
    
    FrameLitrosUd.visible = vParamAplic.Descriptores
    
    
    framePortes.visible = False
    If vParamAplic.NumeroInstalacion = vbFenollar Then
        framePortes.visible = True
    Else
        If vParamAplic.TipoPortes = 1 Then framePortes.visible = True
    End If
    
    
    
    Me.cmdActualizarImportes1(0).Enabled = vUsu.Nivel <= 1
    Me.cmdActualizarImportes1(1).Enabled = vUsu.Nivel <= 1
    
    
    'Si hay algun combo los cargamos
    CargarComboStatus
    CargarComboArticuloVarios
    CargarComboADV
    
    'El tag de los stocks
    AccionesSobreTagText3_ True, True
    
    'Pone el Tag del primer botón de busqueda de cuentas a -1
    'Si tag =-1 abre busqueda en la tabla: sartic, BD: Ariges
    'Si tag>0 abre busqueda en la tabla asociada al indice.
    imgCuentas(0).Tag = "-1"
         
    '## A mano
    NombreTabla = "sartic"
    Ordenacion = " ORDER BY codartic"
  
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = conn
    Data1.RecordSource = "Select * from " & NombreTabla & " where codartic='-1' "
    Data1.Refresh
    
    'Los usuarios/agentes NO pueden ver los precios ni la solapa de movimiento
    'En herbelca
    FrameDatosAlmacen2.visible = True
    SSTab1.TabVisible(6) = True
    If vParamAplic.NumeroInstalacion = 2 Then
        FrameDatosAlmacen2.visible = vUsu.CodigoAgente = 0
        SSTab1.TabVisible(6) = vUsu.CodigoAgente = 0
    End If
    
    If DatosADevolverBusqueda = "" Then
        PonerModo 0
        PonerCamposLineas False
    Else
        If DatosADevolverBusqueda = "@1@" Then 'Poner Modo Busqueda
            BotonBuscar
        Else 'Poner Modo Insertar
            If Mid(DatosADevolverBusqueda, 1, 2) = "::" Then
                'Abrimos el articulo poniendo un articulo especificado a continuacion
                
                'Lo haremos en el ACTIVATE
            Else
                PonerModo 3
                If Mid(DatosADevolverBusqueda, 1, 2) = "··" Then   'INSERTA desde telematel
                    'En el activate
                Else
                      'PonerModo 3  'lo hace arriba del IF
                      Text1(0).Text = DatosADevolverBusqueda
                End If
            End If
        End If
    End If
    
    Label6.visible = (vParamAplic.NumeroInstalacion = vbFenollar)
    txtReser.visible = (vParamAplic.NumeroInstalacion = vbFenollar)
    
    '-- Descriptores especiales y botón de composición (Rafa VRS 4.0.9)
    If vParamAplic.Descriptores Then
        'cmdGenerar.visible = True  estara en poner modo
        Label1(6) = "Cod. Categoria"
        Label1(9) = "Cod. Modelo"
        Label1(17) = "Cod. Formato"
        '-- Aqui cambiamos los tag para evitar lios.
        CambiaTagDescriptores Text1(3), "Cod. Categoria"
        CambiaTagDescriptores Text1(5), "Cod. Formato"
        CambiaTagDescriptores Text1(6), "Cod. Modelo"
    Else
        cmdGenerar.visible = False
    End If
    '--
    ImagenesNavegacion
    If InstalacionEsEulerTaxco Then
        Toolbar2.Buttons(13).Value = tbrPressed
        CargaColumnas 5  'Solo para euler
    Else
        CargaColumnas 0
    End If
End Sub


Private Sub LimpiarCampos()
    limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador.Caption = ""
    'CargaGrid Me.DataGrid3, Me.data4, False  'Desenlazamos el GRID
    PonerCamposLineas False
    'Aqui va el especifico de cada form es
    Me.chkConjunto.Value = 0
    Me.chkSeries.Value = 0
    chkRotacion.Value = 0
    Me.chkctrstock.Value = 0
    Me.chkMateriaPrima.Value = 0
    Me.chkWeb.Value = 0
    Me.cboArticuloVarios.ListIndex = -1
    Me.cboStatus.ListIndex = -1
    txtReser.Text = ""
    If cboADV.visible Then cboADV.ListIndex = -1
    If Me.cboCalidad.visible Then cboCalidad.ListIndex = -1
    If vParamAplic.NumeroInstalacion = 2 Then cboTipoComiArtVario.ListIndex = -1
    cboUnidadCompra.ListIndex = -1
End Sub


Private Sub LimpiarCamposAlmacenes()
Dim i As Byte
    Text3(0).BackColor = vbRed
    For i = 0 To Text3.Count - 1
        Text3(i).Text = ""
    Next i
    Text2(8).Text = ""
    Me.chkInventario.Value = 0
    lblIndicador.Caption = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    Me.parNumTAb = 0
End Sub


Private Sub frmA_DatoSeleccionado(CadenaSeleccion As String)
'Almacenes Propios
    Text3(0).Text = RecuperaValor(CadenaSeleccion, 1)
    FormateaCampo Text3(0)
    Text2(8).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmADR_DatoSeleccionado(CadenaSeleccion As String)
    Text1(32).Text = RecuperaValor(CadenaSeleccion, 1)
    FormateaCampo Text1(32)
    Text2(7).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
Dim cadB As String
Dim Aux As String
Dim indice As Integer
      
    If CadenaDevuelta <> "" Then
        If Val(imgCuentas(0).Tag) >= 0 Then
            'Se llama desde el botón de busqueda del campo Tipos de IVA
            'Recuperar solo el campo código y Descripción
            HaDevueltoDatos = True
            Screen.MousePointer = vbHourglass
    
            indice = Val(Me.imgCuentas(0).Tag)
            Text1(indice + 2).Text = RecuperaValor(CadenaDevuelta, 1)
            Text2(indice).Text = RecuperaValor(CadenaDevuelta, 2)
        Else
            HaDevueltoDatos = True
            Screen.MousePointer = vbHourglass
            
            If Modo <> 6 Then
                'Recupera todo el registro de Artículos
                'Sabemos que campos son los que nos devuelve
                'Creamos una cadena consulta y ponemos los datos
                cadB = ""
                Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
                cadB = Aux
                CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
                PonerCadenaBusqueda
            Else
                'Llamamos desde el boton auxiliar de Conjuntos
                txtAux(0).Text = RecuperaValor(CadenaDevuelta, 1)
                txtAux2.Text = RecuperaValor(CadenaDevuelta, 2)
            End If
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub frmB2_Selecionado(CadenaDevuelta As String)
    BuscaChekc = CadenaDevuelta
End Sub

Private Sub frmCat_DatoSeleccionado(CadenaSeleccion As String)
'Categoria del Articulo
    Text1(22).Text = RecuperaValor(CadenaSeleccion, 1)
    Text2(22).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmF_Selec(vFecha As Date) 'Calendario Fechas

    Select Case Val(imgFecha(0).Tag)
        Case 0
            Text1(10).Text = Format(vFecha, "dd/mm/yyyy")
        Case 1
            Text1(18).Text = Format(vFecha, "dd/mm/yyyy")
        Case 2
            Text3(7).Text = Format(vFecha, "dd/mm/yyyy")
            
        Case 3
            Text1(24).Text = Format(vFecha, "dd/mm/yyyy")
    End Select
End Sub


Private Sub frmFA_DatoSeleccionado(CadenaSeleccion As String)
'Familia de Articulo
    Text1(3).Text = RecuperaValor(CadenaSeleccion, 1)
    FormateaCampo Text1(3)
    Text2(1).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmM_DatoSeleccionado(CadenaSeleccion As String)
'Mantenimiento de Marcas
    Text1(4).Text = RecuperaValor(CadenaSeleccion, 1)
    FormateaCampo Text1(4)
    Text2(2).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMAct_DatoSeleccionado(CadenaSeleccion As String)
    Text5(0).Text = RecuperaValor(CadenaSeleccion, 1)
    FormateaCampo Text5(0)
    Text5(1).Text = RecuperaValor(CadenaSeleccion, 2)
    
End Sub

Private Sub frmP_DatoSeleccionado(CadenaSeleccion As String)
'Proveedores
    Text1(2).Text = RecuperaValor(CadenaSeleccion, 1)
    FormateaCampo Text1(2)
    Text2(0).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmTA_DatoSeleccionado(CadenaSeleccion As String)
'Tipo de Articulo
    Text1(6).Text = RecuperaValor(CadenaSeleccion, 1)
    Text2(4).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmTU_DatoSeleccionado(CadenaSeleccion As String)
'Tipo de Unidad
    Text1(5).Text = RecuperaValor(CadenaSeleccion, 1)
    FormateaCampo Text1(5)
    Text2(3).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmUbic_DatoSeleccionado(CadenaSeleccion As String)
'Mto Ubicaciones de almacen
    Text3(1).Text = RecuperaValor(CadenaSeleccion, 1)
    Text2(6).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub imgAyuda_Click(Index As Integer)
    imgayuda(Index).Tag = ""
    If Index = 0 Then
        imgayuda(Index).Tag = imgayuda(Index).ToolTipText & vbCrLf & vbCrLf
        imgayuda(Index).Tag = imgayuda(Index).Tag & "Stock:   cantidad stock total actual" & vbCrLf
        If vParamAplic.Produccion Then
            imgayuda(Index).Tag = imgayuda(Index).Tag & "En produccion: " & vbCrLf
            imgayuda(Index).Tag = imgayuda(Index).Tag & "        Cantidad en produccion pendiente de cerrar. " & vbCrLf
            imgayuda(Index).Tag = imgayuda(Index).Tag & "               En postivo las cantidades a producir" & vbCrLf
            imgayuda(Index).Tag = imgayuda(Index).Tag & "               En negativo si es componente en produccion"
            If vParamAplic.NumeroInstalacion <> vbAmesa Then
                imgayuda(Index).Tag = imgayuda(Index).Tag & vbCrLf & "               Restará la cantidad pendiente de servir. Ped. cliente"
            End If
        Else
            imgayuda(Index).Tag = imgayuda(Index).Tag & "Reservas:  cantidad pendiente de servir en pedidos cliente"
        End If
        imgayuda(Index).Tag = imgayuda(Index).Tag & vbCrLf & "Pedido prov:  cantidad pedido proveedor "
          
        imgayuda(Index).Tag = imgayuda(Index).Tag & vbCrLf & vbCrLf & "Disponible: la suma de las cantidades"
    End If
    MsgBox imgayuda(Index).Tag, vbInformation
    
    imgayuda(Index).Tag = ""
End Sub

Private Sub imgCuentas_Click(Index As Integer)

    If Modo = 2 Or Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass

    Select Case Index
        Case 0 'Codigo Proveedor
'            Set frmP = New frmComProveedores
'            frmP.DatosADevolverBusqueda = "0"
'            frmP.Show vbModal
            Set frmP = New frmBasico2
            AyudaProveedores frmP, Text1(2)
            Set frmP = Nothing
        Case 1  'Cod. Familia
'            Set frmFA = New frmAlmFamiliaArticulo
'            frmFA.DatosADevolverBusqueda = "0"
'            frmFA.Show vbModal
            Set frmFA = New frmBasico2
            AyudaFamilias frmFA, Text1(Index + 2)
            Set frmFA = Nothing
        Case 2  'Cod. Marca
            Set frmM = New frmAlmMarcas
            frmM.DatosADevolverBusqueda = "0"
            frmM.Show vbModal
            Set frmM = Nothing
        Case 3  'Cod. Tipo Unidad
            Set frmTU = New frmAlmTipoUnidad
            frmTU.DatosADevolverBusqueda = "0"
            frmTU.Show vbModal
            Set frmTU = Nothing
        Case 4  'Cod. Tipo Articulo
            Set frmTA = New frmAlmTipoArticulo
            frmTA.DatosADevolverBusqueda = "0"
            frmTA.Show vbModal
            Set frmTA = Nothing
            
        Case 5  'Tipos de IVA. Tabla de la BD Contabilidad
            imgCuentas(0).Tag = Index
            MandaBusquedaPrevia ""
            imgCuentas(0).Tag = -1
            
        Case 6 'Código de Almacen
            Set frmA = New frmAlmAlPropios
            frmA.DatosADevolverBusqueda = "0"
            frmA.Show vbModal
            Set frmA = Nothing
            
        Case 7 'cod. ubicaciones
            Set frmUbic = New frmAlmUbicaciones
            frmUbic.DatosADevolverBusqueda = "0"
            frmUbic.Show vbModal
            Set frmUbic = Nothing
            
        Case 8 'cod. categoria
            Set frmCat = New frmAlmCategorias
            frmCat.DatosADevolverBusqueda = "0"
            frmCat.Show vbModal
            Set frmCat = Nothing
            
        Case 9
            'ADR
            Set frmADR = New frmAlmADR
            frmADR.DatosADevolverBusqueda = "0"
            frmADR.Show vbModal
            Set frmADR = Nothing
        
        
        Case 10
            AbreFrmBuscaGrid_El2 2
    End Select
    
    If Index = 6 Then
        PonerFoco Text3(0)
    ElseIf Index = 7 Then
        PonerFoco Text3(1)
    ElseIf Index = 8 Then
        PonerFoco Text1(22)
    ElseIf Index = 9 Then
        PonerFoco Text1(32)
    ElseIf Index = 10 Then
        'NADA
    Else
        PonerFoco Text1(Index + 2)
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub imgFecha_Click(Index As Integer)
Dim indice As Byte

   If Modo = 2 Or Modo = 0 Then Exit Sub
   Screen.MousePointer = vbHourglass
   imgFecha(0).Tag = Index
   
   Set frmF = New frmCal
   frmF.Fecha = Now
   
   Select Case Index
     Case 0, 1, 3
        If Index = 0 Then
            indice = 10
        ElseIf Index = 1 Then
            indice = 18
        Else
            indice = 24
        End If
        PonerFormatoFecha Text1(indice)
        If Text1(indice).Text <> "" Then frmF.Fecha = CDate(Text1(indice).Text)

     Case 2
        PonerFormatoFecha Text3(7)
         If Text3(7).Text <> "" Then frmF.Fecha = CDate(Text3(7).Text)
   End Select
   Screen.MousePointer = vbDefault
   frmF.Show vbModal
   Set frmF = Nothing
End Sub



Private Sub lw1_DblClick()
Dim Seleccionado As Long
Dim SQL As String
    If Modo <> 2 Then Exit Sub
    If lw1.ListItems.Count = 0 Then Exit Sub
    If lw1.SelectedItem Is Nothing Then Exit Sub


    If Me.DatosADevolverBusqueda <> "" Then
        'De momento NO dejo continuar
        MsgBox "Esta buscando un cliente. No puede ver los documentos.", vbExclamation
        Exit Sub
    End If
    
    
    Screen.MousePointer = vbHourglass
    
    'Llegados aqui
    Select Case CByte(RecuperaValor(lw1.Tag, 1))
    Case 0, 1, 2
        
    Case 3
        If lw1.SelectedItem.SmallIcon = 6 Then
            'PEDIDO CLIENTE
            If vParamAplic.TipoFormularioClientes = 0 Then
            
                frmFacEntPedidos.DatosADevolverBusqueda2 = lw1.SelectedItem.Text
                frmFacEntPedidos.EsHistorico = False
                frmFacEntPedidos.Show vbModal
            Else
                frmFacEntPedSail.DatosADevolverBusqueda2 = lw1.SelectedItem.Text
                frmFacEntPedSail.EsHistorico = False
                frmFacEntPedSail.Show vbModal
            End If
        Else
            'PROVEEDOR
            If vParamAplic.TipoFormularioClientes = 0 Then
                frmComEntPedidos2.MostrarDatos = lw1.SelectedItem.Text
                frmComEntPedidos2.EsHistorico = False
                frmComEntPedidos2.Show vbModal
            Else
                'SAIL
                 frmComEntPedidosSa.MostrarDatos = lw1.SelectedItem.Text
                frmComEntPedidosSa.EsHistorico = False
                frmComEntPedidosSa.Show vbModal
            End If
        End If
    Case 4
        'Deberia ver el o lo k siese
        DataGrid1EnSMOVAL
    End Select
        
    'Pase lo que pase, por si acaso, cargamos el lw
    lw1.SetFocus
    Seleccionado = lw1.SelectedItem.Index
    CargaDatosLW
    If Not lw1.SelectedItem Is Nothing Then lw1.SelectedItem.Selected = False
    Set lw1.SelectedItem = Nothing
    If lw1.ListItems.Count >= Seleccionado Then
            lw1.ListItems(Seleccionado).Selected = True
            lw1.ListItems(Seleccionado).EnsureVisible
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnCodEAN_Click()
    'EAN
    mnMtoCodigosEAN_Click
End Sub

Private Sub mnEliminar_Click()
    Select Case Modo
        Case 5  'Eliminar lineas Artículos x Almacen
            BotonEliminarLinea
        Case 6 'Eliminar Líneas Conjuntos
            BotonEliminarConjunto
        Case 7 'Eliminar Lineas de Control de Instalacion
            BotonEliminarInstalacion
            
        '---- [23/09/2009] LAURA : Añadir lineas de Cod. EAN
        Case 8 'Eliminar lineas de codigos EAN
            BotonEliminarCodigosEAN
        '----
        Case 9
            BotonEliminarMateriaActiva
        Case 10
            BotonEliminarEquivalencia
        Case Else   'Eliminar Artículo
            BotonEliminar
    End Select
End Sub


Private Sub mnEquivalencias_Click()
    BotonEquivalencias
End Sub

Private Sub mnMateriasActivas_Click()
    'Materias activas
    BotonMateriasActivas
End Sub

Private Sub mnModificar_Click()
Dim cad As String
Dim Aux As String
Dim i As Integer


     If vUsu.Nivel2 = 2 Then Exit Sub

    Select Case Modo
        Case 5  'Modificar lineas Artículos x Almacen
'                cad = Text1(0).Text
'                i = InStr(1, cad, """")
'                If i > 0 Then
'                    Aux = Mid(cad, 1, i)
'                    Aux = Aux & """"
'                    Aux = Aux & Mid(cad, i + 1, Len(cad))
'                End If
'                NombreSQL cad
'                If BloqueoManual(NombreTabla, "'" & cad & "|" & Text3(0).Text & "|'") Then BotonModificarLinea
                Aux = " codartic=" & DBSet(Text1(0).Text, "T")
                If BloqueaRegistro("sartic", Aux) Then BotonModificarConjunto Me.DataGrid3, Me.data4
                
                
        Case 6 'Modificar Líneas Conjuntos
'                If BloqueoManual(NombreTabla, "|'" & Text1(0).Text & "'|" & txtAux(0).Text & "|") Then
'                    BotonModificarConjunto Me.DataGrid1, Me.Data2
'                End If
                Aux = " codartic=" & DBSet(Text1(0).Text, "T")
                If BloqueaRegistro("sartic", Aux) Then BotonModificarConjunto Me.DataGrid1, Me.Data2
                
                
        Case 7  'Modificar Linea de Control de Instalacion
'                If BloqueoManual(NombreTabla, "|'" & Text1(0).Text & "'|" & cmdAceptar.Tag & "|") Then
'                    BotonModificarConjunto Me.DataGrid2, Me.Data3
'                End If
                Aux = " codartic=" & DBSet(Text1(0).Text, "T")
                If BloqueaRegistro("sartic", Aux) Then BotonModificarConjunto Me.DataGrid2, Me.data3
                
        '---- [23/09/2009] LAURA : Añadir lineas de Cod. EAN (se añade modo 8)
        Case 8 'Modificar linea codigos EAN
            Aux = " codartic=" & DBSet(Text1(0).Text, "T")
            If BloqueaRegistro("sartic", Aux) Then BotonModificarConjunto Me.DataGrid4, Me.data5
            
            
        Case 9
            'Modificar linea materias activas. NO se puede
            'MsgBox "Elimine e inserte la nueva", vbExclamation
            Exit Sub
            'Aux = " codartic=" & DBSet(Text1(0).Text, "T")
            'If BloqueaRegistro("sartic", Aux) Then BotonModificarConjunto Me.DataGrid4, Me.Data5
            
        Case 10
            'No se modifica, o de alta o de baja
        Case Else   'Modificar Artículos
            If BLOQUEADesdeFormulario(Me) Then BotonModificar
'            If BloqueaRegistroForm(Me) Then BotonModificar
    End Select
End Sub


Private Sub mnMtoConjuntos_Click()
    BotonConjuntos
End Sub

Private Sub mnMtoInstalaciones_Click()
    BotonInstalaciones
End Sub

Private Sub mnMtoStocksAlm_Click()
    BotonArticulosxAlmac
End Sub


Private Sub mnMtoCodigosEAN_Click()
    BotonCodigosEAN
End Sub


Private Sub mnNuevo_Click()
     Select Case Modo
        'Case 5 'Añadir lineas Artículos x Almacen
         '       BotonAnyadirLinea   'QUITAR EL PROCEDEIMIENTO
         
        '---- [23/09/2009] LAURA : Añadir lineas de Cod. EAN (se añade modo 8)
        Case 5, 6, 7, 8, 9, 10 'Añadir Líneas Conjuntos
                  'Añadir Linea de Control de Instalacion
                BotonAnyadirConjunto2
        Case Else 'Añadir Artículos
                BotonAnyadir
    End Select
End Sub

Private Sub mnSalir_Click()
    Screen.MousePointer = vbDefault
    If Modo = 5 Then
        '------------------------------------------------------
        'Si esta insertando lineas es una cosa, si no es otra
        cmdCancelar_Click
    Else
        If (Modo = 6) Or (Modo = 7) Then 'Modo 5: Mto Lineas
                        'Modo 6: Conjuntos, Modo 7: Instalaciones
                        
            cmdRegresar_Click
        Else
            Unload Me
        End If
    End If
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub

Private Sub SSTab1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


'### A mano
'Los metodos del text tendran que estar
'Los descomentamos cuando esten puestos ya los controles
Private Sub Text1_GotFocus(Index As Integer)
    kCampo = Index
    If (Not Text1(Index).MultiLine) And (Text1(Index).ScrollBars) = 0 Then ConseguirFoco Text1(Index), Modo
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
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

    'Si modo=1 busqueda y pierde el foco el control del nombre articulo
    'entonces pongo el foco en aceptar, ya que el 99 % de las veces
    'buscare por nomartic
    If Modo = 1 And Index = 1 Then
        If Trim(Text1(Index).Text) <> "" Then PonerFocoObjeto cmdAceptar
    End If


    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
    
        
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    

    'Si queremos hacer algo ..
    Select Case Index
        Case 0 'Codigo Artículo
            'Comprobar si ya existe el cod de articulo en la tabla
            If Modo = 3 Then 'Insertar
                If ExisteCP(Text1(Index)) Then PonerFoco Text1(Index)
            End If

        Case 2 'Codigo de Proveedor
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index - 2).Text = PonerNombreDeCod(Text1(Index), conAri, "sprove", "nomprove")
            Else
                Text2(Index - 2).Text = ""
            End If
            
        Case 3 'Código de Familia
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index - 2).Text = PonerNombreDeCod(Text1(Index), conAri, "sfamia", "nomfamia")
                If Text2(Index - 2).Text = "" Then
                    Text1(Index).Text = ""
                Else
                    If Modo = 3 Then
                        If vParamAplic.NumeroInstalacion = vbEuler Then
                            'EULER. El codartic lo monta desde la familia mas un secuencial
                            PonerCodigoArticuloEULER False
                
                        End If
                    End If
                End If
            
                
            Else
                Text2(Index - 2).Text = ""
            End If
            
        Case 4 'Código de Marca
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index - 2).Text = PonerNombreDeCod(Text1(Index), conAri, "smarca", "nommarca")
            Else
                Text2(Index - 2).Text = ""
            End If
            
        Case 5 'Código Tipo Unidad
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index - 2).Text = PonerNombreDeCod(Text1(Index), conAri, "sunida", "nomunida")
            Else
                Text2(Index - 2).Text = ""
            End If
            
        Case 6 'Codigo Tipo Artículo
            Text2(Index - 2).Text = PonerNombreDeCod(Text1(Index), conAri, "stipar", "nomtipar")
            If Text1(Index).Text <> "" And Text2(Index - 2).Text = "" Then PonerFoco Text1(Index)
            
        Case 7 'Tipo de IVA
            'conConta: BD Contabilidad
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index - 2).Text = PonerNombreDeCod(Text1(Index), conConta, "tiposiva", "nombriva")
            Else
                Text2(Index - 2).Text = ""
            End If
            
        Case 10, 18, 24 'Fecha alta, Fecha última compra, FECHA VIGENCIA
            If Text1(Index).Text <> "" Then PonerFormatoFecha Text1(Index)

        '---- [23/09/2009] LAURA : Añadir lineas de Cod. EAN  (se borra campo de la cabecera index 31 pasa a ser el 8)
'        Case 11, 12, 31 'numericos
        Case 11, 12, 8, 34 'numericos
            If Not PonerFormatoEntero(Text1(Index)) Then
                'If index = 33 Then Text1(index).Text = ""
            Else
                'If index = 33 Then
                '    If Val(Text1(index).Text) > 100 Then
                '        MsgBox "Campo porcentaje", vbExclamation
                '        Text1(index).Text = "100"
                '    End If
                'End If
            End If

        Case 13, 14, 15, 16, 17, 35 'Precios
            'Formato tipo 2: Decimal(10,4)
            If Text1(Index).Text <> "" Then PonerFormatoDecimal Text1(Index), 2
        
        Case 21 'Texto Control de instalación
            If (Modo <> 0) Then PonerFocoBtn Me.cmdAceptar
            
        Case 22 'categoria
            Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "scateg", "descateg")
            If Text2(Index).Text = "" And Text1(Index) <> "" Then PonerFoco Text1(Index)
            
        Case 25 'Margen comercial
            'Formato 7: Decimal(5,2)
            
            
            If PonerFormatoDecimal(Text1(Index), 7) Then
                ' ---- [06/11/2009] [LAURA] : calcular el PVP
                If Modo = 3 Then PonerPrecioPVP
            End If
        Case 26, 29, 30
             'Precio anual mantenimiento.  Lo que ponga en su tag
             ' Listros x Unidad
             PonerFormatoDecimal Text1(Index), 8
        
        Case 32
            'NumADR
            If Text1(Index).Text <> "" Then
                Text2(7).Text = PonerNombreDeCod(Text1(Index), conAri, "sadr", "descripcion", "codigoADR")
                If Text2(7).Text = "" Then
                    Text1(Index).Text = ""
                    PonerFoco Text1(Index)
                End If
            Else
                Text2(7).Text = ""
            End If
            
        Case 33
            If Text1(Index).Text <> "" Then
                Text2(9).Text = PonerNombreDeCod(Text1(Index), conAri, "sartic", "nomartic", "codartic")
                If Text2(9).Text = "" Then
                    Text1(Index).Text = ""
                    PonerFoco Text1(Index)
                End If
            Else
                Text2(9).Text = ""
            End If



    End Select
End Sub


Private Sub HacerBusqueda()
Dim cadB As String
    
    cadB = ObtenerBusqueda(Me, False, BuscaChekc)

    If chkVistaPrevia = 1 Then
        MandaBusquedaPrevia cadB
    Else
        'Se muestran en el mismo form
        If cadB <> "" Then
            CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
            PonerCadenaBusqueda
        End If
    End If
End Sub


Private Sub MandaBusquedaPrevia(cadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim cad As String
Dim tabla As String
Dim Titulo As String
Dim Conexion As Byte

    'Llamamos a al form
    '##A mano
    cad = ""
    Select Case Val(Me.imgCuentas(0).Tag)
        Case 5  'Tipo de IVA
            'Se llama a Busqueda desde el campo Tipos IVA
            '#A MANO: Porque busca en la tabla tiposiva
            'de la base de datos de Contabilidad
            cad = cad & "Código|tiposiva|codigiva|N||20·Denominacion|tiposiva|nombriva|T||70·"
            tabla = "tiposiva"
            Titulo = "Tipos de IVA"
            Conexion = conConta    'Conexión a BD: Conta
        Case Else   'Registro de la tabla de cabeceras: sartic
            cad = cad & ParaGrid(Text1(0), 23, "Código")
            cad = cad & ParaGrid(Text1(1), 58, "Denominación")
            cad = cad & ParaGrid(Text1(9), 19, "Cod. asoc.")
            tabla = "sartic"
            Titulo = "Artículos"
            Conexion = conAri    'Conexión a BD: Ariges
    End Select
           
    If cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = cad
        frmB.vTabla = tabla
        frmB.vSQL = cadB
        HaDevueltoDatos = False
        '###A mano
        frmB.vDevuelve = "0|1|"
        frmB.vTitulo = Titulo
        frmB.vselElem = 1
        frmB.vConexionGrid = Conexion
'        frmB.vBuscaPrevia = VPrevia
        frmB.vCargaFrame = False
        '#
        frmB.Show vbModal
        Set frmB = Nothing
        'Si ha puesto valores y tenemos que es formulario de busqueda entonces
        'tendremos que cerrar el form lanzando el evento
        If HaDevueltoDatos Then
            If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
                cmdRegresar_Click
        Else   'de ha devuelto datos, es decir NO ha devuelto datos
            PonerFoco Text1(kCampo)
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
        If Modo = 1 Then 'Busqueda
             MsgBox "No hay ningún registro en la tabla " & NombreTabla & " para ese criterio de Búsqueda.", vbInformation
             PonerFoco Text1(0)
        Else
            MsgBox "No hay ningún registro en la tabla " & NombreTabla, vbInformation
        End If
        Screen.MousePointer = vbDefault
        Exit Sub
    Else
        Data1.Recordset.MoveFirst
        PonerModo 2
        PonerCampos
        PonerCamposAlmacenes2
        'David 28 Nov 2008
        ' Si es conjunto mostrare sus solapa
        If Me.chkConjunto.Value = 1 Or Me.chkConjunto.Value = 2 Then PonerModoOpcionesMenu 2
    
        If DatosADevolverBusqueda <> "" Then PonerFocoBtn Me.cmdRegresar
    End If

    Screen.MousePointer = vbDefault
    Exit Sub
    
EEPonerBusq:
    MuestraError Err.Number, "PonerCadenaBusqueda", Err.Description
    PonerModo 0
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerCampos()
Dim Impor As Currency
Dim C As String

    If Data1.Recordset.EOF Then Exit Sub
    
    lblIndicador.Caption = "Datos articulo"
    lblIndicador.Refresh
    PonerCamposForma Me, Data1
    

    
    
    Text2(0).Text = PonerNombreDeCod(Text1(2), conAri, "sprove", "nomprove")
    Text2(1).Text = PonerNombreDeCod(Text1(3), conAri, "sfamia", "nomfamia")
    Text2(2).Text = PonerNombreDeCod(Text1(4), conAri, "smarca", "nommarca")
    Text2(3).Text = PonerNombreDeCod(Text1(5), conAri, "sunida", "nomunida")
    Text2(4).Text = PonerNombreDeCod(Text1(6), conAri, "stipar", "nomtipar")
    mPorIva = "porceiva"
    Text2(5).Text = DevuelveDesdeBD(conConta, "nombriva", "tiposiva", "codigiva", Text1(7).Text, "N", mPorIva)
    Text2(22).Text = PonerNombreDeCod(Text1(22), conAri, "scateg", "descateg")
    
    
    lblIndicador.Caption = "Importes"
    lblIndicador.Refresh
    PonerSumaStocks 'Poner la suma total de stocks de los almacenes donde esta el artic
    
    BloquearChecks Me, Modo

    PrimeraVez = False

    PonerCamposLineas True 'Pone los datos de las tablas de lineas de Componentes e Instalaciones
    
    'Lista campos
    CargaDatosLW
    
    'Pongo el PVP con IVA
    If mPorIva = "porceiva" Then mPorIva = 0
    Impor = CCur(mPorIva)
    Impor = Round2((Impor * Data1.Recordset!PrecioVe) / 100, 4) + Data1.Recordset!PrecioVe
    Me.txtPVPIVA.Text = Format(Impor, FormatoPrecio)
    
    
    
    
    
    'Si tiene conjuntos
    If Val(Data1.Recordset!Conjunto) = 1 Then ponerDatosConjuntos
    
    lblIndicador.Caption = "Fitosanitarios"
    lblIndicador.Refresh
    If vParamAplic.Ariagro <> "" Then
        
        Text2(7).Text = PonerNombreDeCod(Text1(32), conAri, "sadr", "descripcion", "codigoadr")
        'las materias activas ya las carga en donde corresponde: PonerCamposLineas
    End If
    If vParamAplic.NumeroInstalacion = vbFenollar Then
        
          C = " scaped.numpedcl=sliped.numpedcl and cerrado=0 and codartic = " & DBSet(Text1(0).Text, "T") & " AND 1"
          C = DevuelveDesdeBD(conAri, "sum(cantidad) as cuantos ", "scaped,sliped", C, "1")
          If C = "" Then C = "0"
          txtReser.Text = Format(CCur(C), FormatoCantidad)
    End If
    
    If Text1(33).Text = "" Then
        Text2(9).Text = ""
    Else
        Text2(9).Text = PonerNombreDeCod(Text1(33), conAri, "sartic", "nomartic", "codartic")
    End If
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
End Sub


Private Sub PonerCamposLineas(enlaza As Boolean)
'Carga las Pestañas con las tablas de lineas de Conjunto o Instalaciones
'segun la pestaña de datos a mostrar
On Error GoTo EPonerLineas

    Screen.MousePointer = vbHourglass
    'Conjuntos
    CargaGrid DataGrid1, Data2, enlaza
    'Instalaciones
    CargaGrid DataGrid2, data3, enlaza
    'Stocks
    CargaGrid DataGrid3, data4, enlaza

    '---- [23/09/2009] LAURA : Añadir lineas de Cod. EAN
    'lineas codigos EAN
    CargaGrid DataGrid4, data5, enlaza
    '----

    If vParamAplic.Ariagro <> "" Then CargaGrid DataGrid5, data6, enlaza
        
    CargaGrid DataGrid6, data6, enlaza
    
    Screen.MousePointer = vbDefault
    Exit Sub
    
EPonerLineas:
    MuestraError Err.Number, "PonerCamposLineas", Err.Description
'    PonerModo 2
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerSumaStocks()
Dim rst As ADODB.Recordset
Dim SQL As String
    
    SQL = DevuelveDesdeBDNew(conAri, "salmac", "codartic", "codartic", Text1(0).Text, "T")
    If SQL <> "" Then
        SQL = "select sum(canstock) from salmac where codartic=" & DBSet(Text1(0).Text, "T")
        Set rst = New ADODB.Recordset
        rst.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not rst.EOF Then
            Me.txtSumaStock.Text = rst.Fields(0).Value
        End If
        rst.Close
        Set rst = Nothing
    Else
        Me.txtSumaStock.Text = 0
    End If
End Sub


Private Sub PonerPrecioPVP()
Dim cart As CArticulo

    Set cart = New CArticulo
    cart.Codigo = Text1(0).Text
    cart.PrecioUltCom = ComprobarCero(Text1(15).Text)
    cart.MargenComercial = ComprobarCero(Text1(25).Text)
    cart.PrecioVenta = ComprobarCero(Text1(17).Text)
    Text1(17).Text = cart.AplicarMargenComercial 'obtiene el nuevo PVP
    FormateaCampo Text1(17)
    If Text1(7).Text <> "" Then cart.TipoIVA = Text1(7).Text
    
    
'    If cArt.LeerDatos(Me.parCodArtic) Then
'        Text1(2).Text = cArt.PrecioUltCom
'        Text1(2).Text = Format(Text1(2).Text, FormatoPrecio)
'
'        Text1(3).Text = cArt.PrecioVenta 'precio venta actual
'        Text1(3).Text = Format(Text1(3).Text, FormatoPrecio)
'        Text1(5).Text = cArt.MargenComercial
'        Text1(5).Text = Format(Text1(5).Text, FormatoPorcen)
'
'        Text1(4).Text = cArt.AplicarMargenComercial 'obtiene el nuevo PVP
'        Text1(4).Text = Format(Text1(4).Text, FormatoPrecio)
'    End If
    Set cart = Nothing


End Sub



Private Sub PonerCamposAlmacenes2()
    If data4.Recordset.EOF Then Exit Sub
    PonerCamposFormaFrame Me, "Text3", data4
    
    'Rellenar el nombre correspondiente al código de los TextBox de indice 8
    Text2(8).Text = PonerNombreDeCod(Text3(0), conAri, "salmpr", "nomalmac", "codalmac")
    
    'Rellenar el nombre correspondiente al código de ubicacion
    Text2(6).Text = PonerNombreDeCod(Text3(1), conAri, "subica", "nomubica", "codubica")
    
    'El check del inventario
    chkInventario.Value = DBLet(data4.Recordset!statusin, "N")
    
    '-- Esto permanece para saber donde estamos
'    lblIndicador.Caption = Data4.Recordset.AbsolutePosition & " de " & Data4.Recordset.RecordCount
End Sub


'Private Function ComprobarEsInstalacion() As Boolean
'Dim devuelve As String
'Dim EsInstal As Boolean
'
'    EsInstal = False
'    If Not (vParamAplic.Frecuencias) Then Exit Function ' si no estan activadas las frecuencias no se muestra ná
'    If Text1(3).Text <> "" Then
'        devuelve = DevuelveDesdeBDNew(conAri, "sfamia", "instalac", "codfamia", Text1(3).Text, "N")
'        If devuelve = "1" Then
'            EsInstal = CBool(devuelve)
'        Else
'            EsInstal = False
'        End If
'    End If
'    ComprobarEsInstalacion = EsInstal
'End Function
'
'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
Private Sub PonerModo(Kmodo As Byte)
Dim i As Byte
Dim b As Boolean
Dim NumReg As Byte

    'Actualiza Iconos Insertar,Modificar,Eliminar
    ActualizarToolbarGnral Me.Toolbar1, Modo, Kmodo, btnAnyadir
    
    Modo = Kmodo
    PonerIndicador lblIndicador, Modo
    
    'Modo 2. Hay datos y estamos visualizandolos
    '---- [23/09/2009] LAURA : Añadir lineas de Cod. EAN (se añade modo 8)
    b = (Kmodo = 2) Or (Modo = 5) Or (Modo = 6) Or (Modo = 7) Or (Modo = 8) Or (Modo = 9) Or Modo = 10
    
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.visible = b
        cmdRegresar.Caption = "&Regresar"
    Else
        cmdRegresar.visible = False
    End If
    
    'Poner Flechas de Desplazamiento Visibles o no
    NumReg = 1
    If (Modo = 2) Then
        If Not Data1.Recordset.EOF Then
            If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
        End If
    ElseIf Modo = 5 Then
        If Not data4.Recordset.EOF Then
            If data4.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
        End If
    End If
    b = (Modo = 2) Or (Modo = 5)
    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, NumReg
    
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1
    BloquearText1 Me, Modo
    'campos Precio medio y ponderado bloqueados, pq son calculados
    b = True 'bloq
    If Modo = 1 Then
        b = False
    ElseIf Modo = 4 Then
        If vUsu.Nivel <= 1 Then b = False
    End If
    BloquearTxt Text1(13), b
    BloquearTxt Text1(14), b
    
    'fecha ultimo cambio PVP bloqueado pq se actualiza automaticamente
    BloquearTxt Text1(27), True
    
    If Modo = 2 Then
        Text1(17).BackColor = &HC0C0FF
    Else
        If Modo <> 0 Then Text1(17).BackColor = &H80000005
    End If
    
    Me.FrameArtxAlmac.Enabled = (Modo = 5)
    'Me.FrameArtxAlmac2.visible = (Modo = 5)
    If Me.FrameArtxAlmac.Enabled Then
        If Modo = 5 And ModificaLineas = 2 Then BloquearTxt Text3(0), True
         'Me.FrameArtxAlmac.Height = 2010
         'Me.FrameArtxAlmac.Top = 2260
         'Me.FrameArtxAlmac.Left = 360
    End If
    b = Modo <> 5
    If vParamAplic.NumeroInstalacion = 2 Then
        If vUsu.CodigoAgente > 0 Then b = False
    End If
    Me.FrameDatosAlmacen2.visible = b
        
    b = (Modo = 1 Or Modo = 3 Or Modo = 4) '1:Busqueda, 3:Insertar, 4:Modificar
    cboArticuloVarios.Enabled = b
    cboStatus.Enabled = b
    If vParamAplic.NumeroInstalacion = 1 Then cboADV.Enabled = b
    
    If vParamAplic.NumeroInstalacion = 2 Then cboTipoComiArtVario.Enabled = b
    cboUnidadCompra.Enabled = b
    'Bloquear los checkbox
    BloquearChecks Me, Modo

    cmdCancelar.visible = b
    cmdAceptar.visible = b
    Me.imgFecha(i).Enabled = b
    For i = 0 To 5
        Me.imgCuentas(i).Enabled = b
    Next i
    
    'Numero de orden
    'Busquedas o insertar modificar el supr usuario
    b = vUsu.Nivel = 0 And (Modo = 3 Or Modo = 4)
    b = b Or Modo = 1
    '---- [23/09/2009] LAURA : Añadir lineas de Cod. EAN  (se borra campo de la cabecera index 31 pasa a ser el 8)
'    BloquearTxt Text1(31), Not b
    BloquearTxt Text1(8), Not b
    '----
    
    chkVistaPrevia.Enabled = (Modo <= 2)

    'Bton generar denominacion solo en descriptores y en modo insertar
    Me.cmdGenerar.visible = vParamAplic.Descriptores And Modo = 3
    
    cmdEuler.visible = False
    If Modo = 3 Then cmdEuler.visible = InstalacionEsEulerTaxco Or vParamAplic.NumeroInstalacion = 0



    'Poner el tamaño de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos
    PonerModoOpcionesMenu (Modo) 'Poner opciones de menu según modo
    PonerOpcionesMenu   'Activar opciones de menu según nivel
                        'de permisos del usuario
                        
                        
                        
    'Los tag's de los campos de sctock NO estaran visibles si
    'inserto,modifico o busco en la PPAL
    If Modo = 1 Or Modo = 3 Or Modo = 4 Then
        AccionesSobreTagText3_ True, False
    Else
        'Los vuelvo a poner
        AccionesSobreTagText3_ False, False
    End If
    
    'El listview
    If Modo <> 2 Then lw1.ListItems.Clear


    'cmdACtualizar importes en conjuntos
    cmdActualizarImportes1(0).visible = Modo = 6 And (ModificaLineas <> 1)
    cmdActualizarImportes1(1).visible = Modo = 6 And (ModificaLineas <> 1)
End Sub


Private Sub PonerModoOpcionesMenu(Modo As Byte)
'Activas unas Opciones de Menu y Toolbar según el modo en que estemos
Dim b As Boolean
Dim EsInstal As Boolean

    '---- [23/09/2009] LAURA : Añadir lineas de Cod. EAN (se añade modo 8)
    b = (Modo = 2) Or (Modo = 5) Or (Modo = 6) Or (Modo = 7) Or (Modo = 8) Or (Modo = 9) Or (Modo = 10)
    
    
    'Los que sean AGENTES no pueden entrar
    EsInstal = b  'reutilizo un momento la variable
    
    If vParamAplic.NumeroInstalacion = 2 Then
        If vUsu.CodigoAgente > 0 Then
            b = False
            EsInstal = False
        Else
            'Si el modo es cero u 2
            EsInstal = (EsInstal Or Modo = 0 Or Modo = 1)
        End If
    Else
            EsInstal = (EsInstal Or Modo = 0 Or Modo = 1)
    End If
    'Insertar
    Toolbar1.Buttons(6).Enabled = EsInstal
    Me.mnNuevo.Enabled = Toolbar1.Buttons(6).Enabled
    'Modificar
    Toolbar1.Buttons(7).Enabled = b
    Me.mnModificar.Enabled = b
    'eliminar
    Toolbar1.Buttons(8).Enabled = b
    Me.mnEliminar.Enabled = b
    
    'Imprimir
    Toolbar1.Buttons(14).Enabled = Not DeConsulta

    b = (Modo = 2) And Not DeConsulta
    
    'Lineas Articulos x Almacen
    'Los que sean AGENTES no pueden entrar
    If vParamAplic.NumeroInstalacion = 2 Then
        If vUsu.CodigoAgente > 0 Then b = False
    End If
    
    Toolbar1.Buttons(10).Enabled = b And vUsu.Nivel <= 1
    Me.mnMtoStocksAlm.Enabled = b And vUsu.Nivel <= 1
    
    
    Me.SSTab1.TabVisible(2) = (Me.chkConjunto.Value = 1 Or Me.chkConjunto.Value = 2)


    
   
    
    
    If vParamAplic.Ariagro <> "" Then
        mnMateriasActivas.Enabled = b And vUsu.Nivel <= 1
        Me.Toolbar1.Buttons(14).Enabled = b And vUsu.Nivel <= 1
    End If
    
    'Lineas Instalaciones
    'EsInstal = ComprobarEsInstalacion
    EsInstal = True
    b = b And EsInstal
    'Los que sean AGENTES no pueden entrar
    If vParamAplic.NumeroInstalacion = 2 Then
        If vUsu.CodigoAgente > 0 Then b = False
    End If
    
    Toolbar1.Buttons(12).Enabled = b
    Me.mnMtoInstalaciones.Enabled = b
    Me.SSTab1.TabVisible(3) = EsInstal


    '---- [23/09/2009] LAURA : Añadir lineas de Cod. EAN.   DAVID 07/03/2011
    'Lineas cod. EAN   o insertar intercalando en componentes
    ' antes Toolbar1.Buttons(13).Enabled = B
    'Generar Pedido
    b = Modo = 2 And vUsu.Nivel <= 1
    'Los que sean AGENTES no pueden entrar
    If vParamAplic.NumeroInstalacion = 2 Then
        If vUsu.CodigoAgente > 0 Then b = False
    End If
    
    If Modo = 6 Then
        Toolbar1.Buttons(13).Image = 34
        Toolbar1.Buttons(13).ToolTipText = "Insertar intercalando"
        b = (ModificaLineas = 0)
    Else
        'b=modo=2
        Toolbar1.Buttons(13).Image = 23   '23
        Toolbar1.Buttons(13).ToolTipText = "Cod. EAN"
    End If

    'Codigos EAN y materias activas     Febreor 2012   Dejamos modificar aunque sea de consulta
    Toolbar1.Buttons(13).Enabled = b
    Me.mnCodEAN.Enabled = b
    
    Toolbar1.Buttons(15).Enabled = b
    mnEquivalencias.Enabled = Toolbar1.Buttons(15).Enabled
     
    ' -----

    b = (Modo = 0) Or (Modo = 2) Or (Modo = 1)
    'Buscar
    Toolbar1.Buttons(1).Enabled = b
    Me.mnBuscar.Enabled = b
    'Modificar
    Toolbar1.Buttons(2).Enabled = b
    Me.mnVerTodos.Enabled = b
    
    
    
    
    
End Sub


Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
    PonerLongCamposGnral Me, Modo, 1
End Sub


Private Sub PonerModoFrame(Kmodo As Byte)
Dim i As Byte
Dim b As Boolean
    ModoFrame = Kmodo
    
    Select Case ModoFrame
        Case 0  'MODO INICIAL
                For i = 0 To Me.Text3.Count - 1
                    BloquearTxt Text3(i), True
                Next i
                Me.imgFecha(2).Enabled = False
                Me.imgCuentas(6).Enabled = False
                Me.imgCuentas(7).Enabled = False
                Me.chkInventario.Enabled = False
                PonerBotonCabecera True
                
        Case 3  'Modo INSERTAR
                
                BloquearTxt Text3(0), False
                Text2(8).Text = ""
    End Select
    If ModoFrame = 3 Or ModoFrame = 4 Then
        '3=Insertar,  4=Modificar
        
        'Nuevo Marzo 2010
        ' Ni stock, ni los datos de inventario se pueden insertar
        BloquearTxt Text3(0), ModoFrame = 3
        
        For i = 1 To Me.Text3.Count - 1
        
            If i = 2 Or i >= 6 Then
                b = True
            Else
                b = False
            End If
            BloquearTxt Text3(i), b
            If ModoFrame = 3 Then
                If b And i = 2 Then
                    Text3(i).Text = "0"
                Else
                    Text3(i).Text = ""
                End If
            End If
        Next i
        chkInventario.Enabled = False
        Me.imgFecha(2).Enabled = False
        Me.imgCuentas(6).Enabled = (ModoFrame = 3)
        Me.imgCuentas(7).Enabled = (ModoFrame = 3 Or ModoFrame = 4)
        PonerFoco Text3(1)
        PonerBotonCabecera False
    End If
End Sub


Private Function DatosOk() As Boolean
Dim b As Boolean
Dim i As Byte

    DatosOk = False
    
    'Comprobamos que el campo dias de garantia si no tiene valor lo
    'ponemos a 0 para q no de error que no puede ser nulo
    If Trim(Me.Text1(11).Text) = "" Then Text1(11).Text = "0"
    
    b = CompForm(Me, 1)
    If Not b Then Exit Function
    
    'Para los valores de fam,mar,tipo... es obligado que exista el codigo
    BuscaChekc = ""
    For i = 2 To 7
        If Me.Text1(i).Text = "" Xor Text2(i - 2).Text = "" Then BuscaChekc = BuscaChekc & "  -" & RecuperaValor(Text1(i).Tag, 1) & vbCrLf
    Next
    If BuscaChekc <> "" Then
        MsgBox "Error en campos: " & vbCrLf & BuscaChekc, vbExclamation
        b = False
        Exit Function
    End If
    
    'Comprobar si ya existe el cod en la tabla
    If Modo = 3 Then 'Insertar
        If ExisteCP(Text1(0)) Then
            b = False
        Else
            'No podemos crear este articulo ya que es una constante que utiliza
            If Text1(0).Text = "@1@" Then
                MsgBox "Imposible crear articulo @1@", vbExclamation
                b = False
            End If
            
            If Mid(Text1(0).Text, 1, 2) = "::" Then
                MsgBox "Imposible crear articulo ::", vbExclamation
                b = False
            End If
        
        
        End If
        
        
        
        
        If b Then
            'Comprobamos si ha puesto(insertando) el numero de orden
            'Si es asi, k tiene valor
            '---- [23/09/2009] LAURA : Añadir lineas de Cod. EAN  (se borra campo de la cabecera index 31 pasa a ser el 8)
'            If Text1(31).Text <> "" Then
            If Text1(8).Text <> "" Then
                BuscaChekc = DevuelveDesdeBD(conAri, "codartic", "sartic", "numorden", Text1(8).Text)
                If BuscaChekc <> "" Then
                    MsgBox "Ya existe el numero de orden", vbExclamation
                    b = False
                End If
            Else
                'No ha puesto ninguno. Le asigno
                'ASigno el max mas uno
                BuscaChekc = DevuelveDesdeBD(conAri, "max(numorden)", "sartic", "1", "1")
                If BuscaChekc = "" Then BuscaChekc = "0"
                BuscaChekc = Val(BuscaChekc) + 1
                Text1(8).Text = BuscaChekc
            End If
            '----  modo=3
            
            If vParamAplic.NumeroInstalacion = vbEuler Then PonerCodigoArticuloEULER True
            
        End If
        
        
        
        If b Then
            'Si esta bien, y estamos creando un articulo desde telematel, verifico que los datos
            'de proveedor y referencia proveedor que ha puesto son los que le he pasado desde frmtelematel
            If Mid(Me.DatosADevolverBusqueda, 1, 2) = "··" Then
                'codprove|nomprove|refprove|precio|nomartic|ean|codtelem|
                'Proveedor
                BuscaChekc = Mid(Me.DatosADevolverBusqueda, 3)
                BuscaChekc = Trim(RecuperaValor(BuscaChekc, 1))
                
                'Si es cabel, no indicamos proveedor. Lo puede poner a cualquiera. Con lo cual fuerzo proveedor en la variables
                If BuscaChekc = "" Then BuscaChekc = Val(Text1(2).Text)
                
                
                If Val(BuscaChekc) <> Val(Text1(2).Text) Then
                    Text1(2).Text = BuscaChekc
                    Text2(0).Text = ""
                    MsgBox "No es el proveedor de la ficha TELEMATEL", vbExclamation
                    b = False
                Else
                    'Compruebo la referencia
                    BuscaChekc = Mid(Me.DatosADevolverBusqueda, 3)
                    BuscaChekc = RecuperaValor(BuscaChekc, 3)
                    If BuscaChekc <> Text1(31).Text Then
                        b = False
                        MsgBox "No es la referencia de proveedor de la ficha de TELEMATEL", vbExclamation
                        Text1(9).Text = BuscaChekc
                    End If
                End If
                BuscaChekc = ""
            End If
        End If
        
    End If
    
    'Solo hay ECO para los articulos de VARIOS
    If vParamAplic.NumeroInstalacion = 2 Then
        If cboArticuloVarios.ListIndex = 0 Then cboTipoComiArtVario.ListIndex = 0
    End If
     
    'si se ha cambiado el precio venta PVP actualizamos la fecha de
    'ult. cambio PVP
    If Modo = 4 Then 'modo modificar
        'si se ha modificado el ult. precio compra la fecha ult. compra
        'debe tener valor
        If Text1(15).Text <> "" And Trim(Text1(18).Text) = "" Then
            b = False
            MsgBox "Si hay precio de ult. compra la fecha de ult. compra debe tener valor.", vbInformation
        End If
        
        
        'si se ha modificado el precio venta PVP actualizamos campos
        'para guardarlo correctamente
        If CCur(Me.Text1(17).Text) <> CCur(Me.Data1.Recordset!PrecioVe) Then
            Me.Text1(27).Text = Format(Now, "dd/mm/yyyy")
        End If
        
        
        
        'Cuando modificamos, si pasamos un articulo a CADUCADO, entonces comproaremos
        'si tiene sctock. Si es asi NO dejammos continuar
        If Me.cboStatus.ListIndex = 3 And Val(Data1.Recordset!codstatu) < 3 Then
            If Me.chkctrstock.Value = 1 Then
                'Lleva stcok
                'Comprobamos k valor tiene
                BuscaChekc = TotalRegistros("select sum(canstock) from salmac where codartic='" & DevNombreSQL(Text1(0).Text) & "'")
                If Val(BuscaChekc) > 0 Then
                    MsgBox "No podemos pasar un árticulo a caducado teniendo stock.", vbExclamation
                    Exit Function
                End If
            End If
        End If
        
        
        'Modificar herbelca . Articulos de rotacion
        If vParamAplic.NumeroInstalacion = 2 Then
            If Val(Data1.Recordset!Rotacion) <> Abs(chkRotacion.Value) Then
                'han cambiado la marca del articulo.
                'Si no es superusuario no puedo cambiar la rotacion
                If vUsu.Nivel > 0 Then
                    MsgBox "No puede cambiar la marca de rotacion", vbExclamation
                    Exit Function
                End If
            End If
        End If
        
        
        'Mayo 2017.
        'No me puede cambiar un tipo de IVA si hay facturas pendientes de contabilizar
        If vParamAplic.ContabilidadNueva And b Then
            If Val(Text1(7).Text) <> DBLet(Data1.Recordset!Codigiva, "N") Then
                
                'ME ha cambiado el tipo de IVA. NO dejo si falta por contabilziar
                BuscaChekc = "intconta=0 and (codtipom, numfactu,fecfactu) in "
                BuscaChekc = BuscaChekc & " (select codtipom, numfactu,fecfactu from slifac where fecfactu>='2017-07-01' AND codartic=" & DBSet(Text1(0).Text, "T") & ") AND 1"
                BuscaChekc = DevuelveDesdeBD(conAri, "count(*)", "scafac", BuscaChekc, "1")
                If Val(BuscaChekc) > 0 Then
                    MsgBox "El articulo esta en facturas cliente pendientes de contabilizar (" & BuscaChekc & ")." & vbCrLf & "Contabilícelas primero", vbExclamation
                    b = False
                Else
                    BuscaChekc = "intconta=0 and (codprove,numfactu,fecfactu) in "
                    BuscaChekc = BuscaChekc & " (select codprove,numfactu,fecfactu from slifpc where fecfactu>='2017-07-01' AND codartic=" & DBSet(Text1(0).Text, "T") & ") AND 1"
                    BuscaChekc = DevuelveDesdeBD(conAri, "count(*)", "scafpc", BuscaChekc, "1")
                    If Val(BuscaChekc) > 0 Then
                        MsgBox "El articulo esta en facturas proveedores pendientes de contabilizar (" & BuscaChekc & ")." & vbCrLf & "Contabilícelas primero", vbExclamation
                        b = False
                    End If
                End If
            End If
        End If
        
    End If 'Modificando
    
    
    
    
    DatosOk = b
End Function


Private Function DatosOkConjunto() As Boolean
Dim b As Boolean
Dim devuelve As String

    DatosOkConjunto = False
    b = True
    If txtAux(1).Text = "" Then
         MsgBox "El campo Cantidad no puede ser nulo", vbExclamation, "Artículos"
         b = False
    End If
        
    If Not IsNumeric(txtAux(1).Text) Then
        MsgBox "La cantidad de Artículos tiene que ser numérico", vbExclamation
        b = False
    End If
    
    If Me.txtAux2.Text = "" Then
        MsgBox "Error en articulo", vbExclamation
        b = False
    End If
    
    If Not b Then Exit Function
    
    'Comprobamos  si existe, solo si estamos insertando (ModificaLineas=1)
    'conAri: conexion a BD Ariges
    devuelve = DevuelveDesdeBDNew(conAri, "sarti1", "codartic", "codartic", Text1(0).Text, "T", , "codarti1", txtAux(0).Text, "T")
    If ModificaLineas = 1 And devuelve <> "" Then
        b = False
        devuelve = "Ya existe el Artículo de Conjunto: " & vbCrLf
        devuelve = devuelve & "Codigo: " & txtAux(0).Text & vbCrLf
        devuelve = devuelve & "Descripción: " & txtAux2.Text
        
        MsgBox devuelve, vbExclamation, "Artículos"
    End If
    If Not b Then Exit Function
    
    'Comprobar que el articulo no tiene conjuntos, solo si estamos insertando (ModificaLineas=1)
    'Si tiene conjuntos no puede ser elemento de conjunto de otro articulo
    devuelve = DevuelveDesdeBDNew(conAri, "sartic", "conjunto", "codartic", txtAux(0).Text, "N")
   
    
    
    If ModificaLineas = 1 And devuelve = "1" Then
        b = False
        devuelve = "No es un Artículo de Conjunto: " & vbCrLf
        devuelve = devuelve & "Codigo: " & txtAux(0).Text & vbCrLf
        devuelve = devuelve & "Descripción: " & txtAux2.Text & vbCrLf & vbCrLf
        devuelve = devuelve & "¿Continuar?"
        If MsgBox(devuelve, vbQuestion + vbYesNo) = vbYes Then b = True
    End If
    DatosOkConjunto = b
End Function


Private Function DatosOkLinea() As Boolean
Dim b As Boolean
Dim devuelve As String

    DatosOkLinea = False
    b = True
    
    If Trim(Text3(1).Text) = "" Then 'Campo Ubicación
        MsgBox "El campo Ubicación no puede ser nulo", vbExclamation, "Artículos"
        b = False
    End If
    
    'Campo de cantidad de Stock (Son decimales)
    If Trim(Text3(2).Text) = "" Or IsNull(Text3(2).Text) Then
        MsgBox "El campo Cantidad Stock no puede ser nulo", vbExclamation, "Artículos"
        b = False
    End If
    If Not b Then Exit Function
    
    'Comprobamos  si existe
    devuelve = DevuelveDesdeBDNew(conAri, "salmac", "codartic", "codartic", Text1(0).Text, "T", , "codalmac", Text3(0).Text, "N")
    If ModificaLineas = 1 And devuelve <> "" Then
        b = False
        devuelve = "Ya existe el Artículo en el Almacen: " & vbCrLf
        devuelve = devuelve & "Codigo: " & Text3(0).Text & vbCrLf
        devuelve = devuelve & "Descripción: " & Text2(8).Text
        MsgBox devuelve, vbExclamation, "Artículos"
    End If
    
    
    'Comprobaremos k el punto de pedido maximo y minimo, si estan son mayor y menor respectivamente
    '            SQL = SQL & "  = " & DBSet(Text3(3).Text, "N", "S") & ", "
    '        SQL = SQL & "  = " & DBSet(Text3(4).Text, "N", "S") & ", "
    '        SQL = SQL & " = " & DBSet(Text3(5).Text, "N", "S")
    If b Then
        devuelve = ""
        If Text3(3).Text <> "" And Text3(5).Text <> "" Then
            If ImporteFormateado(Text3(3).Text) > ImporteFormateado(Text3(5).Text) Then devuelve = "Importe stock minimo mayor que el stock maximo"

        End If
        
        If devuelve = "" Then
            If Text3(4).Text <> "" Then
                'Veremos si esta entre maximo y minimo
                If Text3(3).Text <> "" Then
                    If ImporteFormateado(Text3(3).Text) > ImporteFormateado(Text3(4).Text) Then devuelve = "Importe stock minimo mayor que el punto pedido"
                End If
                
                If Text3(5).Text <> "" Then
                    If ImporteFormateado(Text3(4).Text) > ImporteFormateado(Text3(5).Text) Then devuelve = "Importe stock maximo menor que el punto pedido"
                End If
            End If
        End If
        
        If devuelve <> "" Then
            MsgBox devuelve, vbQuestion
            b = False
        End If
    End If
    DatosOkLinea = b
End Function


Private Sub Text3_GotFocus(Index As Integer)
    kCampo = Index
    If ModificaLineas <> 0 Then
        ConseguirFoco Text3(Index), 4
    Else
        ConseguirFoco Text3(Index), 2
    End If
End Sub

Private Sub Text3_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        If Index = 8 Then
            PonerFocoBtn Me.cmdAceptar
        Else
            KeyAscii = 0
            SendKeys "{tab}"
        End If
    End If
End Sub


Private Sub Text3_LostFocus(Index As Integer)
    
     If Screen.ActiveForm.ActiveControl.Name = "cmdCancelar" Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    Select Case Index
        Case 0 'Codigo Almacen
             Text2(8).Text = PonerNombreDeCod(Text3(Index), conAri, "salmpr", "nomalmac")
             If Text2(8).Text = "" Then Text3(0).Text = ""
        Case 1 'Codigo ubicacion
            Text2(6).Text = PonerNombreDeCod(Text3(Index), conAri, "subica", "nomubica", "codubica")
            If Text2(6).Text = "" And Text3(Index) <> "" Then PonerFoco Text3(Index)
                
        Case 2, 3, 4, 5, 6 'Stocks, Punto Pedido
                'Formato tipo 1: Decimal(12,2)
                If Trim(Text3(Index)) <> "" Then PonerFormatoDecimal Text3(Index), 1
        
        Case 7  'Fecha Inventario
            If Text3(Index).Text <> "" Then PonerFormatoFecha Text3(Index)

        Case 8  'Hora Inventario
            If Trim(Text3(Index).Text) <> "" Then PonerFormatoHora Text3(Index)
    End Select
End Sub


Private Sub Text5_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text5_LostFocus(Index As Integer)
    Text5(Index).Text = Trim(Text5(Index).Text)
    
    If Index <> 0 Then Exit Sub
    
    BuscaChekc = ""
    If Me.Text5(0).Text <> "" Then
        If PonerFormatoEntero(Text5(0)) Then
            BuscaChekc = DevuelveDesdeBD(conAri, "nombrema", "smatact", "codigoma", Text5(0).Text, "N")
            If BuscaChekc = "" Then
                MsgBox "No existe materia activa", vbExclamation
                Text5(0).Text = ""
                PonerFoco Text5(0)
            End If
        Else
            Text5(0).Text = ""
        End If
    End If
    Text5(1).Text = BuscaChekc
    BuscaChekc = ""
    If Text5(1).Text <> "" Then PonerFocoBtn Me.cmdAceptar
End Sub

Private Sub Text6_GotFocus(Index As Integer)
    If Index = 0 Then ConseguirFoco Text6(Index), 4
End Sub

Private Sub Text6_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text6_LostFocus(Index As Integer)
    If Index = 0 Then
        BuscaChekc = DevuelveDesdeBDNew(conAri, "sartic", "nomartic", "codartic", Text6(0).Text, "T")
        Text6(1).Text = BuscaChekc
        If BuscaChekc = "" Then
            If Text6(0).Text <> "" Then
                MsgBox "No existe el articulo", vbExclamation
                Text6(0).Text = ""
                PonerFoco Text6(0)
            End If
        End If
        BuscaChekc = ""
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    
    If vParamAplic.NumeroInstalacion = vbFenollar Then
        If Button.Index > 5 And Button.Index < 18 Then Exit Sub
    End If


    Select Case Button.Index
        Case 1  'Buscar
            mnBuscar_Click
        Case 2  'Todos
            mnVerTodos_Click
        Case 6  'Nuevo
           mnNuevo_Click
        Case 7  'Modificar
            mnModificar_Click
        Case 8  'Borrar
            mnEliminar_Click
            
        Case 10  'Stocks Almacenes
            mnMtoStocksAlm_Click
        Case 11 'Conjuntos
            mnMtoConjuntos_Click
        Case 12 'Instalaciones
            mnMtoInstalaciones_Click
            
        '---- [23/09/2009] LAURA : Añadir lineas de Cod. EAN
        Case 13 'Codigos EAN
            If Modo = 6 Then
                If ModificaLineas = 0 Then
                    If Not Me.Data2.Recordset.EOF Then IntercalaComponente = True
                    BotonAnyadirConjunto2
                End If
            Else
                mnMtoCodigosEAN_Click
            End If
        '----
        Case 14 'Materias activas
            mnMateriasActivas_Click
        Case 15
            mnEquivalencias_Click
        Case 18 'Imprimir Listado de Articulos
            BotonImprimir
        Case 19 'Salir
            mnSalir_Click
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


Private Sub CargarComboStatus()
'### Combo Situación Artículo
'Cargaremos el combo, o bien desde una tabla o con valores fijos o como
'se quiera, la cuestion es cargarlo
' El estilo del combo debe de ser 2 - Dropdown List
' Si queremos que este ordenado, o lo ordenamos por la sentencia sql
' o marcamos la opcion sorted del combo
'0-Normal, 1-Bloqueado, 2-Caducado

    cboStatus.Clear
    cboStatus.AddItem "Normal"
    cboStatus.ItemData(cboStatus.NewIndex) = 0
    
    'Abril 2014
    cboStatus.AddItem "Obsoleto"
    cboStatus.ItemData(cboStatus.NewIndex) = 1
    
    cboStatus.AddItem "Bloqueado"
    cboStatus.ItemData(cboStatus.NewIndex) = 2
    
    cboStatus.AddItem "Caducado"
    cboStatus.ItemData(cboStatus.NewIndex) = 3
    
End Sub


Private Sub CargarComboArticuloVarios()
'### Combo Situación Artículo
'Cargaremos el combo, o bien desde una tabla o con valores fijos o como
'se quiera, la cuestion es cargarlo
' El estilo del combo debe de ser 2 - Dropdown List
' Si queremos que este ordenado, o lo ordenamos por la sentencia sql
' o marcamos la opcion sorted del combo
'0-No, 1-Si, 2-Rectificacion
 
    cboArticuloVarios.Clear
    cboArticuloVarios.AddItem "No"
    cboArticuloVarios.ItemData(cboArticuloVarios.NewIndex) = 0
    
    cboArticuloVarios.AddItem "Si"
    cboArticuloVarios.ItemData(cboArticuloVarios.NewIndex) = 1
    
    cboArticuloVarios.AddItem "Rectificación"
    cboArticuloVarios.ItemData(cboArticuloVarios.NewIndex) = 2
    
End Sub


Private Sub CargarComboADV()
'### Partes de trabajo. Articulos que se podran poner solo en partes Internos, externos o en ambos casos
'#
'#  Partes ADV. = cualquiera  1 Internos   2 Externos' after `oftweb`;
'#
    cboADV.Clear
    cboADV.AddItem "Todos"
    cboADV.ItemData(cboADV.NewIndex) = 0
    
    cboADV.AddItem "Internos"
    cboADV.ItemData(cboADV.NewIndex) = 1
    
    cboADV.AddItem "Externos"
    cboADV.ItemData(cboADV.NewIndex) = 2
    
End Sub

Private Sub CargarComboComisionArticulosVarios()

    cboTipoComiArtVario.Clear
    cboTipoComiArtVario.AddItem "Normal"
    cboTipoComiArtVario.ItemData(cboTipoComiArtVario.NewIndex) = 0
    
    cboTipoComiArtVario.AddItem "Eco"
    cboTipoComiArtVario.ItemData(cboTipoComiArtVario.NewIndex) = 1

    
End Sub


Private Sub CargarComboTipoUnidadesCompra()

    CargarCombo_Tabla cboUnidadCompra, "stipudcompra", "tipCompra", "desTipCompra"
    
End Sub


Private Function InsetarArticulosPorAlmacen(Optional cadErr As String) As Boolean
'Inserta en la tabla salmac una fila del artículo que se esta insertando
'para cada uno de los almacenes que existen en la tabla salmpr
Dim vCodartic As String, vcodalmac As Integer
Dim rsAlmPr As ADODB.Recordset
Dim cad As String
    
    On Error GoTo EInsEnAlm

    vCodartic = Text1(0).Text
    Set rsAlmPr = New ADODB.Recordset
    cad = "Select codalmac from salmpr order by codalmac;"
    rsAlmPr.Open cad, conn, adOpenForwardOnly, adLockOptimistic, adCmdText

    While Not rsAlmPr.EOF
        vcodalmac = rsAlmPr.Fields(0).Value
        cad = "INSERT INTO salmac (codartic,codalmac,ubialmac,canstock,stockmin,puntoped,stockmax,stockinv,fechainv,horainve,statusin)"
        cad = cad & " VALUES (" & DBSet(vCodartic, "T") & "," & vcodalmac & ",'',0,0,0,0,0,NULL,NULL,0)"
        conn.Execute cad
        rsAlmPr.MoveNext
    Wend
        
    rsAlmPr.Close
    Set rsAlmPr = Nothing
    InsetarArticulosPorAlmacen = True
    Exit Function
    
EInsEnAlm:
    InsetarArticulosPorAlmacen = False
    'MuestraError Err.Number, "Insertando Artículo en Almacenes.", Err.Description
    cadErr = "Insertando Artículo en Almacenes: " & vbCrLf & Err.Description
End Function
   
   
   
Private Function InsertarPreciosPorTarifa2(Optional cadErr As String) As Boolean
'Insertar en la lista de precios las tarifas para el articulo
Dim SQL As String
Dim RS As ADODB.Recordset
Dim cTar As CTarifaArt
Dim NoOK As Boolean
Dim cad As String
'Dim codlista As Double

    On Error GoTo ErrInsPrecio
    
    'comprobamos que el PVP tenga valor y sea >0 para insertar lista de precios
    InsertarPreciosPorTarifa2 = True
    If Text1(17).Text = "" Then Exit Function
    If Not (CCur(Text1(17).Text) > 0) Then Exit Function
    
    

    InsertarPreciosPorTarifa2 = False
    
    
    
    'David. Enero 2009
    If vParamAplic.CreaTarifasArticulo = 0 Then
        'NO CREO NINGUNA.
        'Salgo dando OK
        InsertarPreciosPorTarifa2 = True
        Exit Function
    End If
    
    '---- [14/09/2009] LAURA
    If vParamAplic.CreaTarifasArticulo = 2 Then 'crear todas las tarifas
    '----
    
        SQL = "SELECT * FROM starif WHERE NOT ISNULL(margecom) "
        
    '---- [14/09/2009] LAURA
    Else 'crear solo la tarifa general
        cad = DevuelveDesdeBD(conAri, "min(codlista)", "starif", "1", "1")
        If cad = "" Then cad = "0"
        SQL = "SELECT * FROM starif WHERE NOT ISNULL(margecom) and codlista = " & Val(cad)

    
    End If
    '----
        
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    'para cada tarifa insertar un linea en la tabla de lista de precios
    'por cada codartic,codtarif
    
    '23 Abril 2008
    'Tb, en funcion sobre donde se aplica el margen se hara una cosa u otra
    ' Sobre PVP o sobre PUC
    'FALTA###
    NoOK = False
    While Not RS.EOF
        Set cTar = New CTarifaArt
        cTar.CodigoArticulo = Text1(0).Text
        cTar.CodigoTarifa = RS!codlista
        'Aqui dependera de una cosa u otra para lo del PVP / UPC
        ' 1.-  "    "  va sobre el UPC
        ' 0.- La tarifa va sobre el PVP
        If DBLet(RS!opcionINC, "N") = 0 Then
            'PVP
            cTar.PrecioActual = CCur(Text1(17).Text) 'precio venta al publico (pvp)
        Else
            If Text1(15).Text = "" Then
                cTar.PrecioActual = 0
            Else
                cTar.PrecioActual = CCur(Text1(15).Text) 'precio venta al publico (pUC)
            End If
        End If
        If cTar.InsertarPrecios = False Then NoOK = True
        Set cTar = Nothing
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
        
    If NoOK Then
        InsertarPreciosPorTarifa2 = False
        cadErr = "Los precios del artículo por tarifa NO se han introducido correctamente."
    Else
        InsertarPreciosPorTarifa2 = True
    End If
        
    Exit Function
    
ErrInsPrecio:
    InsertarPreciosPorTarifa2 = False
    cadErr = "Insertar precios por tarifa: " & Err.Description
End Function
   
   
Private Function BloquearTarifas(codArtic As String) As Boolean
Dim cadWhere As String
    cadWhere = "codartic=" & DBSet(codArtic, "T")
    BloquearTarifas = BloqueaRegistro("slista", cadWhere)
End Function
   
   
Private Function ActualizarPreciosVenta() As Boolean
'si se modifica el precio ult. compra a mano preguntar si quiere modificar
'el PVP y las tarifas de venta desde el formulario de actualizar precios
Dim precioUC As Currency 'precio ult. compra (valor actual)
Dim FechaUC As String
Dim newPrecioUC As Currency
Dim bActualizar As Boolean
Dim cad As String
Dim EnPromocionOPrecioEspecial As String


    'Comprobar si se ha modificado el precio desde la ultima compra
    'y preguntar quiere modificar el PVP del articulo aplicandole su margen
    'y el precio de las TArifas aplicandole el margen
    '-- Laura 19/12/2006: el precio de compra es el precio con los descuentos (importe/cantidad)
    precioUC = CCur(DBLet(Me.Data1.Recordset!precioUC, "N"))
    If Not IsNull(Me.Data1.Recordset!ultfecco) Then FechaUC = DBLet(Me.Data1.Recordset!ultfecco, "F")
    newPrecioUC = ImporteFormateado(Text1(15).Text)
    
    bActualizar = False
    cad = ""
    If precioUC <> newPrecioUC Then
        If FechaUC = "" Then
            bActualizar = True
        ElseIf CDate(Text1(18).Text) >= CDate(FechaUC) Then
            bActualizar = True
        Else
            
        End If
        cad = "precio de última compra"
    End If
    
    
    '## LAURA 25/06/2008
    If Not bActualizar Then
        '-- comprobar si se ha modificado el margen comercial y
        '-- en este caso recalcular tambien el PVP y tarifas
        precioUC = CCur(DBLet(Me.Data1.Recordset!margecom, "N")) 'margen actual
        newPrecioUC = ImporteFormateado(Text1(25).Text) 'margen nuevo
        If precioUC <> newPrecioUC Then bActualizar = True
        cad = "margen comercial"
    End If
    '##
    
    
    
    'Marzo 2011
    ' Avisa si el articulo este en ofertas y/o promociones
    'Por si acaso tiene precios especiales

    
    If bActualizar Then
    
    
            precioUC = 0
            EnPromocionOPrecioEspecial = DevuelveDesdeBD(conAri, "codartic", "spromo", "codartic", CStr(Data1.Recordset!codArtic), "T")
            If EnPromocionOPrecioEspecial <> "" Then precioUC = 1
            EnPromocionOPrecioEspecial = DevuelveDesdeBD(conAri, "codartic", "spromo", "codartic", CStr(Data1.Recordset!codArtic), "T")
            If EnPromocionOPrecioEspecial <> "" Then precioUC = precioUC + 2
            If precioUC = 0 Then
                EnPromocionOPrecioEspecial = ""
            Else
                EnPromocionOPrecioEspecial = "ATENCION. Artículo en:"
                If precioUC = 1 Or precioUC = 3 Then EnPromocionOPrecioEspecial = EnPromocionOPrecioEspecial & vbCrLf & " - PROMOCIONES"
                If precioUC = 3 Then EnPromocionOPrecioEspecial = EnPromocionOPrecioEspecial & vbCrLf & " - PRECIOS ESPECIALES"
                EnPromocionOPrecioEspecial = vbCrLf & String(20, "*") & vbCrLf & vbCrLf & EnPromocionOPrecioEspecial & vbCrLf & String(20, "*")
                EnPromocionOPrecioEspecial = vbCrLf & vbCrLf & vbCrLf & EnPromocionOPrecioEspecial
                
            End If
    
    
    
            EnPromocionOPrecioEspecial = "Se ha modificado el " & cad & "." & vbCrLf & "¿Desea actualizar los precios de venta?" & EnPromocionOPrecioEspecial
     
            If MsgBox(EnPromocionOPrecioEspecial, vbQuestion + vbYesNo) = vbYes Then
                'Comprobar que el artículo tiene margen comercial
                If ArticuloTieneMargen(Text1(0).Text) Then
                    'Llamar al form de actualizar precios venta
                    frmComActPrecios.parCodArtic = Text1(0).Text
                    frmComActPrecios.parNomArtic = Text1(1).Text
                    frmComActPrecios.Show vbModal
                End If
            End If
        End If
    
    



End Function
  
  
  
  
Private Function ActualizarPreciosPorTarifa() As Boolean
Dim QueTipo As Byte
Dim Importe As Currency
Dim Aux As Currency
Dim EnPromocionOPrecioEspecial As String

       'Reutilizo BuscaChekc
       QueTipo = 100
       BuscaChekc = ""
       
       '- ver si se ha modificado el precion venta PVP
       Importe = DBLet(Data1.Recordset!PrecioVe, "N")
       If Importe <> CCur(Text1(17).Text) Then
            BuscaChekc = "-el precio de venta." & vbCrLf
            QueTipo = 0 'que mire tarifas PVP
       End If
        
       '- ver si se ha modificado el precio ultima compra
       Importe = DBLet(Data1.Recordset!precioUC, "N")
       Aux = 0
       If Text1(15).Text <> "" Then Aux = CCur(Text1(15).Text)
       If Importe <> Aux Then
            BuscaChekc = BuscaChekc & "-el precio de ultima compra." & vbCrLf
            If Aux = 0 Then BuscaChekc = BuscaChekc & "*****  Precio ultima compra=  CERO    ****** " & vbCrLf
            If QueTipo = 0 Then
                QueTipo = 2  'Que mire las dos
            Else
                QueTipo = 1  'que mire solo en tarifas U.P.C.
            End If
        End If
            
            
        If QueTipo <> 100 Then
        
            'Por si acaso tiene precios especiales
            Aux = 0
            EnPromocionOPrecioEspecial = DevuelveDesdeBD(conAri, "codartic", "spromo", "codartic", CStr(Data1.Recordset!codArtic), "T")
            If EnPromocionOPrecioEspecial <> "" Then Aux = 1
            
            EnPromocionOPrecioEspecial = DevuelveDesdeBD(conAri, "codartic", "sprees", "codartic", CStr(Data1.Recordset!codArtic), "T")
            If EnPromocionOPrecioEspecial <> "" Then Aux = Aux + 2
            
            If Aux = 0 Then
                EnPromocionOPrecioEspecial = ""
            Else
                EnPromocionOPrecioEspecial = ""
                If Aux = 1 Or Aux = 3 Then EnPromocionOPrecioEspecial = EnPromocionOPrecioEspecial & vbCrLf & " - PROMOCIONES"
                If Aux = 2 Then
                    If Not vParamAplic.ActualizaPrecioEspecial Then EnPromocionOPrecioEspecial = EnPromocionOPrecioEspecial & vbCrLf & " - PRECIOS ESPECIALES"
                End If
                If EnPromocionOPrecioEspecial <> "" Then
                    EnPromocionOPrecioEspecial = "ATENCION. Artículo en:" & EnPromocionOPrecioEspecial
                    EnPromocionOPrecioEspecial = vbCrLf & vbCrLf & vbCrLf & String(20, "*") & vbCrLf & EnPromocionOPrecioEspecial & vbCrLf & String(20, "*")
                End If
            End If
                
            BuscaChekc = vbCrLf & BuscaChekc & vbCrLf
            BuscaChekc = "Se han modificado: " & BuscaChekc & "¿Desea actualizar las tarifas de precios?"
            If vParamAplic.ActualizaPrecioEspecial And Aux >= 2 Then BuscaChekc = BuscaChekc & vbCrLf & " Se actualizarán también los precios especiales "
            BuscaChekc = BuscaChekc & EnPromocionOPrecioEspecial
            If MsgBox(BuscaChekc, vbQuestion + vbYesNo) = vbYes Then
                
                Screen.MousePointer = vbHourglass
                ActualizarPreciosPorTarifaDOS QueTipo
                
                'Si actualiza solo el pvp o o si es pvp y upc
                
                If vParamAplic.ActualizaPrecioEspecial Then
                    lblIndicador.Caption = "Precio esp."
                    lblIndicador.Refresh
                    If QueTipo <> 1 Then ActualizarPrecioEspecialGenerico Text1(0).Text, CCur(Text1(17).Text), True, ""
                End If
                lblIndicador.Caption = "MODIFICAR"
                Screen.MousePointer = vbDefault
            End If
        End If
    
    
End Function
  
  
                    'QueTipoActualiza : 0. PVP
                    '                   1. UPC
                    '                   2. LOS DOS
Private Function ActualizarPreciosPorTarifaDOS(PVP As Byte, Optional cadErr As String) As Boolean
'Actualiza en la lista de precios las tarifas para el articulo
Dim SQL As String
Dim RS As ADODB.Recordset
Dim cTar As CTarifaArt
Dim NoOK As Boolean
Dim menErr As String
Dim newPrecio As Currency


    On Error GoTo ErrActPrecio
    
    '-- comprobamos que el PVP tenga valor y sea >0 para insertar lista de precios
    ActualizarPreciosPorTarifaDOS = True
    
    
    
    '-- comprobar que para ese articulo en la tabla de tarifas no haya ningun registros
    '   con valor en el campo precio_nuevo
    SQL = "SELECT COUNT(*) FROM slista WHERE codartic=" & DBSet(Text1(0).Text, "T")
    SQL = SQL & " AND not isnull(precionu) and precionu>0"
    If RegistrosAListar(SQL) > 0 Then
        MsgBox "No se pueden actualizar las tarifas del artículo." & vbCrLf & "Tiene precios nuevos.", vbExclamation
        Exit Function
    End If
    
    
    ActualizarPreciosPorTarifaDOS = False
    
    If Not BloquearTarifas(Text1(0).Text) Then
        MsgBox "NO se han actualizado las tarifas de precios.", vbExclamation, "Actualizar precios"
        Exit Function
    End If
    
    
    '-- seleccionar todas las posibles tarifas
    SQL = "SELECT * FROM starif WHERE NOT ISNULL(margecom) "
    If PVP < 2 Then
        'Sera de uno de los tipos
        SQL = SQL & " AND opcionINC = " & CStr(PVP)
    End If
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    'para cada tarifa actualizar la linea en la tabla de lista de precios
    'por cada codartic,codtarif
    NoOK = False
    While Not RS.EOF
        If BloquearTarifas(Text1(0).Text) Then
            Set cTar = New CTarifaArt
            If cTar.LeerDatos(Text1(0).Text, RS!codlista) Then
                
                If cTar.TarifaSobre = 0 Then
                    'TARIFAS SOBRE PVP
                    newPrecio = Round2((CCur(Text1(17).Text) * cTar.MargenComercial) / 100, 4)
                    newPrecio = CCur(Text1(17).Text) + newPrecio
                    
                Else
                    'TARIFAS SOBRE UPC
                    newPrecio = Round2((CCur(Text1(15).Text) * cTar.MargenComercial) / 100, 4)
                    newPrecio = CCur(Text1(15).Text) + newPrecio
                End If
                
                If cTar.ActualizarPrecios(Format(Now, "dd/mm/yyyy"), newPrecio, 0, menErr, False) = False Then NoOK = True
            Else
                'si no existe el articulo con esa tarifa la damos de alta
                cTar.CodigoArticulo = Text1(0).Text
                cTar.CodigoTarifa = RS!codlista
                'Si la tarifa es sobre PVP, mando el PVP
                'Si es sobre el UPC mando el UPC
                If DBLet(RS!opcionINC, "N") = 0 Then
                    'PVP
                    cTar.PrecioActual = CCur(Text1(17).Text) 'precio venta al publico (pvp)
                Else
                    cTar.PrecioActual = CCur(Text1(15).Text) 'precio venta al publico (pUC)
                End If
                
                If Not cTar.InsertarPrecios Then NoOK = True
            End If
            Set cTar = Nothing
        Else
            NoOK = True
'            MsgBox "NO se han actualizado correctamente todas las tarifa del artículo.", vbExclamation, "Actualizar precios"
            'Exit Function
        End If
        
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
        
    If NoOK Then
        ActualizarPreciosPorTarifaDOS = False
        cadErr = "NO se han actualizado correctamente todas las tarifa del artículo."
        cadErr = cadErr & vbCrLf & menErr
        MsgBox cadErr, vbExclamation, "Actualizar Precios"
    Else
        ActualizarPreciosPorTarifaDOS = True
    End If
        
    Exit Function
    
ErrActPrecio:
    ActualizarPreciosPorTarifaDOS = False
    cadErr = "Actualizar precios por tarifa: " & Err.Description
    MsgBox cadErr, vbExclamation
End Function
   
    
Private Function InsertarModificarLinea() As Boolean
Dim i As Integer
Dim SQL As String

    On Error GoTo EInsertarModificarLinea

    InsertarModificarLinea = False
    SQL = ""
    Select Case ModificaLineas
    Case 1  'INSERTAR
        If DatosOkLinea Then 'INSERTAR
            SQL = "INSERT INTO salmac (codartic,codalmac,ubialmac,canstock,stockmin,puntoped,stockmax,stockinv,fechainv,horainve,statusin) VALUES ("
            
            SQL = SQL & DBSet(Text1(0).Text, "T") & ", "
            SQL = SQL & Text3(0).Text & ", "
            SQL = SQL & DBSet(Text3(1).Text, "T") & ", "
            
            'Campos Stocks (Son Decimales)
            SQL = SQL & DBSet(Text3(2).Text, "N", "N") & ", "
            For i = 3 To 6
                SQL = SQL & DBSet(Text3(i).Text, "N", "S") & ", "
            Next i
        
            'Campo Fecha
            SQL = SQL & DBSet(Text3(7).Text, "F", "S") & ", "
'            If Trim(Text3(7).Text) <> "" Then
'              SQL = SQL & DBSet(Text3(7).Text, "F") & ", "
'            Else
'              SQL = SQL & "NULL, "
'            End If
        
            If Trim(Text3(8).Text) <> "" Then     'Campo Hora
              SQL = SQL & Format(Text3(8).Text, "hh:mm:ss") & ", "
            Else
              SQL = SQL & "NULL, "
            End If
        
            SQL = SQL & chkInventario.Value & ")"
        End If
        
    Case 2  'MODIFICAR
        If DatosOkLinea Then
            SQL = "UPDATE salmac Set ubialmac = " & DBSet(Text3(1).Text, "T") & ", "
            SQL = SQL & " canstock = " & DBSet(Text3(2).Text, "N") & ", "
            SQL = SQL & " stockmin = " & DBSet(Text3(3).Text, "N", "S") & ", "
            SQL = SQL & " puntoped = " & DBSet(Text3(4).Text, "N", "S") & ", "
            SQL = SQL & " stockmax = " & DBSet(Text3(5).Text, "N", "S") & ", "
            SQL = SQL & " stockinv = " & DBSet(Text3(6).Text, "N", "S")
    '        If Trim(Text3(7).Text) <> "" Then
    '            SQL = SQL & ", fechainv = " & DBSet(Text3(7).Text, "F", "S")
    '            If Trim(Text3(8).Text) <> "" Then
    '                SQL = SQL & ", horainve = '" & Format(Text3(8).Text, "hh:mm:ss") & "'"
    '            Else
    '                SQL = SQL & ", horainve = " & ValorNulo
    '            End If
    '        End If
    '        SQL = SQL & ", statusin = " & (chkInventario.Value)
            SQL = SQL & " WHERE codartic = " & DBSet(Text1(0).Text, "T") & " AND "
            SQL = SQL & " codalmac =" & Val(Text3(0).Text)
            
        End If
    End Select
        
    If SQL <> "" Then
        conn.Execute SQL
        InsertarModificarLinea = True
    Else
        PonerFoco Text3(1)
    End If
    Exit Function

EInsertarModificarLinea:
    MuestraError Err.Number, "Insertar/Modificar Stocks Almacenes", Err.Description
End Function
    
    
Private Function InsertarArticulo() As Boolean
Dim b As Boolean
Dim menErr As String

    On Error GoTo ErrInsArt
    conn.BeginTrans
    
    b = InsertarDesdeForm(Me)
    If Not b Then menErr = "Insertando en tabla articulos"
    'insertar una linea en salmac para cada uno de los almacenes
    If b Then b = InsetarArticulosPorAlmacen(menErr)
    
    'insertar una linea de lista de precios para cada tarifa
    If b Then b = InsertarPreciosPorTarifa2(menErr)
                
    If b Then
        conn.CommitTrans
    Else
        conn.RollbackTrans
        MsgBox menErr, vbExclamation
    End If
    InsertarArticulo = b
    Exit Function
                
ErrInsArt:
    conn.RollbackTrans
    InsertarArticulo = False
    MuestraError Err.Number, "Insertar artículo.", Err.Description
End Function
    
    
    
    

Public Function InsertarModificarConjunto() As Boolean
Dim SQL As String
On Error GoTo EInsertarModificarLinea

    SQL = ""
    InsertarModificarConjunto = False
    
    If DatosOkConjunto Then
        Select Case ModificaLineas
        Case 1 'Insertar
                If IntercalaComponente Then
                    SQL = "UPDATE sarti1 SET numlinea=numlinea +1 "
                    SQL = SQL & " WHERE codartic =" & DBSet(Text1(0).Text, "T") & " AND "
                    SQL = SQL & " numlinea >=" & cmdAceptar.Tag & " ORDER BY numlinea desc"
                    conn.Execute SQL
                    Espera 0.5
                End If
        
        
        
                SQL = "INSERT INTO sarti1 VALUES ("
                SQL = SQL & DBSet(Text1(0).Text, "T") & ", "
                SQL = SQL & cmdAceptar.Tag & ", "
                SQL = SQL & DBSet(txtAux(0).Text, "T") & ", "
                SQL = SQL & DBSet(txtAux(1).Text, "N") & ") "
        Case 2 'Modificar
      
                SQL = "UPDATE sarti1 Set codarti1 = " & DBSet(txtAux(0).Text, "T")
                SQL = SQL & ", cantidad = " & DBSet(txtAux(1).Text, "N")
                SQL = SQL & " WHERE codartic =" & DBSet(Text1(0).Text, "T") & " AND "
                SQL = SQL & " numlinea =" & cmdAceptar.Tag
        End Select
    End If
    
    If SQL <> "" Then
        conn.Execute SQL
        InsertarModificarConjunto = True
        txtAux2.BackColor = &H80000005
        IntercalaComponente = False
    End If
    Exit Function
    
EInsertarModificarLinea:
    MuestraError Err.Number, "Insertar/Modificar Conjuntos", Err.Description
End Function


Public Function InsertarModificarInstalacion() As Boolean
Dim SQL As String
Dim Valor As String

On Error GoTo EInsertarModificarInstalacion
    InsertarModificarInstalacion = False
    
    
    If vParamAplic.NumeroInstalacion = vbFontenas Then
        'CALIDAD
        SQL = ""
        'INSERTAR
        If ModificaLineas = 1 And cboCalidad.ListIndex < 0 Then SQL = "- Seleccione ensayo"
        If Trim(txtAux(9).Text) = "" Then SQL = SQL & vbCrLf & "-indique especificación"
        If Trim(txtAux(10).Text) = "" Xor Trim(txtAux(11).Text) = "" Then SQL = SQL & vbCrLf & "-maximo/minimo. Los dos o ninguno"
        If SQL <> "" Then
            MsgBox "Error: " & vbCrLf & SQL, vbExclamation
            Exit Function
        End If
    
        If ModificaLineas = 1 Then 'INSERTAR
            'Que no exista
            SQL = "codigoensayo =" & cboCalidad.ItemData(cboCalidad.ListIndex) & " AND codartic"
            SQL = DevuelveDesdeBD(conAri, "codigoensayo", "sarti7", SQL, Data1.Recordset!codArtic, "T")
            If SQL <> "" Then
                MsgBox "Ya existe el ensayo " & cboCalidad.List(cboCalidad.ListIndex), vbExclamation
                Exit Function
            End If
        End If
        SQL = "REPLACE INTO sarti7(codartic,codigoensayo,especificaciones ,mini,maxi) VALUES ("
        SQL = SQL & DBSet(Text1(0).Text, "T") & ","
        If ModificaLineas = 1 Then
            'INSERTAR
            SQL = SQL & cboCalidad.ItemData(cboCalidad.ListIndex)
        Else
            'Modificar
            SQL = SQL & data3.Recordset!numlinea
        End If
        SQL = SQL & "," & DBSet(txtAux(9).Text, "T") & "," & DBSet(txtAux(10).Text, "N", "S") & "," & DBSet(txtAux(11).Text, "N", "S") & ")"
    
    Else
        'INSTALACION
        Valor = Trim(txtAux(2).Text)
        If Valor = "" Then Valor = " "
        
        If ModificaLineas = 1 Then 'INSERTAR
            SQL = "INSERT INTO sarti2 VALUES ("
            SQL = SQL & DBSet(Text1(0).Text, "T") & ", "
            SQL = SQL & cmdAceptar.Tag & ", "
            SQL = SQL & DBSet(Valor, "T") & ") "
        ElseIf ModificaLineas = 2 Then 'MODIFICAR
            SQL = "UPDATE sarti2 Set licontro = " & DBSet(Valor, "T")
            SQL = SQL & " WHERE codartic =" & DBSet(Text1(0).Text, "T") & " AND "
            SQL = SQL & " numlinea =" & cmdAceptar.Tag
        End If
    
    
    End If
    
    
    
    conn.Execute SQL
    InsertarModificarInstalacion = True
    Exit Function

EInsertarModificarInstalacion:
    MuestraError Err.Number, "Insertar/Modificar Instalación", Err.Description
End Function


'---- [23/09/2009] LAURA : Añadir lineas de Cod. EAN
Public Function InsertarModificarCodigosEAN() As Boolean
Dim SQL As String
Dim Valor As String

    On Error GoTo ErrInsModEAN
    InsertarModificarCodigosEAN = False
    
    Valor = Trim(txtAux(8).Text)
    If Valor = "" Then Valor = " "
    
    If ModificaLineas = 1 Then 'INSERTAR
        SQL = "INSERT INTO sarti3 VALUES ("
        SQL = SQL & DBSet(Text1(0).Text, "T") & ", "
        SQL = SQL & cmdAceptar.Tag & ", "
        SQL = SQL & DBSet(Valor, "T") & ") "
    ElseIf ModificaLineas = 2 Then 'MODIFICAR
        SQL = "UPDATE sarti3 Set codigoea = " & DBSet(Valor, "T")
        SQL = SQL & " WHERE codartic =" & DBSet(Text1(0).Text, "T") & " AND "
        SQL = SQL & " numlinea =" & cmdAceptar.Tag
    End If
    
    conn.Execute SQL
    InsertarModificarCodigosEAN = True
    Exit Function

ErrInsModEAN:
    MuestraError Err.Number, "Insertar/Modificar codigos EAN", Err.Description
End Function
'----

'---- [19/12/2011] David Materias activas fitosnaitariios
Public Function InsertarModificarMATACT() As Boolean
Dim SQL As String
Dim Valor As String

    On Error GoTo ErrInsModEAN
    InsertarModificarMATACT = False
    
    If Text5(0).Text = "" Then
        MsgBox "Error materia activa", vbExclamation
        Exit Function
    End If

    
    If ModificaLineas = 1 Then 'INSERTAR
        cmdAceptar.Tag = Text5(0).Text
        SQL = "INSERT INTO sarti5(codartic,codigoma) VALUES ("
        SQL = SQL & DBSet(Text1(0).Text, "T") & "," & DBSet(Text5(0).Text, "N") & ") "
    ElseIf ModificaLineas = 2 Then 'MODIFICAR
        'SQL = "UPDATE sarti3 Set codigoea = " & DBSet(Valor, "T")
        'SQL = SQL & " WHERE codartic =" & DBSet(Text1(0).Text, "T") & " AND "
        'SQL = SQL & " numlinea =" & cmdAceptar.Tag
    End If
    
    conn.Execute SQL
    InsertarModificarMATACT = True
    Exit Function

ErrInsModEAN:
    MuestraError Err.Number, "Insertar/Modificar materia activa", Err.Description
    PonerFoco Text5(0)
End Function



Private Sub CargaGrid(ByRef vDataGrid As DataGrid, ByRef vData As Adodc, enlaza As Boolean)
Dim tots As String
Dim SQL As String

    On Error GoTo ECargaGrid
    
      
    If vDataGrid.Name = "DataGrid1" Then
        SQL = MontaSQLCarga(enlaza, 2)
        CargaGridGnral DataGrid1, Me.Data2, SQL, PrimeraVez
        If vParamAplic.ComponentePorcentaje Then
            tots = "%"
        Else
            tots = "cantidad"
        End If
        tots = "N||||0|;N||||0|;S|txtAux(0)|T|Cod. Artículo|1350|;S|cmdAux|B||0|;S|txtAux2|T|Desc. Artículo|3550|;S|txtAux(1)|T|" & tots & "|890|" & FormatoCantidad & "|;"
        tots = tots & "S|txtAux(3)|T|PVP|950|;S|txtAux(4)|T|UPC|950|;S|txtAux(5)|T|Pre.Tarif|950|;"
        'Materia prima
        tots = tots & "S|txtAux(6)|T|M.Pr.|550|;"
        'Dic 2013    Canstock   . Si hay que añadir otro campo desplazar el txtaux(8) y abrir hueco
        tots = tots & "S|txtAux(7)|T|St Ppal|850|;"
        arregla tots, DataGrid1, Me
        DataGrid1.Columns(4).Alignment = dbgCenter
        DataGrid1.ScrollBars = dbgAutomatic
        
    ElseIf vDataGrid.Name = "DataGrid2" Then
        
        
        SQL = MontaSQLCarga(enlaza, 3)
        CargaGridGnral DataGrid2, Me.data3, SQL, PrimeraVez
        
        If vParamAplic.NumeroInstalacion = vbFontenas Then
            'FONTENAS   codigoensayo,ensayo,sarti7.especificaciones,mini,maxi
            tots = "N||||0|;S|cboCalidad|C|Ensayo|1800|;S|txtAux(9)|T|Especificación|3900|;"
            tots = tots & "S|txtAux(10)|T|Mínimo|1200|;S|txtAux(11)|T|Maximo|1200|;"
        Else
            'Teinsa y el resto
            tots = "N||||0|;N||||0|;S|txtAux(2)|T|Control Instalaciones|7100|;"
        End If
        arregla tots, DataGrid2, Me
        DataGrid2.ScrollBars = dbgAutomatic
        
    ElseIf vDataGrid.Name = "DataGrid3" Then
        SQL = MontaSQLCarga(enlaza, 4)
        CargaGridGnral DataGrid3, Me.data4, SQL, PrimeraVez
        tots = "S|Text3(0)|T|Cod.Alm|1200|;S|cmdAlma|B||0|;S|Text2(8)|T|Nombre Almacen|2400|;S|Text3(2)|T|Stock|1200|;"
        'Los campos que no se ven que van FUERA DEL GRID
        tots = tots & "N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;"
        arregla tots, DataGrid3, Me
        DataGrid3.ScrollBars = dbgAutomatic
 
    '---- [23/09/2009] LAURA : Añadir lineas de Cod. EAN
    ElseIf vDataGrid.Name = "DataGrid4" Then 'Lineas cod. EAN
        SQL = MontaSQLCarga(enlaza, 5)
        CargaGridGnral DataGrid4, Me.data5, SQL, PrimeraVez
        tots = "N||||0|;N||||0|;S|txtAux(8)|T|Cod. EAN|2100|;"
        arregla tots, DataGrid4, Me
        DataGrid4.ScrollBars = dbgAutomatic
        
        '19 Diciembre 2011.
    ElseIf vDataGrid.Name = "DataGrid5" Then 'Lineas cod. EAN
        SQL = MontaSQLCarga(enlaza, 6)
        CargaGridGnral DataGrid5, Me.data6, SQL, PrimeraVez
        tots = "S|Text5(0)|T|Codigo|1300|;S|cmdMatAux|B||0|;S|Text5(1)|T|Descripcion|3200|;"
        arregla tots, DataGrid5, Me
        DataGrid5.ScrollBars = dbgAutomatic
        
    '----
    ElseIf vDataGrid.Name = "DataGrid6" Then 'Lineas equivalencias
        SQL = MontaSQLCarga(enlaza, 7)
        CargaGridGnral DataGrid6, Me.data7, SQL, PrimeraVez
        tots = "S|Text6(0)|T|Articulo|1500|;S|cmdEquiv|B||0|;S|Text6(1)|T|Descripcion|4000|;"
        arregla tots, DataGrid6, Me
        DataGrid6.ScrollBars = dbgAutomatic

    
    
    End If
    
     
    
    
ECargaGrid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub


Private Function MontaSQLCarga(enlaza As Boolean, Opcion As Byte) As String
'--------------------------------------------------------------------
' MontaSQlCarga:
'   Basándose en la información proporcionada por el vector de campos
'   crea un SQl para ejecutar una consulta sobre la base de datos que los
'   devuelva.
' Si ENLAZA -> Enlaza con el Data
'           -> Si no lo cargamos sin enlazar a ningun campo
'--------------------------------------------------------------------
Dim SQL As String

    If Opcion = 2 Then
        'cadena SQL para cargar los CONJUNTOS de la tabla sarti1
        'SQL = "SELECT sarti1.codartic,sarti1.numlinea,sarti1.codarti1,sartic.nomartic,sarti1.cantidad "
        'SQL = SQL & " FROM sarti1 INNER JOIN sartic ON sarti1.codarti1=sartic.codartic "
        
        
        SQL = "SELECT sarti1.codartic, numlinea, sarti1.codarti1,sartic.nomartic,"
        SQL = SQL & " sarti1.Cantidad , sartic.preciove, sartic.precioUC, slista.precioac,if (mateprima=1,""*"","" "") materiaprima, canstock"
        SQL = SQL & " FROM   sarti1 INNER JOIN sartic ON"
        SQL = SQL & " sarti1.codarti1 = sartic.codArtic"
        
        'Dic 2013
        'Stock en el almacen ppal
        SQL = SQL & " LEFT OUTER join salmac on sarti1.codarti1=salmac.codartic and codalmac=1"
        
        SQL = SQL & " LEFT OUTER JOIN slista ON sarti1.codarti1=slista.codartic AND slista.codlista = " & vParamAplic.CodTarifa
        SQL = SQL & " where sarti1.codartic="
        If enlaza Then
            SQL = SQL & DBSet(Text1(0).Text, "T")
        Else
            SQL = SQL & "'-1@#'"
        End If
        SQL = SQL & " ORDER BY sarti1.numlinea "
        
        
    ElseIf Opcion = 3 Then 'INSTALACIONES
    
        If vParamAplic.NumeroInstalacion = vbFontenas Then
            SQL = "Select codigoensayo numlinea,ensayo,sarti7.especificaciones,mini,maxi from "
            SQL = SQL & " sarti7,scalidad where scalidad.codigo=sarti7.codigoensayo"
            SQL = SQL & " AND sarti7.codartic= "
            If enlaza Then
                SQL = SQL & DBSet(Text1(0), "T")
            Else
                SQL = SQL & "'-1'"
            End If
            SQL = SQL & " ORDER BY ensayo"
    
    
    
        Else
            'Lo que habia. Teinsa y demas
            SQL = "SELECT sarti2.codartic, sarti2.numlinea, sarti2.licontro "
            SQL = SQL & " FROM sarti2"
            If enlaza Then
                SQL = SQL & " WHERE sarti2.codartic=" & DBSet(Text1(0), "T")
            Else
                SQL = SQL & " WHERE sarti2.codartic= '-1'"
            End If
            SQL = SQL & " ORDER BY sarti2.numlinea"
    
        End If
    ElseIf Opcion = 4 Then 'STOCK
        
        SQL = "select salmac.codalmac,nomalmac,canstock,ubialmac,stockmin,puntoped,stockmax,stockinv,fechainv,horainve,statusin  "
        SQL = SQL & " from salmac,salmpr where salmac.codalmac=salmpr.codalmac AND "
        If enlaza Then
            SQL = SQL & " codartic=" & DBSet(Text1(0), "T")
        Else
            SQL = SQL & " codartic= '-1'"
        End If
    
    '---- [23/09/2009] LAURA : Añadir lineas de Cod. EAN
    ElseIf Opcion = 5 Then
        SQL = "SELECT *"
        SQL = SQL & " FROM sarti3"
        If enlaza Then
            SQL = SQL & " WHERE sarti3.codartic=" & DBSet(Text1(0), "T")
        Else
            SQL = SQL & " WHERE sarti3.codartic= '-1'"
        End If
        SQL = SQL & " ORDER BY sarti3.numlinea"
    '---- [19/12/2011] Materias activas
    ElseIf Opcion = 6 Then
        SQL = "SELECT sarti5.codigoma,nombrema"
        SQL = SQL & " FROM sarti5,smatact WHERE sarti5.codigoma=smatact.codigoma and sarti5.codartic="
        
        If enlaza Then
            SQL = SQL & DBSet(Text1(0), "T")
        Else
            SQL = SQL & " '-1'"
        End If
        SQL = SQL & " ORDER BY sarti5.codigoma"
    '----
    
    '---- [24/02/2012] Equivalencias
    ElseIf Opcion = 7 Then
    
        SQL = "select codarti1 ,nomartic "
        SQL = SQL & " from sarti6,sartic where sartic.codartic=sarti6.codarti1 AND sarti6.codartic= "
        If enlaza Then
            SQL = SQL & DBSet(Text1(0), "T")
        Else
            SQL = SQL & "'-1'"
        End If
    End If
    
    MontaSQLCarga = SQL
End Function


Private Sub LLamaLineas2(alto As Single, xModo As Byte, Opcion As Byte)
Dim b As Boolean
Dim J As Integer

    ModificaLineas = xModo
    '---- [23/09/2009] LAURA : Añadir lineas de Cod. EAN (se añade modo 8)
    b = (Modo >= 5 Or Modo <= 9) And (ModificaLineas = 1 Or ModificaLineas = 2) 'Insertar o Modificar Lineas

    Select Case Opcion
    Case 2 'CONJUNTOS
        DeseleccionaGrid Me.DataGrid1
        
        txtAux(0).Height = DataGrid1.RowHeight
        txtAux(0).visible = b
        txtAux(0).Top = alto
        txtAux(1).Height = DataGrid1.RowHeight
        txtAux(1).visible = b
        txtAux(1).Top = alto
        txtAux2.Height = DataGrid1.RowHeight
        txtAux2.visible = b
        txtAux2.Top = alto
        cmdAux.visible = b
        cmdAux.Top = alto
        cmdAux.Height = DataGrid1.RowHeight
         
    Case 3 'INSTALACIONES
        DeseleccionaGrid Me.DataGrid2
        If vParamAplic.NumeroInstalacion = vbFontenas Then
            CargaComboCalidad
            cboCalidad.Top = alto
            cboCalidad.visible = ModificaLineas = 1
            For J = 9 To 11
                txtAux(J).Height = DataGrid2.RowHeight
                txtAux(J).visible = True
                txtAux(J).Top = alto
            Next
        Else
            txtAux(2).Height = DataGrid2.RowHeight
            txtAux(2).visible = True
            txtAux(2).Top = alto
            
        End If
    Case 4
        'STOCK
        DeseleccionaGrid Me.DataGrid3
        Text3(0).Height = DataGrid3.RowHeight
        Text3(0).visible = b
        Text3(0).Top = alto
        Text3(2).Height = DataGrid3.RowHeight
        Text3(2).visible = b
        Text3(2).Top = alto
        Text2(8).Height = DataGrid3.RowHeight
        Text2(8).visible = b
        Text2(8).Top = alto
        
        If b Then
            If ModificaLineas = 1 Then
                cmdAlma.visible = b And ModificaLineas = 1
                cmdAlma.Top = alto
                cmdAlma.Height = DataGrid1.RowHeight
            Else
                cmdAlma.visible = False
                Text3(0).Width = DataGrid3.Columns(0).Width
            End If
        Else
            cmdAlma.visible = False
        End If
        
    '---- [23/09/2009] LAURA : Añadir lineas de Cod. EAN
    Case 5 'Lineas Cod. EAN
        DeseleccionaGrid Me.DataGrid4
        txtAux(8).Height = DataGrid4.RowHeight
        txtAux(8).visible = True
        txtAux(8).Top = alto
    
    '---- [19/12/2011]
    Case 6 'Materias activas
        DeseleccionaGrid Me.DataGrid5
        Text5(0).Height = DataGrid5.RowHeight
        Text5(0).visible = b
        Text5(0).Top = alto
        Text5(1).Height = DataGrid5.RowHeight
        Text5(1).visible = b
        Text5(1).Top = alto
        cmdMatAux.visible = b
        cmdMatAux.Top = alto
        cmdMatAux.Height = DataGrid5.RowHeight
         '---- [19/12/2011]
    Case 7
        'Materias activas
        DeseleccionaGrid Me.DataGrid6
        Text6(0).Height = DataGrid5.RowHeight
        Text6(0).visible = b
        Text6(0).Top = alto
        Text6(1).Height = DataGrid5.RowHeight
        Text6(1).visible = b
        Text6(1).Top = alto
        cmdEquiv.visible = b
        cmdEquiv.Top = alto
        cmdEquiv.Height = DataGrid6.RowHeight
     
    '----
    End Select
End Sub


Private Sub txtAux_GotFocus(Index As Integer)
    ConseguirFoco txtAux(Index), Modo
    
End Sub


Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        If (Index = 2) Or Index = 1 Then
            KeyAscii = 0
            PonerFocoBtn Me.cmdAceptar
            Exit Sub
        End If
    End If
    KEYpress KeyAscii
End Sub


Private Sub txtAux_LostFocus(Index As Integer)

    txtAux(Index).Text = Trim(txtAux(Index).Text)
    If txtAux(Index).Text = "" Then Exit Sub
    
    Select Case Index
        Case 0 'cod Articulo de conjunto
            TagText3 = "mateprima"
            txtAux2.Text = DevuelveDesdeBDNew(conAri, "sartic", "nomartic", "codartic", txtAux(0).Text, "T", TagText3)
            MateriaPrima = TagText3 = "1"
            TagText3 = ""
            
        Case 1
            'Si es materiaprima then
            If txtAux(1).Text <> "" Then
                If vParamAplic.ComponentePorcentaje And MateriaPrima Then
                    'Formato decimal
                    'Octubre2017. Antes era formato decimal. Dejamos que sea lo que quiera
                    'If Not PonerFormatoDecimal(txtAux(Index), 4) Then txtAux(1).Text = ""
                    If Not PonerFormatoDecimal(txtAux(Index), 2) Then txtAux(1).Text = ""
                    
                    
                Else
                    If Not PonerFormatoDecimal(txtAux(Index), 2) Then txtAux(1).Text = ""
                End If
                If txtAux(1).Text = "" Then PonerFoco txtAux(1)
                
            End If
            
        Case 10, 11
            'Frmato decimal
            If txtAux(Index).Text <> "" Then
                If Not PonerFormatoDecimal(txtAux(Index), 1) Then txtAux(Index).Text = ""
            End If
            'If Index = 11 Then PonerFocoBtn cmdAceptar
            
    End Select
End Sub


Private Sub PonerBotonCabecera(b As Boolean)
    On Error Resume Next

    Me.cmdAceptar.visible = Not b
    Me.cmdCancelar.visible = Not b
    Me.cmdRegresar.visible = b
    Me.cmdRegresar.Caption = "&Cabecera"
    
    If b Then
        cmdRegresar.Cancel = True
        
        
        '   5.-  Mantenimiento Lineas de Articulos x Almacen
        '   6.-  Mantenimiento Lineas de Componentes de Conjuntos
        '   7.-  Mantenimiento Lineas de Control de Instalaciones
        '   8.-  Mantenimiento Lineas de EAN
        '   9.-  Mantenimiento Lineas de Materias activas
        '   10.- equivalencias
        Select Case Modo
        Case 5
            Me.lblIndicador.Caption = "Lineas stock"
        Case 6
            Me.lblIndicador.Caption = "Lineas conjuntos"
        Case 7
            Me.lblIndicador.Caption = "Lineas instalaciones"
        Case 8
            Me.lblIndicador.Caption = "Lineas EAN"
        
        Case 9
            Me.lblIndicador.Caption = "Lin. Materias activas"
        Case 10
            Me.lblIndicador.Caption = "Lin equivalencias"
        Case Else
            Me.lblIndicador.Caption = "Lineas Detalle"
        End Select
        
        
        PonerFocoBtn Me.cmdRegresar
    Else
        cmdCancelar.Cancel = True
        Me.lblIndicador.Caption = ""
    End If
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Function Eliminar() As Boolean
    Set LOG = New cLOG
    
    

    conn.BeginTrans
    
    If EliminarArticulo(Data1.Recordset!codArtic, lblIndicador) Then
        LOG.Insertar 7, vUsu, Data1.Recordset!codArtic & " " & Data1.Recordset!NomArtic
        conn.CommitTrans
        Eliminar = True
    Else
        conn.RollbackTrans
        Eliminar = False
        
    End If
    Set LOG = Nothing
    lblIndicador.Caption = ""
End Function


Private Sub PosicionarData()
Dim cad As String, Indicador As String

    cad = "codartic=" & DBSet(Text1(0).Text, "T")
    If SituarData(Data1, cad, Indicador) Then
        PonerModo 2
        PonerCampos
        
        lblIndicador.Caption = Indicador
    ElseIf Not Data1.Recordset.EOF Then
'        Data1.Recordset.MoveFirst
        PonerCampos
        PonerModo 2
    ElseIf Modo = 3 Then
        'Acabamos de insertar un registro y lo seleccionamos en el recordset
        CadenaConsulta = "Select * from " & NombreTabla & " WHERE codartic =" & DBSet(Text1(0).Text, "T")
        Data1.RecordSource = CadenaConsulta
        If SituarData(Data1, cad, Indicador) Then
            PonerModo 2
            PonerCampos
            lblIndicador.Caption = Indicador
        End If
    Else
        LimpiarCampos
        PonerModo 0
    End If
End Sub


Private Sub BotonImprimir()
    AbrirListado (6) '6: Informe de Articulos
End Sub


Private Sub AccionesSobreTagText3_(Guardar As Boolean, Cargando As Boolean)
Dim i As Integer

  
    If Guardar Then
        If Cargando Then TagText3 = ""
        For i = 0 To Text3.Count - 1
            If Cargando Then TagText3 = TagText3 & Replace(Text3(i).Tag, "|", ";") & "|"
            Text3(i).Tag = ""
        Next i
        
        'AÑADIMOS EL CHECK chkInventario.
        If Cargando Then TagText3 = TagText3 & Replace(chkInventario.Tag, "|", ";") & "|"
        chkInventario.Tag = ""
    Else
        For i = 0 To Text3.Count - 1
            Text3(i).Tag = Replace(RecuperaValor(TagText3, i + 1), ";", "|")
            Text3(i).Locked = True
        Next i
        chkInventario.Tag = Replace(RecuperaValor(TagText3, i + 1), ";", "|")
    End If
End Sub


Private Sub PonerDatosForaGrid(ForzarLimpiar As Boolean)
Dim i As Integer
Dim Limp As Boolean

    Limp = True
    If Not ForzarLimpiar Then
        If Not (data4.Recordset Is Nothing) Then
            If Not data4.Recordset.EOF Then Limp = False
        End If
    End If
    
    
    If Limp Then

        'Limpiamos
        For i = 0 To Text3.Count - 1
            Text3(i).Text = ""
        Next i
        Text2(6).Text = ""
        Text2(8).Text = ""
        chkInventario.Value = 0
        
    Else
        'EL
    End If
End Sub

'DAVID
'Para poner el foco en un objeto y si da error que no se arrastre
Private Sub PonerFocoObjeto(obj As Object)
    On Error Resume Next
    obj.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub







'-----------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------
'
'
'       El listview tendra los datos de albaranes, facturas... que tenga el cliente
'       Con lo cual, a partir de un click tendremos que ser capaces de situarnos en
'       el formulario correspondiente
'
'-----------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------


Private Sub ImagenesNavegacion()
    With Me.Toolbar2
        .ImageList = frmPpal.ImgListPpal
        .Buttons(1).Image = 5
        .Buttons(3).Image = 6
        .Buttons(5).Image = 7
        .Buttons(7).Image = 1
        .Buttons(11).Image = 2
        .Buttons(13).Image = 10
    End With
    
    Set lw1.SmallIcons = frmPpal.ImgListPpal
End Sub


Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    If Button.Tag = "" Then Exit Sub
    Label2(0).Caption = ""
    'Levantamos todos los botones y dejamos pulsado el de ahora
    For NumRegElim = 1 To Toolbar2.Buttons.Count
        If Toolbar2.Buttons(NumRegElim).Tag <> "" Then
            If Toolbar2.Buttons(NumRegElim).Index <> Button.Index Then Toolbar2.Buttons(NumRegElim).Value = tbrUnpressed
        End If
    Next NumRegElim
    CargaColumnas CByte(Button.Tag)
    Me.Toolbar2.Refresh
    
    'Hacemos las acciones
    If Modo = 2 Then CargaDatosLW
End Sub





Private Sub CargaColumnas(OpcionList As Byte)
Dim Columnas As String
Dim Ancho As String
Dim Alinea As String
Dim Formato As String
Dim Ncol As Integer
Dim C As ColumnHeader

    Select Case OpcionList
    Case 0 'TARIFAS
        Label2(0).Caption = "Tarifas"
        Columnas = "Tarifa|Descripcion |Tipo|Importe|"
        Ancho = "800|2900|850|1500|"
        'vwColumnRight =1  left=0   center=2
        Alinea = "0|0|0|2|"
        'Formatos
        Formato = "|||" & FormatoPrecio & "|"
        Ncol = 4
    
    Case 1 'PRECIOS ESPECIALES
        Label2(0).Caption = "Precios especiales"
        Columnas = "Cod. cli.|Nombre |Precio|"
        Ancho = "1200|3500|1300|"
        'vwColumnRight =1  left=0   center=2
        Alinea = "0|0|1|"
        'Formatos
        Formato = "000||" & FormatoImporte & "|"
        Ncol = 3
        
    Case 2
        Label2(0).Caption = "Promociones"
        Columnas = "Tarifa|Descripcion|F. inicio|F. Fin| Precio|"
        Ancho = "900|2300|1100|1100|1150|"
        'vwColumnRight =1  left=0   center=2
        Alinea = "0|0|0|0|1|"
        'Formatos
        Formato = "000||dd/mm/yyyy|dd/mm/yyyy|" & FormatoPrecio & "|"
        Ncol = 5
        
    Case 3 'PEDIDOS
        Label2(0).Caption = "PEDIDOS"
        Columnas = "NºPed|Fecha|Cod.|Nombre|Candtidad|"
        Ancho = "1250|1100|800|2300|1000|"
        'vwColumnRight =1  left=0   center=2
        Alinea = "0|0|0|0|1|"
        'Formatos
        Formato = "|dd/mm/yyyy|||" & FormatoImporte & "|"
        Ncol = 5
        
    Case 4
        'MOVIMIENTOS
        Label2(0).Caption = "MOVIMIENTOS ALMACEN"
        Columnas = "Alm|Fecha|Tipo|Entrada|Documento|Cantidad|C/P/T|"
        Ancho = "600|1100|900|900|1000|1000|900|"
        'vwColumnRight =1  left=0   center=2
        Alinea = "0|0|0|0|0|1|1|"
        'Formatos
        Formato = "|dd/mm/yyyy||||" & FormatoCantidad & "||"
        Ncol = 7
        
    Case 5
        'Precios proveedor
        Label2(0).Caption = "PRECIOS PROVEE."
        Columnas = "Prov.|Nombre|Precio|Cambio|Precio N.|"
        Ancho = "1200|2400|1050|900|1050|"
        'vwColumnRight =1  left=0   center=2
        Alinea = "0|0|1|0|1|"
        'Formatos
        Formato = "000||" & FormatoPrecio & "|dd/mm/yy|" & FormatoPrecio & "|"
        Ncol = 5
    End Select
    
    Me.FrameDisponible.visible = OpcionList = 3

    'Guardo la opcion en el tag
    lw1.Tag = OpcionList & "|" & Ncol & "|"
    
    lw1.ColumnHeaders.Clear
    
    For NumRegElim = 1 To Ncol
         Set C = lw1.ColumnHeaders.Add()
         C.Text = RecuperaValor(Columnas, CInt(NumRegElim))
         C.Width = RecuperaValor(Ancho, CInt(NumRegElim))
         C.Alignment = Val(RecuperaValor(Alinea, CInt(NumRegElim)))
         C.Tag = RecuperaValor(Formato, CInt(NumRegElim))
    Next NumRegElim
End Sub


Private Sub CargaDatosLW()
Dim C As String
Dim bs As Byte
    bs = Screen.MousePointer
    C = Me.lblIndicador.Caption
    lblIndicador.Caption = "Leyendo " & Label2(0).Caption
    lblIndicador.Refresh
    CargaDatosLW2
    Me.lblIndicador.Caption = C
    Screen.MousePointer = bs
End Sub

Private Sub CargaDatosLW2()
Dim cad As String
Dim RS As ADODB.Recordset
Dim IT As ListItem
Dim ElIcono As Integer



    On Error GoTo ECargaDatosLW
    
    If Modo <> 2 Then Exit Sub
    
    For NumRegElim = 1 To Toolbar2.Buttons.Count
        If Toolbar2.Buttons(NumRegElim).Value = tbrPressed Then
            ElIcono = Toolbar2.Buttons(NumRegElim).Image
            Exit For
        End If
    Next
    
    

    
    Select Case CByte(RecuperaValor(lw1.Tag, 1))
    Case 0
        'OFERTAS
        cad = "select l.codlista,nomlista,if(opcionINC=0,""PVP"",""UPC""),precioac from slista l,starif c where c.codlista=l.codlista"

        BuscaChekc = ""
    Case 1
        'Precios especiales
        cad = "select l.codclien,nomclien,precioac from sprees l,sclien s where s.codclien=l.codclien"
        BuscaChekc = ""

        
    Case 2
        'Promociones
        cad = "select l.codlista,nomlista,fechaini,fechafin,precioac from spromo l, starif s where l.codlista=s.codlista"
        BuscaChekc = ""
   
    Case 3
        '*****************************
        'Es una funcion especial
        CargaDatosPedidos
        Exit Sub
        
    Case 4
        'Cargamos movimientos almacen
        cad = "select codalmac,fechamov,detamovi,if(tipomovi=1,""*"","" ""),document,cantidad,codigope from smoval l WHERE 1=1 "
        'BuscaChekc = "ORDER BY fechamov desc,horamovi desc"
        If vParamAplic.NumeroInstalacion = vbHerbelca Then
            BuscaChekc = " ORDER BY codartic,fechamov desc, horamovi desc," & NombreTabla & ".codalmac "
        Else
            BuscaChekc = " ORDER BY codartic," & NombreTabla & ".codalmac, fechamov desc, horamovi desc"
        End If
            
        
        
    Case 5
        cad = "select l.codprove,nomprove,precioac,fechanue,precionu from slispr l inner join sprove on l.codprove=sprove.codprove WHERE 1=1 "
        BuscaChekc = ""
    End Select
    
    
    'La fecha
    
    'EL where del codclien
    cad = cad & " and l.codartic='" & DevNombreSQL(Data1.Recordset!codArtic) & "'"
    
    
    

    
    'El ORDER BY
    If BuscaChekc <> "" Then cad = cad & " ORDER BY fechamov desc,horamovi desc"
    BuscaChekc = ""
    
    lw1.ListItems.Clear
    Set RS = New ADODB.Recordset
    
    RS.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
        Set IT = lw1.ListItems.Add()
        If lw1.ColumnHeaders(1).Tag <> "" Then
            IT.Text = Format(RS.Fields(0), lw1.ColumnHeaders(1).Tag)
        Else
            IT.Text = RS.Fields(0)
        End If
        'El resto de cmpos
        For NumRegElim = 2 To CInt(RecuperaValor(lw1.Tag, 2))
            If IsNull(RS.Fields(NumRegElim - 1)) Then
                IT.SubItems(NumRegElim - 1) = " "
            Else
                If lw1.ColumnHeaders(NumRegElim).Tag <> "" Then
                    IT.SubItems(NumRegElim - 1) = Format(RS.Fields(NumRegElim - 1), lw1.ColumnHeaders(NumRegElim).Tag)
                Else
                    IT.SubItems(NumRegElim - 1) = RS.Fields(NumRegElim - 1)
                End If
            End If
        Next
        IT.SmallIcon = ElIcono
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    
    
    Exit Sub
ECargaDatosLW:
    MuestraError Err.Number, "", Err.Description
    Set RS = Nothing
    
End Sub



Private Sub CargaDatosPedidos()
Dim C As String
Dim Importe As Currency
Dim T As Currency
    
    'Limpiamos
    lw1.ListItems.Clear
    For NumRegElim = 1 To 3
        Text4(NumRegElim).Text = ""
    Next
    
    'Cargamos el primer combo
    Text4(0).Text = txtSumaStock.Text
    T = 0
    If txtSumaStock.Text <> "" Then T = ImporteFormateado(txtSumaStock.Text)
        
        
    
    'Cargamos primero los de cliente
    If vParamAplic.NumeroInstalacion <> vbAmesa Then
        C = "select scaped.numpedcl,fecpedcl,codclien,nomclien,sum(cantidad) as cuantos"
        C = C & " from scaped,sliped where scaped.numpedcl=sliped.numpedcl and cerrado=0 and codartic='"
        C = C & DevNombreSQL(Data1.Recordset!codArtic) & "' GROUP BY 1"
        Importe = CargaListPedidos(6, C)
    
        T = T - Importe
    End If
    
    
    'JULIO 19
    If vParamAplic.Produccion Then
           'Importe: a tienen cargado los datos
          C = " sordprod.codigo=sliordpr.codigo and fecproduccion is null and codartic = " & DBSet(Text1(0).Text, "T") & " AND 1"
          C = DevuelveDesdeBD(conAri, "sum(cantidad) as cuantos ", "sordprod ,sliordpr", C, "1")
          If C = "" Then C = "0"
          Importe = Importe + CCur(C)
          T = T + CCur(C)
          
          C = " sordprod.codigo=sliordpr2.codigo and fecproduccion is null and codarti2 = " & DBSet(Text1(0).Text, "T") & " AND 1"
          C = DevuelveDesdeBD(conAri, "sum(cantidad) as cuantos ", "sordprod ,sliordpr2", C, "1")
          If C = "" Then C = "0"
          Importe = Importe - CCur(C)
          T = T - CCur(C)
          
          
          
          
    End If
    
    'JULIO 19
    If vParamAplic.Produccion Then
           'Importe: a tienen cargado los datos
          C = " sordprod.codigo=sliordpr.codigo and fecproduccion is null and codartic = " & DBSet(Text1(0).Text, "T") & " AND 1"
          C = DevuelveDesdeBD(conAri, "sum(cantidad) as cuantos ", "sordprod ,sliordpr", C, "1")
          If C = "" Then C = "0"
          Importe = Importe + CCur(C)
          T = T + CCur(C)
          
          C = " sordprod.codigo=sliordpr2.codigo and fecproduccion is null and codarti2 = " & DBSet(Text1(0).Text, "T") & " AND 1"
          C = DevuelveDesdeBD(conAri, "sum(cantidad) as cuantos ", "sordprod ,sliordpr2", C, "1")
          If C = "" Then C = "0"
          Importe = Importe - CCur(C)
          T = T - CCur(C)
          
          
          
          
    End If
    
    Text4(1).Text = Format(Importe, FormatoImporte)
    
    
    
    'Cargamos los comprados
    C = "select scappr.numpedpr,fecpedpr,codprove,nomprove,sum(cantidad) as cuantos"
    C = C & " from scappr,slippr where scappr.numpedpr=slippr.numpedpr  and codartic='"
    C = C & DevNombreSQL(Data1.Recordset!codArtic) & "' group by 1"
    Importe = CargaListPedidos(9, C)
    T = T + Importe
    Text4(2).Text = Format(Importe, FormatoImporte)
    'Disponible
    Text4(3).Text = Format(T, FormatoImporte)
End Sub


Private Function CargaListPedidos(ByRef ElIcono As Integer, cad As String) As Currency
Dim RS As ADODB.Recordset
Dim IT As ListItem
Dim cantidad As Currency

    Set RS = New ADODB.Recordset
    
    cantidad = 0
    RS.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
        Set IT = lw1.ListItems.Add()
        If lw1.ColumnHeaders(1).Tag <> "" Then
            IT.Text = Format(RS.Fields(0), lw1.ColumnHeaders(1).Tag)
        Else
            IT.Text = RS.Fields(0)
        End If
        'El resto de cmpos
        For NumRegElim = 2 To CInt(RecuperaValor(lw1.Tag, 2))
            If IsNull(RS.Fields(NumRegElim - 1)) Then
                IT.SubItems(NumRegElim - 1) = " "
            Else
                If lw1.ColumnHeaders(NumRegElim).Tag <> "" Then
                    IT.SubItems(NumRegElim - 1) = Format(RS.Fields(NumRegElim - 1), lw1.ColumnHeaders(NumRegElim).Tag)
                Else
                    IT.SubItems(NumRegElim - 1) = RS.Fields(NumRegElim - 1)
                End If
            End If
        Next
        cantidad = cantidad + DBLet(RS!Cuantos, "N")
        IT.SmallIcon = ElIcono
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    CargaListPedidos = cantidad
End Function




Private Sub ponerDatosConjuntos()
Dim Im1 As Currency
Dim Im2 As Currency
Dim Aux As Currency

    On Error GoTo EponerDatosConjuntos
    'Signo los valores del articulo del UPC y PVP
    txtConjunto(0).Text = Text1(15).Text
    txtConjunto(3).Text = Text1(17).Text
    
    If Data2.Recordset.RecordCount > 0 Then Me.Data2.Recordset.MoveFirst
        
    'Recorrer el RS buscando los importes reales
    While Not Data2.Recordset.EOF
    '
        'COSTE
        Aux = DBLet(Data2.Recordset!cantidad, "N")
        Aux = Aux * DBLet(Data2.Recordset!precioUC, "N")
        
        If vParamAplic.ComponentePorcentaje Then
            If Data2.Recordset!MateriaPrima = "*" Then Aux = Aux / 100
        End If
        Im1 = Im1 + Aux
        
        'PVP
        Aux = DBLet(Data2.Recordset!cantidad, "N")
        Aux = Aux * Data2.Recordset!PrecioVe
        If vParamAplic.ComponentePorcentaje Then
            If Data2.Recordset!MateriaPrima = "*" Then Aux = Aux / 100
        End If
        Im2 = Im2 + Aux
            
        
        
        Data2.Recordset.MoveNext
    Wend
    If Data2.Recordset.RecordCount > 0 Then Me.Data2.Recordset.MoveFirst
    txtConjunto(1).Text = Format(Im1, FormatoPrecio)
    txtConjunto(4).Text = Format(Im2, FormatoPrecio)
    
    'Difernecias
    Im1 = ImporteFormateado(txtConjunto(0).Text) - Im1
    Im2 = ImporteFormateado(txtConjunto(3).Text) - Im2
    txtConjunto(2).Text = Format(Im1, FormatoPrecio)
    txtConjunto(5).Text = Format(Im2, FormatoPrecio)
    
    Exit Sub
EponerDatosConjuntos:
    MuestraError Err.Number, Err.Description
End Sub



Private Function ComprobarPorcentajesCorrectos() As Boolean
    On Error GoTo EComprobarPorcentajesCorrectos
    ComprobarPorcentajesCorrectos = True
    Set miRsAux = New ADODB.Recordset
    BuscaChekc = "SELECT  sum(sarti1.Cantidad) FROM   sarti1 INNER JOIN sartic ON sarti1.codarti1 = sartic.codArtic"
    BuscaChekc = BuscaChekc & " where mateprima=1 and sarti1.codartic=" & DBSet(Text1(0), "T")
    miRsAux.Open BuscaChekc, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        If Not IsNull(miRsAux.Fields(0)) Then
            If miRsAux.Fields(0) <> 100 Then
                MsgBox "La suma de porcentajes de los componenetes no es 100(" & miRsAux.Fields(0) & ")", vbExclamation
                ComprobarPorcentajesCorrectos = False
            End If
        End If
    End If
    miRsAux.Close
EComprobarPorcentajesCorrectos:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "ComprobarPorcentajesCorrectos", Err.Description
        ComprobarPorcentajesCorrectos = False
    End If
    BuscaChekc = ""
    Set miRsAux = Nothing
End Function






Private Sub DataGrid1EnSMOVAL()
'Abrir el formulario del Mantenimiento del que viene el Movimiento
'Se busca en histórico o en Form
Dim SQL As String
    
    Select Case lw1.SelectedItem.SubItems(2)
        Case "TRA" 'traspaso de almacenes
            'Traspaso de Almacen
            With frmAlmTraspaso
                .EsHistorico = True
                .hcoCodMovim = lw1.SelectedItem.SubItems(4)
                .hcoFechaMovim = lw1.SelectedItem.SubItems(1)
                .Show vbModal
            End With
            
        Case "REG" 'Movimientos de Almacen
                    'Movimientos de Almacen
            With frmAlmMovimientos
                .EsHistorico = True
                .hcoCodMovim = lw1.SelectedItem.SubItems(4)
                .hcoFechaMovim = lw1.SelectedItem.SubItems(1)
                .Show vbModal
            End With

        Case "ALV", "ART", "ALM", "ALZ", "ALR", "ALS", "ALO", "ALE", "ALD", "ALB"
                                'ALV:Albaran de Venta (a clientes)
                                'ART: Albaran rectificativo
                                'ALM: ALbaran Mostrador
                                'ALZ: Albaranes "B"
            'comprobar si el Albaran esta facturado o no
            'si no esta facturado abrir el formulario de Entrada de Albaranes: frmFacEntAlbaranes
            'si esta ya facturado abrir el histórico de facturas: frmFacHcoFacturas
            
            'consultamos si existe el albaran en la tabla de albaranes: scaalb
            SQL = DevuelveDesdeBDNew(conAri, "scaalb", "numalbar", "codtipom", lw1.SelectedItem.SubItems(2), "T", , "numalbar", lw1.SelectedItem.SubItems(4), "N")
            If SQL <> "" Then 'existe el Albaran
                    'Abrira un frm u otro
                    If vParamAplic.TipoFormularioClientes = 0 Then
                         With frmFacEntAlbaranes2
                            If EsNumerico(lw1.SelectedItem.SubItems(4)) Then
                                .hcoCodMovim = Format(lw1.SelectedItem.SubItems(4), "0000000")
                            Else
                                .hcoCodMovim = lw1.SelectedItem.SubItems(4)
                            End If
                            .hcoCodTipoM = lw1.SelectedItem.SubItems(2)
                            .Show vbModal
                        End With
                    Else
                        'SAIL
                        With frmFacEntAlbSAIL
                            If EsNumerico(lw1.SelectedItem.SubItems(4)) Then
                                .hcoCodMovim = Format(lw1.SelectedItem.SubItems(4), "0000000")
                            Else
                                .hcoCodMovim = lw1.SelectedItem.SubItems(4)
                            End If
                            .hcoCodTipoM = lw1.SelectedItem.SubItems(2)
                            .Show vbModal
                        End With
                    End If
            Else 'No existe en albaran, abrir Historico Factura
                With frmFacHcoFacturas2
                    .DesdeFichaCliente = False
                    If EsNumerico(lw1.SelectedItem.SubItems(4)) Then
                        .hcoCodMovim = Format(lw1.SelectedItem.SubItems(4), "0000000")
                    Else
                        .hcoCodMovim = lw1.SelectedItem.SubItems(4)
                    End If
                    .hcoCodTipoM = lw1.SelectedItem.SubItems(2)
                    .hcoFechaMov = lw1.SelectedItem.SubItems(1)
                    
                    .Show vbModal
                End With
            End If
            

'             With frmFacEntAlbaranes
'                If EsNumerico(lw1.SelectedItem.SubItems(4)) Then
'                    .hcoCodMovim = Format(lw1.SelectedItem.SubItems(4), "0000000")
'                Else
'                    .hcoCodMovim = lw1.SelectedItem.SubItems(4)
'                End If
'                .hcoCodTipoM = lw1.SelectedItem.SubItems(2)
'                .RecuperarFactu = False
'                .Show vbModal
'            End With
            
        Case "ALC" 'Albaran de Compra (a Proveedores)
            'comprobar si el Albaran esta facturado o no
            'si no esta facturado abrir el formulario de Entrada de Albaranes: frmComEntAlbaranes
            'si esta ya facturado abrir el histórico de facturas: frmComHcoFacturas
            
            'consultamos si existe el albaran en la tabla de albaranes: scaalp
            SQL = DevuelveDesdeBDNew(conAri, "scaalp", "numalbar", "codprove", lw1.SelectedItem.SubItems(6), "N", , "numalbar", lw1.SelectedItem.SubItems(4), "T", "fechaalb", lw1.SelectedItem.SubItems(1), "F")
            
            If SQL <> "" Then 'existe el Albaran
                If vParamAplic.TipoFormularioClientes = 0 Then
                    With frmComEntAlbaranesGR
                        .hcoCodMovim = Trim(lw1.SelectedItem.SubItems(4))
                        .hcoFechaMovim = lw1.SelectedItem.SubItems(1)
                        .hcoCodProve = lw1.SelectedItem.SubItems(6) 'aqui es el proveedor
                        .Show vbModal
                    End With
                 Else
                    With frmComEntAlbaranSA
                        .hcoCodMovim = Trim(lw1.SelectedItem.SubItems(4))
                        .hcoFechaMovim = lw1.SelectedItem.SubItems(1)
                        .hcoCodProve = lw1.SelectedItem.SubItems(6) 'aqui es el proveedor
                        .Show vbModal
                    End With
                 
                 End If
            Else        'No existe en albaran, abrir Historico Factura
                If vParamAplic.TipoFormularioClientes = 0 Then
                    With frmComHcoFacturas2GR
                        .hcoCodMovim = Trim(lw1.SelectedItem.SubItems(4))
                        .hcoFechaMovim = lw1.SelectedItem.SubItems(1)
                        .hcoCodProve = lw1.SelectedItem.SubItems(6) 'aqui es el proveedor
                        .Show vbModal
                    End With
                    
                Else
                    'SAIL
                    With frmComHcoFacturSA
                       .hcoCodMovim = Trim(lw1.SelectedItem.SubItems(4))
                        .hcoFechaMovim = lw1.SelectedItem.SubItems(1)
                        .hcoCodProve = lw1.SelectedItem.SubItems(6) 'aqui es el proveedor
                        .Show vbModal
                    End With
                    
                End If
            End If
            
            
        '**********************************
        'Laura: modificado 11/09/06
'        Case "FTI" 'Factura Ticket de venta
        Case "ATI" 'Albaran Ticket de venta
        '**********************************
            'Abrir el historico de facturas
             With frmFacHcoFacturas2
                .DesdeFichaCliente = False
                If EsNumerico(lw1.SelectedItem.SubItems(4)) Then
                    .hcoCodMovim = Format(lw1.SelectedItem.SubItems(4), "0000000")
                Else
                    .hcoCodMovim = lw1.SelectedItem.SubItems(4)
                End If
                .hcoCodTipoM = lw1.SelectedItem.SubItems(2)
                .hcoFechaMov = lw1.SelectedItem.SubItems(1)
                .Show vbModal
            End With
    Case "DFI"
        MsgBox "Diferencias de inventario.", vbInformation
    End Select
End Sub




Private Sub ActualizarEAN()

    'codprove|nomprove|refprove|precio|nomartic|ean|codtelem|
    
    
    'Acutalizo TELEMATEL
    BuscaChekc = Mid(DatosADevolverBusqueda, 3)
    BuscaChekc = RecuperaValor(BuscaChekc, 7)
    BuscaChekc = " WHERE codtelem = " & BuscaChekc
    BuscaChekc = "UPDATE stelem set codartic =" & DBSet(Text1(0).Text, "T") & BuscaChekc
    ejecutar BuscaChekc, False
    
    BuscaChekc = Mid(DatosADevolverBusqueda, 3)
    BuscaChekc = RecuperaValor(BuscaChekc, 6)
    
    
    BuscaChekc = "INSERT INTO sarti3(codartic, numlinea ,codigoea) VALUES (" & DBSet(Text1(0).Text, "T") & ",1,'" & BuscaChekc & "')"
    ejecutar BuscaChekc, False
    
    CadenaDesdeOtroForm = "I"
End Sub




Private Function InsertarModificarEQUIV() As Boolean
Dim SQL As String
Dim Valor As String

    On Error GoTo ErrInsModEAN
    InsertarModificarEQUIV = False
    
    If Text6(0).Text = "" Or Text6(1).Text = "" Then
        
        MsgBox "Error articulo equivalente", vbExclamation
        Exit Function
    End If
    
        'Por si acaso
    If Text6(0).Text = Text1(0).Text Then
        MsgBox "YA SE QUE SOY EQUIVALENTE A MI MISMO... to  myself", vbExclamation
        Exit Function
    End If
    
    If ModificaLineas = 1 Then 'INSERTAR
        cmdAceptar.Tag = Text6(0).Text
        SQL = "INSERT INTO sarti6(codartic,codarti1) VALUES ("
        SQL = SQL & DBSet(Text1(0).Text, "T") & "," & DBSet(Text6(0).Text, "T") & ") "
        conn.Execute SQL
        
        'Y la "equivalente"
        SQL = "INSERT IGNORE INTO sarti6(codartic,codarti1) VALUES ("
        SQL = SQL & DBSet(Text6(0).Text, "T") & "," & DBSet(Text1(0).Text, "T") & ") "
    ElseIf ModificaLineas = 2 Then 'MODIFICAR
  
    End If
    
    conn.Execute SQL
    InsertarModificarEQUIV = True
    Exit Function

ErrInsModEAN:
    MuestraError Err.Number, "Insertar/Modificar equivalencia", Err.Description
    PonerFoco Text6(0)
End Function





Private Sub PonerCodigoArticuloEULER(DesdeCmdAceptar As Boolean)
Dim cad As String
    If Me.Text1(3).Text <> "" Then
        If Text2(1).Text <> "" Then
            'select codartic,substring(codartic,4)+0 from sartic where codartic like '001%' order by 2 desc
            cad = Mid(Text2(1).Text, 1, 3)
            cad = "codartic like '" & cad & "%' AND 1"
            cad = DevuelveDesdeBD(conAri, "substring(codartic,4)+0", "sartic", cad, "1 ORDER BY 1 DESC")
            If cad = "" Then cad = "0"
            NumRegElim = Val(cad) + 1
            cad = Format(NumRegElim, "000000")
            cad = Mid(Text2(1).Text, 1, 3) & cad
            If DesdeCmdAceptar Then
                If Text1(0).Text <> cad Then
                    If MsgBox("Le corresponde el articulo: " & cad & vbCrLf & "¿Continuar de igual modo?", vbQuestion + vbYesNo) = vbYes Then Exit Sub
                    
                End If
            End If
            Text1(0).Text = cad
        End If
    End If
End Sub


Private Sub CargaComboCalidad()
    CargarCombo_Tabla cboCalidad, "scalidad", "codigo", "ensayo", , True, "ensayo"
End Sub
