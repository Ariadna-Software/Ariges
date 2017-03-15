VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAlmFamiliaArticulo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Familias de artículos"
   ClientHeight    =   8310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13470
   Icon            =   "frmAlmFamiliaArticulo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8310
   ScaleWidth      =   13470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Digitos 1er nivel|N|N|||empresa|numdigi1|||"
   Begin VB.TextBox Text1 
      Height          =   4995
      Index           =   14
      Left            =   8520
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   20
      Tag             =   "Descr|T|S|||sfamia|descripcion||N|"
      Text            =   "frmAlmFamiliaArticulo.frx":000C
      Top             =   1320
      Width           =   4605
   End
   Begin MSAdodcLib.Adodc Data2 
      Height          =   375
      Left            =   3000
      Top             =   7800
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
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
      Height          =   1095
      Left            =   120
      TabIndex        =   45
      Top             =   6480
      Width           =   12975
      Begin VB.CheckBox chkMarcaPropia 
         Height          =   195
         Left            =   10800
         TabIndex        =   16
         Tag             =   "Comunica|N|N|||sfamia|marcapropia||N|"
         Top             =   240
         Width           =   375
      End
      Begin VB.CheckBox chkComunica 
         Height          =   195
         Left            =   8880
         TabIndex        =   15
         Tag             =   "Comunica|N|N|||sfamia|comunica||N|"
         Top             =   300
         Width           =   375
      End
      Begin VB.CheckBox chkBloqTPV 
         Height          =   195
         Left            =   6000
         TabIndex        =   14
         Tag             =   "Bloq TPV|N|N|||sfamia|bloqEnTPV||N|"
         Top             =   300
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   8
         Left            =   1200
         MaxLength       =   4
         TabIndex        =   13
         Tag             =   "Centro de coste|T|S|||sfamia|codccost||N|"
         Top             =   720
         Width           =   630
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   8
         Left            =   1920
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   48
         Text            =   "Text2"
         Top             =   720
         Width           =   3585
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   9
         Left            =   2160
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   46
         Text            =   "Text2"
         Top             =   240
         Width           =   3585
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   9
         Left            =   1200
         TabIndex        =   12
         Tag             =   "proveedor|N|S|0||sfamia|codprove|0000||"
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Marca propia"
         Height          =   255
         Left            =   11280
         TabIndex        =   61
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Comunica"
         Height          =   255
         Left            =   9360
         TabIndex        =   54
         Top             =   270
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Bloquea en TPV"
         Height          =   255
         Left            =   6480
         TabIndex        =   53
         Top             =   270
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "CCoste"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   49
         Top             =   720
         Width           =   615
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   960
         Picture         =   "frmAlmFamiliaArticulo.frx":0012
         ToolTipText     =   "Buscar centro coste"
         Top             =   720
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Proveedor"
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   47
         Top             =   240
         Width           =   855
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   960
         Picture         =   "frmAlmFamiliaArticulo.frx":0114
         ToolTipText     =   "Buscar centro coste"
         Top             =   240
         Width           =   240
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5295
      Left            =   120
      TabIndex        =   30
      Top             =   1080
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   9340
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Contabilidad"
      TabPicture(0)   =   "frmAlmFamiliaArticulo.frx":0216
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Descuentos"
      TabPicture(1)   =   "frmAlmFamiliaArticulo.frx":0232
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1(10)"
      Tab(1).Control(1)=   "DataGrid1"
      Tab(1).Control(2)=   "txtAux(1)"
      Tab(1).Control(3)=   "txtAux(0)"
      Tab(1).Control(4)=   "Combo1"
      Tab(1).Control(5)=   "Text1(10)"
      Tab(1).ControlCount=   6
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   10
         Left            =   -71760
         MaxLength       =   5
         TabIndex        =   51
         Tag             =   "Centro de coste|N|S|0|100|sfamia|maxdtopar|#0.00|N|"
         Top             =   4200
         Width           =   975
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   -74760
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   3000
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   0
         Left            =   -71880
         MaxLength       =   18
         TabIndex        =   18
         Tag             =   "Código Artículo"
         Text            =   "Artic Artic Artic5"
         Top             =   3000
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   1
         Left            =   -70200
         MaxLength       =   15
         TabIndex        =   19
         Tag             =   "Código Almacen"
         Text            =   "codalmac"
         Top             =   3120
         Visible         =   0   'False
         Width           =   615
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   2535
         Left            =   -74880
         TabIndex        =   50
         Top             =   960
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   4471
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
               LCID            =   1034
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
               LCID            =   1034
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
      Begin VB.Frame Frame3 
         Caption         =   "Compras "
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
         Height          =   1575
         Left            =   240
         TabIndex        =   40
         Top             =   3480
         Width           =   7695
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   13
            Left            =   2400
            MaxLength       =   10
            TabIndex        =   11
            Tag             =   "Cta. Compras servicios|T|S|||sfamia|ctacomprser||N|"
            Text            =   "Text1"
            Top             =   1080
            Width           =   1125
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   13
            Left            =   3600
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   59
            Text            =   "Text2"
            Top             =   1080
            Width           =   3885
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   3
            Left            =   3600
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   42
            Text            =   "Text2"
            Top             =   240
            Width           =   3885
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   5
            Left            =   3600
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   41
            Text            =   "Text2"
            Top             =   675
            Width           =   3885
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   5
            Left            =   2400
            MaxLength       =   10
            TabIndex        =   10
            Tag             =   "Cta.Abono Compras|T|N|||sfamia|abocompr||N|"
            Text            =   "Text1"
            Top             =   675
            Width           =   1125
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   3
            Left            =   2400
            MaxLength       =   10
            TabIndex        =   9
            Tag             =   "Cta. Contable compras|T|N|||sfamia|ctacompr||N|"
            Text            =   "Text1"
            Top             =   240
            Width           =   1125
         End
         Begin VB.Label Label1 
            Caption         =   "Cta. compras servicios"
            Height          =   255
            Index           =   13
            Left            =   240
            TabIndex        =   60
            Top             =   1080
            Width           =   1695
         End
         Begin VB.Image imgCuentas 
            Height          =   240
            Index           =   8
            Left            =   2040
            Picture         =   "frmAlmFamiliaArticulo.frx":024E
            ToolTipText     =   "Buscar cuenta contable"
            Top             =   1125
            Width           =   240
         End
         Begin VB.Image imgCuentas 
            Height          =   240
            Index           =   3
            Left            =   2040
            Picture         =   "frmAlmFamiliaArticulo.frx":0350
            ToolTipText     =   "Buscar cuenta contable"
            Top             =   720
            Width           =   240
         End
         Begin VB.Image imgCuentas 
            Height          =   240
            Index           =   1
            Left            =   2040
            Picture         =   "frmAlmFamiliaArticulo.frx":0452
            ToolTipText     =   "Buscar cuenta contable"
            Top             =   315
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Cta. Contable Compras"
            Height          =   255
            Index           =   6
            Left            =   240
            TabIndex        =   44
            Top             =   270
            Width           =   1815
         End
         Begin VB.Label Label1 
            Caption         =   "Cta. Abono Compras"
            Height          =   255
            Index           =   8
            Left            =   240
            TabIndex        =   43
            Top             =   675
            Width           =   1455
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Ventas "
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
         Height          =   2895
         Left            =   240
         TabIndex        =   31
         Top             =   480
         Width           =   7695
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   12
            Left            =   3600
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   57
            Text            =   "Text2"
            Top             =   2400
            Width           =   3885
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   12
            Left            =   2400
            MaxLength       =   10
            TabIndex        =   8
            Tag             =   "Cta. alternativa vta sevicios|T|N|||sfamia|ctavtaseralt||N|"
            Text            =   "Text1"
            Top             =   2400
            Width           =   1125
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   11
            Left            =   3600
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   55
            Text            =   "Text2"
            Top             =   1920
            Width           =   3885
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   11
            Left            =   2400
            MaxLength       =   10
            TabIndex        =   7
            Tag             =   "Cta. Alternativa Ventas|T|N|||sfamia|ctavtaser||N|"
            Text            =   "Text1"
            Top             =   1920
            Width           =   1125
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   6
            Left            =   2400
            MaxLength       =   10
            TabIndex        =   5
            Tag             =   "Cta. Alternativa Ventas|T|N|||sfamia|ctavent1||N|"
            Text            =   "Text1"
            Top             =   1080
            Width           =   1125
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   7
            Left            =   2400
            MaxLength       =   10
            TabIndex        =   6
            Tag             =   "Cta. Alternativa Abonos|T|N|||sfamia|abovent1||N|"
            Text            =   "Text1"
            Top             =   1485
            Width           =   1125
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   7
            Left            =   3600
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   35
            Text            =   "Text2"
            Top             =   1485
            Width           =   3885
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   6
            Left            =   3600
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   34
            Text            =   "Text2"
            Top             =   1080
            Width           =   3885
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   2
            Left            =   3600
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   33
            Text            =   "Text2"
            Top             =   240
            Width           =   3885
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   4
            Left            =   3600
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   32
            Text            =   "Text2"
            Top             =   675
            Width           =   3885
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   4
            Left            =   2400
            MaxLength       =   10
            TabIndex        =   4
            Tag             =   "Cta. Abono Ventas|T|N|||sfamia|aboventa||N|"
            Text            =   "Text1"
            Top             =   675
            Width           =   1125
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   2
            Left            =   2400
            MaxLength       =   10
            TabIndex        =   3
            Tag             =   "Cta. Contable Ventas|T|N|||sfamia|ctaventa||N|"
            Text            =   "Text1"
            Top             =   240
            Width           =   1125
         End
         Begin VB.Image imgCuentas 
            Height          =   240
            Index           =   7
            Left            =   2040
            Picture         =   "frmAlmFamiliaArticulo.frx":0554
            Tag             =   "-1"
            ToolTipText     =   "Buscar cuenta contable"
            Top             =   2445
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Cta. Alternativa servicios"
            Height          =   255
            Index           =   12
            Left            =   240
            TabIndex        =   58
            Top             =   2430
            Width           =   1815
         End
         Begin VB.Image imgCuentas 
            Height          =   240
            Index           =   6
            Left            =   2040
            Picture         =   "frmAlmFamiliaArticulo.frx":0656
            Tag             =   "-1"
            ToolTipText     =   "Buscar cuenta contable"
            Top             =   1965
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Cta. Ventas servicios"
            Height          =   255
            Index           =   11
            Left            =   240
            TabIndex        =   56
            Top             =   2010
            Width           =   1815
         End
         Begin VB.Label Label1 
            Caption         =   "Cta. Alternativa Abonos"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   39
            Top             =   1515
            Width           =   1815
         End
         Begin VB.Label Label1 
            Caption         =   "Cta. Alternativa Ventas"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   38
            Top             =   1110
            Width           =   1815
         End
         Begin VB.Image imgCuentas 
            Height          =   240
            Index           =   4
            Left            =   2040
            Picture         =   "frmAlmFamiliaArticulo.frx":0758
            Tag             =   "-1"
            ToolTipText     =   "Buscar cuenta contable"
            Top             =   1125
            Width           =   240
         End
         Begin VB.Image imgCuentas 
            Height          =   240
            Index           =   5
            Left            =   2040
            Picture         =   "frmAlmFamiliaArticulo.frx":085A
            ToolTipText     =   "Buscar cuenta contable"
            Top             =   1560
            Width           =   240
         End
         Begin VB.Image imgCuentas 
            Height          =   240
            Index           =   2
            Left            =   2040
            Picture         =   "frmAlmFamiliaArticulo.frx":095C
            ToolTipText     =   "Buscar cuenta contable"
            Top             =   705
            Width           =   240
         End
         Begin VB.Image imgCuentas 
            Height          =   240
            Index           =   0
            Left            =   2040
            Picture         =   "frmAlmFamiliaArticulo.frx":0A5E
            Tag             =   "-1"
            ToolTipText     =   "Buscar cuenta contable"
            Top             =   285
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Cta. Contable Ventas"
            Height          =   255
            Index           =   5
            Left            =   240
            TabIndex        =   37
            Top             =   270
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Cta. Abono Ventas"
            Height          =   255
            Index           =   7
            Left            =   240
            TabIndex        =   36
            Top             =   675
            Width           =   1575
         End
      End
      Begin VB.Label Label1 
         Caption         =   "Maximo descuento particulares"
         Height          =   195
         Index           =   10
         Left            =   -74880
         TabIndex        =   52
         Top             =   4320
         Width           =   2835
      End
   End
   Begin VB.CheckBox chkInstalac 
      Caption         =   "¿Es instalación?"
      Height          =   195
      Left            =   6600
      TabIndex        =   2
      Tag             =   "¿Es instalación?|N|N|||sfamia|instalac||N|"
      Top             =   600
      Width           =   1695
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   12120
      TabIndex        =   23
      Top             =   7800
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   1
      Left            =   3000
      MaxLength       =   25
      TabIndex        =   1
      Tag             =   "Denominación familia de Artículo|T|N|||sfamia|nomfamia||N|"
      Text            =   "Text1"
      Top             =   600
      Width           =   3285
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   0
      Left            =   600
      MaxLength       =   5
      TabIndex        =   0
      Tag             =   "Código familia de artículo|N|N|0|60000|sfamia|codfamia|0000|S|"
      Text            =   "Text"
      Top             =   600
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Left            =   120
      TabIndex        =   24
      Top             =   7680
      Width           =   2655
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   2355
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   12120
      TabIndex        =   22
      Top             =   7800
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   10800
      TabIndex        =   21
      Top             =   7800
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   375
      Left            =   2760
      Top             =   7800
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
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
      TabIndex        =   28
      Top             =   0
      Width           =   13470
      _ExtentX        =   23760
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
            Object.ToolTipText     =   "Lineas descuento"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Actualizar descuento/familia-marca"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
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
         Left            =   5880
         TabIndex        =   29
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Descripcion"
      Height          =   195
      Index           =   14
      Left            =   8520
      TabIndex        =   62
      Top             =   1080
      Width           =   840
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Denominación"
      Height          =   255
      Index           =   1
      Left            =   1800
      TabIndex        =   27
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Cod."
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   26
      Top             =   600
      Width           =   375
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu mnBuscar 
         Caption         =   "&Buscar"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnVerTodos 
         Caption         =   "&Ver todos"
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
      Begin VB.Menu mnbarra3 
         Caption         =   "-"
      End
      Begin VB.Menu mnLineas 
         Caption         =   "Mantenimiento lineas"
         Shortcut        =   ^L
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
Attribute VB_Name = "frmAlmFamiliaArticulo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

Private WithEvents frmB As frmBuscaGrid 'Form para busquedas
Attribute frmB.VB_VarHelpID = -1

'-------------------------------------------------------------------------
'-------------------------------------------------------------------------
'  Variables comunes a todos los formularios
Private Modo As Byte
'-----------------------------
'Se distinguen varios modos
'   0.-  Formulario limpio sin nungun campo rellenado
'   1.-  Preparando para hacer la busqueda
'   2.-  Ya tenemos registros y los vamos a recorrer
'        y podemos editarlos Edicion del campo
'   3.-  Insercion de nuevo registro
'   4.-  Modificar
Dim ModificaLineas As Byte
'1.- Añadir,  2.- Modificar,  3.- Borrar,  0.-Pasar control a Lineas

Private CadenaConsulta As String
Private Ordenacion As String
Private NombreTabla As String  'Nombre de la tabla o de la
Private kCampo As Integer

Dim btnPrimero As Byte
'Variable que indica el número del Boton  PrimerRegistro en la Toolbar1

'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean


Private Sub chkBloqTPV_Click()
    ConseguirfocoChk Modo
End Sub

Private Sub chkBloqTPV_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkComunica_Click()
    ConseguirfocoChk Modo
End Sub

Private Sub chkComunica_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkInstalac_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub chkInstalac_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub chkMarcaPropia_Click()
    ConseguirfocoChk Modo
End Sub

Private Sub chkMarcaPropia_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub cmdAceptar_Click()
Dim numlinea As Integer

    Screen.MousePointer = vbHourglass
    On Error GoTo Error1
    Select Case Modo
        Case 1  'BUSQUEDA
            HacerBusqueda
            
        Case 3 'INSERTAR
            If DatosOk Then
                If InsertarDesdeForm(Me) Then
                    TratarCtaContable
                    PosicionarData
                End If
            End If
        
        Case 4  'MODIFICAR
            If DatosOk Then
                If ModificaDesdeFormulario(Me, 1) Then
                    TerminaBloquear
                    PosicionarData
                End If
            End If
            
        Case 5
        
            'Actualizar el registro en la tabla de lineas 'slialb'
            If ModificaLineas = 1 Then 'INSERTAR lineas Albaran
                numlinea = 1
                If Data2.Recordset.EOF Then numlinea = 0
                If TratarLinea(True) Then
                    If numlinea = 0 Then
                        CargaGrid DataGrid1, Data2, True
                    Else
                        CargaGrid2 DataGrid1, Data2
                    End If
                    BotonAnyadirLinea
                End If
                
            Else
                If TratarLinea(False) Then
                    TerminaBloquear
                    NumRegElim = Val(Data2.Recordset!clasifica)
                    CargaTxtAux False, False
                    CargaGrid2 DataGrid1, Data2
                    PosicionarData2
                    ModificaLineas = 0
                    PonerBotonCabecera True
                    Me.DataGrid1.Enabled = True
                End If
            End If
        
    End Select
        
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub cmdCancelar_Click()
Dim Hay As Boolean
    Select Case Modo
        Case 1 'Busqueda
            LimpiarCampos
            PonerModo 0
        
        Case 3 'Insertar
            Hay = False
            If Not Data1.Recordset Is Nothing Then
                If Not Data1.Recordset.EOF Then Hay = True
            End If
            If Not Hay Then
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
            
        Case 5
            TerminaBloquear
            CargaTxtAux False, False
            If ModificaLineas = 1 Then 'INSERTAR
                DataGrid1.AllowAddNew = False
                ModificaLineas = 0  'Fuerzo el cero para que carge la ampliacion
                If Not Data2.Recordset.EOF Then Data2.Recordset.MoveFirst
                DataGrid1.Enabled = True
            End If
            ModificaLineas = 0
            PonerBotonCabecera True
            Me.DataGrid1.Enabled = True
        
    End Select
    PonerFoco Text1(0)
End Sub


Private Sub BotonAnyadir()
    LimpiarCampos
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 3
    Text1(0).Text = SugerirCodigoSiguienteStr("sfamia", "codfamia")
    FormateaCampo Text1(0)
    PonerFoco Text1(0)
End Sub


Private Sub BotonBuscar()
    If Modo <> 1 Then
        LimpiarCampos
        PonerModo 1
        'Si pasamos el control aqui lo ponemos en amarillo
        PonerFoco Text1(0)
        Text1(0).BackColor = vbYellow
    Else 'Modo=1 Busqueda
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
        CadenaConsulta = "Select * from " & NombreTabla
        PonerCadenaBusqueda
    End If
End Sub


Private Sub Desplazamiento(Index As Integer)
'Para desplazarse por los registros de control Data
    DesplazamientoData Data1, Index
    PonerCampos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
End Sub


Private Sub BotonModificar()
    
    'El registro de codigo 0 no se puede Modificar ni Eliminar
    'If EsCodigoCero(CStr(Data1.Recordset.Fields(0).Value), FormatoCampo(Text1(0))) Then Exit Sub
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    SSTab1.Tab = 0
    PonerModo 4
    PonerFoco Text1(1)
End Sub


Private Sub BotonEliminar()
Dim cad As String

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
    'El registro de codigo 0 no se puede Modificar ni Eliminar
    'If EsCodigoCero(CStr(Data1.Recordset.Fields(0).Value), FormatoCampo(Text1(0))) Then Exit Sub
    cad = DevuelveDesdeBD(conAri, "count(*)", "sartic", "codfamia", Text1(0).Text)
    If cad = "" Then cad = "0"
    If Val(cad) > 0 Then
        cad = "Hay " & cad & " artículos pertenecientes a esta familia"
        MsgBox cad, vbExclamation
        Exit Sub
    End If
    
    '### a mano
    cad = "¿Seguro que desea eliminar la Familia de Artículo?:" & vbCrLf
    cad = cad & vbCrLf & "Cod. : " & Format(Data1.Recordset.Fields(0), FormatoCampo(Text1(0)))
    cad = cad & vbCrLf & "Desc.: " & Data1.Recordset.Fields(1)
    
    



    'Borramos
    If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        On Error GoTo Error2
        Screen.MousePointer = vbHourglass
        NumRegElim = Data1.Recordset.AbsolutePosition
        'Borramos en dtofamia
        lblIndicador.Caption = "Eliminando descuentos"
        lblIndicador.Refresh
        cad = "Delete from sdtofm where codfamia = " & Data1.Recordset!Codfamia
        conn.Execute cad
        
        
        cad = "Delete from sdtomp where codfamia = " & Data1.Recordset!Codfamia
        conn.Execute cad
        
        
        cad = "Delete from sfamiadtos where codfamia = " & Data1.Recordset!Codfamia 'despues del DELETE
        conn.Execute cad
        
        lblIndicador.Caption = "Eliminando familia"
        lblIndicador.Refresh
        Data1.Recordset.Delete
        
        ejecutar cad, False
        
        If SituarDataTrasEliminar(Data1, NumRegElim) Then
            PonerCampos
        Else
            LimpiarCampos
            PonerModo 0
        End If
    End If
    
Error2:
    
    If Err.Number <> 0 Then
        lblIndicador.Caption = ""
        Data1.Recordset.CancelUpdate
        MuestraError Err.Number, "Eliminar Familia de Articulo", Err.Description
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub cmdRegresar_Click()
Dim cad As String

    If Modo = 5 Then  'modo 5: Mantenimientos Lineas
    
    
        PonerModo 2
        If Not Data1.Recordset.EOF Then Me.lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
    Else
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

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
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
    If vParamAplic.Descriptores Then Me.Caption = "Categorias Art."
    ' ICONITOS DE LA BARRA
    btnPrimero = 15 'Posicion del Boton Primero en la toolbar (+ 3 siguientes)
    With Me.Toolbar1
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 1   'Botón Buscar
        .Buttons(2).Image = 2   'Botón Todos
        
        .Buttons(5).Image = 3   'Insertar Nuevo
        .Buttons(6).Image = 4   'Modificar
        .Buttons(7).Image = 5   'Borrar
        
        .Buttons(9).Image = 10
        .Buttons(10).Image = 16  ' Imprimir
        .Buttons(11).Image = 42  ' actualizar dto/familia
        .Buttons(13).Image = 15  'Salir
        
        .Buttons(btnPrimero).Image = 6  'Primero
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
        .Buttons(btnPrimero + 3).Image = 9 'Último
    End With
    
    
    Label1(4).visible = vEmpresa.TieneAnalitica
    Me.Text1(8).visible = vEmpresa.TieneAnalitica
    Me.Text2(8).visible = vEmpresa.TieneAnalitica
    imgBuscar(0).visible = vEmpresa.TieneAnalitica
    
    LimpiarCampos   'Limpia los campos TextBox
    
    'Pone el Tag del primer botón de busqueda de cuentas a -1
    'Si tag =-1 abre busqueda en la tabla: sfamia, BD: Ariges
    'Si tag>0 abre busqueda en la tabla: Cuentas, BD: Conta
    imgCuentas(0).Tag = "-1"
    Me.imgBuscar(0).Tag = "-1"
        
  
    '## A mano
    NombreTabla = "sfamia"
    Ordenacion = " ORDER BY codfamia"
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
       
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = conn
    Data1.RecordSource = "Select * from " & NombreTabla & " WHERE codfamia=-1"
    Data1.Refresh
    
    If DatosADevolverBusqueda = "" Then
        PonerModo 0
    Else
        BotonBuscar
    End If
    CargaGrid DataGrid1, Data2, False
    CargaCombo  'de  clasificaciones
    SSTab1.Tab = 0
End Sub


Private Sub LimpiarCampos()
    limpiar Me   'Metodo general: Limpia los controles TextBox
    'Aqui va el especifico de cada form es
    '### a mano
    Me.chkInstalac.Value = 0
    chkBloqTPV.Value = 0
    chkComunica.Value = 0
    chkMarcaPropia.Value = 0
End Sub


Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
End Sub


Private Sub frmB_Selecionado(CadenaDevuelta As String)
Dim cadB As String
Dim Aux As String
Dim indice As Byte
    
    If CadenaDevuelta <> "" Then
        If Val(imgCuentas(0).Tag) >= 0 Then
            'Se llama desde un botón de busqueda de Cuentas
            'Recuperar solo el campo código y Descripción
            HaDevueltoDatos = True
            Screen.MousePointer = vbHourglass
    
            indice = Val(Me.imgCuentas(0).Tag)
            If indice < 6 Then
                indice = indice + 2
            Else
                indice = indice + 5
            End If
            Text1(indice).Text = RecuperaValor(CadenaDevuelta, 1)
            Text2(indice).Text = RecuperaValor(CadenaDevuelta, 2)
        ElseIf Val(imgBuscar(0).Tag) >= 0 Then
            indice = 8 + Val(imgBuscar(0).Tag)
            '0.- Centro de coste   1.- Proveedor
            Text1(indice).Text = RecuperaValor(CadenaDevuelta, 1)
            Text2(indice).Text = RecuperaValor(CadenaDevuelta, 2)
        Else
            'Recupera todo el registro de Banco Propio
            HaDevueltoDatos = True
            Screen.MousePointer = vbHourglass
            'Sabemos que campos son los que nos devuelve
            'Creamos una cadena consulta y ponemos los datos
            cadB = ""
            Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
            cadB = Aux
            '   Como la clave principal es unica, con poner el sql apuntando
            '   al valor devuelto sobre la clave ppal es suficiente
            'Aux = ValorDevueltoFormGrid(Text1(1), CadenaDevuelta, 2)
            'If CadB <> "" Then CadB = CadB & " AND "
            'CadB = CadB & Aux
            'Se muestran en el mismo form
            CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
            PonerCadenaBusqueda
            Screen.MousePointer = vbDefault
        End If
    End If
End Sub


Private Sub imgBuscar_Click(Index As Integer)
    ' ---- [19/10/2009] [LAURA]: añadir campo centro de coste familia
    
    If Modo = 2 Or Modo = 5 Or Modo = 0 Then Exit Sub
    
    Select Case Index
        Case 0 'Centros de coste de la conta
            Screen.MousePointer = vbHourglass
            Me.imgBuscar(0).Tag = Index
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
            imgBuscar(0).Tag = -1
            Screen.MousePointer = vbDefault
            PonerFoco Text1(8)
            
        Case 1
            Screen.MousePointer = vbHourglass
            Me.imgBuscar(0).Tag = Index
            Set frmB = New frmBuscaGrid
            frmB.vCampos = "Codigo|sprove|codprove|N||20·Nombre|sprove|nomprove|T||70·"
            frmB.vTabla = "sprove"
            frmB.vSQL = ""
            HaDevueltoDatos = False
            '###A mano
            frmB.vDevuelve = "0|1|"
            frmB.vTitulo = "Proveedores"
            frmB.vselElem = 0
            frmB.vConexionGrid = conAri
            
            frmB.Show vbModal
            Set frmB = Nothing
            imgBuscar(0).Tag = -1
            Screen.MousePointer = vbDefault
            PonerFoco Text1(9)
    End Select
End Sub


Private Sub imgCuentas_Click(Index As Integer)

    If Modo = 2 Or Modo = 5 Or Modo = 0 Then Exit Sub
 
    Screen.MousePointer = vbHourglass
    imgCuentas(0).Tag = Index
    MandaBusquedaPrevia "apudirec='S'"
    imgCuentas(0).Tag = -1
    PonerFoco Text1(Index + 2)
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    If Modo = 5 Then 'Eliminar lineas de trabajadores
         BotonEliminarLinea
    Else   'Eliminar Trabajador
         BotonEliminar
    End If
End Sub

Private Sub mnLineas_Click()
    BotonMtoLineas
End Sub

Private Sub mnModificar_Click()
    'If BLOQUEADesdeFormulario(Me) Then BotonModificar
    If Modo = 5 Then 'Modificar lineas
         BotonModificarLinea
    Else   'Modificar
         If BLOQUEADesdeFormulario(Me) Then BotonModificar
    End If
    
End Sub

Private Sub mnNuevo_Click()
    If Modo = 5 Then 'Añadir lineas
         BotonAnyadirLinea
    Else 'Añadir Cabecera de Ofertas
         Me.SSTab1.Tab = 0
         BotonAnyadir
    End If
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
    If Index = 14 Then Exit Sub
    ConseguirFoco Text1(Index), Modo
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 14 Then Exit Sub
    If Not (Index = 0 And KeyCode = 38) Then KEYdown KeyCode
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 14 Then Exit Sub
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
        
    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    'Si queremos hacer algo ..
    Select Case Index
        Case 0 'Codigo familia
'            If Text1(Index).Text <> "" Then
             If PonerFormatoEntero(Text1(Index)) Then
                'Comprobar si ya existe el cod de familia en la tabla
                If Modo = 3 Then 'Insertar
                    If ExisteCP(Text1(Index)) Then PonerFoco Text1(Index)
                End If
            End If
        '#### lo hemos puesto en el evento VALIDATE
'         Case 2, 3, 4, 5 'Cuentas
'            Text2(Index).Text = PonerNombreCuenta(Text1(Index), Modo)
        '####
        '---> Por algun motivo habian comentado ese trozo
        Case 2, 3, 4, 5, 6, 7, 11, 12, 13 'Cuentas
            Text2(Index).Text = PonerNombreCuenta(Text1(Index), Modo, Text1(1).Text)
            If Text2(Index).Text = "" Then Text1(Index).Text = ""
            
            If Modo = 3 Then
              If Index = 2 And Text1(11).Text = "" Then
                    Text1(11).Text = Text1(Index).Text
                    Text2(11).Text = Text2(Index).Text
               ElseIf Index = 6 And Text1(12).Text = "" Then
                    Text1(12).Text = Text1(Index).Text
                    Text2(12).Text = Text2(Index).Text
                End If
            End If
        ' ---- [19/10/2009] [LAURA]: añadir campo centro de coste familia
        Case 8: Me.Text2(Index).Text = PonerNombreCCoste(Me.Text1(Index))
        
        Case 9
            If Not PonerFormatoEntero(Text1(Index)) Then
                If Text1(Index).Text <> "" Then Text1(Index).Text = ""
                Text2(Index).Text = ""
            Else
                
                'Text2(Index).Text = DevuelveDesdeBD(conAri, "nomprove", "sprove", "codprove", Text1(Index).Text, "N")
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "sprove", "nomprove")
                If Text2(Index).Text = "" Then
                    Text1(Index).Text = ""
                    PonerFoco Text1(Index)
                End If
            End If
        Case 10
            If Not PonerFormatoDecimal(Text1(10), 4) Then Text1(10).Text = ""
    End Select
End Sub


Private Sub HacerBusqueda()
Dim cadB As String
    
    cadB = ObtenerBusqueda(Me, False)

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
Dim CargaF As Boolean 'Para saber si se carga el frame o no en el BuscaGrid
Dim Conexion As Byte

        'Llamamos a al form
        '##A mano
        cad = ""
        If Val(Me.imgCuentas(0).Tag) >= 0 Then
        'Se llama a Busqueda desde un campo de Cuenta
            '#A MANO: Porque busca en la tabla Cuentas
            'de la base de datos de Contabilidad
            cad = cad & "Código|Cuentas|codmacta|T||15·Denominacion|Cuentas|nommacta|T||70·"
            tabla = "Cuentas"
            Titulo = "Cuentas"
            Conexion = conConta    'Conexión a BD: Conta
            CargaF = True 'Se puede cargar el frame
        Else
            'Busqueda de una Família de Artículo
            cad = cad & ParaGrid(Text1(0), 15, "Código")
            cad = cad & ParaGrid(Text1(1), 80, "Denominacion")
            tabla = "sfamia"
            Titulo = "Família de Artículos"
            If vParamAplic.Descriptores Then Titulo = "Categorias Art."
            Conexion = conAri    'Conexión a BD: Ariges
            CargaF = False 'No se carga el frame
        End If
        
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
            frmB.vCargaFrame = CargaF
            '#
            frmB.Show vbModal
            Set frmB = Nothing
            'Si ha puesto valores y tenemos que es formulario de busqueda entonces
            'tendremos que cerrar el form lanzando el evento
            If HaDevueltoDatos Then
                If kCampo < 5 Then PonerFoco Text1(kCampo + 1)
                If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then
                    If Not (Val(Me.imgCuentas(0).Tag) >= 0) Then cmdRegresar_Click
                End If
            Else   'de ha devuelto datos, es decir NO ha devuelto datos
                PonerFoco Text1(kCampo)
'                If Modo = 1 Then
'                    MsgBox "No hay ningún registro en la tabla " & tabla
'                    PonerFoco Text1(0)
'                End If
            End If
        End If
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
    End If

    Screen.MousePointer = vbDefault
    Exit Sub
    
EEPonerBusq:
    MuestraError Err.Number, "PonerCadenaBusqueda"
    PonerModo 0
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerCampos()
Dim i As Byte
    
    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma Me, Data1
    
    'poner la descripcion de las cuentas
    For i = 2 To 7
        Text2(i).Text = PonerNombreCuenta(Text1(i), Modo)
    Next i
    For i = 11 To 13
        
        Text2(i).Text = PonerNombreCuenta(Text1(i), Modo)
    Next i
    ' ---- [19/10/2009] [LAURA]: añadir campo centro de coste familia
    If vEmpresa.TieneAnalitica Then Me.Text2(8).Text = PonerNombreCCoste(Me.Text1(8))
        
    
    Text2(9).Text = PonerNombreDeCod(Text1(9), conAri, "sprove", "nomprove")
    
    BloquearChecks Me, Modo
    PonerCamposLineas
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
End Sub


'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
Private Sub PonerModo(Kmodo As Byte)
Dim b As Boolean
Dim NumReg As Byte

 
    'Actualiza Iconos Insertar,Modificar,Eliminar
    ActualizarToolbarGnral Me.Toolbar1, Modo, Kmodo, 5
    
    Modo = Kmodo
    PonerIndicador lblIndicador, Modo
    If Modo = 5 Then lblIndicador.Caption = "Lineas dto"
    
    
    Frame4.Enabled = Modo < 5  'En lineas no dejo trabajar
    b = Modo < 5
    If Not b Then ModificaLineas = 0
    'datagrid1.enabled
    
    '-------------------------------------------------
    'Modo 2. Hay datos y estamos visualizandolos
    b = (Kmodo = 2)
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.visible = b
    Else
        'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
        cmdRegresar.visible = Modo = 5 And ModificaLineas = 0
    End If
    
    'Visualizar flechas de desplazamiento en la toolbar si modo=2
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, NumReg
    
    'Poner Boton de Cabecera o Aceptar/Cancelar
    PonerBotonCabecera b Or (Modo = 0)
        
    'Bloquear Registros si modo distinto de Insertar o Modificar
    BloquearText1 Me, Modo
    
    
    
    BloquearChecks Me, Modo
        
    Me.chkBloqTPV.Enabled = Modo = 1 Or Modo = 3 Or Modo = 4
    chkComunica.Enabled = Me.chkBloqTPV.Enabled
    chkMarcaPropia.Enabled = Me.chkBloqTPV.Enabled
    
    chkVistaPrevia.Enabled = (Modo <= 2)
    
    'Poner el tamaño de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos
    
    PonerModoOpcionesMenu 'Activar opciones de menu según MODO
    PonerOpcionesMenu   'Activar opciones de menu según NIVEL
                        'de permisos del usuario
End Sub


Private Sub PonerModoOpcionesMenu()
Dim b As Boolean
On Error Resume Next

    b = Modo < 3
    If Not b Then
        If Modo = 5 And ModificaLineas = 0 Then b = True
    End If
    
    'Añadir
    Toolbar1.Buttons(5).Enabled = b
    Me.mnNuevo.Enabled = b
    
    b = Modo = 2
    Toolbar1.Buttons(11).Enabled = b
    If Not b Then
        If Modo = 5 And ModificaLineas = 0 Then b = True
    End If
    
    
    
    'Modificar
    Toolbar1.Buttons(6).Enabled = b
    mnModificar.Enabled = b
    'eliminar
    Toolbar1.Buttons(7).Enabled = b
    mnEliminar.Enabled = b
    
    
    Toolbar1.Buttons(9).Enabled = Modo = 2
    Me.mnLineas.Enabled = Modo = 2
    
     '---------------------------------
    b = (Modo >= 3)
    'Buscar
    Toolbar1.Buttons(1).Enabled = Not b
    Me.mnBuscar.Enabled = Not b
    'VerTodos
    Toolbar1.Buttons(2).Enabled = Not b
    Me.mnVerTodos.Enabled = Not b
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
    PonerLongCamposGnral Me, Modo, 1
End Sub


Private Function DatosOk() As Boolean
Dim b As Boolean

    DatosOk = False
    b = CompForm(Me, 1)
    If Not b Then Exit Function
    
    'Comprobar si ya existe el cod de familia en la tabla
    If Modo = 3 Then 'Insertar
        If ExisteCP(Text1(0)) Then b = False
    End If
    
    DatosOk = b
End Function






Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1: mnBuscar_Click
        Case 2: mnVerTodos_Click
        Case 5  'Nuevo
                mnNuevo_Click
        Case 6  'Modificar
                mnModificar_Click
        Case 7  'Borrar
                mnEliminar_Click
                
        Case 9
                BotonMtoLineas
            
        Case 10 'Imprimir listado
            BotonImprimir
            
        Case 11
            'Actualizar dto/familia-marca
            ActualizarDtoFamilia
            
        Case 13: mnSalir_Click
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


Private Sub PonerBotonCabecera(b As Boolean)

    Me.cmdAceptar.visible = Not b
    Me.cmdCancelar.visible = Not b
    If b Then
        PonerFocoBtn Me.cmdRegresar
        cmdRegresar.Cancel = True
    Else
        cmdCancelar.Cancel = True
    End If
    
    
    Me.cmdRegresar.Caption = "Cabecera"
    If b Then
        b = False
        If Modo = 2 Then
            b = Me.DatosADevolverBusqueda <> ""
            If b Then Me.cmdRegresar.Caption = "Regresar"
        ElseIf Modo = 5 Then
            b = ModificaLineas = 0
        End If
        
    End If
    Me.cmdRegresar.visible = b
   
    
    'Habilitar las opciones correctas del menu
    PonerModoOpcionesMenu
    PonerOpcionesMenu
    If Err.Number <> 0 Then Err.Clear

    
End Sub


Private Sub PosicionarData()
Dim cad As String, Indicador As String

    cad = "(codfamia=" & Text1(0).Text & ")"
    If SituarData(Data1, cad, Indicador) Then
        PonerModo 2
        lblIndicador.Caption = Indicador
    Else
        PonerModo 0
    End If
End Sub


Private Sub PosicionarData2()
    On Error GoTo EPosicionarData2
    
    Data2.Recordset.Find "clasifica = " & NumRegElim
    If Data2.Recordset.EOF Then Data2.Recordset.MoveFirst
    NumRegElim = 0
    Exit Sub
EPosicionarData2:
    MuestraError Err.Number
End Sub

Private Sub BotonImprimir()
    frmListado3.Opcion = 12
    frmListado3.Show vbModal
End Sub


Private Sub TratarCtaContable()
Dim i As Integer
Dim CtaCreadas As String
    For i = 2 To 7
        If Text2(i).Text = vbCrearNuevaCta Then
            If InStr(1, CtaCreadas, Text1(i).Text & "|") = 0 Then
                InsertarCuentaCble Text1(i).Text, "", "", Text1(1).Text
                CtaCreadas = CtaCreadas & Text1(i).Text & "|"
            End If
            Text2(i).Text = Text1(1).Text
        End If
    Next i
End Sub

Private Sub BotonMtoLineas()
       If Data1.Recordset Is Nothing Then Exit Sub
       If Data1.Recordset.EOF Then Exit Sub
       
        Me.SSTab1.Tab = 1
        ModificaLineas = 0
        PonerModo 5
        PonerBotonCabecera True
End Sub





Private Sub BotonModificarLinea()
'Modificar una linea
Dim vWhere As String
    
    On Error GoTo EModificarLinea


    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Then Exit Sub '1= Insertar
    Me.SSTab1.Tab = 1

    If Data2.Recordset.EOF Then Exit Sub
    vWhere = "codfamia = " & Text1(9).Text & " and clasifica=" & Data2.Recordset!clasifica
    If Not BloqueaRegistro("sfamiadtos", vWhere) Then Exit Sub
    CargaTxtAux True, False
    ModificaLineas = 2 'Modificar

    'Añadiremos el boton de aceptar y demas objetos para insertar
    Me.lblIndicador.Caption = "MODIFICAR"
    PonerBotonCabecera False
    PonerFoco txtAux(0)
    Me.DataGrid1.Enabled = False

EModificarLinea:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub BotonAnyadirLinea()


    'Si no estaba modificando lineas salimos
    ' Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 2 Then Exit Sub

    ModificaLineas = 1 'Ponemos Modo Añadir Linea

    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerBotonCabecera False
    Me.SSTab1.Tab = 1
  
    lblIndicador.Caption = "INSERTAR"

    AnyadirLinea DataGrid1, Data2
    CargaTxtAux True, True

    Combo1.ListIndex = -1
    PonerFocoCbo Combo1
    Me.DataGrid1.Enabled = False
End Sub


Private Sub BotonEliminarLinea()
'Eliminar una linea De Mantenimiento. Tabla: slima1
Dim SQL As String
    
    On Error GoTo EEliminarLinea

    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas > 0 Then Exit Sub '1= Insertar, 2=Modificar

    If Data2.Recordset.EOF Then Exit Sub

    Me.SSTab1.Tab = 1
    

    SQL = "¿Seguro que desea eliminar la línea de descuento?     "
    SQL = SQL & vbCrLf & "Clasificacion:  " & Data2.Recordset!Nombre & vbCrLf
    SQL = SQL & "Descuento1:  " & Format(Data2.Recordset!dtoline1, FormatoDescuento)
    


    If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        NumRegElim = Data2.Recordset.AbsolutePosition
        SQL = " codfamia = " & Text1(0).Text & " AND clasifica=" & Data2.Recordset!clasifica
        SQL = "Delete from sfamiadtos WHERE " & SQL
        conn.Execute SQL

        ModificaLineas = 0
        CargaGrid2 DataGrid1, Data2
        SituarDataTrasEliminar Data2, NumRegElim

    End If
    PonerFocoBtn Me.cmdRegresar

EEliminarLinea:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Lineas descuentos", Err.Description
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
    
    SQL = "select sfamiadtos.clasifica,nombre,dtoline1,dtoline2 from sfamiadtos,sfamiatipodto where"
    SQL = SQL & " sfamiadtos.clasifica=sfamiatipodto.clasifica and codfamia="
    If enlaza Then
        SQL = SQL & Data1.Recordset!Codfamia
       
    Else
        SQL = SQL & "  -1"
    End If
    SQL = SQL & " Order by sfamiadtos.clasifica"
    MontaSQLCarga = SQL
End Function


Private Sub PonerCamposLineas()
'Carga las Pestañas con las tablas de lineas del Trabajador seleccionado para mostrar

    On Error GoTo EPonerLineas

    Screen.MousePointer = vbHourglass

    'Datos de la tabla slipre
    CargaGrid DataGrid1, Data2, True

    Screen.MousePointer = vbDefault
    Exit Sub
    
EPonerLineas:
    MuestraError Err.Number, "PonerCamposLineas"
    PonerModo 2
    Screen.MousePointer = vbDefault
End Sub


Private Sub CargaGrid(ByRef vDataGrid As DataGrid, ByRef vData As Adodc, enlaza As Boolean)
Dim b As Boolean
Dim SQL As String

    On Error GoTo ECargaGrid

    b = DataGrid1.Enabled
    
    SQL = MontaSQLCarga(enlaza)
    CargaGridGnral vDataGrid, vData, SQL, True

    
    CargaGrid2 vDataGrid, vData
    vDataGrid.ScrollBars = dbgAutomatic
        
    b = (Modo = 5) And (ModificaLineas = 1 Or ModificaLineas = 2)
    vDataGrid.Enabled = Not b

    'PrimeraVez = False
    Exit Sub
    
ECargaGrid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub


Private Sub CargaGrid2(ByRef vDataGrid As DataGrid, ByRef vData As Adodc)
Dim i As Integer

    On Error GoTo ECargaGrid

    vData.Refresh

                vDataGrid.Columns(0).Caption = "Código"
                vDataGrid.Columns(0).visible = False
                
                vDataGrid.Columns(1).Caption = "Descripción"
                vDataGrid.Columns(1).Width = 3200
 
                vDataGrid.Columns(2).Caption = "Dto. 1"
                vDataGrid.Columns(2).Width = 900
                vDataGrid.Columns(2).Alignment = dbgRight
                vDataGrid.Columns(2).NumberFormat = FormatoDescuento
                
                vDataGrid.Columns(3).Caption = "Dto. 2"
                vDataGrid.Columns(3).Width = 900
                vDataGrid.Columns(3).Alignment = dbgRight
                vDataGrid.Columns(3).NumberFormat = FormatoDescuento
                
                


    For i = 0 To vDataGrid.Columns.Count - 1
        vDataGrid.Columns(i).Locked = True
        vDataGrid.Columns(i).AllowSizing = False
    Next i
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
Dim i As Byte

    If Not visible Then
        'Fijamos el alto (ponerlo en la parte inferior del form)
        For i = 0 To txtAux.Count - 1 'TextBox
            txtAux(i).Top = 290
            txtAux(i).visible = visible
        Next i
        Combo1.visible = False
    Else
        If limpiar Then 'Vaciar los textBox (Vamos a Insertar)
            DeseleccionaGrid DataGrid1
            For i = 0 To txtAux.Count - 1
                txtAux(i).Text = ""
              '  BloquearTxt txtAux(i), False  'Todos menos el nombre
            Next i

        Else 'Vamos a modificar
            For i = 0 To txtAux.Count - 1
         
                txtAux(i).Text = DataGrid1.Columns(i + 2).Text
            
                txtAux(i).Locked = False
            Next i
        End If
        


        'Fijamos altura(Height) y posición Top
        '-------------------------------
        alto = ObtenerAlto(DataGrid1, 10)
        
        For i = 0 To txtAux.Count - 1
            txtAux(i).Top = alto
            txtAux(i).Height = DataGrid1.RowHeight
        Next i
        Combo1.Top = alto
        'Fijamos anchura y posicion Left
        '--------------------------------
        'Cod. Almac

        'Precio, Dto1, Dto2, Precio
        Combo1.Left = DataGrid1.Left + 340
        Combo1.Width = DataGrid1.Columns(2).Left - DataGrid1.Left - 240
        For i = 0 To txtAux.Count - 1
            txtAux(i).Left = DataGrid1.Columns(i + 2).Left + 10 + 120
            txtAux(i).Width = DataGrid1.Columns(i + 2).Width - 10
        Next i
        

        Combo1.visible = limpiar
        'Los ponemos Visibles o No
        '--------------------------
        For i = 0 To txtAux.Count - 1
            txtAux(i).visible = visible
        Next i
        
    End If

End Sub


Private Sub TxtAux_Change(Index As Integer)
    'Precio y Modo Borrar Lineas
    If Index = 4 And ModificaLineas = 2 Then txtAux(5).Text = "M"
End Sub

Private Sub txtAux_GotFocus(Index As Integer)
Dim cadkey As Integer

    cadkey = ObtenerCadKey(kCampo, Index)
    kCampo = Index
'    ConseguirFoco txtAux(Index), Modo, cadkey
    ConseguirFocoLin txtAux(Index), cadkey
    

    
End Sub

Private Sub TxtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    If Not (Index = 0 And KeyCode = 38) Then KEYdown KeyCode
    

End Sub


Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub CargaCombo()

    Combo1.Clear

    Set miRsAux = New ADODB.Recordset
    miRsAux.Open "Select clasifica elcodigo,nombre elNombre from sfamiatipodto ORDER BY clasifica", conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Combo1.AddItem miRsAux!ElNombre
        Combo1.ItemData(Combo1.NewIndex) = miRsAux!ElCodigo
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
End Sub



Private Function TratarLinea(Insertar As Boolean) As Boolean
'Modifica o Inserta un registro de linea
Dim SQL As String

    On Error GoTo EModificarLinea

    TratarLinea = False
    SQL = ""
    
    'Meto el DATOS OK aqui
    '-------------------------
    If Insertar Then
        If Combo1.ListIndex < 0 Then
            SQL = "- Seleccione una clasificacion" & vbCrLf
        Else
            CadenaConsulta = "codfamia = " & Text1(0).Text & " AND clasifica "
            CadenaConsulta = DevuelveDesdeBD(conAri, "clasifica", "sfamiadtos ", CadenaConsulta, Combo1.ItemData(Combo1.ListIndex), "N")
            If CadenaConsulta <> "" Then SQL = SQL & "- Ya existe para la clasificacion: " & Combo1.Text & vbCrLf
        End If
    End If
    
    'Sea insertar o modificar. Hay cosas que son para los dos
    If txtAux(0).Text = "" Then
        SQL = SQL & "- Descuento 1 obligado" & vbCrLf
    Else
        'NO PUEDE EXISTIR EL DESCUENTO YA
        
        CadenaConsulta = ""
        If Not Insertar Then CadenaConsulta = "clasifica <> " & Data2.Recordset.Fields!clasifica & " AND "
        CadenaConsulta = CadenaConsulta & "dtoline1 = " & DBSet(txtAux(0).Text, "N") & " AND codfamia"
        
        CadenaConsulta = DevuelveDesdeBD(conAri, "clasifica", "sfamiadtos ", CadenaConsulta, Text1(0).Text, "N")
        If CadenaConsulta <> "" Then SQL = SQL & "- Ya existe el descuento en la clasificacion: " & CadenaConsulta & vbCrLf
        
        
        
    End If
    CadenaConsulta = Data1.RecordSource
    
    If SQL <> "" Then
        SQL = "Errores: " & vbCrLf & vbCrLf & SQL
        MsgBox SQL, vbExclamation
        Exit Function
    End If
    
    'Si tieen proveedor asignado
    'Si ha llegado aqui comprobaremos que el dto no supera al maximo que tiene el proveedor
    If Text1(9).Text Then
            SQL = txtAux(0).Text
            If txtAux(1).Text <> "" Then
                If ImporteFormateado(txtAux(1).Text) > ImporteFormateado(SQL) Then SQL = txtAux(1).Text
            End If
            
            CadenaConsulta = "codfamia = " & Text1(0).Text & " AND codmarca is NULL AND codprove "
            
            CadenaConsulta = DevuelveDesdeBD(conAri, "dtoline1+dtoline2", "sdtomp", CadenaConsulta, Text1(9).Text, "N")
            If CadenaConsulta <> "" Then
                If ImporteFormateado(SQL) > CCur(CadenaConsulta) Then
                    SQL = "Descuento mayor al asignado por el proveedor a esta familia." & vbCrLf & "¿Continuar?"
                Else
                    SQL = ""
                End If
            Else
                'No tiene asignado dto
                SQL = ""
            End If
            'Reestablezco
            CadenaConsulta = Data1.RecordSource
            If SQL <> "" Then
                If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then Exit Function
            End If
    End If
            
    'Creamos la sentencia SQL
    If Insertar Then
        SQL = "INSERT INTO sfamiadtos(codfamia,clasifica,dtoline1,dtoline2) VALUES ("
        SQL = SQL & Text1(0).Text & "," & Combo1.ItemData(Combo1.ListIndex) & ","
        SQL = SQL & DBSet(txtAux(0).Text, "N") & "," & DBSet(txtAux(1).Text, "N", "N") & ")"
    Else
        SQL = "UPDATE sfamiadtos Set dtoline1 = " & DBSet(txtAux(0).Text, "N")
        SQL = SQL & ", dtoline2=" & DBSet(txtAux(1).Text, "N", "N")
        SQL = SQL & " WHERE codfamia = " & Text1(0).Text & " AND clasifica=" & Data2.Recordset!clasifica
    End If
    
    
    conn.Execute SQL
    TratarLinea = True
    
        
    Exit Function
    
EModificarLinea:
    MuestraError Err.Number, "Modificar Lineas descuento" & vbCrLf & Err.Description
    CadenaConsulta = ""
End Function

Private Sub txtAux_LostFocus(Index As Integer)
    txtAux(Index).Text = Trim(txtAux(Index).Text)
    If txtAux(Index).Text = "" Then Exit Sub
    If Not PonerFormatoDecimal(txtAux(Index), 4) Then txtAux(Index).Text = ""
    
    
End Sub

Private Sub ActualizarDtoFamilia()
    
    If Data1.Recordset.EOF Then Exit Sub
    
    If vUsu.Nivel > 1 Then
        MsgBox "No tiene permiso", vbExclamation
        Exit Sub
    End If
    
    
    CadenaDesdeOtroForm = Text1(0).Text
    frmVarios.Opcion = 8
    frmVarios.Show vbModal
    
End Sub
