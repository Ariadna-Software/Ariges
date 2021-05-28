VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAlmFamiliaArticuloGr 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Familias de artículos"
   ClientHeight    =   9840
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   14685
   Icon            =   "frmAlmFamiliaArticuloGr.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9840
   ScaleWidth      =   14685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Digitos 1er nivel|N|N|||empresa|numdigi1|||"
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
      Height          =   300
      Left            =   12600
      TabIndex        =   67
      Top             =   315
      Width           =   1605
   End
   Begin VB.Frame FrameDesplazamiento 
      Height          =   705
      Left            =   5085
      TabIndex        =   65
      Top             =   180
      Width           =   2415
      Begin MSComctlLib.Toolbar ToolbarDes 
         Height          =   330
         Left            =   240
         TabIndex        =   66
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
               Object.ToolTipText     =   "ï¿½ltimo"
               ImageIndex      =   9
            EndProperty
         EndProperty
      End
   End
   Begin VB.CheckBox chkInactiva 
      Caption         =   "Inactiva"
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
      Left            =   12600
      TabIndex        =   9
      Tag             =   "Inactiva|N|N|||sfamia|inactiva||N|"
      Top             =   2070
      Width           =   1410
   End
   Begin VB.Frame FrameBotonGnral2 
      Height          =   705
      Left            =   3960
      TabIndex        =   63
      Top             =   180
      Width           =   1020
      Begin MSComctlLib.Toolbar Toolbar5 
         Height          =   330
         Left            =   210
         TabIndex        =   64
         Top             =   180
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Comprobación"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   270
      TabIndex        =   61
      Top             =   180
      Width           =   3585
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   180
         TabIndex        =   62
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
      Left            =   1710
      TabIndex        =   2
      Tag             =   "proveedor|N|S|0||sfamia|codprove|000000||"
      Top             =   1890
      Width           =   945
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
      Index           =   9
      Left            =   2670
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   58
      Text            =   "Text2"
      Top             =   1890
      Width           =   6555
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
      Index           =   8
      Left            =   2655
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   57
      Text            =   "Text2"
      Top             =   2370
      Width           =   6555
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
      Left            =   1710
      MaxLength       =   4
      TabIndex        =   3
      Tag             =   "Centro de coste|T|S|||sfamia|codccost||N|"
      Top             =   2370
      Width           =   945
   End
   Begin VB.CheckBox chkBloqTPV 
      Caption         =   "Bloquea en TPV"
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
      Left            =   9810
      TabIndex        =   6
      Tag             =   "Bloq TPV|N|N|||sfamia|bloqEnTPV||N|"
      Top             =   1575
      Width           =   1905
   End
   Begin VB.CheckBox chkComunica 
      Caption         =   "Comunica"
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
      Left            =   12600
      TabIndex        =   7
      Tag             =   "Comunica|N|N|||sfamia|comunica||N|"
      Top             =   1575
      Width           =   1410
   End
   Begin VB.CheckBox chkMarcaPropia 
      Caption         =   "Marca propia"
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
      Left            =   12600
      TabIndex        =   5
      Tag             =   "Comunica|N|N|||sfamia|marcapropia||N|"
      Top             =   1080
      Width           =   1635
   End
   Begin VB.CheckBox chkPuntos 
      Caption         =   "Permite canje"
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
      Left            =   9810
      TabIndex        =   8
      Tag             =   "Canje|N|N|||sfamia|PtosPermiteCanje||N|"
      Top             =   2070
      Width           =   1770
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
      Height          =   5715
      Index           =   14
      Left            =   9765
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   26
      Tag             =   "Descr|T|S|||sfamia|descripcion||N|"
      Text            =   "frmAlmFamiliaArticuloGr.frx":000C
      Top             =   3195
      Width           =   4605
   End
   Begin MSAdodcLib.Adodc Data2 
      Height          =   375
      Left            =   3150
      Top             =   9135
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   6015
      Left            =   120
      TabIndex        =   32
      Top             =   2880
      Width           =   9390
      _ExtentX        =   16563
      _ExtentY        =   10610
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
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
      TabCaption(0)   =   "Contabilidad"
      TabPicture(0)   =   "frmAlmFamiliaArticuloGr.frx":0012
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Descuentos"
      TabPicture(1)   =   "frmAlmFamiliaArticuloGr.frx":002E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label1(10)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label1(15)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "DataGrid1"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "txtAux(1)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "txtAux(0)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Combo1"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Text1(10)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Text1(15)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "FrameToolAux0"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).ControlCount=   9
      Begin VB.Frame FrameToolAux0 
         Height          =   645
         Left            =   135
         TabIndex        =   68
         Top             =   405
         Width           =   1500
         Begin MSComctlLib.Toolbar ToolAux 
            Height          =   330
            Index           =   0
            Left            =   150
            TabIndex        =   69
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
         Index           =   15
         Left            =   4470
         MaxLength       =   5
         TabIndex        =   23
         Tag             =   "Dto adi. PVM|N|S|0|100|sfamia|dtopmv|#0.00|N|"
         Top             =   4755
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
         Index           =   10
         Left            =   4470
         MaxLength       =   5
         TabIndex        =   22
         Tag             =   "Centro de coste|N|S|0|100|sfamia|maxdtopar|#0.00|N|"
         Top             =   4200
         Width           =   1290
      End
      Begin VB.ComboBox Combo1 
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
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   3000
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   3120
         MaxLength       =   18
         TabIndex        =   20
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
         Left            =   4800
         MaxLength       =   15
         TabIndex        =   21
         Tag             =   "Código Almacen"
         Text            =   "codalmac"
         Top             =   3120
         Visible         =   0   'False
         Width           =   615
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   2535
         Left            =   120
         TabIndex        =   47
         Top             =   1140
         Width           =   8985
         _ExtentX        =   15849
         _ExtentY        =   4471
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   19
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
         Caption         =   "Cuentas Contables Compras "
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
         Height          =   1710
         Left            =   -74760
         TabIndex        =   42
         Top             =   3930
         Width           =   9000
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
            Left            =   2490
            MaxLength       =   10
            TabIndex        =   18
            Tag             =   "Cta. Compras servicios|T|S|||sfamia|ctacomprser||N|"
            Text            =   "Text1"
            Top             =   1215
            Width           =   1350
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
            Index           =   13
            Left            =   3915
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   53
            Text            =   "Text2"
            Top             =   1215
            Width           =   4965
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
            Index           =   3
            Left            =   3915
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   44
            Text            =   "Text2"
            Top             =   375
            Width           =   4965
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
            Index           =   5
            Left            =   3915
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   43
            Text            =   "Text2"
            Top             =   810
            Width           =   4965
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
            Left            =   2490
            MaxLength       =   10
            TabIndex        =   17
            Tag             =   "Cta.Abono Compras|T|N|||sfamia|abocompr||N|"
            Text            =   "Text1"
            Top             =   810
            Width           =   1350
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
            Index           =   3
            Left            =   2490
            MaxLength       =   10
            TabIndex        =   16
            Tag             =   "Cta. Contable compras|T|N|||sfamia|ctacompr||N|"
            Text            =   "Text1"
            Top             =   375
            Width           =   1350
         End
         Begin VB.Label Label1 
            Caption         =   "Servicios"
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
            Left            =   150
            TabIndex        =   54
            Top             =   1215
            Width           =   1695
         End
         Begin VB.Image imgCuentas 
            Height          =   240
            Index           =   8
            Left            =   2220
            ToolTipText     =   "Buscar cuenta contable"
            Top             =   1260
            Width           =   240
         End
         Begin VB.Image imgCuentas 
            Height          =   240
            Index           =   3
            Left            =   2220
            ToolTipText     =   "Buscar cuenta contable"
            Top             =   855
            Width           =   240
         End
         Begin VB.Image imgCuentas 
            Height          =   240
            Index           =   1
            Left            =   2220
            ToolTipText     =   "Buscar cuenta contable"
            Top             =   450
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Compras"
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
            Left            =   150
            TabIndex        =   46
            Top             =   405
            Width           =   1815
         End
         Begin VB.Label Label1 
            Caption         =   "Abono"
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
            TabIndex        =   45
            Top             =   810
            Width           =   1455
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Cuentas Contables Ventas "
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
         Height          =   3210
         Left            =   -74760
         TabIndex        =   33
         Top             =   480
         Width           =   9000
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
            Left            =   3915
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   51
            Text            =   "Text2"
            Top             =   2625
            Width           =   4965
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
            Left            =   2490
            MaxLength       =   10
            TabIndex        =   15
            Tag             =   "Cta. alternativa vta sevicios|T|N|||sfamia|ctavtaseralt||N|"
            Text            =   "Text1"
            Top             =   2625
            Width           =   1350
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
            Index           =   11
            Left            =   3915
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   49
            Text            =   "Text2"
            Top             =   2193
            Width           =   4965
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
            Left            =   2490
            MaxLength       =   10
            TabIndex        =   14
            Tag             =   "Cta. Alternativa Ventas|T|N|||sfamia|ctavtaser||N|"
            Text            =   "Text1"
            Top             =   2193
            Width           =   1350
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
            Left            =   2490
            MaxLength       =   10
            TabIndex        =   12
            Tag             =   "Cta. Alternativa Ventas|T|N|||sfamia|ctavent1||N|"
            Text            =   "Text1"
            Top             =   1329
            Width           =   1350
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
            Left            =   2490
            MaxLength       =   10
            TabIndex        =   13
            Tag             =   "Cta. Alternativa Abonos|T|N|||sfamia|abovent1||N|"
            Text            =   "Text1"
            Top             =   1761
            Width           =   1350
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
            Index           =   7
            Left            =   3915
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   37
            Text            =   "Text2"
            Top             =   1761
            Width           =   4965
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
            Index           =   6
            Left            =   3915
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   36
            Text            =   "Text2"
            Top             =   1329
            Width           =   4965
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
            Left            =   3915
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   35
            Text            =   "Text2"
            Top             =   465
            Width           =   4965
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
            Index           =   4
            Left            =   3915
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   34
            Text            =   "Text2"
            Top             =   897
            Width           =   4965
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
            Left            =   2490
            MaxLength       =   10
            TabIndex        =   11
            Tag             =   "Cta. Abono Ventas|T|N|||sfamia|aboventa||N|"
            Text            =   "Text1"
            Top             =   897
            Width           =   1350
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
            Index           =   2
            Left            =   2490
            MaxLength       =   10
            TabIndex        =   10
            Tag             =   "Cta. Contable Ventas|T|N|||sfamia|ctaventa||N|"
            Text            =   "Text1"
            Top             =   465
            Width           =   1350
         End
         Begin VB.Image imgCuentas 
            Height          =   240
            Index           =   0
            Left            =   2205
            Picture         =   "frmAlmFamiliaArticuloGr.frx":004A
            Tag             =   "-1"
            ToolTipText     =   "Buscar proveedor"
            Top             =   495
            Width           =   240
         End
         Begin VB.Image imgCuentas 
            Height          =   240
            Index           =   7
            Left            =   2220
            Tag             =   "-1"
            ToolTipText     =   "Buscar cuenta contable"
            Top             =   2670
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Alternativa servicios"
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
            Index           =   12
            Left            =   150
            TabIndex        =   52
            Top             =   2655
            Width           =   2040
         End
         Begin VB.Image imgCuentas 
            Height          =   240
            Index           =   6
            Left            =   2220
            Tag             =   "-1"
            ToolTipText     =   "Buscar cuenta contable"
            Top             =   2190
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Ventas servicios"
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
            Left            =   150
            TabIndex        =   50
            Top             =   2235
            Width           =   1815
         End
         Begin VB.Label Label1 
            Caption         =   "Alternativa Abonos"
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
            Left            =   150
            TabIndex        =   41
            Top             =   1785
            Width           =   1950
         End
         Begin VB.Label Label1 
            Caption         =   "Alternativa"
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
            Left            =   150
            TabIndex        =   40
            Top             =   1335
            Width           =   1815
         End
         Begin VB.Image imgCuentas 
            Height          =   240
            Index           =   4
            Left            =   2220
            Tag             =   "-1"
            ToolTipText     =   "Buscar cuenta contable"
            Top             =   1350
            Width           =   240
         End
         Begin VB.Image imgCuentas 
            Height          =   240
            Index           =   5
            Left            =   2220
            ToolTipText     =   "Buscar cuenta contable"
            Top             =   1785
            Width           =   240
         End
         Begin VB.Image imgCuentas 
            Height          =   240
            Index           =   2
            Left            =   2220
            ToolTipText     =   "Buscar cuenta contable"
            Top             =   930
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Ventas"
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
            Left            =   150
            TabIndex        =   39
            Top             =   495
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Abono"
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
            Left            =   150
            TabIndex        =   38
            Top             =   900
            Width           =   1575
         End
      End
      Begin VB.Label Label1 
         Caption         =   "Descuento adicional cálculo precio mínimo"
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
         Index           =   15
         Left            =   255
         TabIndex        =   56
         Top             =   4800
         Width           =   4260
      End
      Begin VB.Label Label1 
         Caption         =   "Máximo descuento particulares"
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
         Index           =   10
         Left            =   255
         TabIndex        =   48
         Top             =   4230
         Width           =   4005
      End
   End
   Begin VB.CheckBox chkInstalac 
      Caption         =   "¿Es instalación?"
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
      Left            =   9810
      TabIndex        =   4
      Tag             =   "¿Es instalación?|N|N|||sfamia|instalac||N|"
      Top             =   1080
      Width           =   2910
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
      Left            =   13365
      TabIndex        =   25
      Top             =   9180
      Visible         =   0   'False
      Width           =   1065
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
      Index           =   1
      Left            =   4215
      MaxLength       =   40
      TabIndex        =   1
      Tag             =   "Denominación familia de Artículo|T|N|||sfamia|nomfamia||N|"
      Text            =   "Text1"
      Top             =   1095
      Width           =   4950
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
      Index           =   0
      Left            =   1410
      MaxLength       =   5
      TabIndex        =   0
      Tag             =   "Código familia de artículo|N|N|0|60000|sfamia|codfamia|0000|S|"
      Text            =   "Text"
      Top             =   1095
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Left            =   180
      TabIndex        =   28
      Top             =   9045
      Width           =   2655
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
         TabIndex        =   29
         Top             =   180
         Width           =   2355
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
      Left            =   13365
      TabIndex        =   27
      Top             =   9180
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
      Left            =   12150
      TabIndex        =   24
      Top             =   9180
      Width           =   1065
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   375
      Left            =   2925
      Top             =   9180
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
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   1
      Left            =   1425
      ToolTipText     =   "Buscar proveedor"
      Top             =   1890
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
      Index           =   9
      Left            =   360
      TabIndex        =   60
      Top             =   1890
      Width           =   1035
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   0
      Left            =   1425
      ToolTipText     =   "Buscar centro coste"
      Top             =   2370
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "CCoste"
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
      TabIndex        =   59
      Top             =   2370
      Width           =   750
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Descripción"
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
      Index           =   14
      Left            =   9720
      TabIndex        =   55
      Top             =   2835
      Width           =   1155
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Denominación"
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
      Left            =   2610
      TabIndex        =   31
      Top             =   1095
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Código"
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
      Left            =   315
      TabIndex        =   30
      Top             =   1095
      Width           =   690
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
Attribute VB_Name = "frmAlmFamiliaArticuloGr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

Private WithEvents frmB As frmBuscaGrid 'Form para busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmProv As frmBasico2 'proveedores
Attribute frmProv.VB_VarHelpID = -1
Private WithEvents frmFamP As frmBasico2 ' mandabusquedaprevia
Attribute frmFamP.VB_VarHelpID = -1
Private WithEvents frmCCos As frmBasico2 'centros de coste
Attribute frmCCos.VB_VarHelpID = -1
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


Dim indCodigo As Integer
Dim cadB As String
Private BuscaChekc As String

Private Sub chkBloqTPV_Click()
    If Modo = 1 Then
        'Buscqueda
        If InStr(1, BuscaChekc, "chkBloqTPV") = 0 Then BuscaChekc = BuscaChekc & "chkBloqTPV|"
    End If

End Sub

Private Sub chkBloqTPV_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub chkBloqTPV_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkComunica_Click()
    If Modo = 1 Then
        'Buscqueda
        If InStr(1, BuscaChekc, "chkComunica") = 0 Then BuscaChekc = BuscaChekc & "chkComunica|"
    End If
End Sub

Private Sub chkComunica_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub chkComunica_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub chkInactiva_Click()
    If Modo = 1 Then
        'Buscqueda
        If InStr(1, BuscaChekc, "chkInactiva") = 0 Then BuscaChekc = BuscaChekc & "chkInactiva|"
    End If
End Sub

Private Sub chkInactiva_GotFocus()
    ConseguirfocoChk Modo
End Sub


Private Sub chkInactiva_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub chkInstalac_Click()
    If Modo = 1 Then
        'Buscqueda
        If InStr(1, BuscaChekc, "chkInstalac") = 0 Then BuscaChekc = BuscaChekc & "chkInstalac|"
    End If
End Sub

Private Sub chkInstalac_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub chkInstalac_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub



Private Sub chkMarcaPropia_Click()
    If Modo = 1 Then
        'Buscqueda
        If InStr(1, BuscaChekc, "chkMarcaPropia") = 0 Then BuscaChekc = BuscaChekc & "chkMarcaPropia|"
    End If
End Sub

Private Sub chkMarcaPropia_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub chkMarcaPropia_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub chkPuntos_Click()
    If Modo = 1 Then
        'Buscqueda
        If InStr(1, BuscaChekc, "chkPuntos") = 0 Then BuscaChekc = BuscaChekc & "chkPuntos|"
    End If
End Sub

Private Sub chkPuntos_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub chkPuntos_KeyPress(KeyAscii As Integer)
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
                    'PonerBotonCabecera True
                    Me.DataGrid1.Enabled = True
                    PonerModo 2
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
            'PonerBotonCabecera True
            PonerModo 2
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
    DesplazamientoData Data1, Index, True
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
Dim Cad As String

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
    'El registro de codigo 0 no se puede Modificar ni Eliminar
    'If EsCodigoCero(CStr(Data1.Recordset.Fields(0).Value), FormatoCampo(Text1(0))) Then Exit Sub
    Cad = DevuelveDesdeBD(conAri, "count(*)", "sartic", "codfamia", Text1(0).Text)
    If Cad = "" Then Cad = "0"
    If Val(Cad) > 0 Then
        Cad = "Hay " & Cad & " artículos pertenecientes a esta familia"
        MsgBox Cad, vbExclamation
        Exit Sub
    End If
    
    '### a mano
    Cad = "¿Seguro que desea eliminar la Familia de Artículo?:" & vbCrLf
    Cad = Cad & vbCrLf & "Cod. : " & Format(Data1.Recordset.Fields(0), FormatoCampo(Text1(0)))
    Cad = Cad & vbCrLf & "Desc.: " & Data1.Recordset.Fields(1)
    
    



    'Borramos
    If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        On Error GoTo Error2
        Screen.MousePointer = vbHourglass
        NumRegElim = Data1.Recordset.AbsolutePosition
        'Borramos en dtofamia
        lblIndicador.Caption = "Eliminando descuentos"
        lblIndicador.Refresh
        Cad = "Delete from sdtofm where codfamia = " & Data1.Recordset!Codfamia
        conn.Execute Cad
        
        
        Cad = "Delete from sdtomp where codfamia = " & Data1.Recordset!Codfamia
        conn.Execute Cad
        
        
        Cad = "Delete from sfamiadtos where codfamia = " & Data1.Recordset!Codfamia 'despues del DELETE
        conn.Execute Cad
        
        lblIndicador.Caption = "Eliminando familia"
        lblIndicador.Refresh
        Data1.Recordset.Delete
        
        ejecutar Cad, False
        
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
Dim Cad As String

    If Modo = 5 Then  'modo 5: Mantenimientos Lineas
    
    
        PonerModo 2
        If Not Data1.Recordset.EOF Then Me.lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
    Else
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
Dim I As Integer
    'Icono del formulario
    Me.Icon = frmPpal.Icon
    If vParamAplic.Descriptores Then Me.Caption = "Categorias Art."

    For I = 1 To imgCuentas.Count - 1
        imgCuentas(I).Picture = imgCuentas(0).Picture
    Next
    For I = 0 To imgBuscar.Count - 1
        imgBuscar(I).Picture = imgCuentas(0).Picture
    Next

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
        .Buttons(1).Image = 42 ' actualizar dto/familia
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
    
    For I = 0 To ToolAux.Count - 1
        With Me.ToolAux(I)
            .HotImageList = frmPpal.imgListComun_OM16
            .DisabledImageList = frmPpal.imgListComun_BN16
            .ImageList = frmPpal.imgListComun16
            .Buttons(1).Image = 3   'Insertar
            .Buttons(2).Image = 4   'Modificar
            .Buttons(3).Image = 5   'Borrar
        End With
    Next I
    
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
    
    
    Me.chkPuntos.visible = vParamAplic.PtosAsignar > 0
'    Label4(1).visible = vParamAplic.PtosAsignar > 0
    
    
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
    chkPuntos.Value = 0
    Me.chkInactiva.Value = 0
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


Private Sub frmFamP_DatoSeleccionado(CadenaSeleccion As String)
Dim Aux As String

    HaDevueltoDatos = True
    Screen.MousePointer = vbHourglass
    cadB = ""
    Aux = ValorDevueltoFormGrid(Text1(0), CadenaSeleccion, 1)
    cadB = Aux
    'Se muestran en el mismo form
    CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
    PonerCadenaBusqueda
    Screen.MousePointer = vbDefault
End Sub

Private Sub frmProv_DatoSeleccionado(CadenaSeleccion As String)
    Text1(9).Text = RecuperaValor(CadenaSeleccion, 1) 'proveedores
    If Text1(9).Text <> "" Then Text1(9).Text = Format(Text1(9).Text, "000000")
    Text2(9).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub imgBuscar_Click(Index As Integer)
    ' ---- [19/10/2009] [LAURA]: añadir campo centro de coste familia
    
    If Modo = 2 Or Modo = 5 Or Modo = 0 Then Exit Sub
    
    Select Case Index
        Case 0 'Centros de coste de la conta
            Screen.MousePointer = vbHourglass
            
            Set frmCCos = New frmBasico2
            AyudaCentroCoste frmCCos, Text1(8).Text
            Set frmCCos = Nothing

            Screen.MousePointer = vbDefault
            PonerFoco Text1(8)
            
        Case 1
            Screen.MousePointer = vbHourglass

            Set frmProv = New frmBasico2
            AyudaProveedores frmProv, Text1(9).Text
            Set frmProv = Nothing

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
    If KeyAscii = teclaBuscar Then
        Select Case Index
            'ventas
            Case 2: KEYBusqueda2 KeyAscii, 0 'cuenta ventas
            Case 4: KEYBusqueda2 KeyAscii, 2 'cuenta abonos
            Case 6: KEYBusqueda2 KeyAscii, 4 'cuenta alternativa
            Case 7: KEYBusqueda2 KeyAscii, 5 'cuenta alternativa abonos
            Case 11: KEYBusqueda2 KeyAscii, 6 'cuenta ventas servicios
            Case 12: KEYBusqueda2 KeyAscii, 7 'cuenta alternativa servicios
            'compras
            Case 3: KEYBusqueda2 KeyAscii, 1 'cuenta compras
            Case 5: KEYBusqueda2 KeyAscii, 3 'cuenta compras abonos
            Case 13: KEYBusqueda2 KeyAscii, 8 'cuenta servicios
            
            Case 8: KEYBusqueda KeyAscii, 0 'centro de coste
            Case 9: KEYBusqueda KeyAscii, 1 'proveedor
        End Select
    Else
        KEYpress KeyAscii
    End If

End Sub

Private Sub KEYBusqueda2(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgCuentas_Click (indice)
End Sub



Private Sub KEYBusqueda(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgBuscar_Click (indice)
End Sub

Private Sub KEYFecha(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgFec_Click (indice)
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
        Case 10, 15
            If Not PonerFormatoDecimal(Text1(Index), 4) Then Text1(Index).Text = ""
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
Dim Cad As String
Dim tabla As String
Dim Titulo As String
Dim CargaF As Boolean 'Para saber si se carga el frame o no en el BuscaGrid
Dim Conexion As Byte

'        'Llamamos a al form
'        '##A mano
'        Cad = ""
'        If Val(Me.imgCuentas(0).Tag) >= 0 Then
'        'Se llama a Busqueda desde un campo de Cuenta
'            '#A MANO: Porque busca en la tabla Cuentas
'            'de la base de datos de Contabilidad
'            Cad = Cad & "Código|Cuentas|codmacta|T||15·Denominacion|Cuentas|nommacta|T||70·"
'            tabla = "Cuentas"
'            Titulo = "Cuentas"
'            Conexion = conConta    'Conexión a BD: Conta
'            CargaF = True 'Se puede cargar el frame
'        Else
'            'Busqueda de una Família de Artículo
'            Cad = Cad & ParaGrid(Text1(0), 15, "Código")
'            Cad = Cad & ParaGrid(Text1(1), 80, "Denominacion")
'            tabla = "sfamia"
'            Titulo = "Família de Artículos"
'            If vParamAplic.Descriptores Then Titulo = "Categorias Art."
'            Conexion = conAri    'Conexión a BD: Ariges
'            CargaF = False 'No se carga el frame
'        End If
'
'        If Cad <> "" Then
'            Screen.MousePointer = vbHourglass
'            Set frmB = New frmBuscaGrid
'            frmB.vCampos = Cad
'            frmB.vTabla = tabla
'            frmB.vSQL = CadB
'            HaDevueltoDatos = False
'            '###A mano
'            frmB.vDevuelve = "0|1|"
'            frmB.vTitulo = Titulo
'            frmB.vselElem = 1
'            frmB.vConexionGrid = Conexion
'            frmB.vCargaFrame = CargaF
'            '#
'            frmB.Show vbModal
'            Set frmB = Nothing
'            'Si ha puesto valores y tenemos que es formulario de busqueda entonces
'            'tendremos que cerrar el form lanzando el evento
'            If HaDevueltoDatos Then
'                If kCampo < 5 Then PonerFoco Text1(kCampo + 1)
'                If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then
'                    If Not (Val(Me.imgCuentas(0).Tag) >= 0) Then cmdRegresar_Click
'                End If
'            Else   'de ha devuelto datos, es decir NO ha devuelto datos
'                PonerFoco Text1(kCampo)
''                If Modo = 1 Then
''                    MsgBox "No hay ningún registro en la tabla " & tabla
''                    PonerFoco Text1(0)
''                End If
'            End If
'        End If

    Set frmFamP = New frmBasico2
    
    AyudaFamilias frmFamP, CStr(Text1(0)), cadB
    
    Set frmFamP = Nothing
    
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
Dim I As Byte
    
    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma Me, Data1
    
    'poner la descripcion de las cuentas
    For I = 2 To 7
        Text2(I).Text = PonerNombreCuenta(Text1(I), Modo)
    Next I
    For I = 11 To 13
        
        Text2(I).Text = PonerNombreCuenta(Text1(I), Modo)
    Next I
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
Dim B As Boolean
Dim NumReg As Byte
Dim I As Integer

 
    'Actualiza Iconos Insertar,Modificar,Eliminar
'    ActualizarToolbarGnral Me.Toolbar1, Modo, Kmodo, 5
    
    Modo = Kmodo
    PonerIndicador lblIndicador, Modo
    If Modo = 5 Then lblIndicador.Caption = "Lineas dto"
    
    
'    Frame4.Enabled = Modo < 5  'En lineas no dejo trabajar
    B = Modo < 5
    If Not B Then ModificaLineas = 0
    'datagrid1.enabled
    
    '-------------------------------------------------
    'Modo 2. Hay datos y estamos visualizandolos
    B = (Kmodo = 2)
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.visible = B
    Else
        'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
        cmdRegresar.visible = False
    End If
    
    'Visualizar flechas de desplazamiento en la toolbar si modo=2
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
    DesplazamientoVisible B And Data1.Recordset.RecordCount > 1
    
    'Poner Boton de Cabecera o Aceptar/Cancelar
'    PonerBotonCabecera B Or (Modo = 0)
    Me.cmdAceptar.visible = (Modo = 1 Or Modo = 3 Or Modo = 4 Or Modo = 5)
    Me.cmdCancelar.visible = (Modo = 1 Or Modo = 3 Or Modo = 4 Or Modo = 5)

        
    'Bloquear Registros si modo distinto de Insertar o Modificar
    BloquearText1 Me, Modo
    
    For I = 0 To Text1.Count - 1
        Text1(I).Enabled = (Modo = 1 Or Modo = 3 Or Modo = 4)
    Next I
    
    BloquearChecks Me, Modo
        
    Me.chkBloqTPV.Enabled = Modo = 1 Or Modo = 3 Or Modo = 4
    chkComunica.Enabled = Me.chkBloqTPV.Enabled
    chkMarcaPropia.Enabled = Me.chkBloqTPV.Enabled
    
    Me.chkInactiva.Enabled = Me.chkBloqTPV.Enabled
    Me.chkInstalac.Enabled = Me.chkBloqTPV.Enabled
    
    chkVistaPrevia.Enabled = (Modo <= 2)
    
    If vParamAplic.PtosAsignar > 0 Then
        B = False
        If Modo = 1 Then
            B = True
        Else
            If Modo > 2 Then B = vUsu.Nivel = 0
        End If
        chkPuntos.Enabled = B
        
        chkPuntos.Enabled = chkPuntos.Enabled And Modo <> 5
    End If
    
    B = (Modo = 1 Or Modo = 3 Or Modo = 4)
    imgBuscar(0).Enabled = B And vEmpresa.TieneAnalitica
    imgBuscar(0).visible = B And vEmpresa.TieneAnalitica
    imgBuscar(1).Enabled = B
    imgBuscar(1).visible = B
    
    For I = 0 To Me.imgCuentas.Count - 1
        imgCuentas(I).Enabled = B
        imgCuentas(I).visible = B
    Next I
    
    
    'Poner el tamaño de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos
    
    PonerModoOpcionesMenu 'Activar opciones de menu según MODO
    PonerOpcionesMenu   'Activar opciones de menu según NIVEL
                        'de permisos del usuario
End Sub

Private Sub DesplazamientoVisible(bol As Boolean)
    FrameDesplazamiento.visible = bol
    FrameDesplazamiento.Enabled = bol
End Sub



Private Sub PonerModoOpcionesMenu()
Dim B As Boolean
Dim I As Integer
Dim bAux As Boolean

On Error Resume Next

    B = (Modo = 0 Or Modo = 2)
    
    'Añadir
    Toolbar1.Buttons(1).Enabled = B
    Me.mnNuevo.Enabled = B
    
    'Buscar
    Toolbar1.Buttons(5).Enabled = B
    Me.mnBuscar.Enabled = B
    'VerTodos
    Toolbar1.Buttons(6).Enabled = B
    Me.mnVerTodos.Enabled = B
    
    B = (Modo = 2)
    Toolbar5.Buttons(1).Enabled = B
    
    'Modificar
    Toolbar1.Buttons(2).Enabled = B
    mnModificar.Enabled = B
    'eliminar
    Toolbar1.Buttons(3).Enabled = B
    mnEliminar.Enabled = B
    
    'imprimir
    Toolbar1.Buttons(8).Enabled = True
    
    B = (Modo = 2) And DatosADevolverBusqueda = ""
    For I = 0 To ToolAux.Count - 1
        ToolAux(I).Buttons(1).Enabled = B
        bAux = (B And Me.Data2.Recordset.RecordCount > 0)
        ToolAux(I).Buttons(2).Enabled = bAux
        ToolAux(I).Buttons(3).Enabled = bAux
    Next I
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
    PonerLongCamposGnral Me, Modo, 1
End Sub


Private Function DatosOk() As Boolean
Dim B As Boolean

    DatosOk = False
    B = CompForm(Me, 1)
    If Not B Then Exit Function
    
    'Comprobar si ya existe el cod de familia en la tabla
    If Modo = 3 Then 'Insertar
        If ExisteCP(Text1(0)) Then B = False
    End If
    
    DatosOk = B
End Function

Private Sub ToolAux_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
    PonerModo 5
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
        Case 1 'Nuevo
            mnNuevo_Click
        Case 2  'Modificar
            mnModificar_Click
        Case 3 'Eliminar
            mnEliminar_Click
        Case 5 'Busqueda
            mnBuscar_Click
        Case 6 'Ver Todos
            mnVerTodos_Click
        Case 8 'Imprimir
            BotonImprimir
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


Private Sub PonerBotonCabecera(B As Boolean)

    Me.cmdAceptar.visible = Not B
    Me.cmdCancelar.visible = Not B
    If B Then
        PonerFocoBtn Me.cmdRegresar
        cmdRegresar.Cancel = True
    Else
        cmdCancelar.Cancel = True
    End If
    
    
    Me.cmdRegresar.Caption = "Cabecera"
    If B Then
        B = False
        If Modo = 2 Then
            B = Me.DatosADevolverBusqueda <> ""
            If B Then Me.cmdRegresar.Caption = "Regresar"
        ElseIf Modo = 5 Then
            B = ModificaLineas = 0
        End If
        
    End If
    Me.cmdRegresar.visible = B
   
    
    'Habilitar las opciones correctas del menu
    PonerModoOpcionesMenu
    PonerOpcionesMenu
    If Err.Number <> 0 Then Err.Clear

    
End Sub


Private Sub PosicionarData()
Dim Cad As String, Indicador As String

    Cad = "(codfamia=" & Text1(0).Text & ")"
    If SituarData(Data1, Cad, Indicador) Then
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
'    frmListado3.Opcion = 12
'    frmListado3.Show vbModal
    Me.Hide
    frmInformesNew.OpcionListado = 7
    frmInformesNew.Show vbModal
    Me.Show vbModal
End Sub


Private Sub TratarCtaContable()
Dim I As Integer
Dim CtaCreadas As String
    For I = 2 To 7
        If Text2(I).Text = vbCrearNuevaCta Then
            If InStr(1, CtaCreadas, Text1(I).Text & "|") = 0 Then
                InsertarCuentaCble Text1(I).Text, "", "", Text1(1).Text
                CtaCreadas = CtaCreadas & Text1(I).Text & "|"
            End If
            Text2(I).Text = Text1(1).Text
        End If
    Next I
End Sub

Private Sub BotonMtoLineas()

       If Data1.Recordset Is Nothing Then Exit Sub
       If Data1.Recordset.EOF Then Exit Sub
       
       
        If vParamAplic.NumeroInstalacion = 2 Then
            CadenaConsulta = " codartic in (select codartic from sartic where codfamia=" & Data1.Recordset!Codfamia & ") AND 1"
            CadenaConsulta = DevuelveDesdeBD(conAri, "count(distinct(codartic)) ", " sprees", CadenaConsulta, "1")
            If Val(CadenaConsulta) > 0 Then MsgBox "Hay " & CadenaConsulta & " articulos en precios especiales pertenecientes a esta familia", vbExclamation
            CadenaConsulta = ""
        End If
       
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
'    PonerBotonCabecera False
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
'    PonerBotonCabecera False
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
Dim B As Boolean
Dim SQL As String

    On Error GoTo ECargaGrid

    B = DataGrid1.Enabled
    
    SQL = MontaSQLCarga(enlaza)
    CargaGridGnral vDataGrid, vData, SQL, True

    
    CargaGrid2 vDataGrid, vData
    vDataGrid.ScrollBars = dbgAutomatic
        
    B = (Modo = 5) And (ModificaLineas = 1 Or ModificaLineas = 2)
    vDataGrid.Enabled = Not B

    'PrimeraVez = False
    Exit Sub
    
ECargaGrid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub


Private Sub CargaGrid2(ByRef vDataGrid As DataGrid, ByRef vData As Adodc)
Dim I As Integer

    On Error GoTo ECargaGrid

    vData.Refresh

    vDataGrid.Columns(0).Caption = "Código"
    vDataGrid.Columns(0).visible = False
    
    vDataGrid.Columns(1).Caption = "Descripción"
    vDataGrid.Columns(1).Width = 5285

    vDataGrid.Columns(2).Caption = "Descuento 1"
    vDataGrid.Columns(2).Width = 1520
    vDataGrid.Columns(2).Alignment = dbgRight
    vDataGrid.Columns(2).NumberFormat = FormatoDescuento
    
    vDataGrid.Columns(3).Caption = "Descuento 2"
    vDataGrid.Columns(3).Width = 1520
    vDataGrid.Columns(3).Alignment = dbgRight
    vDataGrid.Columns(3).NumberFormat = FormatoDescuento
    
    For I = 0 To vDataGrid.Columns.Count - 1
        vDataGrid.Columns(I).Locked = True
        vDataGrid.Columns(I).AllowSizing = False
    Next I
    vDataGrid.HoldFields
    vDataGrid.RowHeight = 350
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
        For I = 0 To txtAux.Count - 1 'TextBox
            txtAux(I).Top = 290
            txtAux(I).visible = visible
        Next I
        Combo1.visible = False
    Else
        If limpiar Then 'Vaciar los textBox (Vamos a Insertar)
            DeseleccionaGrid DataGrid1
            For I = 0 To txtAux.Count - 1
                txtAux(I).Text = ""
              '  BloquearTxt txtAux(i), False  'Todos menos el nombre
            Next I

        Else 'Vamos a modificar
            For I = 0 To txtAux.Count - 1
         
                txtAux(I).Text = DataGrid1.Columns(I + 2).Text
            
                txtAux(I).Locked = False
            Next I
        End If
        
        'Fijamos altura(Height) y posición Top
        '-------------------------------
        alto = ObtenerAlto(DataGrid1, 10)
        
        For I = 0 To txtAux.Count - 1
            txtAux(I).Top = alto
            txtAux(I).Height = DataGrid1.RowHeight
        Next I
        Combo1.Top = alto
        'Fijamos anchura y posicion Left
        '--------------------------------
        'Cod. Almac

        'Precio, Dto1, Dto2, Precio
        Combo1.Left = DataGrid1.Left + 340
        Combo1.Width = DataGrid1.Columns(2).Left - DataGrid1.Left - 240
        For I = 0 To txtAux.Count - 1
            txtAux(I).Left = DataGrid1.Columns(I + 2).Left + 10 + 120
            txtAux(I).Width = DataGrid1.Columns(I + 2).Width - 10
        Next I
        

        Combo1.visible = limpiar
        'Los ponemos Visibles o No
        '--------------------------
        For I = 0 To txtAux.Count - 1
            txtAux(I).visible = visible
        Next I
        
    End If

End Sub


Private Sub Toolbar5_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1:
            'Actualizar dto/familia-marca
            ActualizarDtoFamilia
    End Select

End Sub

Private Sub ToolbarDes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Desplazamiento (Button.Index)
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
    If Text1(9).Text <> "" Then
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



