VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRepNumSerie2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Numeros de Serie"
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   11430
   ClipControls    =   0   'False
   Icon            =   "frmRepNumSerie.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   11430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   3880
      Left            =   240
      TabIndex        =   31
      Top             =   1920
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   6853
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Datos compra/venta"
      TabPicture(0)   =   "frmRepNumSerie.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameActuales"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "FrameNuevos"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "FrameBaja"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "FrameSusti"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Histórico"
      TabPicture(1)   =   "frmRepNumSerie.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "chkAux"
      Tab(1).Control(1)=   "txtAux(2)"
      Tab(1).Control(2)=   "txtAux2(1)"
      Tab(1).Control(3)=   "txtAux2(7)"
      Tab(1).Control(4)=   "txtAux2(6)"
      Tab(1).Control(5)=   "txtAux2(5)"
      Tab(1).Control(6)=   "txtAux2(4)"
      Tab(1).Control(7)=   "txtAux2(3)"
      Tab(1).Control(8)=   "txtAux2(2)"
      Tab(1).Control(9)=   "txtAux2(0)"
      Tab(1).Control(10)=   "txtAux(1)"
      Tab(1).Control(11)=   "txtAux(0)"
      Tab(1).Control(12)=   "DataGrid1"
      Tab(1).Control(13)=   "Label1(15)"
      Tab(1).Control(14)=   "Label1(14)"
      Tab(1).Control(15)=   "Label1(13)"
      Tab(1).Control(16)=   "Label1(12)"
      Tab(1).Control(17)=   "Label1(11)"
      Tab(1).Control(18)=   "Label1(10)"
      Tab(1).Control(19)=   "Label1(7)"
      Tab(1).ControlCount=   20
      Begin VB.CheckBox chkAux 
         Enabled         =   0   'False
         Height          =   195
         Left            =   -66120
         TabIndex        =   79
         Top             =   600
         Width           =   255
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   320
         Index           =   2
         Left            =   -74760
         MaxLength       =   10
         TabIndex        =   78
         Text            =   "fecha"
         Top             =   3000
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtAux2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   1
         Left            =   -70800
         Locked          =   -1  'True
         TabIndex        =   77
         Text            =   "nomdpto"
         Top             =   3000
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.TextBox txtAux2 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   7
         Left            =   -65280
         MaxLength       =   5
         TabIndex        =   76
         Top             =   2880
         Width           =   615
      End
      Begin VB.TextBox txtAux2 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   6
         Left            =   -65280
         MaxLength       =   10
         TabIndex        =   74
         Top             =   2520
         Width           =   1095
      End
      Begin VB.TextBox txtAux2 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   5
         Left            =   -65280
         MaxLength       =   10
         TabIndex        =   72
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox txtAux2 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   4
         Left            =   -65280
         MaxLength       =   10
         TabIndex        =   70
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox txtAux2 
         Height          =   315
         Index           =   3
         Left            =   -65280
         TabIndex        =   68
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox txtAux2 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   2
         Left            =   -65280
         MaxLength       =   10
         TabIndex        =   66
         Top             =   960
         Width           =   1065
      End
      Begin VB.TextBox txtAux2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   0
         Left            =   -73080
         Locked          =   -1  'True
         TabIndex        =   64
         Text            =   "nomclien"
         Top             =   3000
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   320
         Index           =   1
         Left            =   -71640
         MaxLength       =   6
         TabIndex        =   63
         Text            =   "coddpt"
         Top             =   3000
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   320
         Index           =   0
         Left            =   -73920
         MaxLength       =   6
         TabIndex        =   62
         Text            =   "codcli"
         Top             =   3000
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Frame FrameSusti 
         Caption         =   " Sustituido por "
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
         Height          =   680
         Left            =   5760
         TabIndex        =   58
         Top             =   3080
         Visible         =   0   'False
         Width           =   5175
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   17
            Left            =   1350
            MaxLength       =   15
            TabIndex        =   59
            Tag             =   "Nº Serie|T|S|||sserie|numsersu||N|"
            Text            =   "Text1"
            Top             =   240
            Width           =   1710
         End
         Begin VB.Label Label3 
            Caption         =   "Nº Serie"
            Height          =   255
            Index           =   1
            Left            =   600
            TabIndex        =   60
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Frame FrameBaja 
         Caption         =   "Datos de baja"
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
         Height          =   880
         Left            =   5760
         TabIndex        =   53
         Top             =   2120
         Width           =   5175
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   18
            Left            =   150
            MaxLength       =   10
            TabIndex        =   55
            Tag             =   "Fecha Baja|F|S|||sserie|fechabaja|dd/mm/yyyy|N|"
            Text            =   "Text1"
            Top             =   440
            Width           =   1200
         End
         Begin VB.ComboBox cboMotivoBaja 
            Height          =   315
            ItemData        =   "frmRepNumSerie.frx":0044
            Left            =   1560
            List            =   "frmRepNumSerie.frx":0046
            Style           =   2  'Dropdown List
            TabIndex        =   54
            Tag             =   "Motivo de Baja|N|S|||sserie|codmotba|0|N|"
            Top             =   440
            Width           =   3015
         End
         Begin VB.Label Label3 
            Caption         =   "Fecha"
            Height          =   255
            Index           =   2
            Left            =   150
            TabIndex        =   57
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label3 
            Caption         =   "Motivo"
            Height          =   255
            Index           =   3
            Left            =   1560
            TabIndex        =   56
            Top             =   240
            Width           =   1215
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   2
            Left            =   720
            Picture         =   "frmRepNumSerie.frx":0048
            ToolTipText     =   "Buscar fecha"
            Top             =   220
            Width           =   240
         End
      End
      Begin VB.Frame FrameNuevos 
         Caption         =   " Datos Compra "
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
         Height          =   1635
         Left            =   5760
         TabIndex        =   43
         Top             =   440
         Width           =   5175
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   12
            Left            =   1020
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   48
            Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwwwwwwwwww aq"
            Top             =   460
            Width           =   4035
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   12
            Left            =   165
            MaxLength       =   6
            TabIndex        =   47
            Tag             =   "Cod. Proveedor|N|S|0|999999|sserie|codprove|000000|N|"
            Text            =   "Text11"
            Top             =   460
            Width           =   780
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   14
            Left            =   1200
            MaxLength       =   10
            TabIndex        =   46
            Tag             =   "Fecha Compra|F|S|||sserie|fechacom|dd/mm/yyyy|N|"
            Text            =   "Text1"
            Top             =   1200
            Width           =   1200
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   15
            Left            =   4200
            MaxLength       =   5
            TabIndex        =   45
            Tag             =   "Nº linea|N|S|0|99999|sserie|numline2||N|"
            Text            =   "Text1"
            Top             =   1200
            Width           =   735
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   13
            Left            =   1200
            MaxLength       =   10
            TabIndex        =   44
            Tag             =   "Nº Albaran Compra|T|S|||sserie|numalbpr||N|"
            Text            =   "Text1 Text"
            Top             =   840
            Width           =   1200
         End
         Begin VB.Label Label1 
            Caption         =   "Proveedor"
            Height          =   255
            Index           =   4
            Left            =   165
            TabIndex        =   52
            Top             =   260
            Width           =   735
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   4
            Left            =   1020
            ToolTipText     =   "Buscar proveedor"
            Top             =   220
            Width           =   240
         End
         Begin VB.Label Label2 
            Caption         =   "Fecha Compra"
            Height          =   255
            Left            =   120
            TabIndex        =   51
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Nº Albaran"
            Height          =   255
            Index           =   5
            Left            =   165
            TabIndex        =   50
            Top             =   840
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Nº linea Compra"
            Height          =   255
            Index           =   6
            Left            =   2640
            TabIndex        =   49
            Top             =   1200
            Width           =   1215
         End
      End
      Begin VB.Frame FrameActuales 
         Caption         =   " Datos Venta "
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
         Height          =   3320
         Left            =   240
         TabIndex        =   32
         Top             =   440
         Width           =   5400
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   16
            Left            =   1140
            TabIndex        =   9
            Tag             =   "Tipo Mov|T|S|||sserie|codtipom||N|"
            Text            =   "Text3"
            Top             =   2475
            Width           =   615
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   10
            Left            =   3840
            MaxLength       =   10
            TabIndex        =   13
            Tag             =   "Fecha Venta|F|S|||sserie|fechavta|dd/mm/yyyy|N|"
            Text            =   "Text1"
            Top             =   2475
            Width           =   1215
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   6
            Left            =   120
            MaxLength       =   6
            TabIndex        =   5
            Tag             =   "Cod. Cliente|N|S|0|999999|sserie|codclien|000000|N|"
            Text            =   "Text1"
            Top             =   600
            Width           =   780
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   6
            Left            =   980
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   34
            Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwwwwwwwwww aq"
            Top             =   600
            Width           =   4260
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   7
            Left            =   740
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   33
            Text            =   "Text2"
            Top             =   1260
            Width           =   3405
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   7
            Left            =   120
            MaxLength       =   3
            TabIndex        =   6
            Tag             =   "Direccion/Dpto.|N|S|0|999|sserie|coddirec|000|N|"
            Text            =   "Text1"
            Top             =   1260
            Width           =   540
         End
         Begin VB.ComboBox cboTipomov 
            Height          =   315
            ItemData        =   "frmRepNumSerie.frx":00D3
            Left            =   1140
            List            =   "frmRepNumSerie.frx":00D5
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   2840
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   11
            Left            =   3840
            MaxLength       =   5
            TabIndex        =   14
            Tag             =   "Nº Linea Venta|N|S|0|99999|sserie|numline1||N|"
            Text            =   "Text1"
            Top             =   2835
            Width           =   735
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   8
            Left            =   3840
            MaxLength       =   10
            TabIndex        =   11
            Tag             =   "Nº Albaran Venta|N|S|0|9999999|sserie|numalbar|0000000|N|"
            Text            =   "Text1"
            Top             =   1755
            Width           =   1215
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   9
            Left            =   3840
            MaxLength       =   10
            TabIndex        =   12
            Tag             =   "Nº Factura Venta|N|S|0|9999999|sserie|numfactu|0000000|N|"
            Text            =   "Text1"
            Top             =   2115
            Width           =   1215
         End
         Begin VB.CheckBox chkTieneMan 
            Caption         =   "Tiene Mantenimiento"
            Enabled         =   0   'False
            Height          =   195
            Left            =   120
            TabIndex        =   7
            Tag             =   "¿Tiene Mantenimiento?|N|S|||sserie|tieneman||N|"
            Top             =   1750
            Width           =   1815
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   3
            Left            =   1140
            MaxLength       =   10
            TabIndex        =   8
            Tag             =   "Nº Mantenimiento|T|S|||sserie|nummante||N|"
            Text            =   "Text1 Text"
            Top             =   2115
            Width           =   1305
         End
         Begin VB.Label Label6 
            Caption         =   "Nº Factura"
            Height          =   255
            Left            =   2880
            TabIndex        =   42
            Top             =   2115
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Nº Albaran"
            Height          =   255
            Index           =   0
            Left            =   2880
            TabIndex        =   41
            Top             =   1755
            Width           =   855
         End
         Begin VB.Label Label11 
            Caption         =   "Fecha Vta"
            Height          =   255
            Left            =   2880
            TabIndex        =   40
            Top             =   2475
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Cliente"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   39
            Top             =   360
            Width           =   495
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   2
            Left            =   660
            ToolTipText     =   "Buscar cliente"
            Top             =   360
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Direc."
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   38
            Top             =   1020
            Width           =   495
         End
         Begin VB.Image imgBuscar 
            Enabled         =   0   'False
            Height          =   240
            Index           =   3
            Left            =   740
            ToolTipText     =   "Buscar direc./dpto"
            Top             =   1020
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo Movim."
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   37
            Top             =   2475
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Nº linea Vta"
            Height          =   255
            Index           =   3
            Left            =   2880
            TabIndex        =   36
            Top             =   2835
            Width           =   855
         End
         Begin VB.Label Label12 
            Caption         =   "Nº Mantenim."
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   2115
            Width           =   975
         End
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmRepNumSerie.frx":00D7
         Height          =   2625
         Left            =   -74840
         TabIndex        =   61
         Top             =   600
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   4630
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         BorderStyle     =   0
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
         Caption         =   "Tiene Mantenimiento"
         Height          =   255
         Index           =   15
         Left            =   -65800
         TabIndex        =   80
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Nº linea Vta"
         Height          =   255
         Index           =   14
         Left            =   -66195
         TabIndex        =   75
         Top             =   2880
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Vta"
         Height          =   255
         Index           =   13
         Left            =   -66195
         TabIndex        =   73
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Factura"
         Height          =   255
         Index           =   12
         Left            =   -66195
         TabIndex        =   71
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Albaran"
         Height          =   255
         Index           =   11
         Left            =   -66195
         TabIndex        =   69
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Movim."
         Height          =   255
         Index           =   10
         Left            =   -66195
         TabIndex        =   67
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Manteni."
         Height          =   255
         Index           =   7
         Left            =   -66195
         TabIndex        =   65
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1420
      Left            =   240
      TabIndex        =   23
      Top             =   440
      Width           =   11055
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   4
         Left            =   9045
         MaxLength       =   10
         TabIndex        =   3
         Tag             =   "Fecha Ult. Repar.|F|S|||sserie|ultrepar|dd/mm/yyyy|N|"
         Text            =   "Text1"
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   5
         Left            =   9045
         MaxLength       =   10
         TabIndex        =   4
         Tag             =   "Fecha Fin Garantia|F|S|||sserie|fingaran|dd/mm/yyyy|N|"
         Text            =   "Text1"
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   0
         Left            =   1350
         MaxLength       =   15
         TabIndex        =   0
         Tag             =   "Nº Serie|T|N|||sserie|numserie||S|"
         Text            =   "Text1"
         Top             =   240
         Width           =   1710
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   1
         Left            =   1350
         MaxLength       =   16
         TabIndex        =   1
         Tag             =   "Cod. Artículo|T|N|||sserie|codartic||S|"
         Text            =   "Text1"
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   1
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   25
         Text            =   "Text2"
         Top             =   600
         Width           =   3885
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   2
         Left            =   1350
         MaxLength       =   8
         TabIndex        =   2
         Tag             =   "Cod. Tipo Artículo|T|N|||sserie|codtipar||N|"
         Text            =   "Te"
         Top             =   960
         Width           =   1005
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   2
         Left            =   2400
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   24
         Text            =   "Text2"
         Top             =   960
         Width           =   3405
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   8760
         Picture         =   "frmRepNumSerie.frx":00EC
         ToolTipText     =   "Buscar fecha"
         Top             =   600
         Width           =   240
      End
      Begin VB.Label Label9 
         Caption         =   "Ult. Repar."
         Height          =   255
         Left            =   7800
         TabIndex        =   30
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Fin Garantia"
         Height          =   255
         Left            =   7800
         TabIndex        =   29
         Top             =   960
         Width           =   855
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   1
         Left            =   8760
         Picture         =   "frmRepNumSerie.frx":0177
         ToolTipText     =   "Buscar fecha"
         Top             =   980
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Nº Serie"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   28
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Artículo"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   600
         Width           =   735
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1065
         ToolTipText     =   "Buscar artículo"
         Top             =   600
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1065
         Picture         =   "frmRepNumSerie.frx":0202
         ToolTipText     =   "Buscar tipo artículo"
         Top             =   960
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Artíc."
         Height          =   255
         Index           =   9
         Left            =   240
         TabIndex        =   26
         Top             =   960
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   9000
      TabIndex        =   15
      Top             =   6000
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   10200
      TabIndex        =   16
      Top             =   6000
      Width           =   1035
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   10155
      TabIndex        =   17
      Top             =   6015
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Index           =   0
      Left            =   240
      TabIndex        =   21
      Top             =   5835
      Width           =   2535
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   240
         Left            =   240
         TabIndex        =   22
         Top             =   180
         Width           =   2115
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   11430
      _ExtentX        =   20161
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
            Object.ToolTipText     =   "Sustituir"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Recuperar Nº Serie"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Componentes"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
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
         Left            =   9000
         TabIndex        =   20
         Top             =   0
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   3600
      Top             =   6000
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   5640
      Top             =   6000
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
      TabIndex        =   19
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
      Begin VB.Menu mnSustituir 
         Caption         =   "S&ustituir"
         Shortcut        =   ^U
      End
      Begin VB.Menu mnComponentes 
         Caption         =   "&Componentes"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnImprimir 
         Caption         =   "&Imprimir"
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
Attribute VB_Name = "frmRepNumSerie2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)
Public DatoAInsertar As String

Private WithEvents frmB As frmBuscaGrid 'Form para busquedas (frmBuscaGrid)
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmF As frmCal 'Calendario de Fechas
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmTA As frmAlmTipoArticulo  'Form Mantenimiento Tipo Articulo
Attribute frmTA.VB_VarHelpID = -1
Private WithEvents frmA As frmAlmArticu2  'Form Mantenimiento Articulos
Attribute frmA.VB_VarHelpID = -1
Private WithEvents frmCli As frmFacClientes 'Form Mantenimiento Clientes
Attribute frmCli.VB_VarHelpID = -1
Private WithEvents frmProv As frmComProveedores 'Form Mantenimiento Proveedores
Attribute frmProv.VB_VarHelpID = -1

Private HaDevueltoDatos As Boolean
Private Modo As Byte
Private ModoAnterior As Byte


Dim primeravez As Boolean

Dim NombreTabla As String
Dim Ordenacion As String
Dim kCampo As Integer

Dim btnPrimero As Byte
'Variable que indica el número del Boton  PrimerRegistro en la Toolbar1

Dim EsCabecera As Boolean
'Para saber en MandaBusquedaPrevia si busca en la tabla sserie o en la tabla sdirec

Dim CadenaConsulta As String



Private Sub cboMotivoBaja_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkTieneMan_Click()
    If Modo = 3 Or Modo = 4 Then
        BloquearTxt Text1(3), Not CBool(Me.chkTieneMan.Value)
    End If
End Sub

Private Sub chkTieneMan_KeyDown(KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub


Private Sub chkTieneMan_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub chkVistaPrevia_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub cmdAceptar_Click()
    On Error GoTo Error1
    
    Screen.MousePointer = vbHourglass
    
    Select Case Modo
        Case 1 'BUSCAR
            HacerBusqueda
            
        Case 3 'INSERTAR
            
            If DatosOk Then
                If InsertarDesdeForm(Me) Then
                    
                    CadenaConsulta = "numserie=" & DBSet(Text1(0).Text, "T") & "  AND codartic=" & DBSet(Text1(1).Text, "T") & ""
                    CadenaConsulta = "Select * from " & NombreTabla & " WHERE " & CadenaConsulta
                    Data1.RecordSource = CadenaConsulta
                    PosicionarData
                End If
            End If
            
        Case 4 'MODIFICAR
            If DatosOk Then
                 If ModificaDesdeFormulario(Me, 1) Then
                     TerminaBloquear
                     PosicionarData
                 End If
             End If
    End Select
    
    Screen.MousePointer = vbDefault
    Exit Sub
    
Error1:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "", Err.Description
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
                PonerCampos
                PonerModo 2
                
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


Private Sub cmdRegresar_Click()
Dim cad As String

    If Data1.Recordset.EOF Then
        MsgBox "Ningún registro devuelto.", vbExclamation
        Exit Sub
    End If

    cad = Data1.Recordset.Fields(0) & "|" 'num serie
    cad = cad & Data1.Recordset.Fields(1) & "|" 'cod artic
    cad = cad & Text2(1).Text & "|"  'nom artic
    cad = cad & Data1.Recordset.Fields(3) & "|" 'cod cliente
    RaiseEvent DatoSeleccionado(cad)
    Unload Me
End Sub




Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

    If Me.adodc1.Recordset.EOF Then Exit Sub
    
'    If Modo = 2 Then
        Me.chkAux.Value = DBLet(Me.adodc1.Recordset!TieneMan, "N")
        txtAux2(2).Text = DBLet(Me.adodc1.Recordset!nummante, "T")
        txtAux2(3).Text = DBLet(Me.adodc1.Recordset!codtipom, "T")
        
        txtAux2(4).Text = DBLet(Me.adodc1.Recordset!Numalbar, "T")
        txtAux2(5).Text = DBLet(Me.adodc1.Recordset!Numfactu, "T")
        txtAux2(6).Text = DBLet(Me.adodc1.Recordset!FechaVta, "F")
        txtAux2(7).Text = DBLet(Me.adodc1.Recordset!numline1, "T")
'    End If
End Sub

Private Sub Form_activate()
    Screen.MousePointer = vbDefault
    If primeravez Then
        primeravez = False
        If Me.DatoAInsertar <> "" Then
            BotonAnyadir
            Text1(0).Text = DatoAInsertar
        End If
    End If
End Sub


Private Sub Form_Load()

    primeravez = True
    
    'Icono del formulario
    Me.Icon = frmPpal.Icon
    
    'Icono de busqueda
    For kCampo = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(kCampo).Picture = frmPpal.imgListComun.ListImages(19).Picture
    Next kCampo

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
        
        .Buttons(10).Image = 42 'Sustitucion de num serie
        .Buttons(11).Image = 21 'Recuperar num serie
        .Buttons(12).Image = 34 'Componentes
        .Buttons(13).Image = 16 'Imprimir
        .Buttons(16).Image = 15 'Salir
        
        .Buttons(btnPrimero).Image = 6 'Primero
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
        .Buttons(btnPrimero + 3).Image = 9 'Ultimo
    End With
    
    
    LimpiarCampos   'Limpia los campos TextBox
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    'situarnos en el primer tab
    Me.SSTab1.Tab = 0
    'siempre bloqueardos campos fora grid
    For kCampo = 0 To Me.txtAux2.Count - 1
        BloquearTxt txtAux2(kCampo), True
    Next kCampo
    Me.chkAux.Enabled = False
    
    
    '-- cargar combos
    CargarCombo_Tabla Me.cboMotivoBaja, "smotba", "codmotiv", "desmotiv", , True
    
    '-- cargar el Data
    NombreTabla = "sserie" 'Tabla Numero de Serie
    Ordenacion = " ORDER BY codartic, numserie "
    CadenaConsulta = "Select * from " & NombreTabla & " WHERE numserie = -1" 'No recupera datos
    Data1.ConnectionString = conn
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    
    CargaGrid False

    If DatosADevolverBusqueda = "" Then
        PonerModo 0
    Else
        If DatoAInsertar = "" Then
            PonerModo 1
            Text1(0).BackColor = vbYellow
        End If
    End If
    
    If vParamAplic.HayDeparNuevo > 0 Then Label1(2).Caption = "Dpto."
    
    
    Screen.MousePointer = vbDefault
End Sub



Private Sub frmA_DatoSeleccionado(CadenaSeleccion As String)
'Formulario Mantenimiento Articulos
    Text1(1).Text = RecuperaValor(CadenaSeleccion, 1)
    Text2(1).Text = RecuperaValor(CadenaSeleccion, 2)
    'Tipo Articulos
    Text1(2).Text = RecuperaValor(CadenaSeleccion, 3)
    Text2(2).Text = RecuperaValor(CadenaSeleccion, 4)
End Sub


Private Sub frmB_Selecionado(CadenaDevuelta As String)
'Formulario para Busqueda
Dim CadB As String
Dim Aux As String
      
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        If EsCabecera Then 'Llama desde VerTodos del Form
                'Estamos en Cabecera
                'Recupera todo el registro de Nº Serie
                'Sabemos que campos son los que nos devuelve
                'Creamos una cadena consulta y ponemos los datos
                CadB = ""
                '                       El primero es un pipe
                If Mid(CadenaDevuelta, 1, 1) = "|" Then CadenaDevuelta = """""" & CadenaDevuelta
                Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
                CadB = Aux
                Aux = ValorDevueltoFormGrid(Text1(1), CadenaDevuelta, 2)
                CadB = CadB & " and " & Aux
                CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
                PonerCadenaBusqueda
        Else  'Llama desde Prismatico Direcciones/Departamentos
                Text1(7).Text = Format(RecuperaValor(CadenaDevuelta, 1), "000")
                Text2(7).Text = RecuperaValor(CadenaDevuelta, 2)
                
                'Pongo QU NOOOOO ha devuelto datos. Asi no hace el regresar
                HaDevueltoDatos = False
                
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub frmCli_DatoSeleccionado(CadenaSeleccion As String)
'Formulario Mantenimiento Clientes
    Text1(6).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    Text2(6).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub frmF_Selec(vFecha As Date)
'Calendario de Fecha
Dim Indice As Byte
    Indice = Val(Me.imgFecha(0).Tag)
    Text1(Indice).Text = Format(vFecha, "dd/mm/yyyy")
End Sub


Private Sub frmProv_DatoSeleccionado(CadenaSeleccion As String)
'Formulario Mantenimiento Proveedores
    Text1(12).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    Text2(12).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmTA_DatoSeleccionado(CadenaSeleccion As String)
'Formulario Mantenimiento Tipo Articulo
    Text1(2).Text = RecuperaValor(CadenaSeleccion, 1)
    Text2(2).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub imgBuscar_Click(index As Integer)
Dim Indice As Byte

    If Modo = 2 Or Modo = 0 Then Exit Sub
 
    Screen.MousePointer = vbHourglass

    Select Case index
        Case 0 'Codigo Articulo
            Set frmA = New frmAlmArticu2
            'frmA.DatosADevolverBusqueda3 = "@1@" 'Abrir en Modo busqueda
            frmA.DesdeTPV = False
            frmA.Show vbModal
            Set frmA = Nothing
            Indice = 1
        Case 1  'Cod. Tipo Articulo
            Set frmTA = New frmAlmTipoArticulo
            frmTA.DatosADevolverBusqueda = "0"
            frmTA.Show vbModal
            Set frmTA = Nothing
            Indice = 2
        Case 2 'Cod. Cliente
            Set frmCli = New frmFacClientes
            frmCli.DatosADevolverBusqueda = "0"
            frmCli.Show vbModal
            Set frmCli = Nothing
            Indice = 6
        Case 3 'Direc/Dpto del Cliente
             'Mostrar las Direc. o Dptos del cliente seleccionado
             If Trim(Text1(6).Text) = "" Then
                MsgBox "Debe seleccionar un cliente para mostrar sus Direc./Dpto.", vbInformation
                Screen.MousePointer = vbDefault
                Exit Sub
             Else
                EsCabecera = False
                MandaBusquedaPrevia " codclien= " & Val(Text1(6).Text)
                Indice = 7
             End If
        Case 4 'Cod. Proveedor
            Set frmProv = New frmComProveedores
            frmProv.DatosADevolverBusqueda = "0"
            frmProv.Show vbModal
            Set frmProv = Nothing
            Indice = 12
    End Select
    PonerFoco Text1(Indice)
    Screen.MousePointer = vbDefault
End Sub


Private Sub imgFecha_Click(index As Integer)
Dim Indice As Byte

   If Modo = 2 Or Modo = 0 Then Exit Sub
   Screen.MousePointer = vbHourglass
      
   Set frmF = New frmCal
   frmF.Fecha = Now
   
   Select Case index
        Case 0: Indice = 4 'Fecha ult. compra
        Case 1: Indice = 5 'Fecha fin garantia
        Case 2: Indice = 18 'fecha baja equipo
   End Select
   imgFecha(0).Tag = Indice

   PonerFormatoFecha Text1(Indice)
   If Text1(Indice).Text <> "" Then frmF.Fecha = CDate(Text1(Indice).Text)

   Screen.MousePointer = vbDefault
   frmF.Show vbModal
   Set frmF = Nothing
   PonerFoco Text1(Indice)
End Sub


Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnComponentes_Click()
'Mostrar equipos que tiene un cliente, un dpto, un mantenimiento,...
    BotonComponentes
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnImprimir_Click()
     AbrirListado (60) '60: Informe Nº Serie
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

Private Sub mnSustituir_Click()
'Sustituir un Nº de Serie en garantia por otro
    BotonSustituir
End Sub




Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub

Private Sub Text1_GotFocus(index As Integer)
    kCampo = index
    ConseguirFoco Text1(index), Modo
End Sub


Private Sub Text1_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
End Sub

Private Sub Text1_KeyPress(index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub Text1_LostFocus(index As Integer)
Dim devuelve As String

    If Not PerderFocoGnral(Text1(index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub



    Select Case index
        Case 1 'Codigo Articulo
            If Text1(index).Text <> "" Then
                Text2(index).Text = PonerNombreDeCod(Text1(index), conAri, "sartic", "nomartic")
                devuelve = "nseriesn"
                Text1(index + 1).Text = DevuelveDesdeBDNew(conAri, "sartic", "codtipar", "codartic", Text1(index).Text, "T", devuelve)
                If devuelve = "1" Then
                    Text2(index + 1).Text = DevuelveDesdeBDNew(conAri, "stipar", "nomtipar", "codtipar", Text1(index + 1).Text, "T")
                Else
                    Text2(index + 1).Text = ""
                    Text1(index + 1).Text = ""
                    Text2(index).Text = ""
                    MsgBox "El artículo no tiene control de nº de serie.", vbInformation
                    PonerFoco Text1(index)
                End If
            Else
                Text2(index).Text = ""
            End If
            
        Case 2 'Codigo Tipo de Articulo
            Text2(index).Text = PonerNombreDeCod(Text1(index), conAri, "stipar", "nomtipar")
            Text1(index).Text = DevuelveDesdeBD(conAri, "codtipar", "stipar", "codtipar", Text1(index).Text, "T")
            
        Case 6 'Cliente
            
            Text2(index).Text = PonerNombreDeCod(Text1(index), conAri, "sclien", "nomclien")
            
        Case 7 'Direc/dpto del cliente
            If Text1(index).Text = "" Then
                Text2(index).Text = ""
                Exit Sub
            End If
            Text1(index).Text = Format(Text1(index).Text, "000")
            
            'Comprobar que el cliente seleccionado tiene esa direccion o dpto
            devuelve = DevuelveDesdeBDNew(conAri, "sdirec", "nomdirec", "codclien", Text1(6).Text, "N", , "coddirec", Text1(7).Text, "N")
            Text2(index).Text = devuelve 'Nombre direc. o dpto
            If devuelve = "" Then 'No existe el dpto
                
                devuelve = DevuelveTextoDepto(False)
                devuelve = "No existe" & devuelve & Text1(index).Text & " para el cliente: "
                devuelve = devuelve & Text1(6).Text & " - " & Text2(6).Text
                MsgBox devuelve, vbInformation
                PonerFoco Text1(index)
            End If
            
        Case 12 'Proveedor
            Text2(index).Text = PonerNombreDeCod(Text1(index), conAri, "sprove", "nomprove")
            
        Case 4, 5, 10, 14 'Fechas ult. modif., fin garantia
            If Text1(index).Text <> "" And Text1(index).Locked = False Then PonerFormatoFecha Text1(index)
            
            
        Case 18 'fecha de baja
            PonerFormatoFecha Me.Text1(18)
            If Me.Text1(18).Text = "" Then
                Me.cboMotivoBaja.ListIndex = -1
            End If
    End Select
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.index
        Case 1: mnBuscar_Click 'Busqueda
        Case 2: mnVerTodos_Click 'Ver Todos
            
        Case 5: mnNuevo_Click 'Nuevo
        Case 6: mnModificar_Click  'Modificar
        Case 7: mnEliminar_Click 'Eliminar
            
        Case 10: mnSustituir_Click 'Sustitucion num serie
        Case 11: BotonRecuperar 'Recuperar nº serie
        Case 12: mnComponentes_Click 'Componentes
            
        Case 13: mnImprimir_Click 'Imprimir
        Case 16: mnSalir_Click  'Salir
             
        Case btnPrimero To btnPrimero + 3 'Flechas de Desplazamiento
            Desplazamiento (Button.index - btnPrimero)
    End Select
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim Cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, Cerrar
    If Cerrar Then Unload Me
    If KeyAscii = 27 And Modo = 1 Then cmdCancelar_Click 'busqueda
End Sub


Private Sub PonerModo(Kmodo As Byte)
Dim I As Byte
Dim b As Boolean
Dim NumReg As Byte

    Modo = Kmodo
    PonerIndicador lblIndicador, Modo
    
    'Visualizar flechas de desplazamiento en la toolbar si modo=2
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero, (Modo = 2), NumReg
        
    b = (Modo = 2)
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.visible = b
    Else
        cmdRegresar.visible = False
    End If
        
    '-------------------------------------------
    'Bloquear Registros
    BloquearText1 Me, Modo
    
    'Los Datos de Albaran de Compras y Ventas siempre bloqueados
    'se actualizan por codigo de programa al insertar las lineas de Albaran
    Me.cboTipomov.Enabled = False
    
            
    'Modo INSERTAR
    b = (Modo = 3) Or (Modo = 4)
    If Modo = 3 Then Me.chkTieneMan.Value = 1
    Me.chkTieneMan.Enabled = b 'Insertar o Modificar
    If b Then BloquearTxt Text1(3), Not CBool(Me.chkTieneMan.Value)
    Me.cboTipomov.Enabled = False 'Insertar o Modificar

    
    '## LAURA 19/06/2008
    '   añadir datos de baja
    BloquearCmb Me.cboMotivoBaja, Not ((Modo = 1) Or (Modo = 3) Or (Modo = 4))
    '##
    
    '------------------------------------
    b = Modo <> 0 And Modo <> 2
    cmdCancelar.visible = b
    cmdAceptar.visible = b
    
    For I = 0 To Me.imgBuscar.Count - 1
'        Me.imgBuscar(i).Enabled = b
        BloquearImg Me.imgBuscar(I), Not b
    Next I
    
    For I = 0 To Me.imgFecha.Count - 1
        Me.imgFecha(I).Enabled = b 'Si es insertar o modificar
    Next I
    
    'Si Modificar y se ha insertado un nº Albaran no modificar datos
    'del proveedor
    If Trim(Text1(13).Text) <> "" Then
        BloquearTxt Text1(12), True
        Me.imgBuscar(4).Enabled = False
    End If
    
    'Poner el tamaño de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos
    
    PonerModoOpcionesMenu   'Activar opciones de menu según Modo
    PonerOpcionesMenu   'Activar opciones de menu según nivel
                        'de permisos del usuario
End Sub


Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
    PonerLongCamposGnral Me, Modo, 1
End Sub


Private Sub PonerModoOpcionesMenu()
Dim b As Boolean

    'Modo 2. Hay datos y estamos visualizandolos
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

    'Sustituir
    Toolbar1.Buttons(10).Enabled = b
    Me.mnSustituir.Enabled = b
    
    'recuperar nº serie
    Toolbar1.Buttons(11).Enabled = b And Text1(6).Text <> ""

    'Componentes
    Toolbar1.Buttons(12).Enabled = b
    Me.mnComponentes.Enabled = b
    

    '-------------------------------------
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
    Me.cboMotivoBaja.ListIndex = -1
    '### a mano
    Me.chkTieneMan.Value = 0
    
    CargaGrid False
End Sub


Private Sub Desplazamiento(index As Integer)
'Botones de Desplazamiento de la Toolbar
'Para desplazarse por los registros de control Data
    DesplazamientoData Data1, index
    PonerCampos
    
End Sub


Private Sub BotonBuscar()
    'Buscar
    If Modo <> 1 Then
        LimpiarCampos
        Me.SSTab1.Tab = 0
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
'Ver todos
    LimpiarCampos
    If chkVistaPrevia.Value = 1 Then
        EsCabecera = True
        MandaBusquedaPrevia ""
    Else
        CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Sub BotonAnyadir()

    LimpiarCampos 'Vacía los TextBox
    
    Me.SSTab1.Tab = 0
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    ModoAnterior = Modo 'Para el botón Cancelar en Modo Insertar
    PonerModo 3
    If Me.DatoAInsertar = "" Then
        PonerFoco Text1(0)
    Else
        PonerFoco Text1(1)
    End If
End Sub


Private Sub BotonModificar()
    'Escondemos el navegador y ponemos Modo Modificar
    PonerModo 4
    
    'Como el campo1 y campo2 es clave primaria, NO se puede modificar
    BloquearTxt Text1(0), True
    BloquearTxt Text1(1), True
    Me.imgBuscar(0).Enabled = False
    PonerFoco Text1(2)
End Sub


Private Sub BotonEliminar()
Dim SQL As String

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    'Comprobamos si se puede eliminar
    If Not SePuedeEliminar Then Exit Sub
    
    SQL = ""
    SQL = SQL & "Va a Eliminar el Nº Serie del Articulo: " & vbCrLf
    SQL = SQL & vbCrLf & "Nº Serie: " & Text1(0).Text
    SQL = SQL & vbCrLf & "Artic. : " & Text1(1).Text & " - " & Text2(1).Text
    
    SQL = SQL & vbCrLf & vbCrLf & "¿Desea continuar? "
    If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        On Error GoTo Error2
        NumRegElim = Data1.Recordset.AbsolutePosition
        If Not Eliminar Then Exit Sub
        If SituarDataTrasEliminar(Data1, NumRegElim) Then
            PonerCampos
        Else
            LimpiarCampos
            PonerModo 0
        End If
    End If
    
Error2:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar Nº Serie", Err.Description
        Data1.Recordset.CancelUpdate
    End If
End Sub


Private Function Eliminar() As Boolean
Dim SQL As String

    On Error GoTo FinEliminar

      SQL = " WHERE numserie=" & DBSet(Data1.Recordset!numSerie, "T")
      SQL = SQL & " AND codartic = " & DBSet(Data1.Recordset!codArtic, "T")
    
      conn.Execute "Delete  from " & NombreTabla & SQL
                      
FinEliminar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar", Err.Description
        Eliminar = False
    Else
        Eliminar = True
    End If
End Function


Private Function DatosOk() As Boolean
Dim b As Boolean

    b = CompForm(Me, 1)
    If Not b Then Exit Function
 
    'Comprobar que se introduce valor en fecha fin garantia
    If Text1(5).Text = "" Then
        MsgBox "El valor de fecha fin garantia no puede ser nulo.", vbInformation
        b = False
    End If
    
    '## LAURA 19/06/2008
    '- comprobar q si la fecha baja tiene valor el motivo de baja tambien
    '  y viceversa.
    If Me.Text1(18).Text = "" Then
        Me.cboMotivoBaja.ListIndex = -1
    ElseIf Trim(cboMotivoBaja.List(cboMotivoBaja.ListIndex)) = "" Then
        MsgBox "Debe seleccionar un motivo de baja si hay valor en la fecha de baja.", vbInformation
        b = False
    End If
    '##
    
    DatosOk = b
End Function



Private Sub MandaBusquedaPrevia(CadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim cad As String
Dim tabla As String
Dim Titulo As String, Desc As String
Dim selElem As Byte

    'Llamamos a al form
    cad = ""
    If EsCabecera Then
    'Estamos en Modo de Cabeceras
    'Registro de la tabla de cabeceras: sserie
        cad = cad & ParaGrid(Text1(0), 15, "Nº Serie")
        cad = cad & ParaGrid(Text1(1), 20, "Artic.")
        cad = cad & "Desc. Artic.|sartic|nomartic|T||38·"
        cad = cad & ParaGrid(Text1(2), 6, "TArt.")
        cad = cad & "Desc. Tipo|stipar|nomtipar|T||20·"
    
        tabla = "(" & NombreTabla & " LEFT JOIN sartic ON " & NombreTabla & ".codartic=sartic.codartic" & ")"
        tabla = tabla & " LEFT JOIN stipar ON " & NombreTabla & ".codtipar=stipar.codtipar"
    
        Titulo = "Nº Serie"
        selElem = 2
   Else
        If vParamAplic.HayDeparNuevo = 1 Then
            Titulo = "Dptos Cliente: "
            Desc = "Dpto."
        ElseIf vParamAplic.HayDeparNuevo = 0 Then
            Titulo = "Direc. Cliente: "
            Desc = "Direc."
        Else
            Titulo = "obra Cliente: "
            Desc = "Obra"
        End If
        Titulo = Titulo & Text1(6).Text & " - " & Text2(6).Text 'Cod y Desc. Cliente
        cad = cad & "Cod. " & Desc & "|sdirec|coddirec|N||20·"
        cad = cad & "Desc. " & Desc & "|sdirec|nomdirec|T||40·"
        tabla = "sdirec"
        selElem = 1
    End If
           
    If cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = cad
        frmB.vTabla = tabla
        frmB.vSQL = CadB
        HaDevueltoDatos = False
        '###A mano
        frmB.vDevuelve = "0|1|"
        frmB.vTitulo = Titulo
        frmB.vselElem = selElem
        frmB.vConexionGrid = conAri 'Conexion a BD Ariges
'        frmB.vBuscaPrevia = chkVistaPrevia
        '#
        If Not EsCabecera Then frmB.Label1.FontSize = 11
        frmB.Show vbModal
        Set frmB = Nothing
        'Si ha puesto valores y tenemos que es formulario de busqueda entonces
        'tendremos que cerrar el form lanzando el evento
        If HaDevueltoDatos Then
            If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
                cmdRegresar_Click
'        Else   'de ha devuelto datos, es decir NO ha devuelto datos
'            If Modo = 5 Then
'                PonerFoco txtAux(0)
'            Else
                'Esto esta mal
                'Si hace cmdregresar, ahi hay un UNLOAD
                'con lo cual NO podemos poner foco, pq volvera a hacer un LOAD
                'PonerFoco Text1(kCampo)
'            End If
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub HacerBusqueda()
Dim CadB As String

    CadB = ObtenerBusqueda(Me, False)
    EsCabecera = True
    If chkVistaPrevia = 1 Then
        MandaBusquedaPrevia CadB
    ElseIf CadB <> "" Then 'Se muestran en el mismo form
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Sub PonerCadenaBusqueda()
Screen.MousePointer = vbHourglass

    On Error GoTo EEPonerBusq

    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    
    If Data1.Recordset.RecordCount <= 0 Then
        If Modo = 1 Then
            MsgBox "No hay ningún registro en la tabla " & NombreTabla & " para ese criterio de Búsqueda.", vbInformation
        Else
            MsgBox "No hay ningún registro en la tabla " & NombreTabla & ".", vbInformation
        End If
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
    Toolbar1.Buttons(11).Enabled = False
    If Data1.Recordset.EOF Then Exit Sub



     'Si se el campo numsersu tiene valor mostrar el frame de sustitucion
    Me.FrameSusti.visible = DBLet(Data1.Recordset!numsersu, "T") <> ""

    PonerCamposForma Me, Data1

    'Poner el nombre del cod. Articulo
    Text2(1).Text = PonerNombreDeCod(Text1(1), conAri, "sartic", "nomartic")
    'Poner el nombre del cod. Tipo Articulo
    Text2(2).Text = PonerNombreDeCod(Text1(2), conAri, "stipar", "nomtipar")
    'Poner el nombre del cod. Cliente
    Text2(6).Text = PonerNombreDeCod(Text1(6), conAri, "sclien", "nomclien")
    'Poner el nombre del cod. Direc./Dpto
    Text2(7).Text = DevuelveDesdeBDNew(conAri, "sdirec", "nomdirec", "codclien", Text1(6).Text, "N", , "coddirec", Text1(7).Text, "N")
    'Poner el nombre del cod. Proveedor
    Text2(12).Text = PonerNombreDeCod(Text1(12), conAri, "sprove", "nomprove")
    If Trim(Text1(13).Text) <> "" Then BloquearTxt Text1(12), True
    
    If IsNull(Data1.Recordset!codmotba) Then
        Me.cboMotivoBaja.ListIndex = -1
    Else
        PosicionarCombo Me.cboMotivoBaja, Data1.Recordset!codmotba
    End If
    
    '-- cargar las lineas de venta nº serie
    CargaGrid True
    
    Toolbar1.Buttons(11).Enabled = (Modo = 2) And Trim(Text1(6).Text) <> ""
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    Exit Sub
    
EPonerCampos:
    MuestraError Err.Number, "Poniendo Campos", Err.Description
End Sub


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub


Private Sub PosicionarData()
Dim Indicador As String
Dim vWhere As String

    vWhere = "(numserie=" & DBSet(Text1(0).Text, "T") & "  AND codartic=" & DBSet(Text1(1).Text, "T") & ")"
    If SituarDataMULTI(Data1, vWhere, Indicador) Then
        PonerModo 2
        lblIndicador.Caption = Indicador
    Else
'        LimpiarCampos
        PonerModo 0
    End If
End Sub



Private Function SePuedeEliminar() As Boolean

    If Text1(8).Text <> "" Then
        MsgBox "El nº de serie esta asignado a un albaran de venta y no se puede eliminar.", vbInformation
        SePuedeEliminar = False
    Else
        SePuedeEliminar = True
    End If
    
End Function


Private Sub BotonComponentes()
'Muestra un form de Mensaje para seleccionar el tipo de resumen que queremos mostrar:
'Por Mantenimiento, Por Departamento, Por Cliente
Dim vWhere As String

    If Text1(6).Text = "" Then
        MsgBox "No hay Cliente para mostrar Resumen.", vbInformation
        Exit Sub
    End If
    vWhere = " WHERE codclien = " & Text1(6).Text
    frmMensajes.cadWhere = vWhere
    'vCampos= Mantenimiento|coddirec|Desc. coddirec| cadCliente
    vWhere = Text1(6).Text & " - " & Text2(6).Text
    frmMensajes.vCampos = Text1(3).Text & "|" & Text1(7).Text & "|" & Text2(7).Text & "|" & vWhere & "|"
    frmMensajes.OpcionMensaje = 5 'Componentes
    frmMensajes.Show vbModal
End Sub



Private Sub BotonSustituir()
'Muestra un form para pedir el nuevo numero de serie que sustituye al seleccionado

    If Text1(0).Text = "" Then
        MsgBox "No hay un nº de serie seleccionado.", vbInformation
        Exit Sub
    End If
    
    'pedir en un form el nº de serie nuevo
    frmListado.NumCod = Trim(Text1(0).Text)
    frmListado.CadTag = Trim(Text1(1).Text)
    frmListado.OpcionListado = 407
    frmListado.Show vbModal
    
    PosicionarData
    PonerCampos
End Sub


Private Sub BotonRecuperar()
'Recuperar un nº de serie para asignar a otro cliente y pasar datos antiguos a las líneas
Dim cadFecha As String
Dim oNSerie As CNumSerie

    If Text1(0).Text = "" Then
        MsgBox "No hay un nº de serie seleccionado.", vbInformation
        Exit Sub
    End If
    
    '- pedir la fecha de recuperacion
    cadFecha = InputBox("Introduzca la fecha recuperación Nº Serie: ", "Nº Serie", Format(Now, "dd/mm/yyyy"))
    If cadFecha = "" Then
        MsgBox "Debe introducir una fecha para recuperar el nº serie.", vbInformation
        Exit Sub
    End If
    
    '- comprobar q la fecha es correcta
    If Not EsFechaOK(cadFecha) Then
        MsgBox "La fecha introducida no es válida.", vbInformation
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    '- pasar los datos de venta del cliente a las líneas
    '- limpiar los datos de venta del cliente de la cabecera para poder volver a ser asignado
    Set oNSerie = New CNumSerie
    If oNSerie.LeerDatos(Text1(0).Text, Text1(1).Text) Then
        If oNSerie.RecuperarParaVenta(cadFecha) Then
            PosicionarData
            PonerCampos
        End If
    End If
    Set oNSerie = Nothing
    
    Screen.MousePointer = vbDefault
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
Dim tabla As String
    
    tabla = "sserlin"

    SQL = "SELECT numserie,codartic,numlinea,fecharec,s.codclien,c.nomclien,s.coddirec,d.nomdirec,tieneman,nummante,codtipom,numfactu,fechavta,numalbar,numline1"
    SQL = SQL & " FROM (" & tabla & " s INNER JOIN sclien c ON s.codclien=c.codclien)"
    SQL = SQL & " LEFT OUTER JOIN sdirec d ON s.codclien=d.codclien and s.coddirec=d.coddirec"
    If enlaza Then
        SQL = SQL & " WHERE s.numserie=" & DBSet(Data1.Recordset!numSerie, "T") & " AND s.codartic=" & DBSet(Data1.Recordset!codArtic, "T")
    Else
        SQL = SQL & " WHERE s.numserie = '-1' and s.codartic='-1'"
    End If
    SQL = SQL & " ORDER BY s.fecharec desc"
    MontaSQLCarga = SQL
End Function



Private Sub CargaGrid(enlaza As Boolean)
Dim SQL As String
Dim tots As String

    On Error GoTo ErrCarga

'    b = DataGrid1.Enabled
    
    SQL = MontaSQLCarga(enlaza)
    CargaGridGnral DataGrid1, Me.adodc1, SQL, primeravez
    
    tots = "N||||0|;N||||0|;N||||0|;"
    SQL = DevuelveTextoDepto(True)
    tots = tots & "S|txtAux(2)|T|Fecha|990|;S|txtAux(0)|T|Cliente|700|;S|txtAux2(0)|T|Nombre Cliente|3140|;"
    tots = tots & "S|txtAux(1)|T|" & SQL & "|540|;S|txtAux2(1)|T|Nombre " & SQL & "|2460|;"
    tots = tots & "N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;"
       
    arregla tots, DataGrid1, Me
    
    Me.DataGrid1.Columns(4).NumberFormat = "000000"
    Me.DataGrid1.Columns(6).NumberFormat = "000"
    
'    DataGrid1.Enabled = b

    DataGrid1.ScrollBars = dbgAutomatic
    Exit Sub
    
ErrCarga:
    MuestraError Err.Number, "Cargando datos grid: " & DataGrid1.Tag, Err.Description
End Sub

