VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRepNumSerie2GR 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Numeros de Serie"
   ClientHeight    =   10680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15330
   ClipControls    =   0   'False
   Icon            =   "frmRepNumSerieGR.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10680
   ScaleWidth      =   15330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   225
      TabIndex        =   85
      Top             =   135
      Width           =   3585
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   180
         TabIndex        =   86
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
   Begin VB.Frame FrameBotonGnral2 
      Height          =   705
      Left            =   3915
      TabIndex        =   83
      Top             =   135
      Width           =   2055
      Begin MSComctlLib.Toolbar Toolbar5 
         Height          =   330
         Left            =   210
         TabIndex        =   84
         Top             =   180
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Sustituir"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Recuperar Nro.Serie"
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   1
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Componentes"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FrameDesplazamiento 
      Height          =   705
      Left            =   6030
      TabIndex        =   81
      Top             =   135
      Width           =   2415
      Begin MSComctlLib.Toolbar ToolbarDes 
         Height          =   330
         Left            =   240
         TabIndex        =   82
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
      Height          =   315
      Left            =   13365
      TabIndex        =   80
      Top             =   180
      Width           =   1755
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7350
      Left            =   240
      TabIndex        =   29
      Top             =   2475
      Width           =   14970
      _ExtentX        =   26405
      _ExtentY        =   12965
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
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
      TabCaption(0)   =   "Compra / Venta / Histórico"
      TabPicture(0)   =   "frmRepNumSerieGR.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameActuales"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "FrameNuevos"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "FrameBaja"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "FrameSusti"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame3"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      Begin VB.Frame Frame3 
         Caption         =   "Histórico"
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
         Height          =   3705
         Left            =   270
         TabIndex        =   59
         Top             =   3435
         Width           =   14460
         Begin VB.CheckBox chkAux 
            Enabled         =   0   'False
            Height          =   195
            Left            =   11340
            TabIndex        =   79
            Top             =   360
            Width           =   255
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
            Height          =   320
            Index           =   0
            Left            =   920
            MaxLength       =   6
            TabIndex        =   70
            Text            =   "codcli"
            Top             =   2400
            Visible         =   0   'False
            Width           =   615
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
            Height          =   320
            Index           =   1
            Left            =   3200
            MaxLength       =   6
            TabIndex        =   69
            Text            =   "coddpt"
            Top             =   2400
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox txtAux2 
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
            Left            =   1760
            Locked          =   -1  'True
            TabIndex        =   68
            Text            =   "nomclien"
            Top             =   2400
            Visible         =   0   'False
            Width           =   1245
         End
         Begin VB.TextBox txtAux2 
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
            Left            =   12705
            MaxLength       =   10
            TabIndex        =   67
            Top             =   675
            Width           =   1335
         End
         Begin VB.TextBox txtAux2 
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
            Left            =   12705
            TabIndex        =   66
            Top             =   1155
            Width           =   885
         End
         Begin VB.TextBox txtAux2 
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
            Left            =   12705
            MaxLength       =   10
            TabIndex        =   65
            Top             =   1515
            Width           =   1245
         End
         Begin VB.TextBox txtAux2 
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
            Index           =   5
            Left            =   12705
            MaxLength       =   10
            TabIndex        =   64
            Top             =   1875
            Width           =   1245
         End
         Begin VB.TextBox txtAux2 
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
            Left            =   12705
            MaxLength       =   10
            TabIndex        =   63
            Top             =   2235
            Width           =   1365
         End
         Begin VB.TextBox txtAux2 
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
            Left            =   12705
            MaxLength       =   5
            TabIndex        =   62
            Top             =   2595
            Width           =   885
         End
         Begin VB.TextBox txtAux2 
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
            Index           =   1
            Left            =   4040
            Locked          =   -1  'True
            TabIndex        =   61
            Text            =   "nomdpto"
            Top             =   2400
            Visible         =   0   'False
            Width           =   1245
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
            Height          =   320
            Index           =   2
            Left            =   80
            MaxLength       =   10
            TabIndex        =   60
            Text            =   "fecha"
            Top             =   2400
            Visible         =   0   'False
            Width           =   615
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Bindings        =   "frmRepNumSerieGR.frx":0028
            Height          =   3120
            Left            =   135
            TabIndex        =   71
            Top             =   315
            Width           =   10800
            _ExtentX        =   19050
            _ExtentY        =   5503
            _Version        =   393216
            AllowUpdate     =   0   'False
            AllowArrows     =   -1  'True
            BorderStyle     =   0
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
            Caption         =   "Número"
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
            Left            =   11295
            TabIndex        =   78
            Top             =   675
            Width           =   1515
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo Movim."
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
            Left            =   11295
            TabIndex        =   77
            Top             =   1155
            Width           =   1515
         End
         Begin VB.Label Label1 
            Caption         =   "Nº Albaran"
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
            Left            =   11295
            TabIndex        =   76
            Top             =   1515
            Width           =   1395
         End
         Begin VB.Label Label1 
            Caption         =   "Nº Factura"
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
            Left            =   11295
            TabIndex        =   75
            Top             =   1875
            Width           =   1395
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha Vta"
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
            Left            =   11295
            TabIndex        =   74
            Top             =   2235
            Width           =   1275
         End
         Begin VB.Label Label1 
            Caption         =   "Nº linea Vta"
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
            Left            =   11295
            TabIndex        =   73
            Top             =   2595
            Width           =   1395
         End
         Begin VB.Label Label1 
            Caption         =   "Tiene Mantenimiento"
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
            Left            =   11610
            TabIndex        =   72
            Top             =   315
            Width           =   2340
         End
      End
      Begin VB.Frame FrameSusti 
         Caption         =   " Sustituido por "
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
         Height          =   1035
         Left            =   240
         TabIndex        =   56
         Top             =   2265
         Visible         =   0   'False
         Width           =   8865
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
            Height          =   315
            Index           =   17
            Left            =   1125
            MaxLength       =   15
            TabIndex        =   57
            Tag             =   "Nº Serie|T|S|||sserie|numsersu||N|"
            Text            =   "Text1"
            Top             =   420
            Width           =   1710
         End
         Begin VB.Label Label3 
            Caption         =   "Nº Serie"
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
            Left            =   195
            TabIndex        =   58
            Top             =   420
            Width           =   960
         End
      End
      Begin VB.Frame FrameBaja 
         Caption         =   "Datos de baja"
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
         Height          =   1020
         Left            =   9180
         TabIndex        =   51
         Top             =   2265
         Width           =   5625
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
            Left            =   180
            MaxLength       =   10
            TabIndex        =   53
            Tag             =   "Fecha Baja|F|S|||sserie|fechabaja|dd/mm/yyyy|N|"
            Text            =   "Text1"
            Top             =   525
            Width           =   1350
         End
         Begin VB.ComboBox cboMotivoBaja 
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
            ItemData        =   "frmRepNumSerieGR.frx":003D
            Left            =   1785
            List            =   "frmRepNumSerieGR.frx":003F
            Style           =   2  'Dropdown List
            TabIndex        =   52
            Tag             =   "Motivo de Baja|N|S|||sserie|codmotba|0|N|"
            Top             =   525
            Width           =   3015
         End
         Begin VB.Label Label3 
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
            Index           =   2
            Left            =   180
            TabIndex        =   55
            Top             =   240
            Width           =   765
         End
         Begin VB.Label Label3 
            Caption         =   "Motivo"
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
            Left            =   1785
            TabIndex        =   54
            Top             =   240
            Width           =   1215
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   2
            Left            =   1260
            Picture         =   "frmRepNumSerieGR.frx":0041
            ToolTipText     =   "Buscar fecha"
            Top             =   270
            Width           =   240
         End
      End
      Begin VB.Frame FrameNuevos 
         Caption         =   " Datos Compra "
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
         Height          =   1770
         Left            =   9180
         TabIndex        =   41
         Top             =   465
         Width           =   5625
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
            Left            =   1065
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   46
            Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwwwwwwwwww aq"
            Top             =   555
            Width           =   4440
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
            Left            =   165
            MaxLength       =   6
            TabIndex        =   45
            Tag             =   "Cod. Proveedor|N|S|0|999999|sserie|codprove|000000|N|"
            Text            =   "Text11"
            Top             =   555
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
            Index           =   14
            Left            =   1575
            MaxLength       =   10
            TabIndex        =   44
            Tag             =   "Fecha Compra|F|S|||sserie|fechacom|dd/mm/yyyy|N|"
            Text            =   "Text1"
            Top             =   1260
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
            Index           =   15
            Left            =   2970
            MaxLength       =   5
            TabIndex        =   43
            Tag             =   "Nº linea|N|S|0|99999|sserie|numline2||N|"
            Text            =   "Text1"
            Top             =   1260
            Width           =   960
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
            Left            =   180
            MaxLength       =   10
            TabIndex        =   42
            Tag             =   "Nº Albaran Compra|T|S|||sserie|numalbpr||N|"
            Text            =   "Text1 Text"
            Top             =   1260
            Width           =   1350
         End
         Begin VB.Image imgFra 
            Height          =   240
            Index           =   1
            Left            =   5175
            ToolTipText     =   "Buscar Factura Compra"
            Top             =   1260
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
            Index           =   4
            Left            =   165
            TabIndex        =   50
            Top             =   300
            Width           =   1050
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   4
            Left            =   1290
            ToolTipText     =   "Buscar proveedor"
            Top             =   270
            Width           =   240
         End
         Begin VB.Label Label2 
            Caption         =   "Fecha Compra"
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
            Left            =   1575
            TabIndex        =   49
            Top             =   990
            Width           =   1320
         End
         Begin VB.Label Label1 
            Caption         =   "Albaran"
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
            Left            =   165
            TabIndex        =   48
            Top             =   990
            Width           =   1080
         End
         Begin VB.Label Label1 
            Caption         =   "Nº linea Compra"
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
            Left            =   3015
            TabIndex        =   47
            Top             =   990
            Width           =   1440
         End
      End
      Begin VB.Frame FrameActuales 
         Caption         =   " Datos Venta "
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
         Height          =   1785
         Left            =   240
         TabIndex        =   30
         Top             =   450
         Width           =   8865
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
            Left            =   4530
            TabIndex        =   9
            Tag             =   "Tipo Mov|T|S|||sserie|codtipom||N|"
            Text            =   "Text3"
            Top             =   1260
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
            Index           =   10
            Left            =   2115
            MaxLength       =   10
            TabIndex        =   13
            Tag             =   "Fecha Venta|F|S|||sserie|fechavta|dd/mm/yyyy|N|"
            Text            =   "dd/mm/yyyy"
            Top             =   1260
            Width           =   1350
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
            Index           =   6
            Left            =   120
            MaxLength       =   6
            TabIndex        =   5
            Tag             =   "Cod. Cliente|N|S|0|999999|sserie|codclien|000000|N|"
            Text            =   "Text1"
            Top             =   555
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
            Index           =   6
            Left            =   980
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   32
            Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwwwwwwwwww aq"
            Top             =   555
            Width           =   3990
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
            Left            =   5640
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   31
            Text            =   "Text2"
            Top             =   540
            Width           =   3000
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
            Left            =   5025
            MaxLength       =   3
            TabIndex        =   6
            Tag             =   "Direccion/Dpto.|N|S|0|999|sserie|coddirec|000|N|"
            Text            =   "Text1"
            Top             =   540
            Width           =   585
         End
         Begin VB.ComboBox cboTipomov 
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
            ItemData        =   "frmRepNumSerieGR.frx":00CC
            Left            =   4530
            List            =   "frmRepNumSerieGR.frx":00CE
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   1260
            Visible         =   0   'False
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
            Index           =   11
            Left            =   3510
            MaxLength       =   5
            TabIndex        =   14
            Tag             =   "Nº Linea Venta|N|S|0|99999|sserie|numline1||N|"
            Text            =   "Text1"
            Top             =   1260
            Width           =   900
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
            Index           =   8
            Left            =   135
            MaxLength       =   10
            TabIndex        =   11
            Tag             =   "Nº Albaran Venta|N|S|0|9999999|sserie|numalbar|0000000|N|"
            Text            =   "0000000"
            Top             =   1260
            Width           =   945
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
            Index           =   9
            Left            =   1125
            MaxLength       =   10
            TabIndex        =   12
            Tag             =   "Nº Factura Venta|N|S|0|9999999|sserie|numfactu|0000000|N|"
            Text            =   "0000000"
            Top             =   1260
            Width           =   945
         End
         Begin VB.CheckBox chkTieneMan 
            Caption         =   "Tiene Mto"
            Enabled         =   0   'False
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
            Left            =   6000
            TabIndex        =   7
            Tag             =   "¿Tiene Mantenimiento?|N|S|||sserie|tieneman||N|"
            Top             =   1320
            Width           =   1320
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
            Left            =   7320
            MaxLength       =   10
            TabIndex        =   8
            Tag             =   "Nº Mantenimiento|T|S|||sserie|nummante||N|"
            Text            =   "Text1 Text"
            Top             =   1260
            Width           =   1350
         End
         Begin VB.Image imgFra 
            Height          =   240
            Index           =   0
            Left            =   5655
            ToolTipText     =   "Buscar Factura Venta"
            Top             =   1305
            Width           =   240
         End
         Begin VB.Label Label6 
            Caption         =   "Factura"
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
            Left            =   1170
            TabIndex        =   40
            Top             =   990
            Width           =   945
         End
         Begin VB.Label Label1 
            Caption         =   "Albaran"
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
            Left            =   135
            TabIndex        =   39
            Top             =   990
            Width           =   1215
         End
         Begin VB.Label Label11 
            Caption         =   "Fecha Vta"
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
            Left            =   2115
            TabIndex        =   38
            Top             =   990
            Width           =   1185
         End
         Begin VB.Label Label1 
            Caption         =   "Cliente"
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
            Left            =   120
            TabIndex        =   37
            Top             =   270
            Width           =   765
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   2
            Left            =   1020
            ToolTipText     =   "Buscar cliente"
            Top             =   270
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Dirección"
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
            Left            =   5025
            TabIndex        =   36
            Top             =   255
            Width           =   900
         End
         Begin VB.Image imgBuscar 
            Enabled         =   0   'False
            Height          =   240
            Index           =   3
            Left            =   6030
            ToolTipText     =   "Buscar direc./dpto"
            Top             =   225
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo Movim."
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
            Left            =   4530
            TabIndex        =   35
            Top             =   990
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Nº Línea"
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
            Left            =   3510
            TabIndex        =   34
            Top             =   990
            Width           =   945
         End
         Begin VB.Label Label12 
            Caption         =   "Número"
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
            Left            =   7320
            TabIndex        =   33
            Top             =   990
            Width           =   840
         End
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1420
      Left            =   240
      TabIndex        =   21
      Top             =   975
      Width           =   14925
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
         Left            =   11160
         MaxLength       =   10
         TabIndex        =   3
         Tag             =   "Fecha Ult. Repar.|F|S|||sserie|ultrepar|dd/mm/yyyy|N|"
         Text            =   "Text1"
         Top             =   450
         Width           =   1380
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
         Index           =   5
         Left            =   13275
         MaxLength       =   10
         TabIndex        =   4
         Tag             =   "Fecha Fin Garantia|F|S|||sserie|fingaran|dd/mm/yyyy|N|"
         Text            =   "Text1"
         Top             =   450
         Width           =   1380
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
         Index           =   0
         Left            =   225
         MaxLength       =   15
         TabIndex        =   0
         Tag             =   "Nº Serie|T|N|||sserie|numserie||S|"
         Text            =   "000000000000000"
         Top             =   450
         Width           =   1935
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
         Left            =   2250
         MaxLength       =   16
         TabIndex        =   1
         Tag             =   "Cod. Artículo|T1|N|||sserie|codartic||S|"
         Text            =   "0000000000000000"
         Top             =   450
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
         Index           =   1
         Left            =   4365
         Locked          =   -1  'True
         TabIndex        =   23
         Text            =   "Text2"
         Top             =   450
         Width           =   6630
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
         Left            =   2250
         MaxLength       =   8
         TabIndex        =   2
         Tag             =   "Cod. Tipo Artículo|T|N|||sserie|codtipar||N|"
         Text            =   "Te"
         Top             =   900
         Width           =   1050
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
         Left            =   3330
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   22
         Text            =   "Text2"
         Top             =   900
         Width           =   3990
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1935
         ToolTipText     =   "Buscar tipo artículo"
         Top             =   945
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   3105
         Picture         =   "frmRepNumSerieGR.frx":00D0
         Tag             =   "-1"
         ToolTipText     =   "Buscar artículo"
         Top             =   180
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   12690
         Picture         =   "frmRepNumSerieGR.frx":0AD2
         ToolTipText     =   "Buscar fecha"
         Top             =   180
         Width           =   240
      End
      Begin VB.Label Label9 
         Caption         =   "Ult.Reparación"
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
         Left            =   11160
         TabIndex        =   28
         Top             =   195
         Width           =   1890
      End
      Begin VB.Label Label4 
         Caption         =   "Fin Garantia"
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
         Left            =   13275
         TabIndex        =   27
         Top             =   195
         Width           =   1215
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   1
         Left            =   14535
         Picture         =   "frmRepNumSerieGR.frx":0B5D
         ToolTipText     =   "Buscar fecha"
         Top             =   180
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Nº Serie"
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
         TabIndex        =   26
         Top             =   195
         Width           =   1005
      End
      Begin VB.Label Label5 
         Caption         =   "Artículo"
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
         Left            =   2250
         TabIndex        =   25
         Top             =   195
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo de Artículo"
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
         Left            =   225
         TabIndex        =   24
         Top             =   945
         Width           =   1620
      End
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
      Left            =   12915
      TabIndex        =   15
      Top             =   10050
      Width           =   1035
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
      Left            =   14115
      TabIndex        =   16
      Top             =   10050
      Width           =   1035
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
      Left            =   14130
      TabIndex        =   17
      Top             =   10035
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Index           =   0
      Left            =   240
      TabIndex        =   19
      Top             =   9930
      Width           =   2535
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
         Left            =   240
         TabIndex        =   20
         Top             =   180
         Width           =   2115
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
      TabIndex        =   18
      Top             =   8220
      Visible         =   0   'False
      Width           =   3495
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
Attribute VB_Name = "frmRepNumSerie2GR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)
Public DatoAInsertar As String

Private WithEvents frmB As frmBasico2 'Form para busquedas (frmBuscaGrid)
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmB1 As frmBasico2 'Form para busquedas
Attribute frmB1.VB_VarHelpID = -1

Private WithEvents frmF As frmCal 'Calendario de Fechas
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmTA As frmAlmTipoArticulo  'Form Mantenimiento Tipo Articulo
Attribute frmTA.VB_VarHelpID = -1
Private WithEvents frmA As frmBasico2  'Form Mantenimiento Articulos
Attribute frmA.VB_VarHelpID = -1
Private WithEvents frmCli As frmBasico2 'frmFacClientesGr 'Form Mantenimiento Clientes
Attribute frmCli.VB_VarHelpID = -1
Private WithEvents frmProv As frmBasico2 '%=%=frmComProveedores 'Form Mantenimiento Proveedores
Attribute frmProv.VB_VarHelpID = -1

Private HaDevueltoDatos As Boolean
Private Modo As Byte
Private ModoAnterior As Byte


Dim PrimeraVez As Boolean

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
Dim Cad As String

    If Data1.Recordset.EOF Then
        MsgBox "Ningún registro devuelto.", vbExclamation
        Exit Sub
    End If

    Cad = Data1.Recordset.Fields(0) & "|" 'num serie
    Cad = Cad & Data1.Recordset.Fields(1) & "|" 'cod artic
    Cad = Cad & Text2(1).Text & "|"  'nom artic
    Cad = Cad & Data1.Recordset.Fields(3) & "|" 'cod cliente
    RaiseEvent DatoSeleccionado(Cad)
    VariePublic = Text1(0).Text
    Unload Me
End Sub




Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

    If Me.Adodc1.Recordset.EOF Then Exit Sub
    
'    If Modo = 2 Then
        Me.chkAux.Value = DBLet(Me.Adodc1.Recordset!TieneMan, "N")
        txtAux2(2).Text = DBLet(Me.Adodc1.Recordset!nummante, "T")
        txtAux2(3).Text = DBLet(Me.Adodc1.Recordset!codtipom, "T")
        
        txtAux2(4).Text = DBLet(Me.Adodc1.Recordset!Numalbar, "T")
        txtAux2(5).Text = DBLet(Me.Adodc1.Recordset!Numfactu, "T")
        txtAux2(6).Text = DBLet(Me.Adodc1.Recordset!FechaVta, "F")
        txtAux2(7).Text = DBLet(Me.Adodc1.Recordset!numline1, "T")
'    End If
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    If PrimeraVez Then
        PrimeraVez = False
        If Me.DatoAInsertar <> "" Then
            BotonAnyadir
            Text1(0).Text = DatoAInsertar
        End If
    End If
End Sub


Private Sub Form_Load()
Dim i As Integer

    PrimeraVez = True
    
    'Icono del formulario
    Me.Icon = frmPpal.Icon
    
    'Icono de busqueda
    For i = 1 To imgBuscar.Count - 1
        imgBuscar(i).Picture = imgBuscar(0).Picture
    Next i

    For i = 0 To imgFra.Count - 1
        imgFra(i).Picture = imgBuscar(0).Picture
    Next i

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
        .Buttons(1).Image = 36 'Sustitucion de num serie
        .Buttons(2).Image = 47 'Recuperar num serie
        .Buttons(3).Image = 32 'Componentes
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


Private Sub frmB_DatoSeleccionado(CadenaSeleccion As String)
'Formulario para Busqueda
Dim cadB As String
Dim Aux As String
      
    If CadenaSeleccion <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
       'Llama desde Prismatico Direcciones/Departamentos
        Text1(7).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
        Text2(7).Text = RecuperaValor(CadenaSeleccion, 2)
        
        'Pongo QU NOOOOO ha devuelto datos. Asi no hace el regresar
        HaDevueltoDatos = False
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub frmB1_DatoSeleccionado(CadenaSeleccion As String)
Dim cadB As String
Dim Aux As String

    If CadenaSeleccion <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        cadB = ""
        '                       El primero es un pipe
        If Mid(CadenaSeleccion, 1, 1) = "|" Then CadenaSeleccion = """""" & CadenaSeleccion
        Aux = ValorDevueltoFormGrid(Text1(0), CadenaSeleccion, 1)
        cadB = Aux
        Aux = ValorDevueltoFormGrid(Text1(1), CadenaSeleccion, 2)
        cadB = cadB & " and " & Aux
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
        PonerCadenaBusqueda
    End If
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


Private Sub imgBuscar_Click(Index As Integer)
Dim Indice As Byte

    If Modo = 2 Or Modo = 0 Then Exit Sub
 
    Screen.MousePointer = vbHourglass

    Select Case Index
        Case 0 'Codigo Articulo
            Set frmA = New frmBasico2
            'frmA.DatosADevolverBusqueda3 = "@1@" 'Abrir en Modo busqueda
'            frmA.DesdeTPV = False
'            frmA.Show vbModal
            AyudaArticulos frmA, Text1(1)
            Set frmA = Nothing
            Indice = 1
        Case 1  'Cod. Tipo Articulo
            Set frmTA = New frmAlmTipoArticulo
            frmTA.DatosADevolverBusqueda = "0"
            frmTA.Show vbModal
            Set frmTA = Nothing
            Indice = 2
        Case 2 'Cod. Cliente
'            Set frmCli = New frmFacClientesGr
'            frmCli.DatosADevolverBusqueda = "0"
'            frmCli.Show vbModal
            Set frmCli = New frmBasico2
            AyudaClientes frmCli, Text1(6).Text
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
            Indice = 12
'            Set frmProv = New frmComProveedores
'            frmProv.DatosADevolverBusqueda = "0"
'            frmProv.Show vbModal
            Set frmProv = New frmBasico2
            AyudaProveedores frmProv, Text1(Indice)
            Set frmProv = Nothing
    End Select
    PonerFoco Text1(Indice)
    Screen.MousePointer = vbDefault
End Sub


Private Sub imgFecha_Click(Index As Integer)
Dim Indice As Byte

   If Modo = 2 Or Modo = 0 Then Exit Sub
   Screen.MousePointer = vbHourglass
      
   Set frmF = New frmCal
   frmF.Fecha = Now
   
   Select Case Index
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


Private Sub imgFra_Click(Index As Integer)
Dim SQL As String
Dim Numalbar As String
Dim Codtipm As String
Dim FecAlbCompra As String
    
    
                        'ALV:Albaran de Venta (a clientes)
                        'ART: Albaran rectificativo
                        'ALM: ALbaran Mostrador
                        'ALZ: Albaranes "B"
                        'ALI: Albaranes internos
    'comprobar si el Albaran esta facturado o no
    'si no esta facturado abrir el formulario de Entrada de Albaranes: frmFacEntAlbaranes
    'si esta ya facturado abrir el histórico de facturas: frmFacHcoFacturas
        
    If Index = 0 Then
        If Text1(8).Text = "" And Text1(10).Text = "" Then Exit Sub
        
        Numalbar = Text1(8)
        Codtipm = Text1(16) 'Data2.Recordset!detamovi
        
        'consultamos si existe el albaran en la tabla de albaranes: scaalb
        SQL = DevuelveDesdeBDNew(conAri, "scaalb", "numalbar", "codtipom", Codtipm, "T", , "numalbar", Numalbar, "N")
        If SQL <> "" Then 'existe el Albaran
            If vParamAplic.TipoFormularioClientes = 0 Then
            
            
               If vParamAplic.HaciendoFrmulariosGrandes Then
            
                     With frmFacEntAlbaranesGR
                        If EsNumerico(Numalbar) Then
                            .hcoCodMovim = Format(Numalbar, "0000000")
                        Else
                            .hcoCodMovim = Numalbar
                        End If
                        .hcoCodTipoM = Codtipm
                        .Show vbModal
                    End With
            
                Else
                    With frmFacEntAlbaranes2
                        If EsNumerico(Numalbar) Then
                            .hcoCodMovim = Format(Numalbar, "0000000")
                        Else
                            .hcoCodMovim = Numalbar
                        End If
                        .hcoCodTipoM = Codtipm
                        .Show vbModal
                    End With
                End If
            Else
                'FORMULARIO SAIL
                 With frmFacEntAlbSAIL
                 '   If EsNumerico(Data2.Recordset!document) Then
                 '       .hcoCodMovim = Format(Data2.Recordset!document, "0000000")
                 '   Else
                        .hcoCodMovim = Numalbar  ' Data2.Recordset!document
                 '   End If
                    .hcoCodTipoM = Codtipm
                    .Show vbModal
                End With
            End If
        
        Else 'No existe en albaran, abrir Historico Factura
            With frmFacHcoFacturas2
                .DesdeFichaCliente = False
                If EsNumerico(Numalbar) Then
                    .hcoCodMovim = Format(Numalbar, "0000000")
                Else
                    .hcoCodMovim = Numalbar ' Data2.Recordset!document
                End If
                .hcoCodTipoM = Codtipm 'Data2.Recordset!detamovi
                If Codtipm <> "MAT" Then .hcoFechaMov = Text1(10)
                
                .Show vbModal
            End With
        End If
    Else
    
        If Text1(12).Text = "" Or Text1(13).Text = "" Or Text1(14).Text = "" Then Exit Sub
    
        FecAlbCompra = "fechaalb"
        SQL = DevuelveDesdeBDNew(conAri, "scaalp", "numalbar", "codprove", Text1(12).Text, "N", FecAlbCompra, "numalbar", Text1(13), "T", "fentrada", Text1(14), "F")
        If SQL <> "" Then 'existe el Albaran
            If vParamAplic.TipoFormularioClientes = 0 Then
                With frmComEntAlbaranesGR
                    .hcoCodMovim = "ALC"
                    .hcoFechaMovim = FecAlbCompra   'Data2.Recordset!FechaMov
                    .hcoCodProve = Text1(12) 'aqui es el proveedor
                    .EsHistorico = False
                    .Show vbModal
                End With
            Else
                'SAIL
                With frmComEntAlbaranSA
                    .hcoCodMovim = "ALC"
                    .hcoFechaMovim = FecAlbCompra   'Data2.Recordset!FechaMov
                    .hcoCodProve = Text1(12) 'aqui es el proveedor
                    .EsHistorico = False
                    .Show vbModal
                End With
            End If
        Else
            FecAlbCompra = "fechaalb"
            SQL = DevuelveDesdeBDNew(conAri, "schalp", "numalbar", "codprove", Text1(12), "N", FecAlbCompra, "numalbar", Text1(13), "T", "fentrada", Text1(14), "F")
            If SQL <> "" Then 'existe el Albaran
                If vParamAplic.TipoFormularioClientes = 0 Then
                    With frmComEntAlbaranesGR
                        .hcoCodMovim = "ALC"
                        .hcoFechaMovim = FecAlbCompra
                        .hcoCodProve = Text1(12) 'aqui es el proveedor
                        .EsHistorico = True
                        .Show vbModal
                    End With
                Else
                    'SAIL
                    With frmComEntAlbaranSA
                        .hcoCodMovim = "ALC"
                        .hcoFechaMovim = Text1(14)
                        .hcoCodProve = Text1(12) 'aqui es el proveedor
                        .EsHistorico = True
                        .Show vbModal
                    End With
                End If
            Else
        
                'No existe en albaran, abrir Historico Factura
                FecAlbCompra = "fechaalb"
                SQL = "codprove = " & Text1(12) & " AND numalbar=" & DBSet(Text1(13), "T") & " AND fentrada = " & DBSet(Text1(14), "F") & " AND 1 "
                SQL = DevuelveDesdeBD(conAri, "numalbar", "scafpa", SQL, "1", "N", FecAlbCompra)
                If SQL = "" Then FecAlbCompra = Now  'no existe
                
                If vParamAplic.TipoFormularioClientes = 0 Then
                    With frmComHcoFacturas2GR
                        .hcoCodMovim = SQL
                        .hcoFechaMovim = FecAlbCompra  'Data2.Recordset!FechaMov
                        .hcoCodProve = Text1(12) 'aqui es el proveedor
                        
                        .Show vbModal
                    End With
                Else
                    frmComHcoFacturSA.hcoCodMovim = "ALC"
                    frmComHcoFacturSA.hcoCodProve = Text1(12) 'aqui es el proveedor
                    frmComHcoFacturSA.hcoFechaMovim = Text1(14)  ' Data2.Recordset!FechaMov
                    frmComHcoFacturSA.Show vbModal
                End If
            
            End If
        End If
    
    
    End If
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

Private Sub Text1_GotFocus(Index As Integer)
    kCampo = Index
    ConseguirFoco Text1(Index), Modo
End Sub


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 1: KEYBusqueda KeyAscii, 0 'articulo
            Case 2: KEYBusqueda KeyAscii, 1 'tipo de articulo
            Case 6: KEYBusqueda KeyAscii, 2 'cliente
            Case 7: KEYBusqueda KeyAscii, 3 'coddirec
            Case 12: KEYBusqueda KeyAscii, 4 'proveedor
            
            Case 4: KEYFecha KeyAscii, 0 'fec.ultima reparacion
            Case 5: KEYFecha KeyAscii, 1 'fec.fin garantia
            Case 18: KEYFecha KeyAscii, 2 'fec.baja
        End Select
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgBuscar_Click (Indice)
End Sub

Private Sub KEYFecha(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgFecha_Click (Indice)
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Dim devuelve As String

    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub



    Select Case Index
        Case 1 'Codigo Articulo
            If Text1(Index).Text <> "" Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "sartic", "nomartic")
                devuelve = "nseriesn"
                Text1(Index + 1).Text = DevuelveDesdeBDNew(conAri, "sartic", "codtipar", "codartic", Text1(Index).Text, "T", devuelve)
                If devuelve = "1" Then
                    Text2(Index + 1).Text = DevuelveDesdeBDNew(conAri, "stipar", "nomtipar", "codtipar", Text1(Index + 1).Text, "T")
                Else
                    Text2(Index + 1).Text = ""
                    Text1(Index + 1).Text = ""
                    Text2(Index).Text = ""
                    MsgBox "El artículo no tiene control de nº de serie.", vbInformation
                    PonerFoco Text1(Index)
                End If
            Else
                Text2(Index).Text = ""
            End If
            
        Case 2 'Codigo Tipo de Articulo
            Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "stipar", "nomtipar")
            Text1(Index).Text = DevuelveDesdeBD(conAri, "codtipar", "stipar", "codtipar", Text1(Index).Text, "T")
            
        Case 6 'Cliente
            
            Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "sclien", "nomclien")
            
        Case 7 'Direc/dpto del cliente
            If Text1(Index).Text = "" Then
                Text2(Index).Text = ""
                Exit Sub
            End If
            Text1(Index).Text = Format(Text1(Index).Text, "000")
            
            'Comprobar que el cliente seleccionado tiene esa direccion o dpto
            devuelve = DevuelveDesdeBDNew(conAri, "sdirec", "nomdirec", "codclien", Text1(6).Text, "N", , "coddirec", Text1(7).Text, "N")
            Text2(Index).Text = devuelve 'Nombre direc. o dpto
            If devuelve = "" Then 'No existe el dpto
                
                devuelve = DevuelveTextoDepto(False)
                devuelve = "No existe" & devuelve & Text1(Index).Text & " para el cliente: "
                devuelve = devuelve & Text1(6).Text & " - " & Text2(6).Text
                MsgBox devuelve, vbInformation
                PonerFoco Text1(Index)
            End If
            
        Case 12 'Proveedor
            Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "sprove", "nomprove")
            
        Case 4, 5, 10, 14 'Fechas ult. modif., fin garantia
            If Text1(Index).Text <> "" And Text1(Index).Locked = False Then PonerFormatoFecha Text1(Index)
            
            
        Case 18 'fecha de baja
            PonerFormatoFecha Me.Text1(18)
            If Me.Text1(18).Text = "" Then
                Me.cboMotivoBaja.ListIndex = -1
            End If
    End Select
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Index
        Case 1: mnNuevo_Click 'Nuevo
        Case 2: mnModificar_Click  'Modificar
        Case 3: mnEliminar_Click 'Eliminar
            
        Case 5: mnBuscar_Click 'Busqueda
        Case 6: mnVerTodos_Click 'Ver Todos
            
        Case 8: mnImprimir_Click 'Imprimir
    End Select
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
    If KeyAscii = 27 And Modo = 1 Then cmdCancelar_Click 'busqueda
End Sub


Private Sub PonerModo(Kmodo As Byte)
Dim i As Byte
Dim B As Boolean
Dim NumReg As Byte

    Modo = Kmodo
    PonerIndicador lblIndicador, Modo
    
    'Visualizar flechas de desplazamiento en la toolbar si modo=2
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
'    DesplazamientoVisible Me.Toolbar1, btnPrimero, (Modo = 2), NumReg
        
    B = (Modo = 2)
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.visible = B
    Else
        cmdRegresar.visible = False
    End If
    DesplazamientoVisible B And Data1.Recordset.RecordCount > 1
        
    '-------------------------------------------
    'Bloquear Registros
    BloquearText1 Me, Modo
    
    'Los Datos de Albaran de Compras y Ventas siempre bloqueados
    'se actualizan por codigo de programa al insertar las lineas de Albaran
    Me.cboTipomov.Enabled = False
    
            
    'Modo INSERTAR
    B = (Modo = 3) Or (Modo = 4)
    If Modo = 3 Then Me.chkTieneMan.Value = 1
    Me.chkTieneMan.Enabled = B 'Insertar o Modificar
    If B Then BloquearTxt Text1(3), Not CBool(Me.chkTieneMan.Value)
    Me.cboTipomov.Enabled = False 'Insertar o Modificar

    '## LAURA 19/06/2008
    '   añadir datos de baja
    BloquearCmb Me.cboMotivoBaja, Not ((Modo = 1) Or (Modo = 3) Or (Modo = 4))
    '##
    
    '------------------------------------
    B = Modo <> 0 And Modo <> 2
    cmdCancelar.visible = B
    cmdAceptar.visible = B
    
    For i = 0 To Me.imgBuscar.Count - 1
'        Me.imgBuscar(i).Enabled = b
        BloquearImg Me.imgBuscar(i), Not B
    Next i
    
    For i = 0 To Me.imgFecha.Count - 1
        Me.imgFecha(i).Enabled = B 'Si es insertar o modificar
    Next i
    
    'Si Modificar y se ha insertado un nº Albaran no modificar datos
    'del proveedor
    If Trim(Text1(13).Text) <> "" Then
        BloquearTxt Text1(12), True
        Me.imgBuscar(4).Enabled = False
    End If
    
    For i = 0 To imgFra.Count - 1
        imgFra(i).Enabled = (Modo = 2)
        imgFra(i).visible = (Modo = 2)
    Next i
    
    'Poner el tamaño de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos
    
    PonerModoOpcionesMenu   'Activar opciones de menu según Modo
    PonerOpcionesMenu   'Activar opciones de menu según nivel
                        'de permisos del usuario
End Sub


Private Sub DesplazamientoVisible(bol As Boolean)
    FrameDesplazamiento.visible = bol
    FrameDesplazamiento.Enabled = bol
End Sub


Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
    PonerLongCamposGnral Me, Modo, 1
End Sub


Private Sub PonerModoOpcionesMenu()
Dim B As Boolean

    'Modo 2. Hay datos y estamos visualizandolos
    B = (Modo = 2 Or Modo = 0)
    'Insertar
    Toolbar1.Buttons(1).Enabled = B
    Me.mnNuevo.Enabled = B
    
    Toolbar1.Buttons(4).Enabled = B
    Me.mnBuscar.Enabled = B
        
    Toolbar1.Buttons(5).Enabled = B
    mnVerTodos.Enabled = B
    
    
    
    B = (Modo = 2)
    'Modificar
    Toolbar1.Buttons(2).Enabled = B
    Me.mnModificar.Enabled = B
    'eliminar
    Toolbar1.Buttons(3).Enabled = B
    Me.mnEliminar.Enabled = B





    'Sustituir
    Toolbar5.Buttons(1).Enabled = B
    Me.mnSustituir.Enabled = B
    
    'recuperar nº serie
    Toolbar5.Buttons(2).Enabled = B And Text1(6).Text <> ""

    'Componentes
    Toolbar5.Buttons(3).Enabled = B
    Me.mnComponentes.Enabled = B

    '-------------------------------------
    B = (Modo >= 3)
    'Buscar
    Toolbar1.Buttons(5).Enabled = Not B
    Me.mnBuscar.Enabled = Not B
    'Ver Todos
    Toolbar1.Buttons(6).Enabled = Not B
    Me.mnVerTodos.Enabled = Not B
End Sub


Private Sub LimpiarCampos()
    limpiar Me   'Metodo general: Limpia los controles TextBox
    'Aqui va el especifico de cada form es
    Me.cboMotivoBaja.ListIndex = -1
    '### a mano
    Me.chkTieneMan.Value = 0
    
    CargaGrid False
End Sub


Private Sub Desplazamiento(Index As Integer)
'Botones de Desplazamiento de la Toolbar
'Para desplazarse por los registros de control Data
    DesplazamientoData Data1, Index, 350
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
Dim B As Boolean

    B = CompForm(Me, 1)
    If Not B Then Exit Function
 
    'Comprobar que se introduce valor en fecha fin garantia
    If Text1(5).Text = "" Then
        MsgBox "El valor de fecha fin garantia no puede ser nulo.", vbInformation
        B = False
    End If
    
    '## LAURA 19/06/2008
    '- comprobar q si la fecha baja tiene valor el motivo de baja tambien
    '  y viceversa.
    If Me.Text1(18).Text = "" Then
        Me.cboMotivoBaja.ListIndex = -1
    ElseIf Trim(cboMotivoBaja.List(cboMotivoBaja.ListIndex)) = "" Then
        MsgBox "Debe seleccionar un motivo de baja si hay valor en la fecha de baja.", vbInformation
        B = False
    End If
    '##
    
    DatosOk = B
End Function



Private Sub MandaBusquedaPrevia(cadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim Cad As String
Dim tabla As String
Dim Titulo As String, Desc As String
Dim selElem As Byte

    'Llamamos a al form
    Cad = ""
    If EsCabecera Then
    'Estamos en Modo de Cabeceras
    'Registro de la tabla de cabeceras: sserie
'        cad = cad & ParaGrid(Text1(0), 15, "Nº Serie")
'        cad = cad & ParaGrid(Text1(1), 20, "Artic.")
'        cad = cad & "Desc. Artic.|sartic|nomartic|T||38·"
'        cad = cad & ParaGrid(Text1(2), 6, "TArt.")
'        cad = cad & "Desc. Tipo|stipar|nomtipar|T||20·"
'
'        tabla = "(" & NombreTabla & " LEFT JOIN sartic ON " & NombreTabla & ".codartic=sartic.codartic" & ")"
'        tabla = tabla & " LEFT JOIN stipar ON " & NombreTabla & ".codtipar=stipar.codtipar"
'
'        Titulo = "Nº Serie"
'        selElem = 2

        Set frmB1 = New frmBasico2
        AyudaNrosSerie frmB1, Text1(0), cadB
        Set frmB1 = Nothing
        
        Exit Sub
        
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
'        cad = cad & "Cod. " & Desc & "|sdirec|coddirec|N||20·"
'        cad = cad & "Desc. " & Desc & "|sdirec|nomdirec|T||40·"
'        tabla = "sdirec"
'        selElem = 1
        
        Set frmB = New frmBasico2
        AyudaMantenimientosAux frmB, Titulo, Desc, Text1(7), "sdirec.codclien=" & Text1(6)
        Set frmB = Nothing
        
        Exit Sub
    End If
           
'    If cad <> "" Then
'        Screen.MousePointer = vbHourglass
'        Set frmB = New frmBuscaGrid
'        frmB.vCampos = cad
'        frmB.vTabla = tabla
'        frmB.vSQL = cadB
'        HaDevueltoDatos = False
'        '###A mano
'        frmB.vDevuelve = "0|1|"
'        frmB.vTitulo = Titulo
'        frmB.vselElem = selElem
'        frmB.vConexionGrid = conAri 'Conexion a BD Ariges
''        frmB.vBuscaPrevia = chkVistaPrevia
'        '#
'        If Not EsCabecera Then frmB.Label1.FontSize = 11
'        frmB.Show vbModal
'        Set frmB = Nothing
'        'Si ha puesto valores y tenemos que es formulario de busqueda entonces
'        'tendremos que cerrar el form lanzando el evento
'        If HaDevueltoDatos Then
'            If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
'                cmdRegresar_Click
''        Else   'de ha devuelto datos, es decir NO ha devuelto datos
''            If Modo = 5 Then
''                PonerFoco txtAux(0)
''            Else
'                'Esto esta mal
'                'Si hace cmdregresar, ahi hay un UNLOAD
'                'con lo cual NO podemos poner foco, pq volvera a hacer un LOAD
'                'PonerFoco Text1(kCampo)
''            End If
'        End If
'    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub HacerBusqueda()
Dim cadB As String

    cadB = ObtenerBusqueda(Me, False)
    EsCabecera = True
    If chkVistaPrevia = 1 Then
        MandaBusquedaPrevia cadB
    ElseIf cadB <> "" Then 'Se muestran en el mismo form
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
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
    
    Toolbar5.Buttons(2).Enabled = False
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
    
    Toolbar5.Buttons(2).Enabled = (Modo = 2) And Trim(Text1(6).Text) <> ""
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
    CargaGridGnral DataGrid1, Me.Adodc1, SQL, PrimeraVez
    
    tots = "N||||0|;N||||0|;N||||0|;"
    SQL = DevuelveTextoDepto(True)
    tots = tots & "S|txtAux(2)|T|Fecha|1490|;S|txtAux(0)|T|Código|900|;S|txtAux2(0)|T|Nombre Cliente|3770|;"
    tots = tots & "S|txtAux(1)|T|" & SQL & "|1040|;S|txtAux2(1)|T|Nombre " & SQL & "|3030|;"
    tots = tots & "N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;"
       
    arregla tots, DataGrid1, Me, 350
    
    Me.DataGrid1.Columns(4).NumberFormat = "000000"
    Me.DataGrid1.Columns(6).NumberFormat = "000"
    
'    DataGrid1.Enabled = b

    DataGrid1.ScrollBars = dbgAutomatic
    Exit Sub
    
ErrCarga:
    MuestraError Err.Number, "Cargando datos grid: " & DataGrid1.Tag, Err.Description
End Sub

Private Sub Toolbar5_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1: mnSustituir_Click 'Sustitucion num serie
        Case 2: BotonRecuperar 'Recuperar nº serie
        Case 3: mnComponentes_Click 'Componentes
    End Select

End Sub

Private Sub ToolbarDes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Desplazamiento (Button.Index)
End Sub
