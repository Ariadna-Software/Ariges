VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmfacformasenvio2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ENVIOOOOOOOOOOOOOOOOOOOO"
   ClientHeight    =   8085
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   15120
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFacFormasEnvio2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8085
   ScaleWidth      =   15120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   6495
      Left            =   7080
      TabIndex        =   22
      Top             =   720
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   11456
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Datos empresa"
      TabPicture(0)   =   "frmFacFormasEnvio2.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FramePortexExtra"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Conductores"
      TabPicture(1)   =   "frmFacFormasEnvio2.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cboDefecto(0)"
      Tab(1).Control(1)=   "Text3(0)"
      Tab(1).Control(2)=   "Text3(1)"
      Tab(1).Control(3)=   "FrameToolAux(0)"
      Tab(1).Control(4)=   "DataGrid2"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Matrículas"
      TabPicture(2)   =   "frmFacFormasEnvio2.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "DataGrid3"
      Tab(2).Control(1)=   "FrameToolAux(1)"
      Tab(2).Control(2)=   "Text4(0)"
      Tab(2).Control(3)=   "Text4(1)"
      Tab(2).Control(4)=   "cboDefecto(1)"
      Tab(2).ControlCount=   5
      Begin VB.ComboBox cboDefecto 
         Height          =   330
         Index           =   1
         ItemData        =   "frmFacFormasEnvio2.frx":0060
         Left            =   -69000
         List            =   "frmFacFormasEnvio2.frx":006A
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Top             =   2280
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.ComboBox cboDefecto 
         Height          =   330
         Index           =   0
         ItemData        =   "frmFacFormasEnvio2.frx":0076
         Left            =   -70200
         List            =   "frmFacFormasEnvio2.frx":0080
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   2160
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.TextBox Text4 
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   1
         Left            =   -72600
         MaxLength       =   40
         TabIndex        =   38
         Text            =   "14"
         Top             =   2160
         Visible         =   0   'False
         Width           =   4005
      End
      Begin VB.TextBox Text4 
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   0
         Left            =   -74760
         MaxLength       =   40
         TabIndex        =   37
         Text            =   "13"
         Top             =   2160
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.TextBox Text3 
         BorderStyle     =   0  'None
         Height          =   360
         Index           =   0
         Left            =   -74520
         MaxLength       =   15
         TabIndex        =   10
         Text            =   "Text3"
         Top             =   2280
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.TextBox Text3 
         BorderStyle     =   0  'None
         Height          =   360
         Index           =   1
         Left            =   -72720
         MaxLength       =   30
         TabIndex        =   11
         Text            =   "Text3"
         Top             =   2040
         Visible         =   0   'False
         Width           =   3270
      End
      Begin VB.Frame FrameToolAux 
         Height          =   555
         Index           =   1
         Left            =   -74760
         TabIndex        =   34
         Top             =   480
         Width           =   1605
         Begin MSComctlLib.Toolbar ToolbarAux 
            Height          =   330
            Index           =   1
            Left            =   120
            TabIndex        =   35
            Top             =   150
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   582
            ButtonWidth     =   609
            ButtonHeight    =   582
            Style           =   1
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   3
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Insertar"
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Modificar"
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Eliminar"
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame FrameToolAux 
         Height          =   555
         Index           =   0
         Left            =   -74760
         TabIndex        =   31
         Top             =   480
         Width           =   1605
         Begin MSComctlLib.Toolbar ToolbarAux 
            Height          =   330
            Index           =   0
            Left            =   120
            TabIndex        =   32
            Top             =   150
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   582
            ButtonWidth     =   609
            ButtonHeight    =   582
            Style           =   1
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   3
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Insertar"
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Modificar"
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Eliminar"
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame FramePortexExtra 
         BorderStyle     =   0  'None
         Height          =   5415
         Left            =   120
         TabIndex        =   23
         Top             =   600
         Width           =   6735
         Begin VB.TextBox txtAux 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2280
            Index           =   7
            Left            =   1950
            MaxLength       =   250
            MultiLine       =   -1  'True
            TabIndex        =   9
            Tag             =   "O|T|S|||senvio|observa|||"
            Top             =   3120
            Width           =   4590
         End
         Begin VB.TextBox txtAux 
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
            Left            =   1965
            MaxLength       =   20
            TabIndex        =   7
            Tag             =   "Telefono|T|S|||senvio|teltrans1|||"
            Top             =   1920
            Width           =   2190
         End
         Begin VB.TextBox txtAux 
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
            Left            =   1965
            MaxLength       =   6
            TabIndex        =   4
            Tag             =   "Código Postal|T|S|||senvio|cptrans|||"
            Top             =   480
            Width           =   1500
         End
         Begin VB.TextBox txtAux 
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
            Left            =   1965
            MaxLength       =   50
            TabIndex        =   6
            Tag             =   "Provincia|T|S|||senvio|protrans|||"
            Top             =   1440
            Width           =   4590
         End
         Begin VB.TextBox txtAux 
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
            Left            =   1965
            MaxLength       =   50
            TabIndex        =   5
            Tag             =   "Población|T|S|||senvio|pobtrans|||"
            Top             =   960
            Width           =   4590
         End
         Begin VB.TextBox txtAux 
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
            Left            =   1965
            MaxLength       =   50
            TabIndex        =   3
            Tag             =   "Direccion|T|N|||senvio|domtrans|||"
            Top             =   0
            Width           =   4560
         End
         Begin VB.TextBox txtAux 
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
            Left            =   1950
            MaxLength       =   20
            TabIndex        =   8
            Tag             =   "Telefono|T|S|||senvio|teltrans2|||"
            Top             =   2520
            Width           =   2190
         End
         Begin VB.Label Label1 
            Caption         =   "Observaciones"
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
            TabIndex        =   30
            Top             =   3120
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Telefono"
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
            Left            =   255
            TabIndex        =   29
            Top             =   1920
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Domicilio"
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
            Left            =   255
            TabIndex        =   28
            Top             =   0
            Width           =   900
         End
         Begin VB.Label Label2 
            Caption         =   "Provincia"
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
            Left            =   240
            TabIndex        =   27
            Top             =   1440
            Width           =   900
         End
         Begin VB.Label Label1 
            Caption         =   "Población"
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
            Left            =   240
            TabIndex        =   26
            Top             =   960
            Width           =   1125
         End
         Begin VB.Label Label1 
            Caption         =   "C.Postal"
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
            Left            =   240
            TabIndex        =   25
            Top             =   480
            Width           =   1020
         End
         Begin VB.Label Label1 
            Caption         =   "Movil"
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
            Left            =   255
            TabIndex        =   24
            Top             =   2520
            Width           =   1095
         End
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Height          =   4815
         Left            =   -74760
         TabIndex        =   33
         Top             =   1200
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   8493
         _Version        =   393216
         AllowUpdate     =   0   'False
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
      Begin MSDataGridLib.DataGrid DataGrid3 
         Height          =   4815
         Left            =   -74760
         TabIndex        =   36
         Top             =   1200
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   8493
         _Version        =   393216
         AllowUpdate     =   0   'False
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
      ItemData        =   "frmFacFormasEnvio2.frx":008C
      Left            =   5760
      List            =   "frmFacFormasEnvio2.frx":0096
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Tag             =   "Etiqueta|N|N|0||senvio|impetiqu|||"
      Top             =   6000
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtAux 
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
      Height          =   330
      Index           =   1
      Left            =   2040
      MaxLength       =   50
      TabIndex        =   1
      Tag             =   "Nombre|T|N|||senvio|nomenvio|||"
      Top             =   5880
      Width           =   795
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
      Height          =   330
      Index           =   0
      Left            =   240
      MaxLength       =   6
      TabIndex        =   0
      Tag             =   "Código Propio|N|N|0|999|senvio|codenvio|000|S|"
      Top             =   6000
      Width           =   795
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmFacFormasEnvio2.frx":00A2
      Height          =   6345
      Left            =   150
      TabIndex        =   18
      Top             =   840
      Width           =   6795
      _ExtentX        =   11986
      _ExtentY        =   11192
      _Version        =   393216
      AllowUpdate     =   0   'False
      BorderStyle     =   0
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
         AllowFocus      =   0   'False
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Frame FrameBotonGnral 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   135
      TabIndex        =   19
      Top             =   45
      Width           =   3585
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   210
         TabIndex        =   20
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
            NumButtons      =   8
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
            EndProperty
         EndProperty
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
      Left            =   11490
      TabIndex        =   13
      Top             =   7560
      Visible         =   0   'False
      Width           =   1065
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
      Left            =   12765
      TabIndex        =   14
      Top             =   7560
      Visible         =   0   'False
      Width           =   1065
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
      Left            =   12720
      TabIndex        =   15
      Top             =   7560
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   1
      Left            =   120
      TabIndex        =   16
      Top             =   7440
      Width           =   3345
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   45
         TabIndex        =   17
         Top             =   180
         Width           =   2895
      End
   End
   Begin MSAdodcLib.Adodc adodc1 
      Height          =   330
      Left            =   9120
      Top             =   120
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
   Begin MSComctlLib.Toolbar ToolbarAyuda 
      Height          =   330
      Left            =   14325
      TabIndex        =   21
      Top             =   180
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ayuda"
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc data2 
      Height          =   330
      Left            =   11640
      Top             =   0
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
   Begin MSAdodcLib.Adodc data3 
      Height          =   330
      Left            =   12960
      Top             =   0
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
      Begin VB.Menu mnBarra2 
         Caption         =   "-"
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
Attribute VB_Name = "frmfacformasenvio2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor: David                                         +-+-
' +-+- Menú: Intermediarios                                 +-+-
' +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-

' +-+-+-+- DISSENY +-+-+-+-
' 1. Posar tots els controls al formulari
' 2. Posar els index correlativament
' 3. Si n'hi han botons de buscar repasar el ToolTipText
' 4. Alliniar els camps numérics a la dreta i el resto a l'esquerra
' 5. Posar els TAGs
' (si es INTEGER: si PK => mínim 1; si no PK => mínim 0; màxim => 99; format => 00)
' (si es DECIMAL; mínim => 0; màxim => 99.99; format => #,###,###,##0.00)
' (si es DATE; format => dd/mm/yyyy)
' 6. Posar els MAXLENGTHs
' 7. Posar els TABINDEXs

Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

'codi per al registe que s'afegix al cridar des d'altre formulari.
'Obrir en modo Insertar i tornar datos del registre insertat
Public NuevoCodigo As String

Public DeConsulta As Boolean
Public CodigoActual As String

' *** adrede: per a quan busque suplements/desconters des de frmViagrc ***
Public ExpedBusca As Long
Public TipoSuplem As Integer
' *********************************************************************

' *** declarar els formularis als que vaig a cridar ***
'Private WithEvents frmB As frmBuscaGrid

Private CadenaConsulta As String
Private CadB As String


Dim Indice As Byte 'Index del text1 on es poses els datos retornats des d'atres Formularis de Mtos


Private kCampo As Integer

Dim BuscaChekc As String

Dim Modo As Byte
'----------- MODOS --------------------------------
'   0.-  Formulari net sense cap camp ple
'   1.-  Preparant per a fer la búsqueda
'   2.-  Ja tenim registres i els anem a recorrer
'        ,podem editar-los Edició del camp
'   3.-  Inserció de nou registre
'   4.-  Modificar
'--------------------------------------------------

Dim ModificaLineas  As Integer
Dim ModoFrame2 As Byte
Dim PrimVez As Boolean

Private Sub PonerModo(vModo)
Dim b As Boolean
Dim I As Integer
    
    Modo = vModo
'    PonerIndicador lblIndicador, Modo
    
    b = (Modo = 2)
    
    If b Then
        PonerContRegIndicador
    Else
        PonerIndicador lblIndicador, Modo
    End If
    
    
    b = Modo = 1 Or Modo = 3 Or Modo = 4
    txtAux(0).visible = b
    txtAux(1).visible = b
    Me.Combo1.visible = b
    
    ' **** si n'hi han camps fora del grid, bloquejar-los ****
    If vParamAplic.CartaPortes Then
        For I = 2 To 8
            BloquearTxt txtAux(I), Not b
        Next I
    End If
    
    
    
    
    
    ' ********************************************************
    b = Modo <> 2 And Modo <> 0
    
    cmdAceptar.visible = b
    cmdCancelar.visible = b
    DataGrid1.Enabled = Not b
    
    
    If Modo < 5 Then
        Text3(0).visible = False
        Text3(1).visible = False
    End If
    
    'Si es retornar
    If DatosADevolverBusqueda <> "" Then cmdRegresar.visible = Modo = 2

    
    BotonesToolBarAux
    
    
'    PonerLongCampos
    PonerModoOpcionesMenu 'Activar/Desact botons de menu según Modo
'    PonerOpcionesMenu 'Activar/Desact botons de menu según permissos de l'usuari
    

    
 '   BloquearImgBuscar Me, Modo
    
End Sub

Private Sub PonerModoOpcionesMenu()
'Activa/Desactiva botons de la toolbar i del menu, según el modo en que estiguem
Dim b As Boolean

    ' *** adrede: per a que no es puga fer res si estic cridant des de frmViagrc ***

    b = (Modo = 2) And ExpedBusca = 0
    'Busqueda
    Toolbar1.Buttons(5).Enabled = b
    Me.mnBuscar.Enabled = b
    'Vore Tots
    Toolbar1.Buttons(6).Enabled = b
    Me.mnVerTodos.Enabled = b
    
    'Insertar
    Toolbar1.Buttons(1).Enabled = b And Not DeConsulta
    Me.mnNuevo.Enabled = b And Not DeConsulta
        
        
    
    b = (b And adodc1.Recordset.RecordCount > 0) And Not DeConsulta And ExpedBusca = 0
    'Modificar
    Toolbar1.Buttons(2).Enabled = b
    Me.mnModificar.Enabled = b

    'Eliminar
    Toolbar1.Buttons(3).Enabled = b
    Me.mnEliminar.Enabled = b
    'Imprimir
    'Toolbar1.Buttons(8).Enabled = b
    Me.mnImprimir.Enabled = b

    ' ******************************************************************************
End Sub

Private Sub BotonAnyadir()
Dim NumF As String
Dim anc As Single
Dim I As Integer
    
    CargaGrid 'primer de tot carregue tot el grid
    CadB = ""
    '********* canviar taula i camp; repasar codEmpre ************
    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
        NumF = NuevoCodigo
    Else
        NumF = SugerirCodigoSiguienteStr("senvio", "codenvio")
        
    End If
    '***************************************************************
    'Situem el grid al final
    
    PonerModo 3
    
    AnyadirLinea DataGrid1, adodc1

    anc = DataGrid1.Top
    If DataGrid1.Row < 0 Then
        anc = DataGrid1.Top + 220
    Else
        anc = anc + DataGrid1.RowTop(DataGrid1.Row) + 5
    End If
    
    ' *** valors per defecte a l'afegir (dins i fora del grid); repasar codEmpre ***
    txtAux(0).Text = NumF
    For I = 1 To 8
        txtAux(I).Text = ""
    Next I
        
    Combo1.ListIndex = 1

    LLamaLineas anc, 3
       
    ' *** posar el foco ***
    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
        PonerFoco txtAux(1) '**** 1r camp visible que NO siga PK ****
    Else
        PonerFoco txtAux(0) '**** 1r camp visible que siga PK ****
    End If
    ' ******************************************************
End Sub

Private Sub BotonVerTodos()
    CargaGrid ""
    CadB = ""
    PonerModo 2
End Sub

Private Sub BotonBuscar()
    Dim I As Integer
    
    ' *** canviar per la PK (no posar codempre si està a Form_Load) ***
    Modo = 1
    CargaGrid "false"
    CargaLineas False, 0
    '*******************************************************************************

    ' *** canviar-ho pels valors per defecte al buscar (dins i fora del grid);
    For I = 0 To 8
        txtAux(I).Text = ""
    Next I

    LLamaLineas DataGrid1.Top + 240, 1
    
    ' *** posar el foco al 1r camp visible que siga PK ***
    PonerFoco txtAux(1)
    ' ***************************************************************
End Sub

Private Sub BotonModificar()
    Dim anc As Single
    Dim I As Integer

    Screen.MousePointer = vbHourglass
    
    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        I = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, I
        DataGrid1.Refresh
    End If
    
    If DataGrid1.Row < 0 Then
        anc = 320
    Else
        anc = DataGrid1.RowTop(DataGrid1.Row) + DataGrid1.Top  'DataGrid1.RowTop(DataGrid1.Row) + 545
    End If

    ' *** asignar als controls del grid, els valors de les columnes ***
    txtAux(0).Text = DataGrid1.Columns(0).Text
    txtAux(1).Text = DataGrid1.Columns(1).Text
    
    BloquearTxt txtAux(0), True
    If Trim(DataGrid1.Columns(2).Text) = "" Then
        Combo1.ListIndex = 1
    Else
        Combo1.ListIndex = 0
    End If
    ' ********************************************************

    SSTab1.Tab = 0
    LLamaLineas anc, 4 'modo 4
   
    ' *** foco al 1r camp visible que NO siga clau primaria ***
    PonerFoco txtAux(1)
    ' *********************************************************
    Screen.MousePointer = vbDefault
End Sub

Private Sub LLamaLineas(alto As Single, xModo As Byte)
Dim I As Integer

    DeseleccionaGrid Me.DataGrid1
    PonerModo xModo

    ' *** posar el Top a tots els controls del grid (botons també) ***
    'Me.imgFec(2).Top = alto
    For I = 0 To 1
        txtAux(I).Top = alto
    Next I
    Me.Combo1.Top = alto
    
End Sub

Private Sub BotonEliminar()
Dim SQL As String
Dim temp As Boolean

    On Error GoTo Error2
    
    'Certes comprovacions
    If adodc1.Recordset.EOF Then Exit Sub
    If Not SepuedeBorrar Then Exit Sub

    
    '### a mano
    SQL = "¿Seguro que desea eliminar la Forma de Envío?" & vbCrLf
    SQL = SQL & vbCrLf & "Código: " & Format(adodc1.Recordset.Fields(0), "000")
    SQL = SQL & vbCrLf & "Denominación: " & adodc1.Recordset.Fields(1)
    If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
        
        CadB = adodc1.RecordSource
        NumRegElim = InStr(1, CadB, " WHERE ")
        If NumRegElim > 0 Then
            CadB = Mid(CadB, NumRegElim + 7)
            NumRegElim = InStr(1, CadB, " ORDER BY ")
            If NumRegElim > 0 Then CadB = Mid(CadB, 1, NumRegElim)
            
        Else
            CadB = ""
        End If
        
        'N'hi ha que eliminar
        NumRegElim = adodc1.Recordset.AbsolutePosition
        SQL = "Delete from senvio where codenvio = " & adodc1.Recordset!CodEnvio
        
        conn.Execute SQL
        
        
            
        CargaGrid CadB

        temp = SituarDataTrasEliminar(adodc1, NumRegElim, True)
        PonerModoOpcionesMenu
        adodc1.Recordset.Cancel
    End If
    Exit Sub
    
Error2:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Eliminando registro", Err.Description
End Sub

Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
'Si maxlength=6 para codprove, en modo busqueda Maxlenth=13 'el doble + 1
    
    'para los txtAux
    PonerLongCamposGnral Me, Modo, 3
End Sub

Private Sub cboDefecto_KeyPress(index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, Modo, False
End Sub

Private Sub cmdAceptar_Click()
Dim I As Long
Dim b As Boolean
Dim cad As String
    Select Case Modo
        Case 3 'INSERTAR
            If DatosOk Then
                
                If vParamAplic.CartaPortes Then
                    b = InsertarDesdeForm2(Me, 0)
                Else
                    b = InsertarDesdeForm2(Me, 2, Me.Name) 'Solo inserta cod,nom,etiq
                End If
                If b Then
                    CargaGrid
                    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
                        cmdCancelar_Click
                        If Not adodc1.Recordset.EOF Then
                            ' *** filtrar per tota la PK; repasar codEmpre **
                            'adodc1.Recordset.Filter = "codempre = " & txtAux(0).Text & " AND codsupdt = " & txtAux(1).Text
                            adodc1.Recordset.Filter = "codenvio = " & txtAux(0).Text
                            ' ****************************************************
                        End If
                        cmdRegresar_Click
                    Else
                        BotonAnyadir
                    End If
                    CadB = ""
                End If
            End If
            
        Case 4 'MODIFICAR
            If DatosOk Then
                'If ModificaDesdeFormulario(Me) Then
                If ModificaDesdeFormulario(Me, 3) Then    'la opcion 3 es l
                    I = adodc1.Recordset.AbsolutePosition
                    TerminaBloquear
                    PonerModo 2
                    CargaGrid CadB
'                    If CadB <> "" Then
'                        CargaGrid CadB
'                        lblIndicador.Caption = "RESULTADO BUSQUEDA"
'                    Else
'                        CargaGrid
'                        lblIndicador.Caption = ""
'                    End If
                    adodc1.Recordset.Move I - 1
                    PonerFocoGrid Me.DataGrid1
                End If
            End If
            
        Case 1  'BUSQUEDA
            CadB = ObtenerBusqueda(Me, False)
            If CadB <> "" Then
                
                PonerModo 2
                CargaGrid CadB
'                lblIndicador.Caption = "RESULTADO BUSQUEDA"
                PonerFocoGrid Me.DataGrid1
            End If
            
            
            
        Case 5, 6
                If InsertarModificarLinea Then
                    If Modo = 6 Then
                        cad = "matricula = " & DBSet(Text3(0).Text, "T")
                    Else
                        cad = "chofer = " & NumRegElim
                    End If
                    
                    
                    If Modo = 5 Then
                        
                        LLamaLineasChofer 0, 0
                        DataGrid2.AllowAddNew = False
                    
                        CargaLineas True, 1
                    
                        If ModificaLineas = 1 Then
                            Data2.Recordset.MoveLast
                        Else
                            Data2.Recordset.Find cad
                        End If
                        b = True
    
                    
                    Else
                        'If Modo = 6 Then
                        
                        LLamaLineasMatricula 0, 0
                        DataGrid3.AllowAddNew = False
                        CargaLineas True, 2
                    
                        If ModificaLineas = 1 Then
                            data3.Recordset.MoveLast
                        Else
                            data3.Recordset.Find cad
                        End If
                        b = True
                    End If
                    PonerModo 2
                End If
    End Select
End Sub

Private Sub cmdCancelar_Click()
Dim cad As String


    If Modo < 5 Then
        Select Case Modo
            Case 3 'INSERTAR
                DataGrid1.AllowAddNew = False
                'CargaGrid
                If Not adodc1.Recordset.EOF Then adodc1.Recordset.MoveFirst
            Case 4 'MODIFICAR
                TerminaBloquear
            Case 1 'BUSQUEDA
                CargaGrid CadB
                
                
        End Select
        
        If Not adodc1.Recordset.EOF Then
            CargaForaGrid
        Else
            LimpiarCampos
        End If
        
        
        
        
    Else
        PonerModoFrame 0, Modo
        If Modo = 5 Then
            DataGrid2.AllowAddNew = False
            If ModificaLineas = 1 Then '1 = Insertar
                If Not Data2.Recordset.EOF Then Data2.Recordset.MoveFirst
            ElseIf ModificaLineas = 2 Then 'Modificar
                 cad = "(chofer=" & Data2.Recordset!Chofer & ")"
                 CargaLineas True, 1
                 Data2.Recordset.Find cad
            End If
            
            LLamaLineasChofer 0, 0
        
        
        Else
            DataGrid3.AllowAddNew = False
                        
            If ModificaLineas = 1 Then '1 = Insertar
                If Not data3.Recordset.EOF Then data3.Recordset.MoveFirst
            ElseIf ModificaLineas = 3 Then 'Modificar
                 cad = "(codenvio='" & "" & "')"
                 CargaLineas True, 2
                 data3.Recordset.Find cad
            End If
            
            
            LLamaLineasMatricula 0, 0
            
        End If
            
            
            
            
        ModificaLineas = 0
    End If
    
    
    
    PonerModo 2
    PonerFocoGrid Me.DataGrid1

End Sub

Private Sub cmdRegresar_Click()
Dim cad As String

    If adodc1.Recordset.EOF Then
        MsgBox "Ningún registro devuelto.", vbExclamation
        Exit Sub
    End If
    cad = ""
    
    
    
    cad = adodc1.Recordset.Fields(0) & "|"
    cad = cad & adodc1.Recordset.Fields(1) & "|"
    
    
    RaiseEvent DatoSeleccionado(cad)
    Unload Me
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub DataGrid1_DblClick()
    If cmdRegresar.visible Then cmdRegresar_Click
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_activate()
    Screen.MousePointer = vbDefault
    'Posem el foco
    If PrimVez Then
        PrimVez = False
        CadB = ""
        CargaGrid
        
    
        If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
            BotonAnyadir
            
        Else
            If Not adodc1.Recordset.EOF Then CargaForaGrid
            PonerModo 2
        End If
    
       
        
    End If
End Sub

Private Sub Form_Load()
Dim I As Integer


    Me.Icon = frmPpal.Icon
    PrimVez = True
    With Me.Toolbar1
        .HotImageList = frmPpal.imgListComun_OM2
        .DisabledImageList = frmPpal.imgListComun_BN2
        .ImageList = frmPpal.ImgListComun2
        'l'1 i el 2 son separadors
        .Buttons(5).Image = 1   'Buscar
        .Buttons(6).Image = 2   'Totss
        'el 5 i el 6 son separadors
        .Buttons(1).Image = 3   'Insertar
        .Buttons(2).Image = 4   'Modificar
        .Buttons(3).Image = 5   'Borrar

    End With
    
    If vParamAplic.CartaPortes Then
        
            With Me.ToolbarAux(0)
                .HotImageList = frmPpal.imgListComun_OM16
                .DisabledImageList = frmPpal.imgListComun_BN16
                .ImageList = frmPpal.imgListComun16
                
                .Buttons(1).Image = 3
                .Buttons(2).Image = 4
                .Buttons(3).Image = 5
             End With
             With Me.ToolbarAux(1)
                .HotImageList = frmPpal.imgListComun_OM16
                .DisabledImageList = frmPpal.imgListComun_BN16
                .ImageList = frmPpal.imgListComun16
                
                .Buttons(1).Image = 3
                .Buttons(2).Image = 4
                .Buttons(3).Image = 5
             End With
        
    End If

    If vParamAplic.CartaPortes Then
        Caption = "Transportistas"
        Me.FramePortexExtra.visible = True
        I = 15210
    Else
        Caption = "Formas de envio"
        Me.FramePortexExtra.visible = False
        I = 7090
        SSTab1.visible = False
    End If
    

    Me.Width = I
    Me.cmdCancelar.Left = I - 1200
    Me.cmdRegresar.Left = cmdCancelar.Left
    Me.cmdAceptar.Left = I - 2350
    
    
    '****************** canviar la consulta *********************************+
    CadenaConsulta = "SELECT codenvio,nomenvio,if(impetiqu=1,'Si','') etiquetas "
    If vParamAplic.CartaPortes Then CadenaConsulta = CadenaConsulta & ",domtrans,cptrans,pobtrans,protrans,teltrans1,teltrans2,observa"
    CadenaConsulta = CadenaConsulta & " FROM senvio "
    '************************************************************************
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
   
    Screen.MousePointer = vbDefault
    If Modo = 4 Then TerminaBloquear
End Sub


Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnImprimir_Click()
    'printNou
End Sub

Private Sub mnModificar_Click()
    'Comprobaciones
    If adodc1.Recordset.EOF Then Exit Sub
    If adodc1.Recordset.RecordCount < 1 Then Exit Sub

    'Prepara para modificar
    '-----------------------
    'If BloqueaDesdeFormularioTXTAUX(Me) Then
    BotonModificar
End Sub

Private Sub mnNuevo_Click()
    BotonAnyadir
End Sub

Private Sub mnSalir_Click()
    Unload Me
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub

Private Sub Text3_KeyPress(index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, Modo, False
End Sub

Private Sub Text4_KeyPress(index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, Modo, False
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.index
        Case 5
                mnBuscar_Click
        Case 6
                mnVerTodos_Click
        Case 1
                mnNuevo_Click
        Case 2
                mnModificar_Click
        Case 3
                mnEliminar_Click
        Case 8 'Imprimir
                mnImprimir_Click
    End Select
End Sub

Private Sub CargaGrid(Optional vSQL As String)
    Dim I As Integer
    Dim SQL As String
    Dim tots As String
    

    If vSQL <> "" Then
        SQL = CadenaConsulta & " WHERE " & vSQL
    Else
        SQL = CadenaConsulta
    End If
    '********************* canviar el ORDER BY *********************++
    'SQL = SQL & " ORDER BY codempre, codsupdt"
    SQL = SQL & " ORDER BY codenvio"
    '**************************************************************++
    
       
    CargaGridGnral Me.DataGrid1, Me.adodc1, SQL, PrimVez, 330
       
    ' *** posar només els controls del grid ***
    tots = "S|txtAux(0)|T|Código|1050|;S|txtAux(1)|T|Denominación|4157|;S|Combo1|C|Etiqu|700|;"
    If vParamAplic.CartaPortes Then
        For I = 1 To 7
            tots = tots & "N||||0|;"
        Next I
    End If
    arregla tots, DataGrid1, Me, 350
    
    DataGrid1.ScrollBars = dbgAutomatic
    ' **********************************************************
    
    ' *** alliniar les columnes que siguen numèriques a la dreta ***
    'DataGrid1.Columns(1).Alignment = dbgRight
    'DataGrid1.Columns(2).Alignment = dbgRight
    'DataGrid1.Columns(5).Alignment = dbgRight
    ' *****************************
    
    
    ' *** Si n'hi han camps fora del grid ***
   
            If Not adodc1.Recordset.EOF Then
                CargaForaGrid
                CargaLineas True, 0
            Else
                LimpiarCampos
                CargaLineas False, 0
            End If
    ' **************************************
End Sub


Private Sub txtAux_GotFocus(index As Integer)
    ConseguirFocoLin txtAux(index)
End Sub

Private Sub TxtAux_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
'Alvançar/Retrocedir els camps en les fleches de desplaçament del teclat.
    KEYdown KeyCode
End Sub

Private Sub txtAux_KeyPress(index As Integer, KeyAscii As Integer)
'    If Index = 3 And KeyAscii = 43 Then '+
'        KeyAscii = 0
'    Else
'        KEYpress KeyAscii
'    End If
    If KeyAscii = 43 Then
        If Modo = 1 Or Modo = 3 Or Modo = 4 Then
            Select Case index
                Case 12: KEYBusqueda KeyAscii, 0 'cuenta contable
            End Select
        End If
    Else
        KEYpress KeyAscii
    End If
    
End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    
End Sub


Private Sub txtAux_LostFocus(index As Integer)
Dim cadMen As String

    If Not PerderFocoGnral(txtAux(index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    '*** configurar el LostFocus dels camps (de dins i de fora del grid) ***
    Select Case index
        
            
    End Select
    
    
End Sub


Private Function DatosOk() As Boolean
Dim Datos As String
Dim b As Boolean
' *** només per ad este manteniment ***
Dim RS As Recordset
Dim cad As String
Dim Cta As String
Dim cadMen As String

'Dim exped As String
' *************************************

    b = CompForm(Me, 2)
    If Not b Then Exit Function


    If b And (Modo = 3) Then
        Datos = DevuelveDesdeBD(conAri, "codenvio", "senvio", "codenvio", txtAux(0).Text, "N")

         
        If Datos <> "" Then
            MsgBox "Ya existe el transportista: " & txtAux(0).Text, vbExclamation
            DatosOk = False
            PonerFoco txtAux(1) '*** posar el foco al 1r camp visible de la PK de la capçalera ***
            Exit Function
        End If
        '*************************************************************************************
    End If

    DatosOk = b
End Function


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me

End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim Cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, Cerrar
    If Cerrar Then Unload Me
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim I As Byte

    If Modo <> 4 Then 'Modificar
        CargaForaGrid
        
    Else
        For I = 0 To txtAux.Count - 1
            txtAux(I).Text = ""
        Next I
    End If
    
    PonerContRegIndicador
    
End Sub

Private Sub CargaForaGrid()
Dim b As Boolean
    If vParamAplic.CartaPortes Then
        If DataGrid1.Columns.Count <= 2 Then Exit Sub
        ' *** posar als camps de fora del grid el valor de la columna corresponent ***
        txtAux(2) = DataGrid1.Columns(3).Text
        txtAux(3) = DataGrid1.Columns(4).Text
        txtAux(4) = DataGrid1.Columns(5).Text
        txtAux(5) = DataGrid1.Columns(6).Text
        txtAux(6) = DataGrid1.Columns(7).Text
        txtAux(7) = DataGrid1.Columns(9).Text
        txtAux(8) = DataGrid1.Columns(8).Text
        
        txtAux(0).Text = adodc1.Recordset.Fields(0)
        CargaLineas Modo = 2, 0
        BotonesToolBarAux
    End If
 End Sub

Private Sub LimpiarCampos()
Dim I As Integer
On Error Resume Next

    ' *** posar a huit tots els camps de fora del grid ***
    If vParamAplic.CartaPortes Then
        For I = 2 To 8
            txtAux(I).Text = ""
        Next I
        ' ****************************************************
    End If
    Combo1.ListIndex = -1
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub PonerContRegIndicador()
'si estamos en modo ver registros muestra el numero de registro en el que estamos
'situados del total de registros mostrados: 1 de 24
Dim cadReg As String

    If (Modo = 2 Or Modo = 0) Then
        cadReg = PonerContRegistros(Me.adodc1)
        If CadB = "" Then
            lblIndicador.Caption = cadReg
        Else
            lblIndicador.Caption = "BUSQUEDA: " & cadReg
        End If
    End If
End Sub




Private Function SepuedeBorrar() As Boolean
        SepuedeBorrar = False
        'Puede borrar
    'Scppr no lleva referencial
    CadB = DevuelveDesdeBD(conAri, "numpedpr", "scappr", "codenvio", CStr(adodc1.Recordset.Fields(0)))
    If CadB <> "" Then
        MsgBox "Dato  vinculado a pedido de proveedor", vbExclamation
        Exit Function
    End If
    
    SepuedeBorrar = True
    
End Function

'******************************************************************************************
'******************************************************************************************
'******************************************************************************************
'******************************************************************************************

' Poner toolbaraux

'******************************************************************************************
'******************************************************************************************
'******************************************************************************************
'******************************************************************************************

Private Sub BotonesToolBarAux()
Dim b As Boolean

    
    If Not vParamAplic.CartaPortes Then Exit Sub
    
   
    
    
        b = Modo = 2 Or Modo = 5
        ToolbarAux(0).Buttons(1).Enabled = b
        If b Then b = Me.Data2.Recordset.RecordCount > 0
        ToolbarAux(0).Buttons(2).Enabled = b   '(Modo = 2 And Me.Data2.Recordset.RecordCount > 0)
        ToolbarAux(0).Buttons(3).Enabled = b  '(Modo = 2 And Me.Data2.Recordset.RecordCount > 0)
        
    
        b = Modo = 2 Or Modo = 6
        ToolbarAux(1).Buttons(1).Enabled = b
        If b Then b = Me.data3.Recordset.RecordCount > 0
        ToolbarAux(1).Buttons(2).Enabled = b
        ToolbarAux(1).Buttons(3).Enabled = b
    
End Sub



Private Sub ToolbarAux_ButtonClick(index As Integer, ByVal Button As MSComctlLib.Button)

    If Modo <> 2 And Modo < 5 Then Exit Sub

    If Modo >= 5 And ModificaLineas > 0 Then Exit Sub
    
    Select Case index
    Case 0
    
        
        'Departamentos
        If Button.index = 3 Then
            BotonEliminarLinea True
        Else
            PonerModo 5
            If Button.index = 1 Then
                'AÑADIR linea factura
                BotonAnyadirLinea
            Else
                'MODIFICAR linea factura
                BotonModificarLinea
            End If
        End If


    Case 1
        'Direcciones de envio
        If Button.index = 3 Then
            BotonEliminarLinea False
        Else
            PonerModo 6
            If Button.index = 1 Then
                'AÑADIR linea factura
                BotonAnyadirLinea
            Else
                'MODIFICAR linea factura
                BotonModificarLinea
            End If
        End If
    End Select
End Sub



Private Sub BotonAnyadirLinea()
Dim aModo As Byte
Dim vWhere As String
    
   
    If ModificaLineas = 2 Then Exit Sub
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    '   5.-  Mantenimiento Lineas de direcciones/dpto
    '   6.-  "              "     de direcciones de envio
    aModo = Modo
    If aModo = 5 Then
        Me.SSTab1.Tab = 1
    Else
        Me.SSTab1.Tab = 2
    End If
    PonerModoFrame 3, aModo  '3: Insertar
    ModificaLineas = 1 'Insertar
    lblIndicador.Caption = "Insertar línea " & IIf(Modo = 5, "", "")
    PonerModoOpcionesMenu

    'Obtenemos la siguiente numero de Direc./Dpto
    vWhere = "codenvio=" & txtAux(0).Text
    
    cboDefecto(aModo - 5).ListIndex = 0
    If aModo = 5 Then
        AnyadirLinea DataGrid2, Data2
        LLamaLineasChofer ObtenerAlto(DataGrid2, 30), 1
        
        PonerFoco Text3(0)
       
    Else
        
        
        
        AnyadirLinea DataGrid3, data3
        LLamaLineasMatricula ObtenerAlto(DataGrid3, 30), 1
        

        
      
        PonerFoco Text4(0)
    End If
        
End Sub


Private Sub PonerModoFrame(Kmodo As Byte, ModoGral As Byte)
Dim I As Byte
On Error GoTo EPonerModoFr

    ModoFrame2 = Kmodo
    PonerModo ModoGral
    
    'Bloquear TextBox sino modo 3 o 4
    Select Case ModoGral
    Case 5
        For I = 0 To Me.Text3.Count - 1
            If ModoFrame2 = 3 Then Text3(I).Text = ""
            BloquearTxt Text3(I), (ModoFrame2 = 0)
        Next I
        
        
        If ModoFrame2 = 4 Then BloquearTxt Text3(0), True
        
       
    Case 6
        
        For I = 0 To Me.Text4.Count - 1
            If ModoFrame2 = 3 Then Text4(I).Text = ""
            BloquearTxt Text4(I), (ModoFrame2 = 0)
        Next I
        If ModoFrame2 = 4 Then BloquearTxt Text4(0), True
       
    End Select
    BotonesToolBarAux
    Me.cmdCancelar.visible = Kmodo = 3 Or Kmodo = 4
    Me.cmdAceptar.visible = cmdCancelar.visible
EPonerModoFr:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub



'Matricula o chofer
Private Sub BotonEliminarLinea(Chofer As Boolean)
Dim cad As String, cad2 As String
Dim I As Integer

    If Modo <> 2 Then Exit Sub
    If Chofer Then
        If Data2.Recordset.EOF Then Exit Sub
        If Data2.Recordset.RecordCount < 1 Then Exit Sub
    Else
        If data3.Recordset.EOF Then Exit Sub
        If data3.Recordset.RecordCount < 1 Then Exit Sub
    End If
    
    
    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Or ModificaLineas = 2 Then Exit Sub '1= Insertar, 2=Modificar
       
       
    If Chofer Then
        BuscaChekc = ""
        For I = 1 To 3
            
            cad = DevuelveDesdeBD(conAri, "count(*)", RecuperaValor("scaalb|schalb|scafac1|", I), "chofer", CStr(Data2.Recordset!Chofer), "N")
            If cad = "" Then cad = "0"
            If Val(cad) > 0 Then
                BuscaChekc = RecuperaValor("Albaranes|Alb. anulados|Facturas|", I)
                Exit For
            End If
        Next I
        
        If BuscaChekc <> "" Then
            MsgBox "Existen datos en " & BuscaChekc & " con este conductor", vbExclamation
            Exit Sub
        End If
        
    Else
        
            
            cad = DevuelveDesdeBD(conAri, "count(*)", "scaalb_portes", "matricula", CStr(data3.Recordset.Fields(0)), "T")
            If cad = "" Then cad = "0"
            If Val(cad) > 0 Then
                MsgBox "Existen albaranes asociados a esta matricula", vbExclamation
                Exit Sub
            End If
        
    End If
       
   
    
    'Dependiendo del parametro de la aplicacion trabajamos con Dpto o Direc.
    cad = IIf(Chofer, "el conductor", "la matricula")
    cad = "¿Seguro que desea eliminar " & cad & vbCrLf
    If Chofer Then
          
        cad = cad & vbCrLf & "Codigo : " & Data2.Recordset.Fields(0)
        cad = cad & vbCrLf & "Nombre: " & Data2.Recordset.Fields(1)
    Else
        cad = cad & vbCrLf & "Codigo : " & data3.Recordset.Fields(0)
        cad = cad & vbCrLf & "Desc.: " & data3.Recordset.Fields(1)
        
    End If
    'Borramos
    If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then
        
        'Hay que eliminar
        On Error GoTo Error2
        Screen.MousePointer = vbHourglass
        If Chofer Then
            NumRegElim = Data2.Recordset.AbsolutePosition
            cad = "sconductor WHERE  chofer = " & Data2.Recordset!Chofer
            I = Data2.Recordset.AbsolutePosition
        Else
            NumRegElim = data3.Recordset.AbsolutePosition
            cad = "smatriculas WHERE matricula = " & DBSet(data3.Recordset!Matricula, "T")
            I = data3.Recordset.AbsolutePosition
        End If
        
        
        cad = "DELETE FROM " & cad
        conn.Execute cad
        
      
        
        I = I - 1
        
        CargaLineas True, IIf(Chofer, 1, 2)
        
        If I > 0 Then
            If Chofer Then
                Data2.Recordset.Move I
            Else
                data3.Recordset.Move I
            End If
        End If
            

        ModificaLineas = 0

    End If
    
    
Error2:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then
        Data2.Recordset.CancelUpdate
        MsgBox Err.Number & ": " & Err.Description, vbExclamation
    End If
End Sub



Private Sub BotonModificarLinea()
Dim aModo As Byte
'Modificar una linea
    aModo = Modo
    
    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Then Exit Sub '1= Insertar
    '   5.-  chofer
    '   6.-  matriculas
    If aModo = 5 Then
        If Data2.Recordset.EOF Then Exit Sub
        If Data2.Recordset.RecordCount < 1 Then Exit Sub
        Me.SSTab1.Tab = 1
        
    Else
        'If aModo = 6 Then
        If data3.Recordset.EOF Then Exit Sub
        If data3.Recordset.RecordCount < 1 Then Exit Sub
        Me.SSTab1.Tab = 2
    
    End If
    
    
    
    
    
       
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModoFrame 4, aModo 'ModoFrame=4 -> Modificar
    Me.lblIndicador.Caption = "Modificar linea "
    ModificaLineas = 2 'Modificar
    PonerModoOpcionesMenu
    
    'Como el campo 1 es clave primaria, NO se puede modificar
    If aModo = 5 Then
        LLamaLineasChofer ObtenerAlto(DataGrid2, 30), 2
        
        Text3(0).Text = DataGrid2.Columns(0).Value
        Text3(1).Text = DataGrid2.Columns(1).Value
        cboDefecto(0).ListIndex = IIf(UCase(DataGrid2.Columns(2).Value) = "SI", 1, 0)
        BloquearTxt Text3(0), True
        PonerFoco Text3(1)
    Else
        'If aModo = 6 Then
        LLamaLineasMatricula ObtenerAlto(DataGrid3, 30), 2
        Text4(0).Text = DataGrid3.Columns(0).Value
        Text4(1).Text = DataGrid3.Columns(1).Value
        cboDefecto(1).ListIndex = IIf(UCase(DataGrid3.Columns(2).Value) = "SI", 1, 0)
        
        BloquearTxt Text4(0), True
        PonerFoco Text4(1)
    End If
    
End Sub



Private Sub LLamaLineasChofer(alto As Single, xModo As Byte)
Dim b As Boolean
Dim I As Byte

    ModificaLineas = xModo
    
    b = Modo = 5 And (ModificaLineas = 1 Or ModificaLineas = 2) 'Insertar o Modificar Lineas

    DataGrid2.Enabled = Not b
    
    DeseleccionaGrid Me.DataGrid2
   
    For I = 0 To 1
           ' Text3(i).Height = DataGrid6.RowHeight
            Text3(I).visible = b
            Text3(I).Top = alto
    Next
    Me.cboDefecto(0).visible = b
    Me.cboDefecto(0).Top = alto
    
End Sub


Private Sub LLamaLineasMatricula(alto As Single, xModo As Byte)
Dim b As Boolean
Dim I As Byte

    ModificaLineas = xModo
    
    b = Modo = 6 And (ModificaLineas = 1 Or ModificaLineas = 2) 'Insertar o Modificar Lineas

    DataGrid3.Enabled = Not b
    
    DeseleccionaGrid Me.DataGrid3
   
    For I = 0 To 1
        
           ' Text3(i).Height = DataGrid6.RowHeight
            Text4(I).visible = b
            Text4(I).Top = alto
        
    Next
    Me.cboDefecto(1).visible = b
    Me.cboDefecto(1).Top = alto
End Sub


Private Sub CargaLineas(enlaza As Boolean, Cual_ As Byte)
'cual:     1  chofer, 2  matricula
'          0 Todos
Dim SQL As String


        If Cual_ = 0 Or Cual_ = 1 Then
            SQL = "SELECT    dni  ,nombre  , if(defecto =1,'Si','') defecto,chofer FROM sconductor where "
            If enlaza Then
                If Trim(txtAux(0).Text) = "" Then
                    SQL = SQL & " false"
                Else
                    SQL = SQL & "codenvio = " & txtAux(0).Text
                End If
            Else
                SQL = SQL & " false"
            End If
             
            SQL = SQL & " ORDER BY  chofer"
            CargaGridGnral DataGrid2, Me.Data2, SQL, PrimVez, 360
            SQL = "S|Text3(0)|T|DNI|1300|;S|Text3(1)|T|Nombre|3000|;"
            SQL = SQL & "S|cboDefecto(0)|C|Def.|790|;"
            'Los campos que no se ven que van FUERA DEL GRID
            SQL = SQL & "||||150|;"
            arregla SQL, DataGrid2, Me, 360
            DataGrid1.ScrollBars = dbgAutomatic
            
        End If
        
        
        If Cual_ = 0 Or Cual_ = 2 Then
            SQL = "SELECT matricula ,titulo ,if(defecto =1,'Si','') defecto,codenvio FROM smatriculas WHERE "
            If enlaza Then
                If Trim(txtAux(0).Text) = "" Then
                    SQL = SQL & " false"
                Else
                    SQL = SQL & " codenvio= " & txtAux(0).Text
                End If
            Else
                SQL = SQL & " false"
            End If
            SQL = SQL & " ORDER BY  matricula "
            CargaGridGnral DataGrid3, Me.data3, SQL, PrimVez, 330
            
            SQL = "S|Text4(0)|T|matricula|1400|;S|Text4(1)|T|Decr.|3500|;"
            SQL = SQL & "S|cboDefecto(1)|C|Def.|990|;"
            'Los campos que no se ven que van FUERA DEL GRID
            SQL = SQL & "N||||0|;"
            arregla SQL, DataGrid3, Me, 330
            DataGrid3.ScrollBars = dbgAutomatic
            
                            
        End If
        
        
End Sub




Private Function InsertarModificarLinea() As Boolean
Dim I As Integer
    InsertarModificarLinea = False
    
    BuscaChekc = ""
    If Modo = 5 Then
        For I = 0 To 1
            Text3(I).Text = Trim(Text3(I).Text)
            If Text3(I).Text = "" And I = 0 Then BuscaChekc = "B"
        Next I
        If Me.cboDefecto(0).ListIndex = -1 Then BuscaChekc = "N"
    Else
        For I = 0 To 1
            Text4(I).Text = Trim(Text4(I).Text)
            If Text4(I).Text = "" And I = 0 Then BuscaChekc = "B"
        Next I
        If Me.cboDefecto(1).ListIndex = -1 Then BuscaChekc = "N"
    End If
    If BuscaChekc <> "" Then
        MsgBox "Campos obligarorios", vbExclamation
        Exit Function
    End If
    
    
    
    'Un par de comprobaciones
    If Modo = 6 And ModificaLineas = 1 Then
        BuscaChekc = DevuelveDesdeBD(conAri, "codenvio", "smatriculas", "matricula", Text4(0).Text, "T")
        If BuscaChekc <> "" Then
            BuscaChekc = DevuelveDesdeBD(conAri, "concat(codenvio,' ',nomenvio)", "senvio", "codenvio", BuscaChekc)
            MsgBox "Ya existe la matricula en otro transportista" & vbCrLf & vbCrLf & BuscaChekc, vbExclamation
            Exit Function
        End If
    End If
    
    
    If ModificaLineas = 2 Then
    
        If Modo = 6 Then
            BuscaChekc = "UPDATE smatriculas set titulo =" & DBSet(Text4(1).Text, "T") & ", defecto =" & Val(cboDefecto(1).ListIndex)
            BuscaChekc = BuscaChekc & " WHERE matricula=" & DBSet(Text4(0).Text, "T")
            CadB = " matricula <> " & DBSet(Text4(0).Text, "T")
        Else
    ',codenvio,nombre,defecto) VALUES (" & NumRegElim & ","
            BuscaChekc = "UPDATE sconductor SET  dni =" & DBSet(Text3(0).Text, "T") & ", nombre = "
            BuscaChekc = BuscaChekc & DBSet(Text3(1).Text, "T") & ", defecto =" & Val(cboDefecto(0).ListIndex)
            BuscaChekc = BuscaChekc & " WHERE chofer= " & Data2.Recordset!Chofer
            CadB = " chofer <> " & Data2.Recordset!Chofer
        End If
    
    
    Else
        If Modo = 6 Then
            BuscaChekc = "INSERT INTO smatriculas(matricula,codenvio,titulo,defecto) VALUES ("
            BuscaChekc = BuscaChekc & DBSet(Text4(0).Text, "T") & "," & adodc1.Recordset!CodEnvio & ","
            BuscaChekc = BuscaChekc & DBSet(Text4(1).Text, "T") & "," & Val(cboDefecto(1).ListIndex) & ")"
            CadB = " matricula <> " & DBSet(Text4(0).Text, "T")
        Else
            BuscaChekc = DevuelveDesdeBD(conAri, "max(chofer)", "sconductor", "1", "1")
            NumRegElim = Val(BuscaChekc) + 1
            BuscaChekc = "INSERT INTO sconductor(chofer,dni,codenvio,nombre,defecto) VALUES (" & NumRegElim & ","
            BuscaChekc = BuscaChekc & DBSet(Text3(0).Text, "T") & "," & adodc1.Recordset!CodEnvio & ","
            BuscaChekc = BuscaChekc & DBSet(Text3(1).Text, "T") & "," & Val(cboDefecto(0).ListIndex) & ")"
            CadB = " chofer <> " & NumRegElim
        End If
    
    End If
    
    If ejecutar(BuscaChekc, False) Then
        InsertarModificarLinea = True
        If cboDefecto(IIf(Modo = 6, 1, 0)).ListIndex = 1 Then
            'Por defecto marcado
            BuscaChekc = "UPDATE " & IIf(Modo = 6, "smatriculas", "sconductor") & " SET defecto =  0 WHERE "
            BuscaChekc = BuscaChekc & CadB & " AND codenvio = " & adodc1.Recordset!CodEnvio
            ejecutar BuscaChekc, True
        End If
    End If
    CadB = ""
End Function

