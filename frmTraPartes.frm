VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmADVTraPartes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   8220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12945
   Icon            =   "frmTraPartes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   12945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Digitos 1er nivel|N|N|||empresa|numdigi1|||"
   Begin VB.Frame FrameMultparte 
      Height          =   8175
      Left            =   0
      TabIndex        =   67
      Top             =   2520
      Width           =   12975
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
         Index           =   14
         Left            =   1560
         MultiLine       =   -1  'True
         TabIndex        =   83
         Text            =   "frmTraPartes.frx":000C
         Top             =   1800
         Width           =   5580
      End
      Begin MSComctlLib.ListView lwC 
         Height          =   5295
         Left            =   240
         TabIndex        =   82
         Top             =   2280
         Width           =   12375
         _ExtentX        =   21828
         _ExtentY        =   9340
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Codclien"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   6068
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Campo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Variedad"
            Object.Width           =   3422
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Partida"
            Object.Width           =   4304
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "NroCampo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "codvarie"
            Object.Width           =   0
         EndProperty
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
         Height          =   315
         Index           =   13
         Left            =   240
         TabIndex        =   71
         Tag             =   "Trabajador|N|N|||advpartes|codtraba|0||"
         Text            =   "Text1"
         Top             =   1800
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
         Index           =   12
         Left            =   2040
         TabIndex        =   69
         Tag             =   "Trabajador|N|N|||advpartes|codtraba|0||"
         Text            =   "Text1"
         Top             =   1020
         Width           =   900
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
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
         Height          =   360
         Index           =   12
         Left            =   3000
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   79
         Text            =   "Text2"
         Top             =   1020
         Width           =   4125
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Empresa externa"
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
         Index           =   3
         Left            =   7440
         TabIndex        =   78
         Tag             =   "Ext|N|N|||advpartes|EsExterno|0||"
         Top             =   1800
         Width           =   2205
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
         Left            =   7440
         TabIndex        =   70
         Tag             =   "Trabajador|N|N|||advpartes|codtraba|0||"
         Text            =   "Text1"
         Top             =   1020
         Width           =   900
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
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
         Height          =   360
         Index           =   11
         Left            =   8400
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   76
         Text            =   "Text2"
         Top             =   1020
         Width           =   4125
      End
      Begin VB.CommandButton cmdMultiParte 
         Caption         =   "&Cerrar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   11160
         TabIndex        =   73
         Top             =   7680
         Width           =   1335
      End
      Begin VB.CommandButton cmdMultiParte 
         Caption         =   "&Generar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   9600
         TabIndex        =   72
         Top             =   7680
         Width           =   1335
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
         Left            =   240
         MaxLength       =   10
         TabIndex        =   68
         Top             =   1020
         Width           =   1425
      End
      Begin VB.Image imgCampos 
         Height          =   240
         Index           =   2
         Left            =   10560
         Picture         =   "frmTraPartes.frx":0011
         ToolTipText     =   "Eliminar campo"
         Top             =   1920
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   7
         Left            =   3120
         Picture         =   "frmTraPartes.frx":0A13
         ToolTipText     =   "Buscar cliente varios"
         Top             =   1500
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
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
         Height          =   240
         Index           =   14
         Left            =   1560
         TabIndex        =   84
         Top             =   1500
         Width           =   1440
      End
      Begin VB.Image imgCampos 
         Height          =   240
         Index           =   0
         Left            =   10200
         Picture         =   "frmTraPartes.frx":0B15
         ToolTipText     =   "Agregar por numero de campo"
         Top             =   1920
         Width           =   240
      End
      Begin VB.Image imgCampos 
         Height          =   240
         Index           =   1
         Left            =   11160
         Picture         =   "frmTraPartes.frx":7367
         ToolTipText     =   "Eliminar por NUMERO de campo"
         Top             =   1920
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   0
         Left            =   11760
         Picture         =   "frmTraPartes.frx":DBB9
         ToolTipText     =   "Quitar seleccion"
         Top             =   1920
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   1
         Left            =   12120
         Picture         =   "frmTraPartes.frx":DD03
         ToolTipText     =   "Seleccionar todos"
         Top             =   1920
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Cantidad"
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
         Index           =   13
         Left            =   240
         TabIndex        =   81
         Top             =   1500
         Width           =   1065
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   6
         Left            =   3240
         Picture         =   "frmTraPartes.frx":DE4D
         ToolTipText     =   "Buscar cliente varios"
         Top             =   720
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Trabajador"
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
         Left            =   2040
         TabIndex        =   80
         Top             =   720
         Width           =   1065
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   5
         Left            =   8760
         Picture         =   "frmTraPartes.frx":DF4F
         ToolTipText     =   "Buscar cliente varios"
         Top             =   720
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Tratamiento"
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
         Index           =   11
         Left            =   7440
         TabIndex        =   77
         Top             =   720
         Width           =   1200
      End
      Begin VB.Label Label2 
         Caption         =   "Generación multiples partes de trabajo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         TabIndex        =   75
         Top             =   120
         Width           =   8295
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   1
         Left            =   960
         Picture         =   "frmTraPartes.frx":E051
         ToolTipText     =   "Buscar fecha"
         Top             =   720
         Width           =   240
      End
      Begin VB.Label Label1 
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
         Height          =   240
         Index           =   10
         Left            =   240
         TabIndex        =   74
         Top             =   720
         Width           =   600
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Empresa externa"
      Height          =   195
      Index           =   2
      Left            =   3480
      TabIndex        =   10
      Tag             =   "Ext|N|N|||advpartes|EsExterno|0||"
      Top             =   2880
      Width           =   1845
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   315
      Index           =   9
      Left            =   9480
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   65
      Text            =   "Text2"
      Top             =   1440
      Width           =   3165
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   9
      Left            =   7800
      MaxLength       =   20
      TabIndex        =   5
      Tag             =   "Flota|T|S|||advpartes|codflota|||"
      Text            =   "Text1"
      Top             =   1440
      Width           =   1620
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   315
      Index           =   8
      Left            =   8745
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   63
      Text            =   "Text2"
      Top             =   885
      Width           =   4005
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   8
      Left            =   7785
      MaxLength       =   4
      TabIndex        =   3
      Tag             =   "Trabajador|N|N|||advpartes|codtraba|0||"
      Text            =   "Text1"
      Top             =   885
      Width           =   900
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   315
      Index           =   7
      Left            =   2400
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   61
      Text            =   "Text2"
      Top             =   1920
      Width           =   4005
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   7
      Left            =   1440
      MaxLength       =   6
      TabIndex        =   6
      Tag             =   "Coddirec|N|S|0|9999|advpartes|coddirec|||"
      Text            =   "Text1"
      Top             =   1920
      Width           =   900
   End
   Begin VB.TextBox txtAux2 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   2
      Left            =   11040
      MaxLength       =   18
      TabIndex        =   37
      Tag             =   "Código Artículo"
      Text            =   "Artic Artic Artic5"
      Top             =   2400
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtAux2 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   1
      Left            =   9240
      MaxLength       =   18
      TabIndex        =   36
      Tag             =   "Código Artículo"
      Text            =   "Artic Artic Artic5"
      Top             =   2400
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdCampo 
      Caption         =   "+"
      Height          =   375
      Left            =   10320
      TabIndex        =   35
      Top             =   2280
      Visible         =   0   'False
      Width           =   255
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3255
      Left            =   120
      TabIndex        =   34
      Top             =   4440
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   5741
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Artículos"
      TabPicture(0)   =   "frmTraPartes.frx":E0DC
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(5)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(6)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "DataGrid1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtAux(6)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtAux(2)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtAux(5)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtAux(4)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtAux(3)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtAux(1)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtAux(0)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdAux(0)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmdAux(1)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtAux(11)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtAux(10)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtAux(7)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtAux(8)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtAux(9)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Combo1"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).ControlCount=   18
      TabCaption(1)   =   "Trabajadores"
      TabPicture(1)   =   "frmTraPartes.frx":E0F8
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DataGrid3"
      Tab(1).Control(1)=   "txtAuxT(2)"
      Tab(1).Control(2)=   "txtAuxT(0)"
      Tab(1).Control(3)=   "txtAuxT(1)"
      Tab(1).Control(4)=   "cmdTra"
      Tab(1).ControlCount=   5
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmTraPartes.frx":E114
         Left            =   11760
         List            =   "frmTraPartes.frx":E11E
         Style           =   2  'Dropdown List
         TabIndex        =   46
         Top             =   2160
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdTra 
         Caption         =   "+"
         Height          =   375
         Left            =   -73680
         TabIndex        =   58
         Top             =   840
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox txtAuxT 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   1
         Left            =   -73320
         MaxLength       =   18
         TabIndex        =   59
         Tag             =   "Código Artículo"
         Text            =   "c"
         Top             =   840
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtAuxT 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   0
         Left            =   -74400
         MaxLength       =   18
         TabIndex        =   57
         Tag             =   "Código Artículo"
         Text            =   "n"
         Top             =   840
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtAuxT 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   2
         Left            =   -73080
         MaxLength       =   18
         TabIndex        =   60
         Tag             =   "Código Artículo"
         Text            =   "h"
         Top             =   1320
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   9
         Left            =   10320
         MaxLength       =   10
         TabIndex        =   45
         Tag             =   "Bultos"
         Text            =   "42"
         Top             =   2160
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   8
         Left            =   9840
         MaxLength       =   5
         TabIndex        =   44
         Tag             =   "Bultos"
         Text            =   "dtop2"
         Top             =   2160
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   7
         Left            =   9240
         MaxLength       =   5
         TabIndex        =   48
         Tag             =   "Bultos"
         Text            =   "42"
         Top             =   2160
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtAux 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   10
         Left            =   840
         MaxLength       =   12
         TabIndex        =   54
         Tag             =   "Importe"
         Text            =   "Importe"
         Top             =   2760
         Width           =   3135
      End
      Begin VB.TextBox txtAux 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   11
         Left            =   5880
         MaxLength       =   60
         TabIndex        =   47
         Tag             =   "Importe"
         Text            =   "Importe"
         Top             =   2760
         Width           =   6495
      End
      Begin VB.CommandButton cmdAux 
         Appearance      =   0  'Flat
         Caption         =   "+"
         Height          =   315
         Index           =   1
         Left            =   2400
         TabIndex        =   50
         ToolTipText     =   "Buscar artículo"
         Top             =   2220
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.CommandButton cmdAux 
         Appearance      =   0  'Flat
         Caption         =   "+"
         Height          =   315
         Index           =   0
         Left            =   600
         TabIndex        =   51
         ToolTipText     =   "Buscar almacen"
         Top             =   2160
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   0
         Left            =   720
         MaxLength       =   15
         TabIndex        =   39
         Tag             =   "Código Almacen"
         Text            =   "codalmac"
         Top             =   2160
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   1
         Left            =   1080
         MaxLength       =   18
         TabIndex        =   40
         Tag             =   "Código Artículo"
         Text            =   "Artic Artic Artic5"
         Top             =   2160
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   3
         Left            =   6000
         MaxLength       =   16
         TabIndex        =   41
         Tag             =   "Cantidad"
         Text            =   "1,234,567,891.25"
         Top             =   2220
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   4
         Left            =   7440
         MaxLength       =   12
         TabIndex        =   42
         Tag             =   "Precio"
         Text            =   "123,456.7879"
         Top             =   2220
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   5
         Left            =   8160
         MaxLength       =   12
         TabIndex        =   43
         Tag             =   "Importe"
         Text            =   "ad"
         Top             =   2160
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   2
         Left            =   2760
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   52
         Tag             =   "Nombre Artículo"
         Text            =   "nomArtic"
         Top             =   2220
         Visible         =   0   'False
         Width           =   3165
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   6
         Left            =   10680
         MaxLength       =   5
         TabIndex        =   49
         Tag             =   "Bultos"
         Text            =   "12345"
         Top             =   2220
         Visible         =   0   'False
         Width           =   495
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   2280
         Left            =   120
         TabIndex        =   38
         Top             =   360
         Width           =   12315
         _ExtentX        =   21722
         _ExtentY        =   4022
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
      Begin MSDataGridLib.DataGrid DataGrid3 
         Height          =   2535
         Left            =   -72120
         TabIndex        =   56
         Top             =   480
         Width           =   6855
         _ExtentX        =   12091
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
      Begin VB.Label Label1 
         Caption         =   "Almacén"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   55
         Top             =   2760
         Width           =   720
      End
      Begin VB.Label Label1 
         Caption         =   "Ampliación"
         Height          =   255
         Index           =   5
         Left            =   4800
         TabIndex        =   53
         Top             =   2760
         Width           =   840
      End
   End
   Begin VB.TextBox txtAux2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   0
      Left            =   7680
      MaxLength       =   18
      TabIndex        =   32
      Tag             =   "Código Artículo"
      Text            =   "Artic Artic Artic5"
      Top             =   2400
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   975
      Index           =   2
      Left            =   1440
      MaxLength       =   200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Tag             =   "Observaciones|T|S|||advpartes|observac|||"
      Top             =   3240
      Width           =   4935
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Cerrado"
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
      Index           =   1
      Left            =   1440
      TabIndex        =   9
      Tag             =   "Cerrado|N|N|||advpartes|cerrado|0||"
      Top             =   2880
      Width           =   1005
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   4
      Left            =   4320
      MaxLength       =   7
      TabIndex        =   8
      Tag             =   "Litros Reales|N|S|0||advpartes|litrosrea|0||"
      Top             =   2400
      Width           =   1065
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   5
      Left            =   1440
      MaxLength       =   7
      TabIndex        =   7
      Tag             =   "Litros Previstos|N|N|0|999999|advpartes|litrospre|0||"
      Top             =   2400
      Width           =   1065
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000013&
      Height          =   315
      Index           =   0
      Left            =   1440
      MaxLength       =   7
      TabIndex        =   0
      Tag             =   "NºParte|N|S|||advpartes|numparte|0000000|S|"
      Text            =   "Text1 7"
      Top             =   480
      Width           =   885
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   3
      Left            =   1440
      MaxLength       =   4
      TabIndex        =   2
      Tag             =   "Cod.Tratamiento|T|N|||advpartes|codtrata|||"
      Text            =   "Text1"
      Top             =   960
      Width           =   900
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   315
      Index           =   3
      Left            =   2400
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   21
      Text            =   "Text2"
      Top             =   960
      Width           =   4005
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   1
      Left            =   3360
      MaxLength       =   10
      TabIndex        =   1
      Tag             =   "Fecha Parte|F|N|||advpartes|fechapar|dd/mm/yyyy|N|"
      Top             =   480
      Width           =   1065
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   6
      Left            =   1440
      MaxLength       =   6
      TabIndex        =   4
      Tag             =   "Socio|N|N|0|999999|advpartes|codclien|000000||"
      Text            =   "Text1"
      Top             =   1440
      Width           =   900
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   315
      Index           =   6
      Left            =   2400
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   20
      Text            =   "Text2"
      Top             =   1440
      Width           =   4005
   End
   Begin MSAdodcLib.Adodc Data2 
      Height          =   375
      Left            =   3120
      Top             =   7680
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
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   11400
      TabIndex        =   15
      Top             =   7800
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Left            =   120
      TabIndex        =   16
      Top             =   7680
      Width           =   2655
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   2355
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   11400
      TabIndex        =   14
      Top             =   7800
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   10080
      TabIndex        =   13
      Top             =   7800
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   375
      Left            =   3480
      Top             =   7680
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
      TabIndex        =   18
      Top             =   0
      Width           =   12945
      _ExtentX        =   22834
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
            Object.ToolTipText     =   "Lineas Campos"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Lineas articulos"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Lineas trabajadores"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Generar albaran"
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir "
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir multiple"
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
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
         Left            =   11640
         TabIndex        =   19
         Top             =   120
         Width           =   1215
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Height          =   1935
      Left            =   6600
      TabIndex        =   31
      Top             =   2280
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   3413
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
   Begin MSAdodcLib.Adodc data3 
      Height          =   615
      Left            =   5880
      Top             =   7560
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1085
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
   Begin MSAdodcLib.Adodc data4 
      Height          =   615
      Left            =   4560
      Top             =   7680
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1085
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
   Begin VB.CheckBox Check1 
      Caption         =   "Facturar"
      Height          =   195
      Index           =   0
      Left            =   6600
      TabIndex        =   12
      Tag             =   "Facturar S/N|N|N|||advpartes|factursn|0||"
      Top             =   2280
      Width           =   1005
   End
   Begin VB.Label Label1 
      Caption         =   "Vehiculo"
      Height          =   255
      Index           =   9
      Left            =   6600
      TabIndex        =   66
      Top             =   1485
      Width           =   840
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   4
      Left            =   7440
      Picture         =   "frmTraPartes.frx":E12A
      ToolTipText     =   "Buscar cliente varios"
      Top             =   1440
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Trabajador"
      Height          =   255
      Index           =   8
      Left            =   6600
      TabIndex        =   64
      Top             =   930
      Width           =   840
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   3
      Left            =   7440
      Picture         =   "frmTraPartes.frx":E22C
      ToolTipText     =   "Buscar cliente varios"
      Top             =   885
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Departamento"
      Height          =   195
      Index           =   7
      Left            =   120
      TabIndex        =   62
      Top             =   1965
      Width           =   1005
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   2
      Left            =   1200
      Picture         =   "frmTraPartes.frx":E32E
      ToolTipText     =   "Buscar cliente varios"
      Top             =   1920
      Width           =   240
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   0
      Left            =   1200
      Picture         =   "frmTraPartes.frx":E430
      ToolTipText     =   "Buscar cliente varios"
      Top             =   960
      Width           =   240
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   1
      Left            =   1200
      Picture         =   "frmTraPartes.frx":E532
      ToolTipText     =   "Buscar cliente varios"
      Top             =   1440
      Width           =   240
   End
   Begin VB.Label Label29 
      Caption         =   "Campos"
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
      Index           =   1
      Left            =   6600
      TabIndex        =   33
      Top             =   2040
      Width           =   675
   End
   Begin VB.Label Label29 
      Caption         =   "Observaciones"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   30
      Top             =   3240
      Width           =   1125
   End
   Begin VB.Label Label1 
      Caption         =   "Nº parte"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   29
      Top             =   480
      Width           =   840
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha"
      Height          =   255
      Index           =   3
      Left            =   2520
      TabIndex        =   28
      Top             =   480
      Width           =   540
   End
   Begin VB.Label Label1 
      Caption         =   "NºParte"
      Height          =   255
      Index           =   28
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Tratamiento"
      Height          =   255
      Index           =   0
      Left            =   135
      TabIndex        =   26
      Top             =   1005
      Width           =   840
   End
   Begin VB.Image imgFec 
      Height          =   240
      Index           =   0
      Left            =   3120
      Picture         =   "frmTraPartes.frx":E634
      ToolTipText     =   "Buscar fecha"
      Top             =   480
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha "
      Height          =   255
      Index           =   29
      Left            =   1200
      TabIndex        =   25
      Top             =   30
      Width           =   585
   End
   Begin VB.Label Label1 
      Caption         =   "Cliente"
      Height          =   255
      Index           =   17
      Left            =   120
      TabIndex        =   24
      Top             =   1485
      Width           =   540
   End
   Begin VB.Label Label1 
      Caption         =   "Cantidad real"
      Height          =   255
      Index           =   1
      Left            =   3240
      TabIndex        =   23
      Top             =   2430
      Width           =   1110
   End
   Begin VB.Label Label1 
      Caption         =   "Cantidad prevista"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   22
      Top             =   2430
      Width           =   1500
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
         Caption         =   "Mantenimiento &lineas"
         Begin VB.Menu mnLineas1 
            Caption         =   "Campos"
            Index           =   0
         End
         Begin VB.Menu mnLineas1 
            Caption         =   "Articulos"
            Index           =   1
         End
         Begin VB.Menu mnLineas1 
            Caption         =   "Trabajadores"
            Index           =   2
         End
      End
      Begin VB.Menu mnBarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu mnFiltro 
      Caption         =   "Filtro(V.Todos)"
      Begin VB.Menu mnFiltro1 
         Caption         =   "Todos"
         Index           =   0
      End
      Begin VB.Menu mnFiltro1 
         Caption         =   "Abiertos"
         Index           =   1
      End
      Begin VB.Menu mnFiltro1 
         Caption         =   "Cerrados"
         Index           =   2
      End
      Begin VB.Menu mnFiltro1 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnFiltro1 
         Caption         =   "Cualquier trabajador"
         Index           =   4
      End
      Begin VB.Menu mnFiltro1 
         Caption         =   "Trabajador conectado"
         Index           =   5
      End
   End
End
Attribute VB_Name = "frmADVTraPartes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

Private WithEvents frmB As frmBuscaGrid 'Form para busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmTra As frmADVTratamientos
Attribute frmTra.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents FrmArt As frmBasico2
Attribute FrmArt.VB_VarHelpID = -1
Private WithEvents frmAlm As frmAlmAlPropios
Attribute frmAlm.VB_VarHelpID = -1
Private WithEvents frmCli As frmBasico2
Attribute frmCli.VB_VarHelpID = -1
Private WithEvents frmList As frmListadoPed 'Listados para Pedidos (pasar pedido a albaran)
Attribute frmList.VB_VarHelpID = -1
Private WithEvents frmDptoEnvio As frmFacCliEnvDpto
Attribute frmDptoEnvio.VB_VarHelpID = -1
Private WithEvents frmT As frmBasico2 'frmAdmTrabajadores
Attribute frmT.VB_VarHelpID = -1
Private WithEvents frmF As frmFlotas
Attribute frmF.VB_VarHelpID = -1


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

Private CadenaConsulta2 As String
Private Ordenacion As String
Private NombreTabla As String  'Nombre de la tabla o de la
Private kCampo As Integer

Dim btnPrimero As Byte
'Variable que indica el número del Boton  PrimerRegistro en la Toolbar1

'-------------------------------------------------------------------------
'Private HaDevueltoDatos As Boolean
Dim CadenaDevuelta2 As String
Dim D As String  'string varios usos


'Pasar a albaran
Dim CodZona As Integer
Dim CadenaSQL As String
Dim FechaAlb As String




'***************************************
' Geslab---> Aripres.      Febrero 2014
'       Esta en MYSQL.
'       Ya no necesitamos la conexion

'Variables
'Private connPres As Connection
Private BDAripres As String


Dim cCli As CCliente

'Dim ArticulosAgrupados As String
'Dim PorGrupo_(2) As String  'Que articulos pertenecen a cada grupo

Dim ArtiHORAS As String 'articulo horas, NO permite ponerlo en lineas de articulos

Dim B As Boolean
Dim ALbInterno As Boolean
Dim FacturacionColectiva As Boolean
Dim BuscaChekc As String
Dim AlmacenLin As Integer


Private Sub Check1_Click(Index As Integer)
     If Modo = 1 Then CheckCadenaBusqueda Check1(Index), BuscaChekc
End Sub

Private Sub Check1_GotFocus(Index As Integer)
    ConseguirfocoChk Modo
End Sub

Private Sub Check1_KeyPress(Index As Integer, KeyAscii As Integer)
        KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()
Dim numlinea As Integer
Dim CambiaLitros As Boolean
    
    Screen.MousePointer = vbHourglass
    On Error GoTo Error1
    Select Case Modo
        Case 1  'BUSQUEDA
            HacerBusqueda
            
        Case 3 'INSERTAR
            If DatosOk Then
                If InsertarDesdeForm(Me) Then
                    InsertaReferenciaTratamientos
                    CadenaConsulta2 = "select * from " & NombreTabla & " WHERE numparte = " & Text1(0).Text & Ordenacion
                    Data1.RecordSource = CadenaConsulta2
                    PosicionarData
                    PonerCampos
                    'Pasamos a meter los campos
                    mnLineas1_Click 0
                    BotonAnyadirLinea
                End If
            End If
        
        Case 4  'MODIFICAR
            If DatosOk Then
                CambiaLitros = False
                If Val(ImporteFormateado(Me.Text1(5).Text)) <> Data1.Recordset!litrospre Then CambiaLitros = True
                
                If ModificaDesdeFormulario(Me, 1) Then
                    If CambiaLitros Then UpdateaCambioLitros
                    TerminaBloquear
                    PosicionarData
                End If
            End If
            
        Case 5, 6, 7
            
            
            If ModificaLineas = 1 Then 'INSERTAR lineas de algo
                numlinea = 1
                If Modo = 5 Then
                    
                    If Data2.Recordset.EOF Then numlinea = 0
                ElseIf Modo = 6 Then
                    
                    If data3.Recordset.EOF Then numlinea = 0
                    
                Else
                    'modo=7
                    If data4.Recordset.EOF Then numlinea = 0
                End If
                
                If InsertarModificar() Then
                    If Modo = 5 Then
                        If numlinea = 0 Then
                            CargaGrid True, 1
                        Else
                            Data2.Refresh
                            CargaGrid2 DataGrid2, Data2   'solo caraga el suyo
                        End If
                    ElseIf Modo = 6 Then
                        If numlinea = 0 Then
                            CargaGrid True, 2
                        Else
                            data3.Refresh
                            CargaGridArt DataGrid1, data3     'solo caraga el suyo
                            DataGrid1.Enabled = True
                        End If
                        
                    Else
                        If numlinea = 0 Then
                            CargaGrid True, 3
                        Else
                            data4.Refresh
                            CargaGridTrab DataGrid3, data4     'solo caraga el suyo
                        End If
                    
                    
                    End If
                    BotonAnyadirLinea
                End If
                
            Else
                If InsertarModificar() Then
                    TerminaBloquear
                    If Modo = 6 Then
                        NumRegElim = Val(data3.Recordset!numlinea)
                        CargaTxtAuxArt False, False
                        data3.Refresh
                        CargaGridArt DataGrid1, data3
                        
                    Else
                        If Modo = 7 Then
                            NumRegElim = Val(data4.Recordset!numlinea)
                            CargatxtAuxT False, False
                            data4.Refresh
                            CargaGridTrab DataGrid3, data4
                        End If
                    End If
                    PosicionarData3
                    ModificaLineas = 0
                    PonerBotonCabecera True
                    If Modo = 6 Then
                        Me.DataGrid1.Enabled = True
                    ElseIf Modo = 7 Then
                        Me.DataGrid3.Enabled = True
                    End If
                End If
            End If
        
    End Select
        
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub cmdAux_Click(Index As Integer)
    CadenaDevuelta2 = ""
    Select Case Index
        Case 0 'Busqueda de Cod. Almacen
            
            Set frmAlm = New frmAlmAlPropios
            frmAlm.DatosADevolverBusqueda = "0"
            frmAlm.Show vbModal
            Set frmAlm = Nothing
            If CadenaDevuelta2 <> "" Then
                Me.txtAux(0).Text = RecuperaValor(CadenaDevuelta2, 1)
                Me.txtAux(10).Text = RecuperaValor(CadenaDevuelta2, 2)
            End If
            PonerFoco txtAux(Index)
            
        Case 1 'Busqueda de Cod. Artic
            Set FrmArt = New frmBasico2
'            FrmArt.DesdeTPV = False
'            FrmArt.Show vbModal
            AyudaArticulos FrmArt, txtAux(Index)
            Set FrmArt = Nothing
            If CadenaDevuelta2 <> "" Then
                Me.txtAux(1).Text = RecuperaValor(CadenaDevuelta2, 1)
                Me.txtAux(2).Text = RecuperaValor(CadenaDevuelta2, 2)
                Me.txtAux(5).Text = ""
                Me.txtAux(9).Text = ""
            End If
            PonerFoco txtAux(Index)
    End Select
End Sub

Private Sub cmdCampo_Click()
    If Modo <> 5 Then Exit Sub
    If ModificaLineas <> 1 Then Exit Sub
    
    CadenaDesdeOtroForm = ""
    frmADVvarios.Opcion = 0
    frmADVvarios.vCampos = Text1(6).Text
    frmADVvarios.Show vbModal
    If CadenaDesdeOtroForm <> "" Then
        
        If Mid(CadenaDesdeOtroForm, 1, 1) = "@" Then
            
            MultiInsercionCampos False
        Else
            'quito el ultimo ·
            CadenaDesdeOtroForm = Mid(CadenaDesdeOtroForm, 1, Len(CadenaDesdeOtroForm) - 1)
        
            'Han seleccionado un campo
            txtAux2(0).Text = RecuperaValor(CadenaDesdeOtroForm, 1)
            txtAux2(1).Text = RecuperaValor(CadenaDesdeOtroForm, 2)
            txtAux2(2).Text = RecuperaValor(CadenaDesdeOtroForm, 3)
            CadenaDesdeOtroForm = ""
        End If
    End If
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
            
        Case 5, 6, 7
            TerminaBloquear
            If Modo = 5 Then
                CargaTxtAux2 False, False
                If ModificaLineas = 1 Then 'INSERTAR
                    DataGrid2.AllowAddNew = False
                    ModificaLineas = 0  'Fuerzo el cero para que carge la ampliacion
                    If Not Data2.Recordset.EOF Then Data2.Recordset.MoveFirst
                    
                End If
                ModificaLineas = 0
            
                Me.DataGrid2.Enabled = True
            ElseIf Modo = 6 Then
                CargaTxtAuxArt False, False
                If ModificaLineas = 1 Then 'INSERTAR
                    DataGrid1.AllowAddNew = False
                    ModificaLineas = 0  'Fuerzo el cero para que carge la ampliacion
                    If Not data3.Recordset.EOF Then data3.Recordset.MoveFirst
                    
                End If
                ModificaLineas = 0
                
                DataGrid1.Enabled = True
            Else
                'Trabakadores
                CargatxtAuxT False, False
                If ModificaLineas = 1 Then 'INSERTAR
                    DataGrid3.AllowAddNew = False
                    ModificaLineas = 0  'Fuerzo el cero para que carge la ampliacion
                    If Not data4.Recordset.EOF Then data4.Recordset.MoveFirst
                    
                End If
                ModificaLineas = 0
            
                DataGrid3.Enabled = True
                
            End If
            PonerBotonCabecera True
    End Select
    PonerFoco Text1(0)
End Sub


Private Sub BotonAnyadir()
    LimpiarCampos
    CargaGrid False, 0
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 3
    Text1(1).Text = Format(Now, "dd/mm/yyyy")
    
    
    
    Text1(8).Text = PonerTrabajadorConectado(CadenaDevuelta2)
    Text2(8).Text = CadenaDevuelta2
    
    
    FormateaCampo Text1(0)
    BloquearTxt Text1(0), True, True
    PonerFoco Text1(1)
End Sub


Private Sub BotonBuscar()
    If Modo <> 1 Then
        LimpiarCampos
        CargaGrid False, 0
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
Dim Aux As String

'Ver todos
    LimpiarCampos
    CargaGrid False, 0
    
    'Filtro abiertos-cerrados
    If mnFiltro1(0).Checked Then
        Aux = ""
    Else
        Aux = "cerrado = "
        If mnFiltro1(1).Checked Then
            Aux = Aux & "0"
        Else
            Aux = Aux & "1"
        End If
    End If
    'Filtro trabajador
    If mnFiltro1(5).Checked Then
        If Aux <> "" Then Aux = Aux & " AND "
        Aux = Aux & "codtraba = " & mnFiltro1(5).Tag
    End If
    
    If chkVistaPrevia.Value = 1 Then
        MandaBusquedaPrevia Aux
    Else
        CadenaConsulta2 = "Select * from " & NombreTabla
        If Aux <> "" Then CadenaConsulta2 = CadenaConsulta2 & " WHERE " & Aux
        
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
    
    
    'Da fallos. Pongo los litros
    'text1(5).Text=dblet(data1.Recordset!litrospre
End Sub


Private Sub BotonEliminar()
Dim Cad As String

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    If Modo <> 2 Then Exit Sub
    If Data1.Recordset.EOF Then Exit Sub
    If Val(Data1.Recordset!cerrado) = 1 Then
        MsgBox "Parte cerrado", vbExclamation
        Exit Sub
    End If
    
    'El registro de codigo 0 no se puede Modificar ni Eliminar
    'If EsCodigoCero(CStr(Data1.Recordset.Fields(0).Value), FormatoCampo(Text1(0))) Then Exit Sub
    
    
    '### a mano
    Cad = "Va a eliminar el parte:" & vbCrLf
    Cad = Cad & vbCrLf & "ID. parte : " & Format(Data1.Recordset.Fields(0), FormatoCampo(Text1(0)))
    Cad = Cad & vbCrLf & "Cliente.: " & Me.Text2(6).Text & vbCrLf & "¿Continuar?"

    'Borramos
    If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        On Error GoTo Error2
        Screen.MousePointer = vbHourglass
        NumRegElim = Data1.Recordset.AbsolutePosition
        'Borramos en advpartes_campos    advpartes_trabajador     advparteslineas
        Cad = "Delete from advpartes_campos where numparte = " & Data1.Recordset!numparte
        conn.Execute Cad
        
        Cad = "Delete from advpartes_trabajador where numparte = " & Data1.Recordset!numparte
        conn.Execute Cad
        
        Cad = "Delete from advparteslineas where numparte = " & Data1.Recordset!numparte
        conn.Execute Cad
        
        Data1.Recordset.Delete
        
 
        
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
        Data1.Recordset.CancelUpdate
        MuestraError Err.Number, "Eliminar Familia de Articulo", Err.Description
    End If
End Sub


Private Sub cmdMultiParte_Click(Index As Integer)
Dim B As Boolean
    If Index = 0 Then
        
        If Not datosOk_Multi Then Exit Sub
        
        Screen.MousePointer = vbHourglass
        conn.BeginTrans
        B = GenerarPartes
        If B Then
            conn.CommitTrans
        Else
            conn.RollbackTrans
        End If
        
        
        If B Then BotonImprimirNuevo2 2  'Imprimimos los partes seleccionados
        
        Screen.MousePointer = vbDefault
        If Not B Then Exit Sub
    Else
        
    End If
    POnerMultiParte False
    If Me.Data1.Recordset.EOF Then
        PonerModo 0
    Else
    
        PonerModo 2
        PonerCampos
    End If
End Sub

Private Sub cmdRegresar_Click()
Dim Cad As String

    If Modo >= 5 Then  'modo 5: Mantenimientos Lineas
    
    
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


Private Sub cmdTra_Click()
    If Modo <> 7 Then Exit Sub
    Set frmB = New frmBuscaGrid
    
    If BDAripres = "" Then
        frmB.vCampos = "Codigo|straba|codtraba|N||20·Nombre|straba|nomtraba|T||70·"
        frmB.vTabla = "straba"
    Else
        frmB.vCampos = "Codigo|trabajadores|idtrabajador|N||20·Nombre|trabajadores|nomtrabajador|T||60·Seccion|trabajadores|seccion|N||10·"
        frmB.vTabla = BDAripres & ".trabajadores"
    End If
    frmB.vSQL = ""
    CadenaDevuelta2 = ""
    '###A mano
    frmB.vDevuelve = "0|1|"
    frmB.vTitulo = "Trabajadores"
    frmB.vselElem = 1
    frmB.vConexionGrid = conAri
    
    frmB.Show vbModal
    Set frmB = Nothing
    
    If CadenaDevuelta2 <> "" Then
        txtAuxT(0).Text = RecuperaValor(CadenaDevuelta2, 1)
        txtAuxT(1).Text = RecuperaValor(CadenaDevuelta2, 2)
        CadenaDevuelta2 = ""
        PonerFoco Me.txtAuxT(2)
    End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
     KEYpress KeyAscii
End Sub

Private Sub data3_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    txtAux(10).Text = "": txtAux(11).Text = ""
    If Modo = 2 Or Modo = 6 Then
        If Not data3.Recordset.EOF Then
            On Error Resume Next
            txtAux(10).Text = DBLet(data3.Recordset!nomalmac, "T")
            txtAux(11).Text = DBLet(data3.Recordset!Ampliaci, "T")
            Err.Clear
        End If
    End If
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
    ArtiHORAS = DevuelveDesdeBD(conAri, "codartichoras", "advparametros", "1", "1")
    PreparaConexionGeslab
    
    'advparametros` add column `codalmac
    NombreTabla = DevuelveDesdeBD(conAri, "codalmac", "advparametros", "1", "1")
    If NombreTabla = "" Then NombreTabla = "1"
    AlmacenLin = Val(NombreTabla)
    NombreTabla = ""
    

    ' ICONITOS DE LA BARRA
    btnPrimero = 23 'Posicion del Boton Primero en la toolbar (+ 3 siguientes)
    With Me.Toolbar1
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 1   'Botón Buscar
        .Buttons(2).Image = 2   'Botón Todos
        
        .Buttons(5).Image = 3   'Insertar Nuevo
        .Buttons(6).Image = 4   'Modificar
        .Buttons(7).Image = 5   'Borrar
        
        .Buttons(9).Image = 10
        .Buttons(10).Image = 23  '
        .Buttons(11).Image = 36  ' actualizar dto/familia
        
        .Buttons(14).Image = 21  'genera
        .Buttons(17).Image = 42  'genera
        
        .Buttons(19).Image = 16  ' Imprimir
        .Buttons(20).Image = 40  ' Imprimir
        .Buttons(21).Image = 15  'Salir
        
        .Buttons(btnPrimero).Image = 6  'Primero
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
        .Buttons(btnPrimero + 3).Image = 9 'Último
    End With
    
    

    mnFiltro1_Click 1
   
    
    
    NombreTabla = PonerTrabajadorConectado(Ordenacion)  'reutilizacion variables
    If NombreTabla = "" Then NombreTabla = "-1"
    If Ordenacion = "" Then Ordenacion = "Trabajador conectado incorrecto"
    Me.mnFiltro1(5).Tag = NombreTabla
    Me.mnFiltro1(5).Caption = "Trab: " & Ordenacion
    
    If NombreTabla = "-1" Then
        mnFiltro1_Click 4 'Cualquier trabajador
    Else
        mnFiltro1_Click 5 'trabajador conectado
    End If
    
    LimpiarCampos   'Limpia los campos TextBox
    
        
  
    '## A mano
    NombreTabla = "advpartes"
    Ordenacion = " ORDER BY numparte"
    FrameMultparte.visible = False
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
       
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = conn
    Data1.RecordSource = "Select * from " & NombreTabla & " WHERE numparte=-1"
    Data1.Refresh
    
    If DatosADevolverBusqueda = "" Then
        PonerModo 0
    Else
        BotonBuscar
    End If
    CargaGrid False, 0 '0:todos

    SSTab1.Tab = 0
End Sub


Private Sub LimpiarCampos()
    limpiar Me   'Metodo general: Limpia los controles TextBox
    'Aqui va el especifico de cada form es
    '### a mano
    Check1(0).Value = 0
    Check1(1).Value = 0
    Check1(2).Value = 0
    txtAux(10).Text = "": txtAux(11).Text = ""
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If Modo >= 5 Then
        cmdCancelar_Click
        
        Cancel = 1
    Else

        CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    
        
        
        
        
        
        Set cCli = Nothing
    End If

End Sub


Private Sub frmAlm_DatoSeleccionado(CadenaSeleccion As String)
    CadenaDevuelta2 = CadenaSeleccion
End Sub

Private Sub frmArt_DatoSeleccionado(CadenaSeleccion As String)
    CadenaDevuelta2 = CadenaSeleccion
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
    CadenaDevuelta2 = CadenaDevuelta
End Sub


Private Sub frmC_Selec(vFecha As Date)
    CadenaDevuelta2 = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmCli_DatoSeleccionado(CadenaSeleccion As String)
    CadenaDevuelta2 = CadenaSeleccion
End Sub

Private Sub frmDptoEnvio_DatoSeleccionado(CadenaSeleccion As String)
    CadenaDevuelta2 = CadenaSeleccion
End Sub

Private Sub frmF_DatoSeleccionado(CadenaSeleccion As String)
     CadenaDevuelta2 = CadenaSeleccion
End Sub

Private Sub frmList_DatoSeleccionado(CadenaSeleccion As String)
'Cuando pasa de Pedido -> Albaran
'Aqui devuelve los valores que se introducen desde el Form de Listado de Pedido
'para generar el Albaran
Dim vSQL As String
Dim A As String
Dim CambiaZona As Boolean
Dim i As Integer
Dim J As Integer

    'Construimos parte de la SQL para insertar en tabla de Albaranes(scaalb)
    FechaAlb = RecuperaValor(CadenaSeleccion, 4)
    vSQL = ""
    vSQL = " '" & Format(FechaAlb, FormatoFecha) & "' as fechaalb, " 'Fecha Albaran
    
    '26/11/2010
    'Si facturamos Si o NO
    'Abril 2012. Martin Quiere la marca SI
    'vSQL = vSQL & CStr(Abs(vParamAplic.MarcarAlbaranFacturar))
    vSQL = vSQL & "1"
    
    vSQL = vSQL & " as factursn, " 'facturar s/n
    vSQL = vSQL & "advpartes.codclien,nomclien,"
    If IsNull(Data1.Recordset!CodDirec) Then
        vSQL = vSQL & "domclien , codpobla, pobclien, proclien,  "
    Else
        Set miRsAux = New ADODB.Recordset
        miRsAux.Open "Select * from sdirec where codclien=" & CStr(Data1.Recordset!codClien) & " AND coddirec = " & CStr(Data1.Recordset!CodDirec), conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If miRsAux.EOF Then
            vSQL = vSQL & "domclien , codpobla, pobclien, proclien,  "
        Else
            If Not IsNull(miRsAux!domdirec) Then vSQL = vSQL & DBSet(miRsAux!domdirec, "T") & " AS "
            vSQL = vSQL & " domclien,"
            
            If Not IsNull(miRsAux!codpobla) Then vSQL = vSQL & DBSet(miRsAux!codpobla, "T") & " AS "
            vSQL = vSQL & " codpobla,"
            
            If Not IsNull(miRsAux!pobdirec) Then vSQL = vSQL & DBSet(miRsAux!pobdirec, "T") & " AS "
            vSQL = vSQL & " pobclien,"
            
            If Not IsNull(miRsAux!prodirec) Then vSQL = vSQL & DBSet(miRsAux!prodirec, "T") & " AS "
            vSQL = vSQL & " proclien,"
            '"domclien , , pobclien, proclien,  "
        
        
        End If
        miRsAux.Close
        Set miRsAux = Nothing
    End If
    vSQL = vSQL & "nifClien,telclie1,"
    
    If IsNull(Data1.Recordset!CodDirec) Then
        vSQL = vSQL & "NULL, NULL"
    Else
        vSQL = vSQL & DBSet(Data1.Recordset!CodDirec, "N") & "," & DBSet(Text2(7).Text, "T")
    End If
    vSQL = vSQL & ", ""Parte: " & Data1.Recordset!numparte & ""","
    vSQL = vSQL & RecuperaValor(CadenaSeleccion, 1) & " as codtraba, " 'Trabajador de Albaran
    'Mayo 2012
    'Sera el trabajador del Parte
    vSQL = vSQL & Me.Text1(8).Text & " as codtrab1, "
    vSQL = vSQL & RecuperaValor(CadenaSeleccion, 2) & " as codtrab2, " 'Material Preparado por
    vSQL = vSQL & "codagent, codforpa, "
    vSQL = vSQL & RecuperaValor(CadenaSeleccion, 3) & " as codenvio, " 'Cod Envio
    
    'Antes
    'vSQL = vSQL & "dtoppago, dtognral, tipofact, observa01, observa02, observa03, observa04, observa05, "
    'vSQL = vSQL & "0,0,tipofact,null,null,null,null,null,"
    'Marzo 2012
    'Observaciones llevaran varias cosas     Abril 2012. Tipofact tb lo pide ne le frmlist
    vSQL = vSQL & "0,0,@tipofact@,"
    If Text1(2).Text = "" Then
        i = 0 'ninguna observacion añadidda
    
    Else
        A = QuitarCaracterEnter(Text1(2).Text)
        
        'Metemos el tratamiento tb
        A = Trim(Text2(3).Text & "( " & Text1(3).Text & ")    " & A)
        If Len(A) <= 80 Then
            'Solo una linea
            vSQL = vSQL & DBSet(A, "T") & ","
            i = 1
        ElseIf Len(A) <= 160 Then
            
            vSQL = vSQL & DBSet(Mid(A, 1, 80), "T") & ","
            vSQL = vSQL & DBSet(Mid(A, 81), "T") & ","
            i = 2
        
        Else
            vSQL = vSQL & DBSet(Mid(A, 1, 80), "T") & ","
            vSQL = vSQL & DBSet(Mid(A, 81, 160), "T") & ","
            vSQL = vSQL & DBSet(Mid(A, 161), "T") & ","
            i = 3
        End If
    End If
    
    
    'He llenado i campos de observacion.
    If Not Me.Data2.Recordset.EOF Then
        J = 5 - i 'cuantas lineas me caben
        If Data2.Recordset.RecordCount > J Then
            'No me cabe una linea por campo
            Data2.Recordset.MoveFirst
            A = ""
            While Not Data2.Recordset.EOF
                A = Trim(A & "     " & Data2.Recordset!codCampo & "- " & Mid(DBLet(Data2.Recordset!nomparti, "T"), 1, 20))
                If Len(A) > 50 Then
                    If i < 5 Then
                        i = i + 1
                        vSQL = vSQL & DBSet(A, "T") & ","
                        A = ""
                    End If
                End If
                Data2.Recordset.MoveNext
            Wend
            Data2.Recordset.MoveFirst
        
            'Los pongo en lineas
        
        
        
        Else
            'Me cabe una linea por campo
            Data2.Recordset.MoveFirst
            A = ""
            While Not Data2.Recordset.EOF
                A = Trim(Data2.Recordset!codCampo & "- " & DBLet(Data2.Recordset!nomparti, "T")) & " -" & DBLet(Data2.Recordset!nomvarie, "T")
                i = i + 1
                vSQL = vSQL & DBSet(A, "T") & ","
                A = ""
                Data2.Recordset.MoveNext
                
            Wend
            Data2.Recordset.MoveFirst
            
            
        End If
    End If
    While i < 5
        vSQL = vSQL & "null,"
        i = i + 1
    Wend
    
    
    'vSQL = vSQL & "numofert, fecofert, "  'Nº Oferta, fecha de la Oferta
    vSQL = vSQL & "null,null,"
    'vSQL = vSQL & Text1(0).Text & " as numpedcl, '" 'Nº Pedido
    vSQL = vSQL & "null,"
    'vSQL = vSQL & Format(Text1(1).Text, FormatoFecha) & "' as fecpedcl, '" 'Fecha Pedido
    'vSQL = vSQL & Format(Text1(2).Text, FormatoFecha) & "' as fecentre, " 'Fecha Prevista Entrega
    vSQL = vSQL & "null,null,"
    vSQL = vSQL & "0 as sementre " 'Semana entrega Pedido
    
    CadenaSQL = vSQL
    
''''''''    'Se almacena aqui si el usuario quiere imprimir el Albaran tras generarlo
    B = CBool(RecuperaValor(CadenaSeleccion, 5))
''''''''    ImprimeEtiq = CBool(RecuperaValor(CadenaSeleccion, 6))
''''''''    ImprimeHojaExp = CBool(RecuperaValor(CadenaSeleccion, 7))
''''''''
''''''''
''''''''    'Solo para la facturacion
''''''''    CtaBancoPropi = RecuperaValor(CadenaSeleccion, 8)
''''''''
''''''''
''''''''    'Enero 2011
''''''''    vSQL = RecuperaValor(CadenaSeleccion, 10)
''''''''    EsAMostrador = vSQL = "1"

End Sub

Private Sub frmT_DatoSeleccionado(CadenaSeleccion As String)
    CadenaDevuelta2 = CadenaSeleccion
End Sub

Private Sub frmTra_DatoSeleccionado(CadenaSeleccion As String)
    CadenaDevuelta2 = CadenaSeleccion
End Sub

Private Sub imgBuscar_Click(Index As Integer)
Dim J As Integer

    If Modo = 2 Or Modo = 5 Or Modo = 0 Then Exit Sub
    
    Select Case Index
        Case 0, 5 'Centros de coste de la conta
            Screen.MousePointer = vbHourglass
            J = 3
            If Index = 5 Then J = 11
            
            Set frmTra = New frmADVTratamientos
            frmTra.DatosADevolverBusqueda = True
            frmTra.Show vbModal
            Set frmTra = Nothing
            If CadenaDevuelta2 <> "" Then
                Text1(J).Text = RecuperaValor(CadenaDevuelta2, 1)
                Text2(J).Text = RecuperaValor(CadenaDevuelta2, 2)
                CadenaDevuelta2 = ""
            End If
            Screen.MousePointer = vbDefault
            If Index = 0 Then PonerFoco Text1(8)
            
        Case 1
            Screen.MousePointer = vbHourglass
            Me.imgBuscar(0).Tag = Index
            CadenaDevuelta2 = ""
            Set frmCli = New frmBasico2
            AyudaClientes frmCli, Text1(6)
            Set frmCli = Nothing
            
            
            
            If CadenaDevuelta2 <> "" Then
                    D = RecuperaValor(CadenaDevuelta2, 1)
                    Set cCli = New CCliente
                    If cCli.LeerDatos(D) Then
                        If Not cCli.ClienteBloqueado(2, False) Then
                            Text2(6).Text = cCli.Nombre
                            If cCli.Observaciones <> "" Then MsgBox cCli.Observaciones, vbInformation
                        Else
                            CadenaDevuelta2 = ""
                        End If
                    
                    Else
                        'NO exiswte el cliente
                        CadenaDevuelta2 = ""
                    End If
                    Set cCli = Nothing
                    If CadenaDevuelta2 <> "" Then
            
                        
                        If Val(D) <> Val(Text1(6).Text) Then
                            'Cambia cliente
                            Text1(7).Text = ""
                            Text2(7).Text = ""
                        End If
                        Text1(6).Text = D
                        Text2(6).Text = RecuperaValor(CadenaDevuelta2, 2)

                    End If
                    CadenaDevuelta2 = ""
                    Set cCli = Nothing
            End If
            
            
            Screen.MousePointer = vbDefault
            PonerFoco Text1(5)
            
        Case 2
                If Text1(6).Text = "" Then
                    MsgBox "Indique el cliente", vbExclamation
                    PonerFoco Text1(6)
                    Exit Sub
                End If
                CadenaDevuelta2 = ""
                Set frmDptoEnvio = New frmFacCliEnvDpto
                frmDptoEnvio.DireccionesEnvio = False
'                If Text1(indice).Text <> "" Then
'                    frmDptoEnvio.VerDatoDpto = CInt(Text1(indice).Text)
'                Else
                    frmDptoEnvio.VerDatoDpto = -1
                'End If
                frmDptoEnvio.codClien = CLng(Text1(6).Text)
                frmDptoEnvio.NomClien = Text2(6).Text
                frmDptoEnvio.Show vbModal
                Set frmDptoEnvio = Nothing
        
                If CadenaDevuelta2 <> "" Then
                    Text1(7).Text = RecuperaValor(CadenaDevuelta2, 1)
                    Text2(7).Text = RecuperaValor(CadenaDevuelta2, 2)
                
                    PonerFoco Text1(5)
                End If
                
        Case 3, 6
            J = 8
            If Index = 6 Then J = 12
            CadenaDevuelta2 = ""
'            Set frmT = New frmAdmTrabajadores
'            frmT.DatosADevolverBusqueda = "0|1|"
'            frmT.Show vbModal
            Set frmT = New frmBasico2
            AyudaTrabajadores frmT, Text1(J)
            Set frmT = Nothing


            If CadenaDevuelta2 <> "" Then
                Text1(J).Text = RecuperaValor(CadenaDevuelta2, 1)
                Text2(J).Text = RecuperaValor(CadenaDevuelta2, 2)
                If Index = 3 Then PonerFoco Text1(4)
            End If
            
        Case 4
            CadenaDevuelta2 = ""
            Set frmF = New frmFlotas
            frmF.DatosADevolverBusqueda = "0|1|"
            frmF.Show vbModal
            Set frmF = Nothing
            If CadenaDevuelta2 <> "" Then
                Text1(9).Text = RecuperaValor(CadenaDevuelta2, 1)
                Text2(9).Text = RecuperaValor(CadenaDevuelta2, 2)
            
                PonerFoco Text1(9)
            End If
            
            
        Case 7
            CadenaDesdeOtroForm = Text1(14).Text
            frmFacClienteObser.Modificar = True
            frmFacClienteObser.Text1 = CadenaDesdeOtroForm
            frmFacClienteObser.Show vbModal
            'Llevara DOS VALORES.
            'Si modifica y el texto
            If CadenaDesdeOtroForm <> "" Then
                If RecuperaValor(CadenaDesdeOtroForm, 1) = "1" Then Text1(14).Text = Mid(CadenaDesdeOtroForm, 3)
            End If
            CadenaDesdeOtroForm = ""
    End Select
End Sub




Private Sub imgCampos_Click(Index As Integer)
    If Index = 0 Then
        CadenaDesdeOtroForm = ""
        frmADVvarios.Opcion = 1   'AGRUPADOS
        frmADVvarios.vCampos = -1
        frmADVvarios.Show vbModal
        If CadenaDesdeOtroForm <> "" Then
            If Mid(CadenaDesdeOtroForm, 1, 1) <> "@" Then CadenaDesdeOtroForm = "@" & CadenaDesdeOtroForm
            Screen.MousePointer = vbHourglass
            MultiInsercionCampos True
            CargaListMulti
            If CadenaSQL <> "" Then
                If MsgBox("Clientes bloqueados" & vbCrLf & "¿Eliminarlos de la seleccion?", vbQuestion + vbYesNo) = vbYes Then
                    CadenaSQL = Mid(CadenaSQL, 2)
                    CadenaSQL = "DELETE FROM tmpnlotes where codusu =" & vUsu.Codigo & " AND (codprove,numalbar) in (" & CadenaSQL & ")"
                    conn.Execute CadenaSQL
                    CargaListMulti
                End If
            End If
            Screen.MousePointer = vbDefault
        End If
    ElseIf Index = 1 Then
        'Borrar
        CadenaSQL = ""
        FechaAlb = ""
        For NumRegElim = 1 To lwC.ListItems.Count
            If lwC.ListItems(NumRegElim).Checked Then
                CadenaSQL = CadenaSQL & ", (" & lwC.ListItems(NumRegElim).SubItems(5) & "," & lwC.ListItems(NumRegElim).SubItems(6) & ")"
                FechaAlb = FechaAlb & "X"
            End If
        Next NumRegElim
        If CadenaSQL <> "" Then
            If MsgBox("Eliminar AGRUPACION seleccionada(" & Len(FechaAlb) & ") ?", vbQuestion + vbYesNoCancel) = vbYes Then
                Screen.MousePointer = vbHourglass
                CadenaSQL = Mid(CadenaSQL, 2)
                CadenaSQL = "DELETE FROM tmpnlotes where codusu =" & vUsu.Codigo & " AND (codprove,numalbar) in (" & CadenaSQL & ")"
                conn.Execute CadenaSQL
                Espera 0.2
                CargaListMulti
                Screen.MousePointer = vbDefault
            End If
        Else
            If lwC.ListItems.Count > 0 Then MsgBox "Seleccione algun dato", vbExclamation
        End If
        
        
    Else
        FechaAlb = ""
        For NumRegElim = 1 To lwC.ListItems.Count
            If lwC.ListItems(NumRegElim).Checked Then
                FechaAlb = FechaAlb & "X"
            End If
        Next NumRegElim
        If FechaAlb <> "" Then
            CadenaSQL = vbCrLf & vbCrLf & "** Si vuelve a añadir por numero de campo, puede volver a aparecer ** "
            CadenaSQL = "Eliminar campos seleccionados(" & Len(FechaAlb) & ") ?" & CadenaSQL
            If MsgBox(CadenaSQL, vbQuestion + vbYesNoCancel) = vbYes Then
            
                For NumRegElim = lwC.ListItems.Count To 1 Step -1
                    If lwC.ListItems(NumRegElim).Checked Then
                        If NumRegElim <> lwC.ListItems.Count Then
                            'SI no es el ultimo
                            If Trim(lwC.ListItems(NumRegElim + 1).SubItems(1)) = """" Then
                                lwC.ListItems(NumRegElim + 1).SubItems(1) = lwC.ListItems(NumRegElim).SubItems(1)
                                lwC.ListItems(NumRegElim + 1).Text = lwC.ListItems(NumRegElim).Text
                            End If
                        End If
                        lwC.ListItems.Remove NumRegElim
                    End If
                Next NumRegElim
            
            End If
        End If
    End If
    CadenaSQL = ""
    FechaAlb = ""
End Sub

Private Sub CargaListMulti()
Dim Cad As String
Dim cCli As Integer
Dim Bloq As Boolean
Dim IT As ListItem

    lwC.ListItems.Clear
    CadenaEnlazeAriagro Cad, True
    Cad = Cad & " WHERE (nrocampo,rcampos.codvarie) IN "   'Ponia codcampo  y no estaba codvarie Julio 19
    Cad = Cad & "(select codprove,numalbar from tmpnlotes where codusu =" & vUsu.Codigo & ")"
    Cad = Cad & " ORDER BY nrocampo,codclien,codcampo"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    cCli = -1
    CadenaSQL = ""
    While Not miRsAux.EOF
        Set IT = Me.lwC.ListItems.Add()
        
        'QUE NO SEA NULL
        If IsNull(miRsAux!codClien) Then
            CadenaSQL = ""
            For cCli = 0 To miRsAux.Fields.Count - 1
                CadenaSQL = CadenaSQL & miRsAux.Fields(cCli).Name & ":   " & DBLet(miRsAux.Fields(cCli), "T") & vbCrLf
            Next
            CadenaSQL = "ERROR grave: " & vbCrLf & vbCrLf & CadenaSQL
            
            MsgBox CadenaSQL, vbCritical
            CadenaSQL = ""
            cmdMultiParte(0).Enabled = False
            miRsAux.Close
            Exit Sub
        End If
        If miRsAux!codClien <> cCli Then
            cCli = miRsAux!codClien
            IT.Text = Format(cCli, "0000")
            IT.SubItems(1) = DBLet(miRsAux!NomClien, "T")
            Bloq = False
            If IT.SubItems(1) = "" Or miRsAux!codsitua > 1 Then Bloq = True
            IT.Tag = miRsAux!codClien
        Else
            IT.Text = " "
            IT.SubItems(1) = "     "" "
            If Not Bloq Then
                IT.Tag = miRsAux!codClien
            Else
                IT.Tag = -1
            End If
        End If
        
        IT.SubItems(2) = Format(miRsAux!codCampo, "0000")
        IT.SubItems(3) = Format(miRsAux!nomvarie, "0000")
        IT.SubItems(5) = Format(miRsAux!nrocampo, "0000")
        IT.SubItems(6) = Format(miRsAux!codvarie, "0000")
        IT.Checked = True
        If Bloq Then
            CadenaSQL = CadenaSQL & ", (" & miRsAux!nrocampo & "," & miRsAux!codvarie & ")"  'codcampo
            IT.ForeColor = vbRed
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    cmdMultiParte(0).Enabled = True
End Sub



Private Sub imgCheck_Click(Index As Integer)
    For NumRegElim = 1 To Me.lwC.ListItems.Count
        lwC.ListItems(NumRegElim).Checked = Index = 1
    Next NumRegElim
End Sub

Private Sub imgFec_Click(Index As Integer)
    
    If Modo = 0 Or Modo = 2 Or Modo >= 5 Then Exit Sub
    
    
    Set frmC = New frmCal
    frmC.Fecha = Now
    If Index = 0 Then
        If Me.Text1(1).Text <> "" Then frmC.Fecha = CDate(Text1(1).Text)
    Else
        If Me.Text1(10).Text <> "" Then frmC.Fecha = CDate(Text1(10).Text)
    End If
    frmC.Show vbModal
    If CadenaDevuelta2 <> "" Then Text1(IIf(Index = 0, 1, 10)).Text = CadenaDevuelta2
    
    
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    If Modo >= 5 Then 'Eliminar lineas de trabajadores
         BotonEliminarLinea
    Else   'Eliminar Trabajador
         BotonEliminar
    End If
End Sub

Private Sub mnFiltro1_Click(Index As Integer)
    'mayo 2012.
    '2 Filtros. Abierto/cerrado y Trabajador conectado
    If Index < 3 Then
        Me.mnFiltro1(0).Checked = False
        Me.mnFiltro1(1).Checked = False
        Me.mnFiltro1(2).Checked = False
        
    Else
        Me.mnFiltro1(4).Checked = False
        Me.mnFiltro1(5).Checked = False
            
    End If
    Me.mnFiltro1(Index).Checked = True
    
    
    
    
    
    
    
    
End Sub

Private Sub mnLineas1_Click(Index As Integer)
    '9, 10, 11
    BotonMtoLineas 9 + Index 'index=0,1,2
End Sub

Private Sub mnModificar_Click()
    'If BLOQUEADesdeFormulario(Me) Then BotonModificar
    If Modo >= 5 Then 'Modificar lineas
         BotonModificarLinea
    Else   'Modificar
    
         If Modo <> 2 Then Exit Sub
         If Data1.Recordset.EOF Then Exit Sub
         If Val(Data1.Recordset!cerrado) = 1 Then
            MsgBox "Parte cerrado", vbExclamation
            Exit Sub
        End If
    
         If BLOQUEADesdeFormulario(Me) Then BotonModificar
    End If
    
End Sub

Private Sub mnNuevo_Click()
    If Modo >= 5 Then 'Añadir lineas
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
    ConseguirFoco Text1(Index), Modo
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Not (Index = 0 And KeyCode = 38) Then KEYdown KeyCode
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
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
Dim devuelve As String
Dim cadMen As String
Dim Sql As String
Dim NRegs As Long

    
    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
        
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
          
    'Si queremos hacer algo ..
    Select Case Index
        Case 1, 10 'Fecha albaran
            If Text1(Index).Text <> "" Then PonerFormatoFecha Text1(Index)
            

                
            
        Case 3, 11 ' Tratamiento
            If Modo = 1 Then Exit Sub
            Text2(Index).Text = ""
            If Text1(Index).Text <> "" Then
                Text1(Index).Text = UCase(Text1(Index).Text)
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "advtrata", "nomtrata", "codtrata", "Tratamiento", "T")
                If Text2(Index).Text = "" Then
                   Text1(Index).Text = ""
                  
                    PonerFoco Text1(Index)
                
                End If
            End If
            
         Case 4, 5, 13 ' litros
            If Text1(Index).Text <> "" Then
                If Not IsNumeric(Text1(Index)) Then
                    MsgBox "Debe ser numero", vbExclamation
                    Text1(Index).Text = ""
                    PonerFoco Text1(Index)
                End If
            End If
         
         Case 6
            If Modo = 1 Then Exit Sub
            Text2(Index).Text = ""
            If Text1(Index).Text <> "" Then
                If Not IsNumeric(Text1(Index).Text) Then
                    MsgBox "Campo numerico", vbExclamation
                Else
                    Set cCli = New CCliente
                    If cCli.LeerDatos(Text1(6).Text) Then
                        If Not cCli.ClienteBloqueado(2, False) Then
                            Text2(Index).Text = cCli.Nombre
                            If cCli.Observaciones <> "" And Modo = 3 Then MsgBox cCli.Observaciones, vbInformation
                        End If
                    End If
                    Set cCli = Nothing
                End If
                
                If Text2(Index).Text = "" Then
                   Text1(Index).Text = ""
                   PonerFoco Text1(Index)
                
                Else
                    'cliente correcto
                    
                        If Text1(7).Text <> "" Then
                           Sql = "codclien = " & Text1(6).Text & " AND coddirec "
                           Text2(7).Text = DevuelveDesdeBD(conAri, "nomdirec", "sdirec", Sql, Text1(7).Text)
                           If Text2(7).Text = "" Then Text1(7).Text = ""
                        End If
                    
                End If
            End If
         
         
         Case 7 '
         
            If Modo = 1 Then Exit Sub
            
            Sql = ""
            devuelve = ""
            NRegs = 7
            If Text1(Index).Text <> "" Then
                If Not IsNumeric(Text1(Index).Text) Then
                    devuelve = "Debe ser numerico"
                Else
                    If Text1(6).Text = "" Then
                        devuelve = "Indique el cliente"
                        NRegs = 6
                    Else
                        Sql = "codclien = " & Text1(6).Text & " AND coddirec "
                        Sql = DevuelveDesdeBD(conAri, "nomdirec", "sdirec", Sql, Text1(Index).Text)
                        If Sql = "" Then devuelve = "No existe el departamento"
                    End If
                End If
            End If
            Text2(7).Text = Sql
            If Sql = "" Then
                If devuelve <> "" Then
                    MsgBox devuelve, vbExclamation
                    If NRegs = 6 Then Text1(Index).Text = ""
                    PonerFoco Text1(NRegs)
                End If
            End If
            
        Case 8, 12
            '
            Sql = ""
            If PonerFormatoEntero(Me.Text1(Index)) Then
                Sql = DevuelveDesdeBD(conAri, "nomtraba", "straba", "codtraba", Text1(Index).Text)
                If Sql = "" Then
                    MsgBox "Trabajador NO encontrado" & Text1(Index), vbExclamation
                    Text1(Index).Text = ""
                End If
                
            End If
            Text2(Index).Text = Sql
        Case 2 'observaciones
            PonerFocoBtn cmdAceptar
            
            
        Case 9 ' FLOTAS
            If Modo = 1 Then Exit Sub
            Text2(Index).Text = ""
            If Text1(Index).Text <> "" Then
                Text1(Index).Text = UCase(Text1(Index).Text)
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "sflotas", "nomflota", "Codflota", "Flotas-Vehiculos", "T")
                If Text2(Index).Text = "" Then
                   Text1(Index).Text = ""
                  
                    PonerFoco Text1(Index)
                
                End If
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
            CadenaConsulta2 = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
            PonerCadenaBusqueda
        End If
    End If
End Sub


Private Sub MandaBusquedaPrevia(cadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim Cad As String
            
        'Busqueda de una Família de Artículo
        Cad = ParaGrid(Text1(0), 15, "Código")
        Cad = Cad & ParaGrid(Text1(1), 15, "Fecha")
        Cad = Cad & "Nombre|sclien|nomclien|T||55·"
        Cad = Cad & "Cerrado|advpartes|if(cerrado=1,""*"","""")|T||10·"
        
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = Cad
        frmB.vTabla = "advpartes,sclien"
        If cadB <> "" Then cadB = " AND " & cadB
        cadB = "advpartes.codclien=sclien.codclien" & cadB
        frmB.vSQL = cadB
        CadenaDevuelta2 = ""
        '###A mano
        frmB.vDevuelve = "0|1|"
        frmB.vTitulo = "Partes trabajo"
        frmB.vselElem = 1
        frmB.vConexionGrid = conAri
        frmB.vCargaFrame = False
        '#
        
        frmB.Show vbModal
        Set frmB = Nothing
        'Si ha puesto valores y tenemos que es formulario de busqueda entonces
        'tendremos que cerrar el form lanzando el evento
        If CadenaDevuelta2 <> "" Then
            CadenaConsulta2 = "select * from " & NombreTabla & " WHERE numparte= " & RecuperaValor(CadenaDevuelta2, 1) & " " & Ordenacion
            PonerCadenaBusqueda
        End If
End Sub


Private Sub PonerCadenaBusqueda()
    Screen.MousePointer = vbHourglass
    On Error GoTo EEPonerBusq

    Data1.RecordSource = CadenaConsulta2
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
    
    Modo = 4
    Text1_LostFocus 3
    Text2(7).Text = ""
    Text1_LostFocus 6
    Text1_LostFocus 8
    Text1_LostFocus 9
    Modo = 2


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

 
    'Actualiza Iconos Insertar,Modificar,Eliminar
    ActualizarToolbarGnral Me.Toolbar1, Modo, Kmodo, 5
    
    Modo = Kmodo
    PonerIndicador lblIndicador, Modo
    If Modo = 5 Then lblIndicador.Caption = "Lineas dto"
    
    BuscaChekc = ""
  
    B = Modo < 5
    If Not B Then ModificaLineas = 0
    'DataGrid2.enabled
    
    '-------------------------------------------------
    'Modo 2. Hay datos y estamos visualizandolos
    B = (Kmodo = 2)
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.visible = B
    Else
        'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
        cmdRegresar.visible = Modo = 5 And ModificaLineas = 0
    End If
    
    'Visualizar flechas de desplazamiento en la toolbar si modo=2
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero, B, NumReg
    
    'Poner Boton de Cabecera o Aceptar/Cancelar
    PonerBotonCabecera B Or (Modo = 0)
        
    'Bloquear Registros si modo distinto de Insertar o Modificar
    BloquearText1 Me, Modo
    
    BloquearTxt Text1(4), Modo <> 1
    
    BloquearChecks Me, Modo
        
    Me.Check1(0).Enabled = Modo = 1 Or Modo = 3 Or Modo = 4
    Me.Check1(1).Enabled = Modo = 1
    B = False
    If Me.Check1(1).Value = 0 Then B = Modo = 1 Or Modo = 3 Or Modo = 4
    Me.Check1(2).Enabled = B
        
        
    chkVistaPrevia.Enabled = (Modo <= 2)
    
    'Poner el tamaño de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos
    
    PonerModoOpcionesMenu 'Activar opciones de menu según MODO
    PonerOpcionesMenu   'Activar opciones de menu según NIVEL
                        'de permisos del usuario
End Sub


Private Sub PonerModoOpcionesMenu()
Dim B As Boolean
On Error Resume Next

    B = Modo < 3
    If Not B Then
        If Modo >= 5 And ModificaLineas = 0 Then B = True
    End If
    
    'Añadir
    Toolbar1.Buttons(5).Enabled = B
    Me.mnNuevo.Enabled = B
    
    B = Modo = 2
    Toolbar1.Buttons(9).Enabled = B
    Toolbar1.Buttons(10).Enabled = B
    Toolbar1.Buttons(11).Enabled = B
    Toolbar1.Buttons(14).Enabled = B Or Modo = 0 'mutiparte
    Toolbar1.Buttons(18).Enabled = B  'facturar
    If Not B Then
        If Modo >= 5 And ModificaLineas = 0 Then B = True
    End If
    
    'Modificar
    Toolbar1.Buttons(6).Enabled = B
    mnModificar.Enabled = B
    'eliminar
    Toolbar1.Buttons(7).Enabled = B
    mnEliminar.Enabled = B
    
    
    
    
    
    Toolbar1.Buttons(9).Enabled = Modo = 2
    Me.mnLineas.Enabled = Modo = 2
    
    
    
     '---------------------------------
    B = (Modo >= 3)
    'Buscar
    Toolbar1.Buttons(1).Enabled = Not B
    Me.mnBuscar.Enabled = Not B
    'VerTodos
    Toolbar1.Buttons(2).Enabled = Not B
    Me.mnVerTodos.Enabled = Not B
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
        Text1(0).Text = SugerirCodigoSiguienteStr("advpartes", "numparte")
        If ExisteCP(Text1(0)) Then B = False
    End If
    
    
    'No puede poner coddirec y no estar nomdirec
    If Text1(7).Text <> "" Xor Text2(7).Text <> "" Then
        MsgBox "Error en el departamento", vbExclamation
        PonerFoco Text1(7)
        B = False
    End If
    
    'No puede poner codtraba y no estar nomtraba
    If Text1(8).Text <> "" Xor Text2(8).Text <> "" Then
        MsgBox "Error en el trabajador", vbExclamation
        PonerFoco Text1(8)
        B = False
    End If
    'No puede poner FLTOA
    If Text1(9).Text <> "" Xor Text2(9).Text <> "" Then
        MsgBox "Error en el vehiculo", vbExclamation
        PonerFoco Text1(9)
        B = False
    End If
    
    
    
    'Diciembre 2014
    'partesADV EsExterno
    If Modo = 4 And B Then
        If Val(Data1.Recordset!EsExterno) <> Val(Me.Check1(2).Value) Then
            'Ha cambiado el valor
            B = ComprobarParteExternoInterno(Check1(2).Value = 0)
        End If
    End If
    DatosOk = B
End Function

Private Function ComprobarParteExternoInterno(Interno As Boolean) As Boolean
    ComprobarParteExternoInterno = True
    '0 = cualquiera  1 Internos   2 Externos'
    ' el codartic...
    'Si es interno, no puede haber un codartic con partesADV a 2
    'si es externo  " ""  a 1
    CodZona = "2"
    FechaAlb = "externos"
    If Not Interno Then
        CodZona = "1"
        FechaAlb = "internos"
    End If
    D = " codartic in (select codartic from advparteslineas where numparte=" & Text1(0).Text & ")"
    D = D & " AND partesADV "
    D = DevuelveDesdeBD(conAri, "count(*)", "sartic", D, CStr(CodZona))
    If Val(D) > 0 Then
        D = " hay " & D & " articulo(s) con la marca de 'Solo " & FechaAlb & "'"
        MsgBox D, vbExclamation
        ComprobarParteExternoInterno = False
    End If
    
    FechaAlb = ""
    CodZona = 0
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
                
        Case 9, 10, 11
                BotonMtoLineas Button.Index
            
        Case 14
            POnerMultiParte True
        Case 17
                pasarAAlbaranes
                
        Case 19
            BotonImprimirNuevo2 1
        Case 20 'Imprimir listado
            If MsgBox("Imprimir todos los registros seleccionados?", vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
            
            BotonImprimirNuevo2 0
            
       
        Case 21: mnSalir_Click
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
        ElseIf Modo >= 5 Then
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

    Cad = "(numparte=" & Text1(0).Text & ")"
    If SituarData(Data1, Cad, Indicador) Then
        CargaGrid True, 0
        PonerModo 2

        lblIndicador.Caption = Indicador
    Else
         CargaGrid False, 0
        PonerModo 0
    End If
End Sub


Private Sub PosicionarData3()
    On Error GoTo EPosicionarData2
    
    If Modo = 5 Then
        'NO HAY
    
    ElseIf Modo = 6 Then
        data3.Recordset.Find "numlinea = " & NumRegElim
    
        If data3.Recordset.EOF Then data3.Recordset.MoveFirst
    Else
         data4.Recordset.Find "numlinea = " & NumRegElim
    
        If data4.Recordset.EOF Then data4.Recordset.MoveFirst
    End If
    NumRegElim = 0
    
    Exit Sub
EPosicionarData2:
    MuestraError Err.Number
End Sub

'Private Sub BotonImprimir()
'    '
'    If Modo <> 2 Then Exit Sub
'    If Data1.Recordset.EOF Then Exit Sub
'
'
'
'    'Meto los campos en ADV
'    '-------------------------------
'    D = "DELETE FROM tmpsliped where codusu = " & vUsu.Codigo
'    conn.Execute D
'
'    D = "insert into `tmpsliped` (`codusu`,`numpedcl`,`numlinea`,`codalmac`,"
'    D = D & "`nomartic`,`ampliaci`,`numbultos`,`cantpedprov`,`fecpedprov`,"
'    D = D & "`stockalm`,`stocktot`,`referart`,`codclien`,`codzona`,`importel`,`cantidad`,`codartic`) values "
'    ''1','0','0','0','','',NULL,'0.00','0','0.00','0','','0','0','',NULL,NULL)"
'    FechaAlb = ""
'    If Not Data2.Recordset.EOF Then
'        If Data2.Recordset.RecordCount > 0 Then
'            Data2.Recordset.MoveFirst
'            NumRegElim = 0
'            FechaAlb = ""
'            While Not Data2.Recordset.EOF
'
'
'                'OCTUBRE 2014
'                ' CAMBIA sitaucion por naturane(MARTIN)
'                'BuscaChekc = DBLet(Data2.Recordset!codsitua)
'                'If BuscaChekc <> "" Then BuscaChekc = DevuelveDesdeBD(conAri, "nomsitua", vParamAplic.Ariagro & ".rsituacioncampo", "codsitua", BuscaChekc)
'                BuscaChekc = "  "
'                If Val(DBLet(Data2.Recordset!esnaturane, "N")) = 1 Then BuscaChekc = "Naturane"
'
'
'                NumRegElim = NumRegElim + 1
'                ''1','0','0','0','','',NULL,'0.00','0','0.00','0','','0','0','',NULL,NULL
'                FechaAlb = FechaAlb & ", (" & vUsu.Codigo & "," & Data2.Recordset!codCampo & "," & NumRegElim & ",0,"
'                FechaAlb = FechaAlb & DBSet(Data2.Recordset!nomparti, "T")
'                FechaAlb = FechaAlb & "," & DBSet(Data2.Recordset!nomvarie, "T")
'                FechaAlb = FechaAlb & ",'0','0','','0','0'," & DBSet(BuscaChekc, "T") & ",NULL,NULL,"
'
'                'Superficie cooperativa  supsigpa  supcoope   8,4
'                'Poligono parcela recinto
'                'poligono,parcela,recintos
'                CadenaConsulta2 = "concat(supcoope,""|"",supsigpa,""|"",poligono,""#"",parcela,""#"",recintos,""#|"")"
'                BuscaChekc = DevuelveDesdeBD(conAri, CadenaConsulta2, vParamAplic.Ariagro & ".rcampos", "codcampo", Data2.Recordset!codCampo)
'                CadenaConsulta2 = RecuperaValor(BuscaChekc, 3)
'                CadenaConsulta2 = Replace(CadenaConsulta2, "#", "|")
'
'
'                FechaAlb = FechaAlb & DBSet(TransformaPuntosComas(RecuperaValor(BuscaChekc, 1)), "N", "N") & "," & DBSet(TransformaPuntosComas(RecuperaValor(BuscaChekc, 2)), "N", "N") & ",'"
'                'Poligono parcela recinto
'                BuscaChekc = RecuperaValor(CadenaConsulta2, 1)
'                If BuscaChekc = "" Then BuscaChekc = "0"
'                BuscaChekc = Format(Val(BuscaChekc), "000")
'                FechaAlb = FechaAlb & BuscaChekc & "-"
'                BuscaChekc = RecuperaValor(CadenaConsulta2, 2)
'                If BuscaChekc = "" Then BuscaChekc = "0"
'                BuscaChekc = Format(Val(BuscaChekc), "000000")
'                FechaAlb = FechaAlb & BuscaChekc & "-"
'                BuscaChekc = RecuperaValor(CadenaConsulta2, 3)
'                If BuscaChekc = "" Then BuscaChekc = "0"
'                BuscaChekc = Format(Val(BuscaChekc), "000")
'                FechaAlb = FechaAlb & BuscaChekc
'
'
'
'                FechaAlb = FechaAlb & "')"
'
'                Data2.Recordset.MoveNext
'            Wend
'            Data2.Recordset.MoveFirst
'            FechaAlb = Mid(FechaAlb, 2) 'qito la coma
'            D = D & FechaAlb
'            ejecutar D, False
'        End If
'    End If
'    BuscaChekc = ""
'    CadenaConsulta2 = ""
'    'ALZIRA MARZO 2012
'    'tENEMOS QUE METER TB LOS DATOS DEL CAMPO PARA EL INFORME
'    'Para ello, ya que no tenemos una tabla , y no la voy a crear, inserto en tmpslipreu con
'    'tmpslipreu`   `codusu`,`numofert` ,`numlinea`,`codartic`,`nomartic`,`ampliaci`
'    '                        codcampo    codclien   nomclien
'    D = "DELETE FROM tmpslipreu where codusu = " & vUsu.Codigo
'    conn.Execute D
'
'    D = "insert into tmpslipreu(`codusu`,`numofert` ,`codartic`,`nomartic`) values "
'
'    FechaAlb = ""
'    If Not Data2.Recordset.EOF Then
'        If Data2.Recordset.RecordCount > 0 Then
'            Data2.Recordset.MoveFirst
'            NumRegElim = 0
'            FechaAlb = ""
'            While Not Data2.Recordset.EOF
'                NumRegElim = NumRegElim + 1
'                'tmpslipreu`   `codusu`,`numofert` ,`numlinea`,`codartic`,`nomartic`,`ampliaci`
'                FechaAlb = FechaAlb & ", (" & vUsu.Codigo & "," & DBLet(Data2.Recordset!codCampo, "N") & "," & DBSet(CStr(Data2.Recordset!codsocio), "T")
'                FechaAlb = FechaAlb & "," & DBSet(Data2.Recordset!nomsocio, "T") & ")"
'
'                Data2.Recordset.MoveNext
'            Wend
'            Data2.Recordset.MoveFirst
'            FechaAlb = Mid(FechaAlb, 2) 'qito la coma
'            D = D & FechaAlb
'            ejecutar D, False
'        End If
'    End If
'
'
'
'
'
'
'
'    'Trabajadores
'    D = "delete from tmpcommandest where codusu = " & vUsu.Codigo
'    conn.Execute D
'
'    If Not data4.Recordset.EOF Then
'        If data4.Recordset.RecordCount > 0 Then
'            data4.Recordset.MoveFirst
'            NumRegElim = 0
'            FechaAlb = ""
'            While Not data4.Recordset.EOF
'                NumRegElim = NumRegElim + 1
'                'usu,'0','0','','',NULL,'0.00','0000-00-00','0',NULL,NULL,NULL
'                'codusu`,`codclien`,`codfamia`,`nomclien`,`nomfamia`,`cantidad`,`importel`,`fechaalb`,`codprove`,`nomprove`,`codartic`,`nomartic
'                FechaAlb = FechaAlb & ", (" & vUsu.Codigo & "," & data4.Recordset!CodTraba & ",0,"
'                FechaAlb = FechaAlb & DBSet(data4.Recordset!NomTraba, "T") & ",''"
'                FechaAlb = FechaAlb & "," & DBSet(data4.Recordset!Horas, "N") & ",0,'0000-00-00','0',NULL,NULL,NULL)"
'
'                data4.Recordset.MoveNext
'            Wend
'            data4.Recordset.MoveFirst
'            FechaAlb = Mid(FechaAlb, 2) 'qito la coma
'            D = " insert into `tmpcommandest` (`codusu`,`codclien`,`codfamia`,`nomclien`,`nomfamia`,`cantidad`,`importel`,`fechaalb`,`codprove`,`nomprove`,`codartic`,`nomartic`) values "
'            D = D & FechaAlb
'            ejecutar D, False
'        End If
'    End If
'
'
'
'
'
'
'    'nomrpt
'    FechaAlb = DevuelveDesdeBD(conAri, "documrpt", "scryst", "codcryst", "55")
'    'sql
'    CadenaSQL = "{advpartes.numparte}=" & Data1.Recordset!numparte
'    'parametros
'    D = "|pEmpresa=""" & vParam.NombreEmpresa & """|vCodUsu=" & vUsu.Codigo & "|"
'    NumRegElim = 2
'
'
'
'
'
'
'    With frmImprimir
'
'        .FormulaSeleccion = CadenaSQL
'
'        'parametros
'
'        .OtrosParametros = D
'        .NumeroParametros = NumRegElim
'        .NumeroCopias = 2
'        .SoloImprimir = False
'        .EnvioEMail = False
'        .Opcion = 3002
'        .Titulo = "Partes ADV"
'        .NombreRPT = FechaAlb
'        .ConSubInforme = True
'        .Show vbModal
'    End With
'    CadenaSQL = ""
'    FechaAlb = ""
'
'
'
'End Sub



'Opcion
' 0 la busqueda
' 1 solo actual
' 2 viene de generacion masiva
Private Sub BotonImprimirNuevo2(QueOpcion As Byte)
    '
Dim C2 As String
Dim RsPpal As ADODB.Recordset
Dim N As Long
Dim cadenapartes As String
Dim CamposIm As String


    'Para la impresion desde boton , el modo debe ser 2 y no debe ser EOF
    If QueOpcion <> 2 Then
        If Modo <> 2 Then Exit Sub
        If Data1.Recordset.EOF Then Exit Sub
    End If
    
    'Meto los campos en ADV
    '-------------------------------
    If QueOpcion = 1 Then
    
        D = " WHERE numparte =" & Data1.Recordset!numparte
    Else
        If QueOpcion = 0 Then
            D = Data1.RecordSource
            NumRegElim = InStr(1, D, " WHERE ")
            If NumRegElim = 0 Then Exit Sub
            
            D = Mid(D, NumRegElim)
    
        Else
            'Desde generacion masiva de partes
    
            D = " WHERE numparte IN (select codclien FROM tmpcrmclien WHERE codusu=" & vUsu.Codigo & ")"
    
    
        End If
    End If
    
    Set RsPpal = New ADODB.Recordset
    RsPpal.Open "Select numparte FROM advpartes " & D, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    
    
    D = "DELETE FROM tmpsliped where codusu = " & vUsu.Codigo
    conn.Execute D
    D = "DELETE FROM tmpslipreu where codusu = " & vUsu.Codigo
    conn.Execute D
    
    D = "DELETE FROM tmpcommandest where codusu = " & vUsu.Codigo
    conn.Execute D
    
    cadenapartes = ""
    CamposIm = ""
    While Not RsPpal.EOF
    
        cadenapartes = cadenapartes & ", " & RsPpal!numparte
        
        D = "insert into `tmpsliped` (`codusu`,`numpedcl`,`numlinea`,`codalmac`,"
        D = D & "`nomartic`,`ampliaci`,`codclien`,`cantpedprov`,`fecpedprov`,"
        D = D & "`stockalm`,`stocktot`,`referart`,`numbultos`,`codzona`,`importel`,`cantidad`,`codartic`) values "
        ''1','0','0','0','','',NULL,'0.00','0','0.00','0','','0','0','',NULL,NULL)"
        FechaAlb = ""
        
        C2 = Data2.RecordSource
        N = InStr(1, C2, " WHERE ")
        C2 = Mid(C2, 1, N) & " WHERE codcampo IN (select codcampo from advpartes_campos where numparte=" & RsPpal!numparte & ")"
        Set miRsAux = New ADODB.Recordset
        miRsAux.Open C2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
            
        NumRegElim = 0
        FechaAlb = ""
        While Not miRsAux.EOF
                
                
                    'OCTUBRE 2014
                    ' CAMBIA sitaucion por naturane(MARTIN)
                    'BuscaChekc = DBLet(Data2.Recordset!codsitua)
                    'If BuscaChekc <> "" Then BuscaChekc = DevuelveDesdeBD(conAri, "nomsitua", vParamAplic.Ariagro & ".rsituacioncampo", "codsitua", BuscaChekc)
                    BuscaChekc = "  "
                    If Val(DBLet(miRsAux!esnaturane, "N")) = 1 Then BuscaChekc = "Naturane"
                
                
                    NumRegElim = NumRegElim + 1
                    ''1','0','0','0','','',NULL,'0.00','0','0.00','0','','0','0','',NULL,NULL
                    FechaAlb = FechaAlb & ", (" & vUsu.Codigo & "," & RsPpal!numparte & "," & NumRegElim & ",0,"
                    FechaAlb = FechaAlb & DBSet(miRsAux!nomparti, "T")
                    FechaAlb = FechaAlb & "," & DBSet(miRsAux!nomvarie, "T") & ","
                    FechaAlb = FechaAlb & miRsAux!codCampo
                    FechaAlb = FechaAlb & ",'0','','0','0'," & DBSet(BuscaChekc, "T") & ",0,NULL,"
                    
                    'Superficie cooperativa  supsigpa  supcoope   8,4
                    'Poligono parcela recinto
                    'poligono,parcela,recintos
                    CadenaConsulta2 = "concat(supcoope,""|"",supsigpa,""|"",poligono,""#"",parcela,""#"",recintos,""#|"")"
                    BuscaChekc = DevuelveDesdeBD(conAri, CadenaConsulta2, vParamAplic.Ariagro & ".rcampos", "codcampo", miRsAux!codCampo)
                    CadenaConsulta2 = RecuperaValor(BuscaChekc, 3)
                    CadenaConsulta2 = Replace(CadenaConsulta2, "#", "|")
                    
                                            
                    FechaAlb = FechaAlb & DBSet(TransformaPuntosComas(RecuperaValor(BuscaChekc, 1)), "N", "N") & "," & DBSet(TransformaPuntosComas(RecuperaValor(BuscaChekc, 2)), "N", "N") & ",'"
                    'Poligono parcela recinto
                    BuscaChekc = RecuperaValor(CadenaConsulta2, 1)
                    If BuscaChekc = "" Then BuscaChekc = "0"
                    BuscaChekc = Format(Val(BuscaChekc), "000")
                    FechaAlb = FechaAlb & BuscaChekc & "-"
                    BuscaChekc = RecuperaValor(CadenaConsulta2, 2)
                    If BuscaChekc = "" Then BuscaChekc = "0"
                    BuscaChekc = Format(Val(BuscaChekc), "000000")
                    FechaAlb = FechaAlb & BuscaChekc & "-"
                    BuscaChekc = RecuperaValor(CadenaConsulta2, 3)
                    If BuscaChekc = "" Then BuscaChekc = "0"
                    BuscaChekc = Format(Val(BuscaChekc), "000")
                    FechaAlb = FechaAlb & BuscaChekc
                    
                    
                    
                    FechaAlb = FechaAlb & "')"
                     
                    miRsAux.MoveNext
        Wend
        If FechaAlb <> "" Then
            FechaAlb = Mid(FechaAlb, 2) 'qito la coma
            D = D & FechaAlb
            ejecutar D, False
        End If
        miRsAux.Close
        
        
        
        BuscaChekc = ""
        CadenaConsulta2 = ""
        'ALZIRA MARZO 2012
        'tENEMOS QUE METER TB LOS DATOS DEL CAMPO PARA EL INFORME
        'Para ello, ya que no tenemos una tabla , y no la voy a crear, inserto en tmpslipreu con
        'tmpslipreu`   `codusu`,`numofert` ,`numlinea`,`codartic`,`nomartic`,`ampliaci`
        '                        codcampo    codclien   nomclien

        D = miRsAux.Source
        miRsAux.Open D, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        D = "insert IGNORE into tmpslipreu(`codusu`,`numofert` ,`codartic`,`nomartic`) values "
        
        FechaAlb = ""
        NumRegElim = 0
       
        While Not miRsAux.EOF
            
               ' NumRegElim = NumRegElim + 1
               ' 'tmpslipreu`   `codusu`,`numofert` ,`numlinea`,`codartic`,`nomartic`,`ampliaci`
               ' FechaAlb = FechaAlb & ", (" & vUsu.Codigo & "," & DBLet(miRsAux!codCampo, "N") & "," & DBSet(CStr(miRsAux!codsocio), "T")
               ' FechaAlb = FechaAlb & "," & DBSet(miRsAux!nomsocio, "T") & ")"
                
                CamposIm = CamposIm & ", " & miRsAux!codCampo

                miRsAux.MoveNext
        Wend
        miRsAux.Close
        

        
        
        
        
        C2 = data4.RecordSource
        N = InStr(1, C2, " WHERE 1=1  ")
        C2 = Mid(C2, 1, N) & " WHERE 1=1  and numparte =" & RsPpal!numparte & " ORDER BY numlinea"
        Set miRsAux = New ADODB.Recordset
        miRsAux.Open C2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        NumRegElim = 0
        FechaAlb = ""
                
        
       While Not miRsAux.EOF
              
                
                
                'usu,'0','0','','',NULL,'0.00','0000-00-00','0',NULL,NULL,NULL
                'codusu`,`codclien`,`codfamia`,`nomclien`,`nomfamia`,`cantidad`,`importel`,`fechaalb`,`codprove`,`nomprove`,`codartic`,`nomartic
                FechaAlb = FechaAlb & ", (" & vUsu.Codigo & "," & RsPpal!numparte & ","
                FechaAlb = FechaAlb & DBSet(miRsAux!NomTraba, "T")
                FechaAlb = FechaAlb & "," & DBSet(miRsAux!Horas, "N") & "," & miRsAux!CodTraba & ")"
                
                miRsAux.MoveNext
           
          
           
        Wend
        miRsAux.Close
        If FechaAlb <> "" Then
             FechaAlb = Mid(FechaAlb, 2) 'qito la coma
                D = " insert into `tmpcommandest` (`codusu`,`codclien`,`nomclien`,`cantidad`,codartic) values "
                D = D & FechaAlb
                ejecutar D, False
        End If
            
        RsPpal.MoveNext
    
    Wend 'rsppal
    RsPpal.Close
        
    
    If CamposIm <> "" Then
        CamposIm = Mid(CamposIm, 2)
        C2 = Data2.RecordSource
        N = InStr(1, C2, " WHERE ")
        C2 = Mid(C2, 1, N) & " WHERE codcampo IN (" & CamposIm & ")"
        
       
        
        FechaAlb = ""
        miRsAux.Open C2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            
               ' NumRegElim = NumRegElim + 1
               ' 'tmpslipreu`   `codusu`,`numofert` ,`numlinea`,`codartic`,`nomartic`,`ampliaci`
                FechaAlb = FechaAlb & ", (" & vUsu.Codigo & "," & DBLet(miRsAux!codCampo, "N") & "," & DBSet(CStr(miRsAux!codsocio), "T")
                FechaAlb = FechaAlb & "," & DBSet(miRsAux!nomsocio, "T") & ")"
                miRsAux.MoveNext
        Wend
        miRsAux.Close
        
        If FechaAlb <> "" Then
            FechaAlb = Mid(FechaAlb, 2)
             D = "insert IGNORE into tmpslipreu(`codusu`,`numofert` ,`codartic`,`nomartic`) values " & FechaAlb
             conn.Execute D
        End If
    End If
    
    
    
    'nomrpt
    FechaAlb = DevuelveDesdeBD(conAri, "documrpt", "scryst", "codcryst", "55")
    'sql
    cadenapartes = Mid(cadenapartes, 2)
    CadenaSQL = "{advpartes.numparte} IN [" & cadenapartes & "]"
    'parametros
    D = "|pEmpresa=""" & vParam.NombreEmpresa & """|vCodUsu=" & vUsu.Codigo & "|"
    NumRegElim = 2

    
    
    
    
    
    With frmImprimir
    
        .FormulaSeleccion = CadenaSQL
        
        'parametros
    
        .OtrosParametros = D
        .NumeroParametros = NumRegElim
        .NumeroCopias = 2
        .SoloImprimir = False
        .EnvioEMail = False
        .Opcion = 3002
        .Titulo = "Partes ADV"
        .NombreRPT = FechaAlb
        .ConSubInforme = True
        .Show vbModal
    End With
    CadenaSQL = ""
    FechaAlb = ""
    
    
    
End Sub






Private Sub BotonMtoLineas(Kboton As Integer)
    If Data1.Recordset Is Nothing Then Exit Sub
    If Data1.Recordset.EOF Then Exit Sub
    
    If Val(Data1.Recordset!cerrado) = 1 Then
            MsgBox "Parte cerrado", vbExclamation
            Exit Sub
    End If
    
    
    ModificaLineas = 0
    '9, 10, 11
    If Kboton = 9 Then
        'Campos socio
        BotonMtoLineasCampos
        Me.lblIndicador.Caption = "Campos"
    ElseIf Kboton = 10 Then
        Me.SSTab1.Tab = 0
        PonerModo 6
        Me.lblIndicador.Caption = "Articulos"
    Else
        Me.SSTab1.Tab = 1
        PonerModo 7
        Me.lblIndicador.Caption = "Trabajadores"
    End If
    PonerBotonCabecera True
    
End Sub

Private Sub BotonMtoLineasCampos()
        PonerModo 5
End Sub





Private Sub BotonModificarLinea()
'Modificar una linea
Dim vWhere As String
    
    On Error GoTo EModificarLinea
    If Modo = 5 Then
        MsgBox "Quite el campo(no se puede modificar)", vbExclamation
        Exit Sub
    End If
    
    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Then Exit Sub '1= Insertar
    
    'bloqueamos
    'If Not BloqueaRegistro("advpartesdtos", vWhere) Then Exit Sub
    
    If Modo = 6 Then
        Me.SSTab1.Tab = 0
        If data3.Recordset.EOF Then Exit Sub
        CargaTxtAuxArt True, False
        PonerFoco txtAux(4)
        DataGrid1.Enabled = False
    Else
        Me.SSTab1.Tab = 1
        If data4.Recordset.EOF Then Exit Sub
        CargatxtAuxT True, False
        PonerFoco txtAuxT(0)
        
    End If
    
    
    
    
    ModificaLineas = 2 'Modificar

    'Añadiremos el boton de aceptar y demas objetos para insertar
    Me.lblIndicador.Caption = "MODIFICAR"
    PonerBotonCabecera False
    If Modo = 6 Then
        Me.DataGrid2.Enabled = False
    ElseIf Modo = 7 Then
        Me.DataGrid3.Enabled = False
    End If
    
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
    
    
    Me.SSTab1.Tab = 0
    If Modo = 7 Then Me.SSTab1.Tab = 1
    

    If Modo = 5 Then
        
        lblIndicador.Caption = "INSERTAR CAM"
        AnyadirLinea DataGrid2, Data2
        CargaTxtAux2 True, True

        PonerFoco txtAux2(0)
        Me.DataGrid2.Enabled = False

    ElseIf Modo = 6 Then
        lblIndicador.Caption = "INSERTAR ART"
        AnyadirLinea DataGrid1, data3
        CargaTxtAuxArt True, True
        Combo1.ListIndex = 1
        txtAux(0).Text = AlmacenLin
        PonerFoco txtAux(0)
        Me.DataGrid1.Enabled = False
    Else
        lblIndicador.Caption = "INSERTAR Trab"
        AnyadirLinea DataGrid3, data4
        CargatxtAuxT True, True

        PonerFoco txtAuxT(0)
        Me.DataGrid3.Enabled = False

    End If

End Sub

Private Sub BotonEliminarLinea()
    Select Case Modo
    Case 5
        BotonEliminarLineaCampo
    Case 6
        BotonEliminarLineaArticulo
    Case 7
        BotonEliminarLineaTrabajador
    End Select
End Sub

Private Sub BotonEliminarLineaCampo()
'Eliminar una linea De Mantenimiento. Tabla: slima1
Dim Sql As String
    
    On Error GoTo EEliminarLinea

    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas > 0 Then Exit Sub '1= Insertar, 2=Modificar

    If Data2.Recordset.EOF Then Exit Sub

    
    

    Sql = "¿Seguro que desea eliminar el campo?     "
    Sql = Sql & vbCrLf & "Codigo:  " & Data2.Recordset!codCampo & vbCrLf
    Sql = Sql & "Descripcion:  " & Data2.Recordset!nomparti & "(" & Data2.Recordset!nomvarie & ")"
    


    If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        NumRegElim = Data2.Recordset.AbsolutePosition
        Sql = " numparte = " & Text1(0).Text & " AND codCampo=" & Data2.Recordset!codCampo
        Sql = "Delete from advpartes_campos WHERE " & Sql
        conn.Execute Sql

        ModificaLineas = 0
        CargaGrid2 DataGrid2, Data2
        SituarDataTrasEliminar Data2, NumRegElim

    End If
    PonerFocoBtn Me.cmdRegresar

EEliminarLinea:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Lineas descuentos", Err.Description
End Sub

Private Sub BotonEliminarLineaArticulo()
'Eliminar una linea De Mantenimiento. Tabla: slima1
Dim Sql As String
    
    On Error GoTo EEliminarLinea

    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas > 0 Then Exit Sub '1= Insertar, 2=Modificar

    If data3.Recordset.EOF Then Exit Sub

    Me.SSTab1.Tab = 0
    

    Sql = "¿Seguro que desea eliminar el articulo?     "
    Sql = Sql & vbCrLf & "Articulo:  " & data3.Recordset!NomArtic & vbCrLf
    Sql = Sql & "Cantidad:  " & data3.Recordset!cantidad
    


    If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        NumRegElim = data3.Recordset.AbsolutePosition
        Sql = " numparte = " & Text1(0).Text & " AND numlinea=" & data3.Recordset!numlinea
        Sql = "Delete from advparteslineas WHERE " & Sql
        conn.Execute Sql

        ModificaLineas = 0
        CargaGridArt DataGrid1, data3
        SituarDataTrasEliminar data3, NumRegElim

    End If
    PonerFocoBtn Me.cmdRegresar

EEliminarLinea:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Lineas descuentos", Err.Description
End Sub

Private Sub BotonEliminarLineaTrabajador()
'Eliminar una linea De Mantenimiento. Tabla: slima1
Dim Sql As String
    
    On Error GoTo EEliminarLinea

    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas > 0 Then Exit Sub '1= Insertar, 2=Modificar

    If Me.data4.Recordset.EOF Then Exit Sub

    Me.SSTab1.Tab = 1
    

    Sql = "¿Seguro que desea eliminar el trabajador?     "
    Sql = Sql & vbCrLf & "Trab:  " & data4.Recordset!CodTraba & " " & data4.Recordset!NomTraba & vbCrLf
    Sql = Sql & "Horas:  " & data4.Recordset!Horas
    


    If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        NumRegElim = data4.Recordset.AbsolutePosition
        Sql = " numparte = " & Text1(0).Text & " AND numlinea=" & data4.Recordset!numlinea
        Sql = "Delete from advpartes_trabajador WHERE " & Sql
        conn.Execute Sql

        ModificaLineas = 0
        CargaGridTrab DataGrid3, data4
        SituarDataTrasEliminar data4, NumRegElim

    End If
    PonerFocoBtn Me.cmdRegresar

EEliminarLinea:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Lineas trabajadores", Err.Description
End Sub



Private Function MontaSQLCarga(enlaza As Boolean, Cual As Byte) As String
'--------------------------------------------------------------------
'   1.- Campos
'   2.- Articulos
'   3.- Trabajadores
'--------------------------------------------------------------------
Dim TraEnTmp As Boolean
Dim Sql As String
Dim Orderby As String

    Select Case Cual
    Case 1
        'CAMPOS.
        'El select enlaza con ariagro
        'pero where codcampo a los de aqui. Con lo cual este es un poco distinto
        CadenaEnlazeAriagro Sql, False
        Sql = Sql & " WHERE codcampo IN "
        If enlaza Then
            Sql = Sql & "(select codcampo from advpartes_campos where numparte=" & Data1.Recordset!numparte & ")"
        Else
            Sql = Sql & "(-1)"
        End If
        
        'Nos salimos ya
        MontaSQLCarga = Sql
        Exit Function
    Case 2
        'Articulos
        Sql = "select numlinea,advparteslineas.codalmac,advparteslineas.codartic,nomartic,dosishab,cantidad,"
        Sql = Sql & " advparteslineas.preciove,origpre,dtoline1,dtoline2,importel,nomalmac,ampliaci,if(fijo=1,""Si"","""") fijo2 from"
        Sql = Sql & " advparteslineas,sartic,salmpr where advparteslineas.codartic=sartic.codartic "
        Sql = Sql & " AND advparteslineas.codalmac = salmpr.codalmac "
        Orderby = " ORDER BY numlinea"
    Case 3
        'trabajadores
        TraEnTmp = False  'leera nombre en advpartes, si esta abierto en la temporal
        If enlaza Then
            If Me.Data1.Recordset!cerrado = 0 Then TraEnTmp = True
        End If
        
         If TraEnTmp Then
            If BDAripres = "" Then
                'No hay inidicado ARIPRES
                Sql = "select numlinea,advpartes_trabajador.codtraba,"
                'el nombre. Si no lo encuentra que pnga no enontrado
                Sql = Sql & "if(trabajadores.nomtraba is null,""*** No encontrad  o"",trabajadores.nomtraba) nomtraba,horas "
                Sql = Sql & " FROM advpartes_trabajador left join straba trabajadores on advpartes_trabajador.codtraba=trabajadores.codtraba"
            
            Else
                Sql = "select numlinea,advpartes_trabajador.codtraba,"
                'el nombre. Si no lo encuentra que pnga no enontrado
                Sql = Sql & "if(trabajadores.nomtrabajador is null,""*** No encontrad  o"",trabajadores.nomtrabajador) nomtraba,horas "
                Sql = Sql & " FROM advpartes_trabajador left join " & BDAripres & ".trabajadores on advpartes_trabajador.codtraba=trabajadores.idtrabajador"
            End If
        Else
            Sql = "SELECT numlinea,codtraba,nomtraba,horas from advpartes_trabajador"
        End If
        Sql = Sql & " WHERE 1=1 "
        Orderby = " ORDER BY numlinea"
    End Select
    Sql = Sql & " and numparte = "
    If enlaza Then
        Sql = Sql & Data1.Recordset!numparte
      
    Else
        Sql = Sql & "  -1"
    End If
    Sql = Sql & Orderby
    MontaSQLCarga = Sql
End Function


Private Sub CadenaEnlazeAriagro(ByRef Sql As String, DesdeMulti As Boolean)
    'Para no meter MUCHOS ariagro.ss
    'Pongo @# y luego lo reemplazo por vparamaplic.Ariagro.
'    SQL = "select rcampos.codcampo, rpartida.nomparti, variedades.nomvarie"
'    SQL = SQL & " from (@#rcampos inner join @#rpartida on rcampos.codparti = rpartida.codparti)"
'    SQL = SQL & " inner join @#variedades on rcampos.codvarie = variedades.codvarie"
'    'where socio
'    SQL = Replace(SQL, "@#", vParamAplic.Ariagro & ".")
'
    
    'SQL = "select rcampos.codcampo, rpartida.nomparti, variedades.nomvarie,rcampos.codclien,rsocios.codsocio,rsocios.nomsocio,rcampos.esnaturane"
    Sql = "select rcampos.codcampo, rpartida.nomparti, variedades.nomvarie,rcampos.codclien,rsocios.codsocio,rsocios.nomsocio,rcampos.esnaturane"
    If DesdeMulti Then Sql = Sql & " ,nomclien,sclien.codsitua,nrocampo,rcampos.codvarie    "
    
    Sql = Sql & " from ((@#rcampos inner join @#rpartida on rcampos.codparti = rpartida.codparti)"
    Sql = Sql & " inner join @#variedades on rcampos.codvarie = variedades.codvarie)"
    Sql = Sql & " inner join @#rsocios on rsocios.codsocio=rcampos.codsocio"
    If DesdeMulti Then Sql = Sql & " left join ariges" & vEmpresa.codempre & ".sclien on rcampos.codclien=sclien.codclien  "
    Sql = Replace(Sql, "@#", vParamAplic.Ariagro & ".")
    
    
End Sub


Private Sub PonerCamposLineas()
'Carga las Pestañas con las tablas de lineas del Trabajador seleccionado para mostrar

    On Error GoTo EPonerLineas

    Screen.MousePointer = vbHourglass

    'Datos de la tabla slipre
    CargaGrid True, 0  '0:todos

    Screen.MousePointer = vbDefault
    Exit Sub
    
EPonerLineas:
    MuestraError Err.Number, "PonerCamposLineas"
    PonerModo 2
    Screen.MousePointer = vbDefault
End Sub


'Cual:
' 0: todos
' 1: solo el 1
' 2          2
' 3          3
Private Sub CargaGrid(enlaza As Boolean, Cual As Byte)
Dim B As Boolean


Dim Sql As String

    On Error GoTo ECargaGrid

    
        
        
    
    If Cual = 0 Or Cual = 1 Then
        Sql = MontaSQLCarga(enlaza, 1)
        CargaGridGnral DataGrid2, Data2, Sql, True

    
        CargaGrid2 DataGrid2, Data2
        DataGrid1.ScrollBars = dbgAutomatic
        
    
        B = (Modo = 5) And (ModificaLineas = 1 Or ModificaLineas = 2)
    
        DataGrid1.Enabled = B
    End If
    
    
    If Cual = 0 Or Cual = 2 Then
        'Articulos
         Sql = MontaSQLCarga(enlaza, 2)
         CargaGridGnral DataGrid1, Me.data3, Sql, True

         CargaGridArt DataGrid1, data3
'
        B = (Modo = 6) And (ModificaLineas = 1 Or ModificaLineas = 2)
        
        DataGrid2.Enabled = Not B
        DataGrid2.ScrollBars = dbgAutomatic
        
    End If
    
    If Cual = 0 Or Cual = 3 Then
        'Trabajadores
         Sql = MontaSQLCarga(enlaza, 3)
         CargaGridGnral DataGrid3, Me.data4, Sql, True

         CargaGridTrab DataGrid3, data4
'
         B = (Modo = 7) And (ModificaLineas = 1 Or ModificaLineas = 2)
        
         DataGrid3.Enabled = Not B
         DataGrid3.ScrollBars = dbgAutomatic
        
    End If

    
    Exit Sub
    
ECargaGrid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub


Private Sub CargaGrid2(ByRef vDataGrid As DataGrid, ByRef vData As Adodc)
Dim i As Integer

    On Error GoTo ECargaGrid

    'vData.Refresh

                vDataGrid.Columns(0).Caption = "Código"
                vDataGrid.Columns(0).Width = 800
                vDataGrid.Columns(0).Alignment = dbgRight
                vDataGrid.Columns(0).NumberFormat = "0000"
                
                vDataGrid.Columns(1).Caption = "Partida"
                vDataGrid.Columns(1).Width = 2500
 
                vDataGrid.Columns(2).Caption = "Variedad"
                vDataGrid.Columns(2).Width = 2300
                
                
                vDataGrid.Columns(3).visible = False
                vDataGrid.Columns(4).visible = False
                vDataGrid.Columns(5).visible = False
                vDataGrid.Columns(6).visible = False  'es naturane

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
Private Sub CargaTxtAux2(visible As Boolean, limpiar As Boolean)
'IN: visible: si es true ponerlos visibles en la posición adecuada
'    limpiar: si es true vaciar los txtAux
Dim alto As Single
Dim i As Byte

    If Not visible Then
        'Fijamos el alto (ponerlo en la parte inferior del form)
        For i = 0 To txtAux2.Count - 1 'TextBox
            txtAux2(i).Top = 290
            txtAux2(i).visible = visible
        Next i
    
    Else
        If limpiar Then 'Vaciar los textBox (Vamos a Insertar)
            DeseleccionaGrid DataGrid2
            For i = 0 To txtAux2.Count - 1
                txtAux2(i).Text = ""
                BloquearTxt txtAux2(i), i > 0
            Next i

        Else 'Vamos a modificar
            For i = 0 To txtAux2.Count - 1
         
                txtAux2(i).Text = DataGrid2.Columns(i + 2).Text
            
                txtAux2(i).Locked = False
            Next i
        End If
        


        'Fijamos altura(Height) y posición Top
        '-------------------------------
        alto = ObtenerAlto(DataGrid2, 10)
        
        For i = 0 To txtAux2.Count - 1
            txtAux2(i).Top = alto
            txtAux2(i).Height = DataGrid2.RowHeight
        Next i
        cmdCampo.Top = alto
        cmdCampo.Height = DataGrid2.RowHeight
        
        
        'Fijamos anchura y posicion Left
        '--------------------------------
        'Cod. Almac

        'Precio, Dto1, Dto2, Precio

        alto = DataGrid2.Left + 120
        txtAux2(0).Left = alto + DataGrid2.Columns(0).Left
        txtAux2(0).Width = DataGrid2.Columns(0).Width - 150
        
        txtAux2(1).Left = alto + DataGrid2.Columns(1).Left + 90
        txtAux2(1).Width = DataGrid2.Columns(1).Width - 210
        txtAux2(2).Left = alto + DataGrid2.Columns(2).Left
        txtAux2(2).Width = DataGrid2.Columns(2).Width - 210
        
        cmdCampo.Left = alto + DataGrid2.Columns(1).Left - 120



        'Los ponemos Visibles o No
        '--------------------------
        For i = 0 To txtAux2.Count - 1
            txtAux2(i).visible = visible
        Next i
        
        
        
        
    End If
    cmdCampo.visible = visible
End Sub


Private Sub CargaTxtAuxArt(visible As Boolean, limpiar As Boolean)
'IN: visible: si es true ponerlos visibles en la posición adecuada
'    limpiar: si es true vaciar los txtAux
Dim alto As Single
Dim i As Byte

    If Not visible Then
        'Fijamos el alto (ponerlo en la parte inferior del form)
        For i = 0 To txtAux.Count - 3 'TextBox la ampliacion siempre esta visible
            txtAux(i).Top = 290
            txtAux(i).visible = visible
        Next i
            
        BloquearTxt txtAux(11), True
        Combo1.visible = False
        Combo1.Enabled = False
    Else
    
        DeseleccionaGrid DataGrid1
        If limpiar Then 'Vaciar los textBox (Vamos a Insertar)
            
            For i = 0 To txtAux.Count - 1
                txtAux(i).Text = ""
                BloquearTxt txtAux(i), Not (i <> 2 And i <> 10 And i <> 6)
            Next i
            Combo1.ListIndex = -1
        Else 'Vamos a modificar
            For i = 0 To txtAux.Count - 1
                txtAux(i).Text = DataGrid1.Columns(i + 1).Text
                BloquearTxt txtAux(i), Not (i <> 0 And i <> 2 And i <> 10 And i <> 6)
            Next i
            If DBLet(data3.Recordset!fijo2, "T") = "" Then
                Combo1.ListIndex = 1
            Else
                Combo1.ListIndex = 0
            End If
        End If
        
        

        'Fijamos altura(Height) y posición Top
        '-------------------------------
        alto = ObtenerAlto(DataGrid1, 10)
        
        For i = 0 To txtAux.Count - 1
            If i < 10 Then txtAux(i).Top = alto
            txtAux(i).Height = DataGrid1.RowHeight
        Next i
        Me.cmdAux(0).Top = alto
        cmdAux(0).Height = DataGrid1.RowHeight
        
        Me.cmdAux(1).Top = alto
        cmdAux(1).Height = DataGrid1.RowHeight
        
        Combo1.Top = alto
       
        Combo1.Enabled = True
        Combo1.visible = True
        alto = DataGrid1.Left + 15
        txtAux(0).Left = alto + DataGrid1.Columns(1).Left
        txtAux(0).Width = DataGrid1.Columns(1).Width

        txtAux(1).Left = alto + DataGrid1.Columns(2).Left + 30
        txtAux(1).Width = DataGrid1.Columns(2).Width - 60
        
 
        txtAux(2).Left = alto + DataGrid1.Columns(3).Left
        txtAux(2).Width = DataGrid1.Columns(3).Width - 60
        cmdAux(0).Left = txtAux(1).Left - 240
        cmdAux(1).Left = txtAux(2).Left - 90
        

        For i = 3 To txtAux.Count - 3 'ampliaci es visible
            
            txtAux(i).Left = alto + DataGrid1.Columns(i + 1).Left
            txtAux(i).Width = DataGrid1.Columns(i + 1).Width - 15
        
        Next i
        i = 10
        Combo1.Left = DataGrid1.Columns(i).Left + DataGrid1.Columns(i).Width + 150
        
        'Los ponemos Visibles o No
        '--------------------------
        For i = 0 To txtAux.Count - 3
            txtAux(i).visible = visible
        Next i
        
        
        
    End If
    Me.cmdAux(0).visible = visible
    Me.cmdAux(1).visible = visible
End Sub




'Esta funcion sustituye a LlamaLineas
Private Sub CargatxtAuxT(visible As Boolean, limpiar As Boolean)
'IN: visible: si es true ponerlos visibles en la posición adecuada
'    limpiar: si es true vaciar los txtAux
Dim alto As Single
Dim i As Byte

    If Not visible Then
        'Fijamos el alto (ponerlo en la parte inferior del form)
        For i = 0 To txtAuxT.Count - 1 'TextBox
            txtAuxT(i).Top = 290
            txtAuxT(i).visible = visible
        Next i
    
    Else
        If limpiar Then 'Vaciar los textBox (Vamos a Insertar)
            DeseleccionaGrid DataGrid3
            For i = 0 To txtAuxT.Count - 1
                txtAuxT(i).Text = ""
                BloquearTxt txtAuxT(i), i = 1
            Next i

        Else 'Vamos a modificar
            For i = 0 To txtAuxT.Count - 1
         
                txtAuxT(i).Text = DataGrid3.Columns(i + 1).Text
            
                txtAuxT(i).Locked = False
            Next i
            BloquearTxt txtAuxT(1), True
        End If
        


        'Fijamos altura(Height) y posición Top
        '-------------------------------
        alto = ObtenerAlto(DataGrid3, 10)
        
        For i = 0 To txtAuxT.Count - 1
            txtAuxT(i).Top = alto
            txtAuxT(i).Height = DataGrid3.RowHeight
        Next i
        cmdTra.Top = alto
        cmdTra.Height = DataGrid3.RowHeight
        
        
        'Fijamos anchura y posicion Left
        '--------------------------------
        'Cod. Almac

        'Precio, Dto1, Dto2, Precio

        alto = DataGrid3.Left
        txtAuxT(0).Left = alto + DataGrid3.Columns(1).Left
        txtAuxT(0).Width = DataGrid3.Columns(1).Width - 120
        
        txtAuxT(1).Left = alto + DataGrid3.Columns(2).Left + 90
        txtAuxT(1).Width = DataGrid3.Columns(2).Width - 15
        txtAuxT(2).Left = alto + DataGrid3.Columns(3).Left
        txtAuxT(2).Width = DataGrid3.Columns(3).Width - 5
        
        cmdTra.Left = alto + DataGrid3.Columns(2).Left - 120



        'Los ponemos Visibles o No
        '--------------------------
        For i = 0 To txtAuxT.Count - 1
            txtAuxT(i).visible = visible
        Next i
        
        
        
        
    End If
    cmdTra.visible = visible
End Sub





Private Sub CargaGridTrab(ByRef vDataGrid As DataGrid, ByRef vData As Adodc)
Dim i As Integer

    On Error GoTo ECargaGrid

    'vData.Refresh
                vDataGrid.Columns(0).visible = False
                
                vDataGrid.Columns(1).Caption = "Código"
                vDataGrid.Columns(1).Width = 1300
                vDataGrid.Columns(1).Alignment = dbgRight
                vDataGrid.Columns(1).NumberFormat = "0000"
                
                vDataGrid.Columns(2).Caption = "Trabajador"
                vDataGrid.Columns(2).Width = 2500
 
                vDataGrid.Columns(3).Caption = "Horas"
                vDataGrid.Columns(3).Width = 1300
                vDataGrid.Columns(3).Alignment = dbgRight
                vDataGrid.Columns(3).NumberFormat = FormatoImporte
                
                


    For i = 0 To vDataGrid.Columns.Count - 1
        vDataGrid.Columns(i).Locked = True
        vDataGrid.Columns(i).AllowSizing = False
    Next i
    vDataGrid.HoldFields
    Exit Sub
ECargaGrid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub





Private Sub TxtAux_Change(Index As Integer)
    If Index = 6 And ModificaLineas = 2 Then 'Precio y Modo Modificar Lineas
        txtAux(7).Text = "M"
    End If
End Sub

Private Sub txtAux_GotFocus(Index As Integer)

    ConseguirFoco txtAux(Index), Modo
  
    
End Sub

Private Sub TxtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Not (Index = 0 And KeyCode = 38) Then KEYdown KeyCode
    
End Sub

Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 11 And KeyAscii = 13 Then
        PonerFocoBtn Me.cmdAceptar
    Else
        KEYpress KeyAscii
    End If
End Sub



Private Sub txtAux_LostFocus(Index As Integer)
Dim Cad As String
Dim CPrecioFact As CPreciosFact
Dim OrigP As String
Dim Precio As String
Dim vC As Currency
    
    
    txtAux(Index).Text = Trim(txtAux(Index).Text)
    Select Case Index
    Case 0
        Cad = ""
        If txtAux(Index).Text <> "" Then
            If PonerFormatoEntero(txtAux(Index)) Then
                Cad = DevuelveDesdeBD(conAri, "nomalmac", "salmpr", "codalmac", txtAux(Index).Text)
                If Cad = "" Then
                    MsgBox "No existe almacen", vbExclamation
                End If
            End If
        End If
        txtAux(10).Text = Cad
        If Cad = "" And txtAux(Index).Text <> "" Then
            txtAux(Index).Text = ""
            PonerFoco txtAux(Index)
        End If
    Case 1
        Cad = ""
        If txtAux(Index).Text <> "" Then
            Precio = "codunida"
            Cad = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", txtAux(Index).Text, "T", Precio)
            If Cad = "" Then
                MsgBox "No existe artículo", vbExclamation
            Else
                txtAux(5).Text = ""
                txtAux(9).Text = ""
                    
            End If
            Precio = ""
        End If
        txtAux(2).Text = "  " & Cad
        If Cad = "" And txtAux(Index).Text <> "" Then
            txtAux(Index).Text = ""
            PonerFoco txtAux(Index)
        Else
            'Ok. Borramos importe y pone
        End If
    
    Case 3, 4, 5
        'dosishab  'cantidad 'preciove
        If Not PonerFormatoDecimal(Me.txtAux(Index), 2) Then
            Me.txtAux(Index).Text = ""
        Else
            If Index = 3 Then 'dosis
                If txtAux(1).Text <> "" Then
                    Cad = " sunida.codunida=sartic.codunida AND codartic"
                    Cad = DevuelveDesdeBD(conAri, "estrabajo", "sunida,sartic", Cad, txtAux(1).Text, "T")
                    If Cad = "0" Then
                        'MsgBox "No debe indicar dosis", vbExclamation
                    Else
                        vC = ImporteFormateado(txtAux(Index).Text)
                        vC = Val(Data1.Recordset!litrospre) * vC
                        vC = Round(vC / 1000, 3)
                        txtAux(4).Text = Format(vC, FormatoPrecio)
                    End If
                End If
            End If
        End If
    
    
        
    Case 7, 8
        'descuentos
        If Not PonerFormatoDecimal(Me.txtAux(Index), 6) Then Me.txtAux(Index).Text = ""
    
    Case 9
        If Not PonerFormatoDecimal(Me.txtAux(Index), 1) Then Me.txtAux(Index).Text = ""
    End Select
    
    
    'precio articulo
    If Index = 0 Or Index = 1 Or Index = 4 Then
        If Me.txtAux(5).Text = "" Then
            If Me.txtAux(0).Text <> "" And Me.txtAux(1).Text <> "" And txtAux(4).Text <> "" Then
                'Ha puesto almacen y articulo y no hay precio
                    
                            Set CPrecioFact = New CPreciosFact
                    
                            CPrecioFact.CodigoArtic = txtAux(1).Text
                            CPrecioFact.CodigoClien = Text1(6).Text
                            CPrecioFact.FijarTarifaActividad
                            
                            
                            Precio = CPrecioFact.ObtenerPrecio(False, Text1(1).Text, OrigP, "")
                            'Si PorCaja vuelve de ObtenerPrecio a false se calcula con precio unidad aunque NumCajas>0
                            'Ya que a regresado con pvp del Articulo

                            txtAux(5).Text = Precio
                            txtAux(6).Text = OrigP 'De donde viene el precio
                    
                            
                            txtAux(7).Text = CPrecioFact.Descuento1
                            PonerFormatoDecimal txtAux(8), 4
                            txtAux(8).Text = CPrecioFact.Descuento2
                            PonerFormatoDecimal txtAux(9), 4

                            Set CPrecioFact = Nothing
                
                
            End If
        End If
    End If
    
    
    If Modo = 6 Then
    
        If (Index = 4 Or Index = 5 Or Index = 7 Or Index = 8) Then 'Cant., Precio, dto1, dto2
            If txtAux(1).Text = "" Then Exit Sub 'Cod artic
            txtAux(9).Text = CalcularImporte(txtAux(4).Text, txtAux(5).Text, txtAux(7).Text, txtAux(8).Text, vParamAplic.TipoDtos)
            PonerFormatoDecimal txtAux(9), 1
        End If
    
    End If
End Sub

Private Sub txtAux2_GotFocus(Index As Integer)
Dim cadkey As Integer

    cadkey = ObtenerCadKey(kCampo, Index)
    kCampo = Index
'    ConseguirFoco txtAux2(Index), Modo, cadkey
    ConseguirFocoLin txtAux2(Index), cadkey
    

    
End Sub

Private Sub TxtAux2_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    If Not (Index = 0 And KeyCode = 38) Then KEYdown KeyCode
    

End Sub


Private Sub txtAux2_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub







Private Sub txtAux2_LostFocus(Index As Integer)

    txtAux2(Index).Text = Trim(txtAux2(Index).Text)
    
    If Index = 0 Then
        D = ""
        If txtAux2(Index).Text <> "" Then
            If PonerFormatoEntero(txtAux2(Index)) Then
                Set miRsAux = New ADODB.Recordset
                CadenaEnlazeAriagro D, False
                miRsAux.Open D, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If miRsAux.EOF Then
                    D = ""
                    MsgBox "No existe el campo: " & txtAux2(0).Text, vbExclamation
                    txtAux2(0).Text = ""
                Else
                    txtAux2(1).Text = miRsAux!nomparti
                    txtAux2(2).Text = miRsAux!nomvarie
                
                End If
                miRsAux.Close
                Set miRsAux = Nothing
            Else
                txtAux2(0).Text = ""
            End If
        End If
        If D = "" Then
            txtAux2(1).Text = ""
            txtAux2(2).Text = ""
        Else
            PonerFocoBtn Me.cmdAceptar
        End If
    End If
    
    
End Sub


Private Function InsertarModificar() As Boolean
    On Error GoTo EIns   'general
    InsertarModificar = False
    
    
    If Modo = 5 Then InsertarModificar = InsertarModificarCampo
    If Modo = 6 Then InsertarModificar = InsertarModificarArticulo
    If Modo = 7 Then InsertarModificar = InsertarModificarTrabajador
    
    Exit Function
EIns:
    MuestraError Err.Number, "Insertar linea"
End Function


Private Function InsertarModificarCampo() As Boolean
    'Solo hay INSERT
    If txtAux2(0).Text = "" Then
        MsgBox "Falta campo", vbExclamation
        InsertarModificarCampo = False
    Else

        D = "INSERT INTO advpartes_campos( numparte, codcampo) VALUES (" & Me.Text1(0).Text & "," & txtAux2(0).Text & ")"
        InsertarModificarCampo = ejecutar(D, False)
    End If
        
End Function


Private Function InsertarModificarArticulo() As Boolean
Dim i As Integer
Dim Cad As String
Dim DeDosis As Boolean
    On Error GoTo eInsertarModificarArticulo
    InsertarModificarArticulo = False
    For i = 0 To Me.txtAux.Count - 2 'ampliacion NO requerida
        txtAux(i).Text = Trim(txtAux(i).Text)
        If i <> 3 And i <> 4 Then
            If txtAux(i).Text = "" Then
                PonerFoco txtAux(i)
                MsgBox "Campos requeridos", vbExclamation
                Exit Function
            End If
        End If
    Next i

    If ArtiHORAS = Me.txtAux(1).Text Then
        MsgBox "No puede poner el articulo HORAS(" & ArtiHORAS & ")", vbExclamation
        Exit Function
    End If
    
    
    
    
    Cad = " sunida.codunida=sartic.codunida AND codartic"
    Cad = DevuelveDesdeBD(conAri, "estrabajo", "sunida,sartic", Cad, txtAux(1).Text, "T")
    DeDosis = Cad = "1"

    If Not DeDosis Then
        If txtAux(3).Text <> "" Then
            If MsgBox("No deberia indicar dosis para este artículo.¿Continuar?", vbQuestion + vbYesNo) = vbNo Then
                
                PonerFoco txtAux(3)
                Exit Function
            End If
        End If
    End If

'    If Not (txtAux(3).Text = "" Xor txtAux(4).Text = "") Then
'        MsgBox "Dosis o cantidad", vbExclamation
'        Exit Function
'    End If
    
    If txtAux(3).Text <> "" And Combo1.ListIndex = 0 Then
        'Pone cantidad fija
        MsgBox "No se puede poner cantidad fija a una dosis", vbExclamation
        Exit Function
    End If
    
    
    
    'Diciembre 2014
    'Los articulos pueden llevar una marca de que sea solo para partes externos o internos
    '0 = cualquiera  1 Internos   2 Externos'
    Cad = DevuelveDesdeBD(conAri, "partesADV", "sartic", "codartic", Me.txtAux(1).Text, "T")
    If Val(Cad) > 0 Then
        'Externos
        If Val(Data1.Recordset!EsExterno) = 1 Then
            'EXTERNO.
            If Val(Cad) = 1 Then
                Cad = "externos"
            Else
                Cad = ""
            End If
        Else
            'INTERNOS
            If Val(Cad) = 2 Then
                Cad = "internos"
            Else
                Cad = ""
            End If
        End If
        If Cad <> "" Then
            Cad = "Artículo con la marca para partes de trabajo " & Cad
            MsgBox Cad, vbExclamation
            Exit Function
        End If
              
    
    End If
    
    
    If ModificaLineas = 1 Then
        'Insertar
        Cad = DevuelveDesdeBD(conAri, "max(numlinea)", "advparteslineas", "numparte", CStr(Data1.Recordset!numparte))
        If Cad = "" Then Cad = "0"
        i = Val(Cad) + 1
        Cad = "insert into `advparteslineas` (`numparte`,`numlinea`,`codalmac`,`codartic`,`dosishab`,`cantidad`,`preciove`,`importel`,`ampliaci`,`dtoline1`,`dtoline2`,`origpre`,fijo) values ("
        Cad = Cad & Data1.Recordset!numparte & "," & i & "," & txtAux(0).Text & ","
        'codartic,dosis,cantidad
        Cad = Cad & DBSet(txtAux(1).Text, "T") & "," & DBSet(txtAux(3).Text, "N", "S") & "," & DBSet(txtAux(4).Text, "N", "S") & ","
        '`preciove`,`importel`,`ampliaci`,
        Cad = Cad & DBSet(txtAux(5).Text, "N") & "," & DBSet(txtAux(9).Text, "N") & "," & DBSet(txtAux(11).Text, "T") & ","
        '`dtoline1`,`dtoline2`,`origpre`
        Cad = Cad & DBSet(txtAux(7).Text, "N") & "," & DBSet(txtAux(8).Text, "N") & "," & DBSet(txtAux(6).Text, "T")
        Cad = Cad & "," & Combo1.ItemData(Combo1.ListIndex) & ")"
    
    Else
        Cad = "UPDATE advparteslineas SET "
        'codartic,dosis,cantidad
        
        Cad = Cad & " codartic= " & DBSet(txtAux(1).Text, "T")
        For i = 4 To 10
            If i = 7 Then
                Cad = Cad & ", " & Me.data3.Recordset.Fields(i).Name & " = " & DBSet(txtAux(i - 1).Text, "T")
            Else
                Cad = Cad & ", " & Me.data3.Recordset.Fields(i).Name & " = " & DBSet(txtAux(i - 1).Text, "N", "S")
            End If
        Next i
        Cad = Cad & ", Ampliaci = " & DBSet(txtAux(11).Text, "T")
        Cad = Cad & ", fijo = " & Combo1.ItemData(Combo1.ListIndex)
        
        Cad = Cad & " WHERE numparte = " & Data1.Recordset!numparte & " AND numlinea = " & data3.Recordset!numlinea
    End If
    
    If ejecutar(Cad, False) Then InsertarModificarArticulo = True
    Exit Function
eInsertarModificarArticulo:
    MuestraError Err.Number, Err.Description
End Function




Private Function InsertarModificarTrabajador() As Boolean
    'Solo hay INSERT
    If txtAuxT(0).Text = "" Or txtAuxT(1).Text = "" Or txtAuxT(2).Text = "" Then
        MsgBox "Campos obligatorios", vbExclamation
        InsertarModificarTrabajador = False
        Exit Function
    End If


    If ModificaLineas = 1 Then
        D = DevuelveDesdeBD(conAri, "max(numlinea)", "advpartes_trabajador", "numparte", CStr(Data1.Recordset!numparte))
        If D = "" Then D = "0"
        D = CStr(Val(D) + 1)
        D = "INSERT INTO advpartes_trabajador(numparte,numlinea,codtraba,nomtraba,horas) VALUES (" & Data1.Recordset!numparte & "," & D
        D = D & "," & txtAuxT(0).Text & ",'FALTA CERRAR'," & DBSet(txtAuxT(2), "N") & ")"
        
    Else
        D = "UPDATE advpartes_trabajador SET codtraba=" & txtAuxT(0).Text & ", horas =" & DBSet(txtAuxT(2), "N")
        D = D & " WHERE numparte = " & Data1.Recordset!numparte & " AND numlinea = " & data4.Recordset!numlinea
    End If
    InsertarModificarTrabajador = ejecutar(D, False)

        
End Function




'Articulos
Private Sub CargaGridArt(ByRef vDataGrid As DataGrid, ByRef vData As Adodc)
Dim i As Byte
    
    On Error GoTo ECargaGrid

    vData.Refresh
   
    vDataGrid.Columns(0).visible = False


            i = 1
            vDataGrid.Columns(i).Caption = "Alm."
            vDataGrid.Columns(i).Width = 470
            vDataGrid.Columns(i).NumberFormat = "000"
            
            i = i + 1 '4
            vDataGrid.Columns(i).Caption = "Articulo"
            vDataGrid.Columns(i).Width = 1450
            
            i = i + 1 '5
            vDataGrid.Columns(i).Caption = "Desc. Artículo"
            vDataGrid.Columns(i).Width = 3400

            i = i + 1
            vDataGrid.Columns(i).Caption = "Dos/hab"
            vDataGrid.Columns(i).Width = 1000
            vDataGrid.Columns(i).Alignment = dbgRight
            vDataGrid.Columns(i).NumberFormat = FormatoImporte

            i = i + 1
            vDataGrid.Columns(i).Caption = "Cantidad"
            vDataGrid.Columns(i).Width = 950
            vDataGrid.Columns(i).Alignment = dbgRight
            vDataGrid.Columns(i).NumberFormat = FormatoImporte
            
            i = i + 1
            vDataGrid.Columns(i).Caption = "Precio"
            vDataGrid.Columns(i).Width = 1000
            vDataGrid.Columns(i).Alignment = dbgRight
            vDataGrid.Columns(i).NumberFormat = FormatoPrecio
                   
            i = i + 1
            vDataGrid.Columns(i).Caption = "OP"
            vDataGrid.Columns(i).Width = 340
            vDataGrid.Columns(i).Alignment = dbgCenter
                
            i = i + 1
            vDataGrid.Columns(i).Caption = "Dto.1"
            vDataGrid.Columns(i).Width = 540
            vDataGrid.Columns(i).Alignment = dbgRight
            vDataGrid.Columns(i).NumberFormat = FormatoDescuento
                
            i = i + 1
            vDataGrid.Columns(i).Caption = "Dto.2"
            vDataGrid.Columns(i).Width = 550
            vDataGrid.Columns(i).Alignment = dbgRight
            vDataGrid.Columns(i).NumberFormat = FormatoDescuento
                
            i = i + 1
            vDataGrid.Columns(i).Caption = "Importe"
            vDataGrid.Columns(i).Width = 1290
            vDataGrid.Columns(i).Alignment = dbgRight
            vDataGrid.Columns(i).NumberFormat = FormatoImporte
            
            i = i + 1
            vDataGrid.Columns(i).visible = False
            vDataGrid.Columns(i + 1).visible = False
            
            'Fijo
            i = i + 2
            vDataGrid.Columns(i).Caption = "Fijo"
            vDataGrid.Columns(i).Width = 600
            vDataGrid.Columns(i).Alignment = dbgCenter
            
            
            
    For i = 0 To vDataGrid.Columns.Count - 1
        vDataGrid.Columns(i).Locked = True
        vDataGrid.Columns(i).AllowSizing = False
    Next i
    vDataGrid.HoldFields
    Exit Sub
    
ECargaGrid:
    MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub


'***************************************
' Geslab---> Aripres.      Febrero 2014
'       Esta en MYSQL.
'
'
'---------------------- GESLAB
Private Sub PreparaConexionGeslab()
Dim F As Date

    'Cuando cerramos el parte guardaremos el nombre del trabajador
    'meintras tanto linkara sobre una temporal advtmptrabajadores
    
    'leemos de parametros
    '`advparametros` (`pathgeslab`,`UltimaActualizacion`
    Set miRsAux = New ADODB.Recordset
    D = "SELECT * from advparametros"
    miRsAux.Open D, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    'NO PUIEDE SER EOF
    
    'CadenaDevuelta2 = miRsAux!UltimaActualizacion
    'F = CDate(CadenaDevuelta2)
    BDAripres = miRsAux!pathgeslab
    
    
    
    
    
    
    
        
    
        
End Sub



Private Sub txtAuxT_GotFocus(Index As Integer)
Dim cadkey As Integer

    cadkey = ObtenerCadKey(kCampo, Index)
    kCampo = Index
'    ConseguirFoco txtAux2(Index), Modo, cadkey
    ConseguirFocoLin txtAuxT(Index), cadkey

End Sub

Private Sub txtAuxT_KeyPress(Index As Integer, KeyAscii As Integer)
     KEYpress KeyAscii
End Sub

Private Sub txtAuxT_LostFocus(Index As Integer)
    txtAuxT(Index).Text = Trim(txtAuxT(Index).Text)
    
    If Index = 0 Then
        D = ""
        If txtAuxT(Index).Text <> "" Then
            If PonerFormatoEntero(txtAuxT(Index)) Then
                Set miRsAux = New ADODB.Recordset
                D = DevuelveDesdeBD(conAri, "nomtrabajador", BDAripres & ".trabajadores", "idtrabajador", txtAuxT(Index).Text)
                If D = "" Then
                    MsgBox "No existe el trabajador: " & txtAuxT(0).Text, vbExclamation
                    txtAuxT(0).Text = ""
                    PonerFoco txtAuxT(0)
                Else
                    PonerFoco txtAuxT(2)
                End If
            End If
        End If
        txtAuxT(1).Text = D
        
    Else
        If Index = 2 Then
            If Not PonerFormatoDecimal(txtAuxT(Index), 1) Then
                txtAuxT(Index).Text = ""
            Else
                'PonerFocoBtn Me.cmdAceptar
            End If
        End If
    End If
    
 
End Sub



Private Sub pasarAAlbaranes()
    If Modo <> 2 Then Exit Sub
    
    If Data1.Recordset.EOF Then Exit Sub


    'Por si acaso
    D = DevuelveDesdeBD(conAri, "cerrado", "advpartes", "numparte", CStr(Data1.Recordset!numparte))
    If D = "1" Then
        MsgBox "El parte ya ha sido cerrado", vbExclamation
        Exit Sub
    End If
    
    

    'Comprobarciones basicas
    D = ""
    If Me.Check1(1).Value = 1 Then D = D & "-Ya facturado"

    'Si no tiene lineas de arti o de trabajadores..
    If data3.Recordset.EOF And data4.Recordset.EOF Then D = D & "-Sin lineas a facturar" & vbCrLf
    
    'Debe tener trabajador de parte
    If Text1(8).Text = "" Then D = D & "- Falta trabajador  introduccion parte"
    
       
    If D <> "" Then
        MsgBox "Error: " & vbCrLf & D, vbExclamation
        Exit Sub
    End If
    
    
    'Comrpobaremos que todas los articulos son correctos
    If Not ComprobarParteExternoInterno(Check1(2).Value = 0) Then Exit Sub
    
    
    
    Set cCli = New CCliente
    D = CStr(Data1.Recordset!codClien)
    If cCli.LeerDatos(D) Then
    
        If cCli.ClienteBloqueado(2, False) Then D = ""
           
    End If
    
    

    
    
    Set cCli = Nothing
    If D = "" Then Exit Sub
    
    
    
    
    If Not data4.Recordset.EOF Then
        'Tiene horas trabajadores
        D = DevuelveDesdeBD(conAri, "codartichoras", "advparametros", "1", "1")
        If D = "" Then
            MsgBox "Falta configurar articulo horas. Llame soporte tecnico", vbExclamation
            Exit Sub
        End If
        D = ""
    Else
        D = "Ningún trabajador de campo asignado al parte.  Continuar?"
        If MsgBox(D, vbQuestion + vbYesNo) = vbNo Then Exit Sub
    End If
    
    
    
    
    
    'Preguntamos
    If D = "" Then
        D = "Seguro que desea cerrar el parte de trabajo?"
        If MsgBox(D, vbQuestion + vbYesNo) = vbNo Then Exit Sub
    End If
        
    
    Screen.MousePointer = vbHourglass
    
    
    GenerarAlbaran
    
  
    Screen.MousePointer = vbDefault
    
    
End Sub


Private Sub UpdateaCambioLitros()
   
    
    Dim cantidad As Currency
     D = ""
    D = "select * from advparteslineas where numparte=" & DBSet(Text1(0).Text, "N")
    D = D & " And dosishab >=0 "
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open D, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    D = ""
    While Not miRsAux.EOF
        
            cantidad = Val(Text1(5).Text)
            'Dosis hab
            cantidad = (cantidad * miRsAux!dosishab) / 1000
            D = D & DBSet(cantidad, "N") & ","
            D = "UPDATE advparteslineas set cantidad=" & DBSet(cantidad, "N")
            
            'Mayo 2012
            'Error calculando lineas con dto
            BuscaChekc = CalcularImporte(CStr(cantidad), CStr(miRsAux!PrecioVe), CStr(miRsAux!dtoline1), CStr(miRsAux!dtoline2), vParamAplic.TipoDtos)
            
            'importel

            D = D & ",importel = " & DBSet(BuscaChekc, "N")
            D = D & " WHERE  numparte=" & DBSet(Text1(0).Text, "N") & " AND numlinea = " & miRsAux!numlinea
            conn.Execute D
            miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
        
    If D <> "" Then CargaGrid True, 2
        
    
End Sub


Private Sub InsertaReferenciaTratamientos()
Dim Cad As String
Dim CPrFact As CPreciosFact
Dim Au As String
Dim cantidad As Currency
    Cad = "select * from advtrata_lineas where codtrata=" & DBSet(Text1(3).Text, "T") & " order by numlinea"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    Cad = ""
    While Not miRsAux.EOF
        NumRegElim = NumRegElim + 1
        'd = "insert into `advparteslineas` (`numparte`,`numlinea`,`codalmac`,`codartic`,"
        'd = d & "`dosishab`,`cantidad`,`ampliaci`,`dtoline1`,`dtoline2`,`origpre`,`preciove`,`importel`) values ("
        
        Cad = Cad & ", (" & Text1(0).Text & "," & NumRegElim & ",1,"
        Cad = Cad & DBSet(miRsAux!codArtic, "T") & ","
        'codartic,dosishab,cantidad,
        cantidad = ImporteFormateado(Text1(5).Text)
        If IsNull(miRsAux!dosishab) Then
            'Es cantidad
            Cad = Cad & "NULL,"
            Cad = Cad & DBSet(miRsAux!cantidad, "N") & ","
            cantidad = miRsAux!cantidad
        Else
            'Dosis hab
            cantidad = (cantidad * miRsAux!dosishab) / 1000
            
            Cad = Cad & DBSet(miRsAux!dosishab, "N") & ","
            Cad = Cad & DBSet(cantidad, "N") & ","
            
        End If
        
        
        'precio
        
            Set CPrFact = New CPreciosFact
    
            CPrFact.CodigoArtic = miRsAux!codArtic
            CPrFact.CodigoClien = Text1(6).Text
            CPrFact.FijarTarifaActividad
            
            
            D = CPrFact.ObtenerPrecio(False, Text1(1).Text, Au, "")
            
            '`ampliaci`,`dtoline1`,`dtoline2`,`origpre`,`preciove`,`importel`
            Cad = Cad & "NULL," & DBSet(CPrFact.Descuento1, "N") & "," & DBSet(CPrFact.Descuento2, "N")
            Cad = Cad & ",'" & Au & "'," & DBSet(D, "N")
            

            


            Au = CalcularImporte(CStr(cantidad), D, CPrFact.Descuento1, CPrFact.Descuento2, vParamAplic.TipoDtos)
            Cad = Cad & "," & DBSet(Au, "N") & ")"
        
            Set CPrFact = Nothing
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    If Cad <> "" Then
        Cad = Mid(Cad, 2)
        D = "insert into `advparteslineas` (`numparte`,`numlinea`,`codalmac`,`codartic`,"
        D = D & "dosishab,`cantidad`,`ampliaci`,`dtoline1`,`dtoline2`,`origpre`,`preciove`,`importel`) values "
        D = D & Cad
        ejecutar D, False
    End If
    
End Sub


'---------------------------------------------------------------------------
'Pasar parte a albaran


Private Sub GenerarAlbaran()
Dim numPed As Long 'Nº Pedido
Dim NumAlb As String 'Nº Albaran
Dim Sql As String
Dim AlbaranGenerado As Boolean

    CadenaSQL = PonerTrabajadorConectado("") 'en el frm lo pong a ""
   
    If CadenaSQL = "" Then
        MsgBox "Error trabajador conectado", vbExclamation
        Exit Sub
    End If
    
    
    CodZona = -1
    Sql = ""

    
    Sql = DevuelveDesdeBDNew(conAri, "sclien", "codzonas", "codclien", Text1(6).Text, "N") 'zona por defecto
    If Sql <> "" Then
        'Vale, ya tengo la zona
        CodZona = Val(Sql)
        Sql = DevuelveDesdeBDNew(conAri, "szonas", "nomzonas", "codzonas", Sql, "N") 'zona por defecto
        If Sql = "" Then
            CodZona = -1
        Else
            Sql = CodZona & "|" & Sql & "|"
        End If
    End If
    

    If Sql = "" Then
        Sql = "||"
    Else
        CodZona = CInt(RecuperaValor(Sql, 1))
    End If
    
    Sql = Sql & Abs(0) & "|" 'En esta poscion maracaremos si SE VE el frame de zona
    
    'Variabale SQL
    'codzona|nomzona|visible famezona|
    
    
    
    Set frmList = New frmListadoPed
    'Datos para ofertar por defecto
    BuscaChekc = "tipofact"
    NumAlb = Trim(DevuelveDesdeBD(conAri, "nifclien", "sclien", "codclien", Text1(6).Text, "N", BuscaChekc))
    CadenaDesdeOtroForm = DevuelveDesdeBD(conAri, "cifempre", "sparam", "1", "1")
    If NumAlb = Trim(CadenaDesdeOtroForm) Then
        NumAlb = "1"
    Else
        NumAlb = "0"
    End If
    
    'Mayo2013.
    'Vemos tambien el tipofacturacion del cliente
    'trab,fecha,interna   y tipofact
    CadenaDesdeOtroForm = CadenaSQL & "|" & Text1(1).Text & "|" & NumAlb & "|" & BuscaChekc & "|"
    BuscaChekc = ""
    
    NumAlb = ""
    CadenaSQL = ""
    frmList.OpcionListado = 1043  'es como el 43
    frmList.codClien = Sql
    frmList.NumCod = Data1.Recordset!numparte
    frmList.Show vbModal
    
    Set frmList = Nothing
    Sql = ""
    
    If CadenaSQL = "" Then Exit Sub
    
    
    NumRegElim = Data1.Recordset.AbsolutePosition
    numPed = Data1.Recordset!numparte
    

    
    'CadenaSQL, se obtiene desde frmList
    lblIndicador.Caption = "Gen. albaran"
    lblIndicador.Refresh


    
    AlbaranGenerado = PasarPedidoAAlbaran(CadenaSQL, NumAlb)

    If AlbaranGenerado Then
   
            
        'ACTUALIZAR EL RIESGO
        '''ActualizaRiesgoCliente CLng(Text1(6).Text)
         
    
    
    
    
'        ComprobarNSeriesLineas (NumAlb)
        
        Espera 0.4
        
        If B Then
            ImprimirAlbaran NumAlb
        Else
        
            MsgBox "El parte de trabajo  Nº: " & Format(numPed, "0000000") & vbCrLf & vbCrLf & "ha generado el Albaran Nº: " & Format(NumAlb, "0000000"), vbInformation
            
        End If
        PosicionarData
        PonerCampos
        Me.Check1(1).Value = 1
    
        PonerModo 2
'        If Modo = 2 Then
'            If Not Data1.Recordset.EOF Then lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
'        End If
        Screen.MousePointer = vbDefault
        
    
        
    End If
End Sub




Private Function PasarPedidoAAlbaran(vSQL As String, NumAlb As String) As Boolean
'IN -> vSQL: cadena para el Select con los datos obtenidos en frmList
'OUT -> numAlb: Nº de Albaran de Venta que se ha insertado
Dim bol As Boolean
Dim MenError As String
Dim devuelve As String
Dim Sql As String
Dim RS As ADODB.Recordset
Dim Agrupados As Byte '3

    On Error GoTo EGenPedido

    bol = False

        
    'Aqui empieza transaccion
    conn.BeginTrans
    
    
    
    'Actualizamos
    ActualizarCantidadReales
        
    
    
    'Acutalizamos las cantidades en funcion de los litros reales
    'Metera una linea con el articulo de parametros , el importe HORAS4
    MenError = "Geslab." & vbCrLf
    CalculoImportesHora
    
     
    
    'Insertar en tablas de Albaranes el Pedido (scaalb, slialb)
    Sql = RecuperaValor(CadenaDesdeOtroForm, 4)  'fracion colectiva esta en el 4
    'Reemplzazaremos la cadena ,@tipofact@,  por ,0 tipofact, o por ,1 tipofact,
    Sql = ", " & Sql & " tipofact,"
    CadenaSQL = Replace(CadenaSQL, ",@tipofact@,", Sql)
    
    
    CadenaDesdeOtroForm = RecuperaValor(CadenaDesdeOtroForm, 2)  'albaran interno esta en el 2
    ALbInterno = CadenaDesdeOtroForm = "1"
    bol = InsertarAlbaran(vSQL, MenError, NumAlb, ALbInterno)
    
    'Actualizar Stock en salmac, e introducir movimiento en smoval
    If bol Then
        MenError = "Error al insertar movimientos de stock."
        bol = InsertarMovStock2(NumAlb, ALbInterno)
    End If
    
   
    
    If bol Then
        'Las horas

    
    End If
    
    
    If bol Then
        'Actualizar el campo fechamov (ult. movimiento) de la tabla de clientes (sclien)
        'si la fecha es posterior a la que tiene
        Set cCli = New CCliente
        If cCli.LeerDatos(Text1(6).Text) Then
            bol = cCli.ActualizaUltFecMovim(FechaAlb)
        Else
            bol = False
        End If
        Set cCli = Nothing
        

    End If
    
    
    
    If bol Then
        Sql = "UPDATE advpartes set cerrado=1 where numparte=" & Data1.Recordset!numparte
        conn.Execute Sql
        
        'Borramos la linea del trabjo
        
        Sql = "DELETE FROM advparteslineas where numparte=" & Data1.Recordset!numparte & " AND codartic = " & DBSet(ArtiHORAS, "T")
        conn.Execute Sql
    End If
    
EGenPedido:
    If Err.Number <> 0 Or Not bol Then
        MenError = "Pasando parte a Albaran." & vbCrLf & "----------------------------" & vbCrLf & MenError
        MuestraError Err.Number, MenError, Err.Description
        bol = False
    End If
    If bol Then
        conn.CommitTrans
    Else
        conn.RollbackTrans
      
        
    End If
    
    PasarPedidoAAlbaran = bol
End Function




Private Function InsertarAlbaran(vSQL As String, MenError As String, NumAlb As String, AlbaranInterno As Boolean) As Boolean
'Devuelve el mensaje de error si se produce
Dim bol As Boolean, Existe As Boolean
Dim devuelve As String
Dim vTipoMov As CTiposMov 'Clase Tipo Movimiento
Dim codtipom As String

    On Error GoTo EInsertarAlbaran
    
    bol = False
    InsertarAlbaran = bol
    
    'Obtener el Contador de PEDIDO
    If AlbaranInterno Then
        codtipom = "ALI"
    Else
        codtipom = "ALS"
    End If
    
    Set vTipoMov = New CTiposMov
    If vTipoMov.Leer(codtipom) Then
        'Comprobar si mientras tanto se incremento el contador de Pedidos
        'para ello vemos si existe una oferta con ese contador y si existe la incrementamos
        Do
            NumAlb = vTipoMov.ConseguirContador(codtipom)
            devuelve = DevuelveDesdeBDNew(conAri, "scaalb", "numalbar", "codtipom", codtipom, "T", , "numalbar", NumAlb, "N")
            If devuelve <> "" Then
                'Ya existe el contador incrementarlo
                Existe = True
                vTipoMov.IncrementarContador (codtipom)
                NumAlb = vTipoMov.ConseguirContador(codtipom)
            Else
                Existe = False
            End If
        Loop Until Not Existe
            
    Else 'No existe el tipo de Movimiento
        Set vTipoMov = Nothing
        Exit Function
    End If
    
    
    'Nuevo OCTUBRE 2010
    vSQL = vSQL & ",NULL,0 as tipAlbaran ,"  '1-con trasporte  0-sin trasporte
    
    'codzona
    vSQL = vSQL & "NULL"
    
    'Campo nuevo observacrm  Febrero 2011
    vSQL = vSQL & ",NULL "
    
    'Acabar la sql con el contador seleccionado
    devuelve = vSQL
    vSQL = "INSERT INTO scaalb (codtipom,numalbar,fechaalb,factursn,codclien,nomclien,domclien,codpobla,pobclien,proclien,nifclien,telclien,"
    vSQL = vSQL & "coddirec,nomdirec,referenc,codtraba,codtrab1,codtrab2,codagent,codforpa,codenvio,dtoppago,dtognral,tipofact,"
    vSQL = vSQL & "observa01,observa02,observa03,observa04,observa05,numofert,fecofert,numpedcl,fecpedcl,fecentre,sementre,coddiren,tipAlbaran,codzonas,observacrm) "
    vSQL = vSQL & "SELECT '" & codtipom & "' as codtipom, " & NumAlb & " as numalbar, " & devuelve
    vSQL = vSQL & " FROM " & NombreTabla & ",sclien WHERE advpartes.codclien=sclien.codclien AND numparte=" & Text1(0).Text

    'Insertar Cabecera
    MenError = "Error al insertar en la tabla Cabecera de Albaranes (scaalb )."
    conn.Execute vSQL, , adCmdText
    
    'Insertar Lineas Pedido
    MenError = "Error al insertar en la tabla Lineas de Albaran (slialb)."
    If Not InsertarLineasAlbaran(codtipom, NumAlb) Then Exit Function
    
    
    
    'Hay varios articulos que se agruparan en una sola linea del albaran.
    'Como...
    
    'ProcesoAgrupacionLineasAlbaranAlzira
    
    
    
    
    
    MenError = "Error al actualizar el contador del ALbaran."

    vTipoMov.IncrementarContador (codtipom)
    Set vTipoMov = Nothing
    bol = True
    
EInsertarAlbaran:
        If Err.Number <> 0 Then bol = False
        InsertarAlbaran = bol
End Function


Private Function InsertarLineasAlbaran(TipoM As String, NumAlb As String) As Boolean
'Inserta en la tabla de lineas de albaran (slialb)
Dim Sql As String
Dim RS As ADODB.Recordset
Dim linea As Integer
Dim NumBulto As String
Dim CantiT As Currency
Dim RT As ADODB.Recordset
Dim J As Integer
Dim ErroresLotes As String
Dim canAux As Currency
    
    On Error GoTo erroresInstLine

    'ENERO 2008.   codprove en slialb para traza de proveedores en lineas

   
     
         
    lblIndicador.Caption = "Lineas sin grup"
    lblIndicador.Refresh
     
     Sql = ""
     Sql = "SELECT '" & TipoM & "', " & NumAlb & " as numalbar, numlinea, codalmac,"
     Sql = Sql & "advparteslineas.codartic, nomartic, ampliaci, "
     Sql = Sql & "cantidad, 0,advparteslineas.preciove, dtoline1, dtoline2, importel, origpre"
     'traza
     Sql = Sql & ",codprove,NULL,NULL"
     Sql = Sql & " FROM advparteslineas,sartic WHERE advparteslineas.codartic = sartic.codartic"
     Sql = Sql & " AND numparte=" & Text1(0).Text
     'If ArticulosAgrupados <> "" Then SQL = SQL & " AND not advparteslineas.codartic IN (" & ArticulosAgrupados & ")"
     'Los de agrupacion
     Set miRsAux = New ADODB.Recordset
     
     
     Sql = "INSERT INTO slialb (codtipom,numalbar,numlinea,codalmac,codartic,nomartic,ampliaci,cantidad,numbultos,precioar,dtoline1,dtoline2,importel,origpre,codproveX,numlote,codccost) " & Sql
     conn.Execute Sql


    'Insetamos en la lineas de los campos

     Sql = "SELECT '" & TipoM & "', " & NumAlb & " as numalbar, @rownum:=@rownum+1 AS rownum, codcampo"
     Sql = Sql & " FROM advpartes_campos,(SELECT @rownum:=0) r WHERE numparte=" & Text1(0).Text
     Sql = "INSERT INTO slialbcampos (codtipom,numalbar,numlinea,codcampo) " & Sql
     conn.Execute Sql






     Espera 0.5
    
     Sql = DevuelveDesdeBD(conAri, "max(numlinea)", "advparteslineas", "numparte", Text1(0).Text, "N")
     If Sql = "" Then Sql = "0"
     linea = Val(Sql)
        

              
              
              
        'UPdateamos la columna de numlotes del albaran si lleva fitosanitarios
        lblIndicador.Caption = "Lotes"
        lblIndicador.Refresh
        Espera 0.5
        
        ErroresLotes = ""
        Set RS = New ADODB.Recordset
        Set RT = New ADODB.Recordset
        
        'Llevamos control de lotes
        Sql = "Select slialb.codartic,cantidad,numlinea from slialb,sartic where slialb.codartic=sartic.codartic AND slialb.codtipom='" & TipoM & "' AND numalbar=" & NumAlb
        Sql = Sql & " AND numserie<>'' AND codcateg in (select codcateg FROM scateg where ctrlotes=1) "
        RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
       
        While Not RS.EOF
                    
                    
                    
            Sql = "select numlotes,fecentra,Codartic,canentra - vendida"
            Sql = Sql & "  disponible from slotes where "
            Sql = Sql & " codartic=" & DBSet(RS!codArtic, "T") & " and canentra - vendida  >0 order by fecentra "
                  
        
            RT.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            J = 0
            CantiT = RS!cantidad
            
            
            While Not RT.EOF
                J = J + 1
                If RT!disponible >= CantiT Then
                    canAux = CantiT
                    CantiT = 0
                Else
                    canAux = RT!disponible
                    CantiT = CantiT - canAux
                End If
                
                
                Sql = "INSERT INTO slialblotes(codtipom,numalbar,numlinea,sublinea,numlote,cantidad,fecentra,codartic) VALUES ("
                Sql = Sql & "'" & TipoM & "'," & NumAlb & "," & RS!numlinea & "," & J & "," & DBSet(RT!numlotes, "T") & ","
                Sql = Sql & DBSet(canAux, "N") & "," & DBSet(RT!fecentra, "F") & "," & DBSet(RT!codArtic, "T") & ")"
                conn.Execute Sql
                
                'Aticulo nuevo. La primera entrada es la que vale
                Sql = "UPDATE slialb SET numlote = '*'" & " WHERE codtipom= '" & TipoM & "' AND numalbar= " & NumAlb
                Sql = Sql & " AND numlinea = " & RS!numlinea
                conn.Execute Sql
            
                'El lote lo incremento (decremento)
                Sql = "UPDATE slotes set vendida = vendida + " & DBSet(canAux, "N")
                Sql = Sql & " WHERE numlotes =" & DBSet(RT!numlotes, "T") & " AND codartic= " & DBSet(RT!codArtic, "T")
                Sql = Sql & " AND fecentra= " & DBSet(RT!fecentra, "F")
                conn.Execute Sql
                
            
                If CantiT = 0 Then
                    While Not RT.EOF
                        RT.MoveNext
                    Wend
                Else
                    RT.MoveNext
                End If
            Wend
            RT.Close
            
            
            If CantiT > 0 Then
                Sql = "Articulo: " & RS!codArtic & "   Cantidad: " & RS!cantidad & "     Pendiente:" & CantiT
                ErroresLotes = ErroresLotes & Sql & vbCrLf
            End If
            
            RS.MoveNext
        Wend
        RS.Close
        Set RS = Nothing
        Set RT = Nothing
    
        If ErroresLotes <> "" Then
            Sql = "Error obteniendo lotes articulos: " & vbCrLf & String(20, "*") & vbCrLf & vbCrLf & ErroresLotes
            Sql = Sql & vbCrLf & vbCrLf & " El programa continuará"
            MsgBox Sql, vbExclamation
        End If
    
erroresInstLine:
    If Err.Number <> 0 Then
         'Hay error , almacenamos y salimos
        InsertarLineasAlbaran = False
    Else
        InsertarLineasAlbaran = True
    End If
End Function




Private Function InsertarMovStock2(NumAlb As String, Interno As Boolean) As Boolean
Dim vCStock As CStock
Dim B As Boolean
Dim RS As ADODB.Recordset
Dim Sql As String

    On Error Resume Next

    InsertarMovStock2 = False
    
    Set vCStock = New CStock
    B = True
    
    Sql = "select * from advparteslineas WHERE numparte =" & Data1.Recordset!numparte
    
    'If ArticulosAgrupados <> "" Then SQL = SQL & " AND not advparteslineas.codartic IN (" & ArticulosAgrupados & ")"
    
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    vCStock.FechaMov = FechaAlb
    vCStock.Trabajador = Val(PonerTrabajadorConectado(""))
    'para cada linea del Pedido Insertar en smoval y Actualizar Stock en salmac
    While (Not RS.EOF) And B
        'si hay control de stock
'        SQL = DevuelveDesdeBDNew(conAri, "sartic", "ctrstock", "codartic", RS!codartic, "T")
'        If Val(SQL) = 1 Then
            If Not InicializarCStockAlbar(vCStock, "S", CStr(RS!numlinea), RS, Interno) Then Exit Function

            'vCStock.Documento = numAlb
            vCStock.Documento = Format(NumAlb, "0000000")
            If vCStock.cantidad <> 0 Then
                'en actualizar stock comprobamos si el articulo tiene control de stock
                    B = vCStock.ActualizarStock(False, False)
            End If
'        End If
        RS.MoveNext
    Wend
    Set vCStock = Nothing
    RS.Close
    Set RS = Nothing
    
    InsertarMovStock2 = B
    
End Function


'Private Function InicializarCStockAlbar2(ByRef vCStock As CStock, TipoM As String, Optional numlinea As String, Optional ByRef RS As ADODB.Recordset) As Boolean
Private Function InicializarCStockAlbar(ByRef vCStock As CStock, TipoM As String, numlinea As String, ByRef RS As ADODB.Recordset, Interno As Boolean) As Boolean
'Para comprobar stock al pasar de Pedido a Albaran de Venta
On Error Resume Next
    
    vCStock.tipoMov = TipoM
    If Interno Then
        vCStock.DetaMov = "ALI"
    Else
        vCStock.DetaMov = "ALS"
    End If
    vCStock.Trabajador = Text1(6).Text
    vCStock.Documento = Text1(0).Text
    vCStock.codArtic = RS!codArtic
    vCStock.codAlmac = CInt(RS!codAlmac)
    
    
    vCStock.cantidad = CSng(RS!cantidad)
    If RS.Fields.Count > 3 Then 'Si no se selecciona el campo importe de la tabla es que solo vamos a comprobar stock y no se necesita
        vCStock.Importe = CCur(RS!ImporteL)
    End If

    vCStock.LineaDocu = CInt(ComprobarCero(numlinea))
    
    If Err.Number <> 0 Then
        MsgBox "No se han podido inicializar la clase para actualizar Stock", vbExclamation
        InicializarCStockAlbar = False
    Else
        InicializarCStockAlbar = True
    End If
End Function





'Modificacion Marzo 2012.
'Las horas las metere en el proceso antes de pasarlas al albaran.
'Las metere en advpartes_linea, y luego, el pase a albaran ya hara lo que tenga que hacer
Private Sub CalculoImportesHora()

'Dim total As Currency
Dim Ampliacion As String
Dim Aux As String
Dim IdCate As Integer
Dim Horas  As Currency
Dim Importe2  As Currency



    'no pongo errores, si da errores que salte alli

    If data4.Recordset.EOF Then Exit Sub
    
    
    Set miRsAux = New ADODB.Recordset
    
    
    
    D = "Select codtraba from advpartes_trabajador WHERE numparte=" & Data1.Recordset!numparte
    miRsAux.Open D, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    D = ""
    While Not miRsAux.EOF
        D = D & ", " & miRsAux!CodTraba
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    D = Mid(D, 2)
    
    
    'Modificacion, se van sumando las horas y punto, no hacemos nada mas

    
    Aux = "SELECT Trabajadores.IdTrabajador, NomTrabajador,Trabajadores.PorcAntiguedad, Trabajadores.porcIRPF, Trabajadores.PorcSS, Categorias.Importe1, Categorias.Importe2,  Trabajadores.idCategoria"
    Aux = Aux & " FROM " & BDAripres & ".Categorias INNER JOIN " & BDAripres & ".Trabajadores ON Categorias.IdCategoria = Trabajadores.idCategoria WHERE "
    Aux = Aux & " IdTrabajador IN (" & D & ") ORDER BY Trabajadores.idcategoria,trabajadores.IdTrabajador"
    miRsAux.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    IdCate = -1
    Ampliacion = "Trab: "
    Horas = 0
    While Not miRsAux.EOF
        
        If miRsAux!idCategoria <> IdCate Then
            'If IdCate > 0 Then
            '    'Cambio categoria, insertamos linea albaran
            '    InsertaLineaAlbaranHoras Importe2, Horas, Ampliacion
            '
            '
            'End If
            'total = 0
            'Horas = 0
            
            'IdCate = miRsAux!idCategoria
            'Ampliacion = "Cat: " & IdCate & " - "
        End If
        'Ampliacion = "Cat: " & IdCate & " - "  'para el codtraba
        
        data4.Recordset.Find "codtraba = " & miRsAux!idtrabajador, , adSearchForward, 1
        Ampliacion = Ampliacion & miRsAux!idtrabajador & " ,"
       ' Importe2 = 0
        If Not data4.Recordset.EOF Then
            
    '        If Not IsNull(miRsAux!porcSS) Then Importe2 = Importe2 + miRsAux!porcSS
    '        If Not IsNull(miRsAux!porcirpf) Then Importe2 = Importe2 + miRsAux!porcirpf
    '        Importe2 = Importe2 / 100
    '        Importe2 = miRsAux!Importe1 - Importe2
            Horas = Horas + data4.Recordset!Horas
    '        Importe2 = (Importe2 * data4.Recordset!Horas)
    '        Importe2 = Round(Importe2, 2)
            'Ampliacion = Ampliacion & miRsAux!idtrabajador & " - "
                
            InsertaEnMarcajes 'en geslab
                
        End If
    '    total = total + Importe2
        
    
        D = "UPDATE advpartes_trabajador set nomtraba = " & DBSet(miRsAux!NomTrabajador, "T")
        D = D & " WHERE numparte = " & Data1.Recordset!numparte & " AND codtraba = " & miRsAux!idtrabajador
        conn.Execute D
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    'Abril 2012
    'Las horas YA NO entran desde aqui
    'd = DevuelveDesdeBD(conAri, "preciove", "sartic", "codartic", ArtiHORAS, "T")
    'Importe2 = CCur(d)
    'InsertaLineaAlbaranHoras Importe2, Horas, Ampliacion
    
    Set miRsAux = Nothing
    
    'ver como inserta horas
    
End Sub

Private Sub InsertaLineaAlbaranHoras(Importe As Currency, Horas As Currency, Ampliaci As String)
Dim Lin  As Integer
Dim Cad As String

    
        Ampliaci = Mid(Ampliaci, 1, Len(Ampliaci) - 1)
    
        Cad = DevuelveDesdeBD(conAri, "max(numlinea)", "advparteslineas", "numparte", CStr(Data1.Recordset!numparte))
        If Cad = "" Then Cad = "0"
        Lin = Val(Cad) + 1
        Cad = "insert into `advparteslineas` (`numparte`,`numlinea`,`codalmac`,`codartic`,`dosishab`,`cantidad`,`preciove`,`importel`,`ampliaci`,`dtoline1`,`dtoline2`,`origpre`) values ("
        Cad = Cad & Data1.Recordset!numparte & "," & Lin & ",1,"
        'codartic,dosis,cantidad
        Cad = Cad & DBSet(ArtiHORAS, "T") & ",NULL," & DBSet(Horas, "N", "S") & ","
        
        '`preciove`,`importel`,`ampliaci`,

        Horas = Round2(Importe * Horas, 2)
        Cad = Cad & DBSet(Importe, "N") & ","
        Cad = Cad & DBSet(Horas, "N") & "," & DBSet(Mid(Ampliaci, 1, 60), "T")
        '`dtoline1`,`dtoline2`,`origpre`
        Cad = Cad & ",0,0,'A')"
    
    
     conn.Execute Cad

   


End Sub

Private Function InsertaEnMarcajes() As Boolean
Dim RT As ADODB.Recordset
Dim Marcaje As Long
Dim Aux As String
Dim Nuevo As Long
Dim Hora1 As Date
Dim vHoras As Integer
Dim Minutos As Currency

    'Entrada idTrabajador Fecha Fecha HorasTrabajadas HorasIncid
    Set RT = New ADODB.Recordset
    Aux = "Select * from " & BDAripres & ".marcajes where idtrabajador =" & miRsAux!idtrabajador
    Aux = Aux & " AND fecha = '" & Format(FechaAlb, FormatoFecha) & "'"
    RT.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Marcaje = -1
    If RT.EOF Then
        Nuevo = -1
        RT.Close
        Aux = "Select max(entrada) from " & BDAripres & ".marcajes"
        RT.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Aux = "0"
        If Not RT.EOF Then Aux = DBLet(RT.Fields(0), "N")
        Marcaje = CStr(Val(Aux) + 1)
        Aux = "INSERT INTO " & BDAripres & ".marcajes(Entrada,idTrabajador,Fecha,correcto,HorasTrabajadas,HorasIncid) VALUES (" & Marcaje & ","
        Aux = Aux & miRsAux!idtrabajador & "," & DBSet(FechaAlb, "F") & ",1,"
        Aux = Aux & TransformaComasPuntos(data4.Recordset!Horas) & ",0)"
        
    Else
        Marcaje = RT!entrada
        Aux = "UPDATE " & BDAripres & ".marcajes set HorasTrabajadas=HorasTrabajadas + " & TransformaComasPuntos(data4.Recordset!Horas)
        Aux = Aux & " WHERE entrada = " & Marcaje
        Nuevo = 0
    End If
    RT.Close
    conn.Execute Aux
    
    
    'Las lineas.....

    Hora1 = "04:55:00" 'empezamos a las cinco de la mañana
    If Nuevo = 0 Then
        'NO Es nuevo, cogeremos la ultima hora trabajada del colega
        Aux = "Select max(Hora) as Datos1 from " & BDAripres & ".EntradaMarcajes WHERE idtrabajador= " & miRsAux!idtrabajador
        Aux = Aux & " AND fecha = " & DBSet(FechaAlb, "F")
        RT.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        If Not RT.EOF Then
            If Not IsNull(RT.Fields(0)) Then Hora1 = RT.Fields(0)
        End If
        RT.Close
    End If
    Minutos = 5
    Hora1 = DateAdd("n", Minutos, Hora1)
        
        
    Aux = "Select max(Secuencia) from " & BDAripres & ".EntradaMarcajes  "
    RT.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Nuevo = 0
    If Not RT.EOF Then Nuevo = DBLet(RT.Fields(0), "N")
    RT.Close
    Nuevo = Nuevo + 1
    
    'Fichada 1
    Aux = "INSERT INTO " & BDAripres & ".EntradaMarcajes(Secuencia ,idTrabajador ,idMarcaje ,Fecha ,Hora ,idInci ,HoraReal) VALUES ("
    Aux = Aux & Nuevo & "," & miRsAux!idtrabajador & "," & Marcaje & "," & DBSet(FechaAlb, "F")
    Aux = Aux & "," & DBSet(Hora1, "H") & ",0," & DBSet(Hora1, "H") & ")"
    conn.Execute Aux
    
    'Fichada 1
    Nuevo = Nuevo + 1
    vHoras = Int(data4.Recordset!Horas)
    Hora1 = DateAdd("h", vHoras, Hora1)
    Minutos = data4.Recordset!Horas - Int(data4.Recordset!Horas)
    Minutos = Minutos * 100
    Minutos = Round2(Minutos * 0.6, 0)
    Hora1 = DateAdd("n", Minutos, Hora1)
    
    Aux = "INSERT INTO " & BDAripres & ".EntradaMarcajes(Secuencia ,idTrabajador ,idMarcaje ,Fecha ,Hora ,idInci ,HoraReal)  VALUES ("
    Aux = Aux & Nuevo & "," & miRsAux!idtrabajador & "," & Marcaje & "," & DBSet(FechaAlb, "F")
    Aux = Aux & "," & DBSet(Hora1, "H") & ",0," & DBSet(Hora1, "H") & ")"
    conn.Execute Aux
    
End Function


Private Sub ActualizarCantidadReales()
Dim cantidad As Currency
Dim CantidadNODosis As Currency

Dim Aux As String
    
    
    Aux = RecuperaValor(CadenaDesdeOtroForm, 3)
    If Trim(Aux) = "" Then Aux = "1"
    CantidadNODosis = ImporteFormateado(Aux)
    
    
    Aux = RecuperaValor(CadenaDesdeOtroForm, 1)
    cantidad = Val(ImporteFormateado(Aux))
    
    D = "UPDATE advpartes set litrosrea = " & DBSet(cantidad, "N")
    D = D & " where numparte=" & DBSet(Text1(0).Text, "N")
    conn.Execute D
     
    D = "select * from advparteslineas where numparte=" & DBSet(Text1(0).Text, "N")
    D = D & " And dosishab >=0 "
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open D, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    D = ""
    While Not miRsAux.EOF
            Aux = RecuperaValor(CadenaDesdeOtroForm, 1)
            cantidad = Val(ImporteFormateado(Aux))
            'Dosis hab
            cantidad = (cantidad * miRsAux!dosishab) / 1000
            D = D & DBSet(cantidad, "N") & ","
            D = "UPDATE advparteslineas set cantidad=" & DBSet(cantidad, "N")
            
            
            Aux = CalcularImporte(CStr(cantidad), CStr(miRsAux!PrecioVe), CStr(miRsAux!dtoline1), CStr(miRsAux!dtoline2), vParamAplic.TipoDtos)
            
            D = D & ", importel= " & DBSet(Aux, "N")
            D = D & " WHERE  numparte=" & DBSet(Text1(0).Text, "N") & " AND numlinea = " & miRsAux!numlinea
            conn.Execute D
            miRsAux.MoveNext
            
            '------------------


            
    Wend
    miRsAux.Close
    
    'Cantidad se multiplica por cantidad real
    If CantidadNODosis > 0 Then
        D = "select * from advparteslineas where numparte=" & DBSet(Text1(0).Text, "N")
        
        'MAYO 2012.  Existen cantidades fijas
        D = D & " AND fijo = 0" 'Es decir, no cogere las cantidades que tengan un 1 (fijas) en campo BD .fijo
        D = D & " And (dosishab = 0 or dosishab is null)"
        Set miRsAux = New ADODB.Recordset
        miRsAux.Open D, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        NumRegElim = 0
        D = ""
        While Not miRsAux.EOF
                Aux = RecuperaValor(CadenaDesdeOtroForm, 3)
                CantidadNODosis = ImporteFormateado(Aux)
                'Dosis hab
                cantidad = (CantidadNODosis * miRsAux!cantidad)
                
                D = "UPDATE advparteslineas set cantidad=" & DBSet(cantidad, "N")
                
                
                Aux = CalcularImporte(CStr(cantidad), CStr(miRsAux!PrecioVe), CStr(miRsAux!dtoline1), CStr(miRsAux!dtoline2), vParamAplic.TipoDtos)
                
                D = D & ", importel= " & DBSet(Aux, "N")
                D = D & " WHERE  numparte=" & DBSet(Text1(0).Text, "N") & " AND numlinea = " & miRsAux!numlinea
                conn.Execute D
                miRsAux.MoveNext
                
                '------------------
    
    
                
        Wend
        miRsAux.Close
    End If
    
    Set miRsAux = Nothing
End Sub




Private Sub ImprimirAlbaran(Numalbar As String)
Dim cadFormula As String
Dim cadSelect As String 'select para insertar en tabla temporal
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"

    cadFormula = ""
    cadSelect = ""
    CadenaSQL = ""
   
   
    '===================================================
    '============ PARAMETROS ===========================
    If ALbInterno Then
        indRPT = 56
    Else
        indRPT = 39
    End If
    
    If Not PonerParamRPT2(indRPT, CadenaSQL, CByte(NumRegElim), cadFormula, pImprimeDirecto, pPdfRpt, pRptvMultiInforme) Then
        Exit Sub
    End If

    
    'Nombre fichero .rpt a Imprimir
    frmImprimir.NombreRPT = cadFormula
    frmImprimir.NombrePDF = cadFormula
    frmImprimir.SeleccionaRPTCodigo = pRptvMultiInforme
    
    cadFormula = ""
                
    
    'Añadir el codigo de usuario como parametro para link con tabla Temporal (tmptiposiva) en el Report
    'tabla temporal para el calculo del bruto total para cada tipo de IVA
    CadenaSQL = CadenaSQL & "pCodUsu=" & vUsu.Codigo & "|"
                
                
                

    'Cod Tipo Movimiento
    If ALbInterno Then
        D = "ALI"
    Else
        D = "ALS"
    End If
    CadenaDevuelta2 = "{scaalb.codtipom}='" & D & "'" 'Val(txtCodigo(0).Text)
    
    If Not AnyadirAFormula(cadFormula, CadenaDevuelta2) Then Exit Sub
    'Nº Albaran
    CadenaDevuelta2 = "{scaalb.numalbar}=" & Numalbar
    If Not AnyadirAFormula(cadFormula, CadenaDevuelta2) Then Exit Sub
    'select para insertar en tabla temporal
    cadSelect = QuitarCaracterACadena(cadFormula, "{")
    cadSelect = QuitarCaracterACadena(cadSelect, "}")

   
    '=========================================================================

    'Aqui sabemos que valor tiene CodClien y añadimos a los parametros el tipo de IVA
    'que se aplica a ese cliente
    CadenaDevuelta2 = DevuelveDesdeBDNew(conAri, "scaalb", "codclien", "codtipom", D, "T", , "numalbar", Numalbar, "N")
    If CadenaDevuelta2 <> "" Then
        FechaAlb = "albarcon"   'Albaran valorado
        CadenaDevuelta2 = DevuelveDesdeBDNew(conAri, "sclien", "tipoiva", "codclien", CadenaDevuelta2, "N", FechaAlb)
        If CadenaDevuelta2 <> "" Then CadenaSQL = CadenaSQL & "pTipoIVA=" & CadenaDevuelta2 & "|"
            
        
        If FechaAlb = "" Or FechaAlb = "albarcon" Then FechaAlb = "0"
        ' 0 "Todo"
        ' 1 "Cantidad y Precio"
        ' 2 "Cantidad"
        CadenaSQL = CadenaSQL & "Albarcon=" & FechaAlb & "|"

    End If
     

     With frmImprimir
            .FormulaSeleccion = cadFormula
            .OtrosParametros = CadenaSQL
            .NumeroParametros = 12
            .SoloImprimir = False
            .EnvioEMail = False
            .Opcion = 45
            .Titulo = "Albaran de Cliente"
            .ConSubInforme = True
            .Show vbModal
    End With
            
    
    

    
End Sub


Private Sub MultiInsercionCampos(DesdeMutiparte As Boolean)
Dim i As Integer


        'Quito el indicador # de multi campo
        CadenaDesdeOtroForm = Mid(CadenaDesdeOtroForm, 2)

        'D = "INSERT INTO advpartes_campos( numparte, codcampo) VALUES "
        ' Me.Text1(0).Text & "," & txtAux2(0).Text & ")"
        CadenaSQL = ""
        While CadenaDesdeOtroForm <> ""
            i = InStr(1, CadenaDesdeOtroForm, "·#")
            
            If i = 0 Then
                CadenaDesdeOtroForm = ""
            Else
                FechaAlb = Mid(CadenaDesdeOtroForm, 1, i - 1)
                If DesdeMutiparte Then
                    NumRegElim = RecuperaValor(FechaAlb, 4)
                    CadenaDesdeOtroForm = Mid(CadenaDesdeOtroForm, i + 2)
                    FechaAlb = RecuperaValor(FechaAlb, 1) 'cdocampo
                Else
                    NumRegElim = RecuperaValor(FechaAlb, 1)
                    CadenaDesdeOtroForm = Mid(CadenaDesdeOtroForm, i + 2)
                    FechaAlb = RecuperaValor(FechaAlb, 1) 'cdocampo
                
                End If
                
                
                If DesdeMutiparte Then
                    CadenaSQL = CadenaSQL & ", (" & vUsu.Codigo & "," & FechaAlb & "," & NumRegElim & ")"
                Else
                    CadenaSQL = CadenaSQL & ", (" & Text1(0).Text & "," & FechaAlb & ")"
                End If
            End If
        Wend
        CadenaSQL = Mid(CadenaSQL, 2) 'quito la primera coma
        
        If DesdeMutiparte Then
            CadenaSQL = "INSERT IGNORE  tmpnlotes(codusu,codprove,numalbar) VALUES " & CadenaSQL
            conn.Execute CadenaSQL
        Else
            CadenaSQL = "INSERT INTO advpartes_campos( numparte, codcampo) VALUES " & CadenaSQL
        
            If ejecutar(CadenaSQL, False) Then
                'Hay que refrescar y boton anyadir
                Data2.Refresh
                CargaGrid2 DataGrid2, Data2   'solo caraga el suyo
                BotonAnyadirLinea
            End If
        
        End If
        
        FechaAlb = ""
        CadenaSQL = ""
        
        '
        
        
End Sub


Private Sub POnerMultiParte(visible As Boolean)

    Screen.MousePointer = vbHourglass

    conn.Execute "DELETE FROM tmpnlotes WHERE codusu =" & vUsu.Codigo
    


    Me.FrameMultparte.Top = 0
    FrameMultparte.visible = visible
    mnOpciones.Enabled = Not visible
    If visible Then PonerModo 3
    If visible Then
        cmdMultiParte(1).Cancel = True
        Me.lwC.ListItems.Clear
    Else
        Me.cmdCancelar.Cancel = True
    End If
    
    
    
    Screen.MousePointer = vbDefault
End Sub



Private Function datosOk_Multi() As Boolean
Dim Agru As String
Dim K As Integer
    datosOk_Multi = False
    
    CadenaSQL = ""
    If lwC.ListItems.Count = 0 Then CadenaSQL = CadenaSQL & vbCrLf & "-Seleccione algun campo"
    If Me.Text1(10).Text = "" Then CadenaSQL = CadenaSQL & vbCrLf & "-Fecha parte "
    If Me.Text1(11).Text = "" Then CadenaSQL = CadenaSQL & vbCrLf & "-Trabajador "
    If Me.Text1(12).Text = "" Then CadenaSQL = CadenaSQL & vbCrLf & "-Tratamiento "
    
    If CadenaSQL <> "" Then
        MsgBox "Campos requeridos: " & vbCrLf & CadenaSQL, vbExclamation
        Exit Function
    End If
    
    CodZona = 0
    For NumRegElim = 1 To lwC.ListItems.Count
        If lwC.ListItems(NumRegElim).Checked Then
            CodZona = CodZona + 1
            If lwC.ListItems(NumRegElim).ForeColor = vbRed Then
                
                MsgBox "Cliente bloqueado", vbExclamation
                Exit Function
            
            End If
        End If
    Next
    If CodZona = 0 Then
        MsgBox "Ningun campo seleccionado", vbExclamation
        Exit Function
    End If
    
    CadenaSQL = "codtrata= " & DBSet(Text1(11).Text, "T") & " AND dosishab>0 AND 1"
    CadenaSQL = DevuelveDesdeBD(conAri, "*", "advtrata_lineas", CadenaSQL, 1, "N")
    If CadenaSQL <> "" Then
        'lleva dosis , Debe indicar cantidad
        CadenaSQL = ""
        If Text1(13).Text = "" Then
            CadenaSQL = "Debe indicar cantidad"
        Else
            If ImporteFormateado(Text1(13).Text) = 0 Then CadenaSQL = "La cantidad no puede ser cero"
        End If
        If CadenaSQL <> "" Then
            MsgBox CadenaSQL, vbExclamation
            Exit Function
        End If
    End If
    'Se van a crear tantos partes como clientes
    CadenaSQL = "|"
    FechaAlb = ""
    CodZona = 0
    Agru = ""
    K = 0
    
    For NumRegElim = 1 To lwC.ListItems.Count
        If lwC.ListItems(NumRegElim).Checked Then
            If InStr(1, CadenaSQL, "|" & lwC.ListItems(NumRegElim).Tag & "|") = 0 Then
            
                CadenaSQL = CadenaSQL & lwC.ListItems(NumRegElim).Tag & "|"
                FechaAlb = FechaAlb & "X"
            End If
            If lwC.ListItems(NumRegElim).SubItems(5) <> Agru Then
               Agru = lwC.ListItems(NumRegElim).SubItems(5)
                K = K + 1
            End If
            CodZona = CodZona + 1
        End If
    Next
    CadenaSQL = "Partes: " & K & vbCrLf
    CadenaSQL = CadenaSQL & "Clientes: " & Len(FechaAlb) & vbCrLf
    CadenaSQL = CadenaSQL & "Total campos: " & CodZona & vbCrLf & vbCrLf & "¿Continuar?"
    CadenaSQL = "Se van a generar " & vbCrLf & vbCrLf & CadenaSQL
    If MsgBox(CadenaSQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Function
    
    datosOk_Multi = True
    
End Function




Private Function GenerarPartes() As Boolean
Dim pos As Integer
Dim Cli As Long
Dim Campos As String
Dim CrearParte As Boolean
Dim Agrupacion As Long


On Error GoTo eGenerarPartes
    GenerarPartes = False
    
    NumRegElim = SugerirCodigoSiguienteStr("advpartes", "numparte")
    
    'Para la impresion
    conn.Execute "DELETE FROM tmpcrmclien WHERE codusu =" & vUsu.Codigo
    
    pos = 1
    'numparte,fechapar,codtrata,codclien,litrospre,factursn,observac,nrohoras,cerrado,coddirec,codtraba,EsExterno
    Cli = -1
    Agrupacion = -1
    Campos = ""
    Do
        If lwC.ListItems(pos).Checked Then
        
            If Val(lwC.ListItems(pos).SubItems(5)) <> Agrupacion Then
                If Agrupacion >= 0 Then
                    GeneraUnParte Campos, Cli, Agrupacion
                    NumRegElim = NumRegElim + 1
                End If
                Agrupacion = Val(lwC.ListItems(pos).SubItems(5))
                 Cli = lwC.ListItems(pos).Tag
                 
                Campos = ""
            Else
                'Esta seleccionado
                If lwC.ListItems(pos).Tag <> Cli Then
                    'NO deberia haber pasado. Es una misma numerodecampo-variedad . NO deberia tener otro cliente
                   If MsgBox("Distinto cliente para un mismo numero de campo . ¿Continuar?", vbQuestion + vbYesNoCancel) <> vbYes Then Err.Raise 513, "", "Proceso cancelado por el usuario"
                    
                   Cli = lwC.ListItems(pos).Tag
                    
                End If
            End If
            Campos = Campos & ", (" & NumRegElim & "," & lwC.ListItems(pos).SubItems(2) & ")"
        End If
        pos = pos + 1
    Loop Until pos > lwC.ListItems.Count
    
    If Agrupacion >= 0 Then
        GeneraUnParte Campos, Cli, Agrupacion 'El ultimo
        GenerarPartes = True
    End If
    
    Exit Function
eGenerarPartes:
    MuestraError Err.Number, Err.Description



End Function


'Si lleva campops es que quiero generar la cabecera y las lineas campo
Private Function GeneraUnParte(ByRef Campos As String, Cliente As Long, Agrupacion As Long) As String
Dim Cad As String
Dim CPrFact As CPreciosFact
Dim Au As String
Dim cantidad As Currency
Dim L As Integer

        ' 'Generamos parte
        Au = "INSERT INTO advpartes(numparte , fechapar, codtrata, codClien, litrospre, litrosrea , factursn, observac, nrohoras, cerrado, CodDirec, CodTraba, EsExterno)"
        Au = Au & " VALUES (" & NumRegElim & "," & DBSet(Text1(10).Text, "F") & "," & DBSet(Text1(11).Text, "T") & "," & Cliente & ","
        Au = Au & DBSet(Text1(13).Text, "N") & "," & DBSet(Text1(13).Text, "N") & ",1,"
        Au = Au & DBSet("[GEN " & vUsu.Login & "     NºCampo: " & Agrupacion & " ]" & vbCrLf & Text1(14).Text, "T") & ","
        Au = Au & "null,0,NULL," & Text1(12).Text & "," & Me.Check1(1).Value & ")"
        conn.Execute Au
        
        'Para la impresion posterior
        Au = "INSERT INTO tmpcrmclien(codusu, codclien) VALUES (" & vUsu.Codigo & "," & NumRegElim & ")"
        conn.Execute Au
        
        Au = Mid(Campos, 2) 'quitamos la primera coma
        Au = "INSERT INTO advpartes_campos (numparte ,codcampo) VALUES " & Au
        conn.Execute Au


    Cad = "select * from advtrata_lineas where codtrata=" & DBSet(Text1(11).Text, "T") & " order by numlinea"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    L = 0
    Cad = ""
    While Not miRsAux.EOF
        L = L + 1
        'd = "insert into `advparteslineas` (`numparte`,`numlinea`,`codalmac`,`codartic`,"
        'd = d & "`dosishab`,`cantidad`,`ampliaci`,`dtoline1`,`dtoline2`,`origpre`,`preciove`,`importel`) values ("
        
        'Marzo 2022
        'Para alzira
        Cad = Cad & ", (" & NumRegElim & "," & L & "," & AlmacenLin & ","
        
        
        
        Cad = Cad & DBSet(miRsAux!codArtic, "T") & ","
        'codartic,dosishab,cantidad,
        cantidad = ImporteFormateado(Text1(13).Text)
        If IsNull(miRsAux!dosishab) Then
            'Es cantidad
            Cad = Cad & "NULL,"
            Cad = Cad & DBSet(miRsAux!cantidad, "N") & ","
            cantidad = miRsAux!cantidad
        Else
            'Dosis hab
            cantidad = (cantidad * miRsAux!dosishab) / 1000
            
            Cad = Cad & DBSet(miRsAux!dosishab, "N") & ","
            Cad = Cad & DBSet(cantidad, "N") & ","
            
        End If
        
        
        'precio
        
            Set CPrFact = New CPreciosFact
    
            CPrFact.CodigoArtic = miRsAux!codArtic
            CPrFact.CodigoClien = Cliente
            CPrFact.FijarTarifaActividad
            
            
            D = CPrFact.ObtenerPrecio(False, Text1(10).Text, Au, "")
            
            '`ampliaci`,`dtoline1`,`dtoline2`,`origpre`,`preciove`,`importel`
            Cad = Cad & "NULL," & DBSet(CPrFact.Descuento1, "N") & "," & DBSet(CPrFact.Descuento2, "N")
            Cad = Cad & ",'" & Au & "'," & DBSet(D, "N")
            

            


            Au = CalcularImporte(CStr(cantidad), D, CPrFact.Descuento1, CPrFact.Descuento2, vParamAplic.TipoDtos)
            Cad = Cad & "," & DBSet(Au, "N") & ")"
        
            Set CPrFact = Nothing
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    If Cad <> "" Then
        Cad = Mid(Cad, 2)
        D = "insert into `advparteslineas` (`numparte`,`numlinea`,`codalmac`,`codartic`,"
        D = D & "dosishab,`cantidad`,`ampliaci`,`dtoline1`,`dtoline2`,`origpre`,`preciove`,`importel`) values "
        D = D & Cad
        conn.Execute D
    End If
    
End Function

