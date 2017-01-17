VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTelematVarios 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8640
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   12930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameDescuadreRefencias 
      Height          =   4455
      Left            =   5760
      TabIndex        =   42
      Top             =   480
      Visible         =   0   'False
      Width           =   6135
      Begin VB.CheckBox chkCabel 
         Caption         =   "Proveedor CABEL"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   62
         Top             =   3600
         Width           =   2775
      End
      Begin VB.CommandButton cmdListadoSinCruadrar 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3360
         TabIndex        =   47
         Top             =   3840
         Width           =   1215
      End
      Begin VB.TextBox txtDescFamia 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   0
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   58
         Text            =   "Text5"
         Top             =   2520
         Width           =   3495
      End
      Begin VB.TextBox txtFamia 
         Height          =   285
         Index           =   0
         Left            =   1320
         TabIndex        =   45
         Text            =   "Text1"
         Top             =   2520
         Width           =   975
      End
      Begin VB.TextBox txtDescFamia 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   1
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   55
         Text            =   "Text5"
         Top             =   3000
         Width           =   3495
      End
      Begin VB.TextBox txtFamia 
         Height          =   285
         Index           =   1
         Left            =   1320
         TabIndex        =   46
         Text            =   "Text1"
         Top             =   3000
         Width           =   975
      End
      Begin VB.TextBox txtProve 
         Height          =   285
         Index           =   5
         Left            =   1320
         TabIndex        =   44
         Text            =   "Text1"
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox txtDescProve 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   2400
         TabIndex        =   52
         Text            =   "Text2"
         Top             =   1680
         Width           =   3495
      End
      Begin VB.TextBox txtProve 
         Height          =   285
         Index           =   4
         Left            =   1320
         TabIndex        =   43
         Text            =   "Text1"
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox txtDescProve 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   2400
         TabIndex        =   49
         Text            =   "Text2"
         Top             =   1200
         Width           =   3495
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   3
         Left            =   4680
         TabIndex        =   48
         Top             =   3840
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Label6"
         Height          =   255
         Left            =   240
         TabIndex        =   60
         Top             =   3960
         Width           =   2655
      End
      Begin VB.Image imgFamilia 
         Height          =   240
         Index           =   0
         Left            =   960
         Picture         =   "frmTelematVarios.frx":0000
         Top             =   2520
         Width           =   240
      End
      Begin VB.Label Label33 
         Caption         =   "Desde"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   59
         Top             =   2520
         Width           =   465
      End
      Begin VB.Image imgFamilia 
         Height          =   240
         Index           =   1
         Left            =   960
         Picture         =   "frmTelematVarios.frx":0102
         Top             =   3000
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Familia"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   81
         Left            =   120
         TabIndex        =   57
         Top             =   2280
         Width           =   600
      End
      Begin VB.Label Label33 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   44
         Left            =   240
         TabIndex        =   56
         Top             =   3000
         Width           =   465
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Referencias sin cruzar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   2
         Left            =   480
         TabIndex        =   54
         Top             =   240
         Width           =   4845
      End
      Begin VB.Image imgPorv 
         Height          =   240
         Index           =   5
         Left            =   960
         Picture         =   "frmTelematVarios.frx":0204
         ToolTipText     =   "Buscar centro coste"
         Top             =   1680
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   53
         Top             =   1680
         Width           =   615
      End
      Begin VB.Image imgPorv 
         Height          =   240
         Index           =   4
         Left            =   960
         Picture         =   "frmTelematVarios.frx":0306
         ToolTipText     =   "Buscar centro coste"
         Top             =   1200
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Proveedor"
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
         Index           =   6
         Left            =   120
         TabIndex        =   51
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   50
         Top             =   1200
         Width           =   615
      End
   End
   Begin VB.Frame FrameImprimirTel 
      Height          =   4455
      Left            =   240
      TabIndex        =   27
      Top             =   0
      Width           =   6135
      Begin VB.TextBox txtFamia 
         Height          =   285
         Index           =   3
         Left            =   1440
         TabIndex        =   32
         Text            =   "Text1"
         Top             =   2640
         Width           =   975
      End
      Begin VB.TextBox txtDescFamia 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   3
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   68
         Text            =   "Text5"
         Top             =   2640
         Width           =   3495
      End
      Begin VB.TextBox txtFamia 
         Height          =   285
         Index           =   2
         Left            =   1440
         TabIndex        =   30
         Text            =   "Text1"
         Top             =   2280
         Width           =   975
      End
      Begin VB.TextBox txtDescFamia 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   2
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   65
         Text            =   "Text5"
         Top             =   2280
         Width           =   3495
      End
      Begin VB.CheckBox chkCabel 
         Caption         =   "Proveedor CABEL"
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   31
         Top             =   3240
         Width           =   2775
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmTelematVarios.frx":0408
         Left            =   360
         List            =   "frmTelematVarios.frx":0415
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   3720
         Width           =   2055
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "Imprimir"
         Height          =   375
         Left            =   3360
         TabIndex        =   34
         Top             =   3840
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   2
         Left            =   4680
         TabIndex        =   35
         Top             =   3840
         Width           =   1215
      End
      Begin VB.TextBox txtDescProve 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   2520
         TabIndex        =   37
         Text            =   "Text2"
         Top             =   1440
         Width           =   3495
      End
      Begin VB.TextBox txtProve 
         Height          =   285
         Index           =   3
         Left            =   1440
         TabIndex        =   29
         Text            =   "Text1"
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox txtDescProve 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   2520
         TabIndex        =   36
         Text            =   "Text2"
         Top             =   1080
         Width           =   3495
      End
      Begin VB.TextBox txtProve 
         Height          =   285
         Index           =   2
         Left            =   1440
         TabIndex        =   28
         Text            =   "Text1"
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Listado telematel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   0
         Left            =   480
         TabIndex        =   70
         Top             =   120
         Width           =   4845
      End
      Begin VB.Label Label33 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   2
         Left            =   360
         TabIndex        =   69
         Top             =   2640
         Width           =   465
      End
      Begin VB.Image imgFamilia 
         Height          =   240
         Index           =   3
         Left            =   1080
         Picture         =   "frmTelematVarios.frx":0438
         Top             =   2640
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Familia"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   67
         Top             =   2040
         Width           =   600
      End
      Begin VB.Label Label33 
         Caption         =   "Desde"
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   66
         Top             =   2280
         Width           =   465
      End
      Begin VB.Image imgFamilia 
         Height          =   240
         Index           =   2
         Left            =   1080
         Picture         =   "frmTelematVarios.frx":053A
         Top             =   2280
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   40
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   39
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Proveedor"
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
         Index           =   5
         Left            =   240
         TabIndex        =   38
         Top             =   720
         Width           =   1215
      End
      Begin VB.Image imgPorv 
         Height          =   240
         Index           =   3
         Left            =   1080
         Picture         =   "frmTelematVarios.frx":063C
         ToolTipText     =   "Buscar centro coste"
         Top             =   1440
         Width           =   240
      End
      Begin VB.Image imgPorv 
         Height          =   240
         Index           =   2
         Left            =   1080
         Picture         =   "frmTelematVarios.frx":073E
         ToolTipText     =   "Buscar centro coste"
         Top             =   1080
         Width           =   240
      End
   End
   Begin VB.Frame FrameCruzar 
      Height          =   4695
      Left            =   240
      TabIndex        =   16
      Top             =   120
      Width           =   10815
      Begin VB.CheckBox chkCabel 
         Caption         =   "Proveedor CABEL"
         Height          =   195
         Index           =   2
         Left            =   7320
         TabIndex        =   63
         Top             =   360
         Width           =   2775
      End
      Begin VB.CommandButton cmdBuscaArticulos 
         Height          =   495
         Left            =   6360
         Picture         =   "frmTelematVarios.frx":0840
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Busca articulos"
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdCruzar 
         Caption         =   "Actualizar"
         Height          =   375
         Left            =   7920
         TabIndex        =   21
         Top             =   4200
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   1
         Left            =   9360
         TabIndex        =   20
         Top             =   4200
         Width           =   1215
      End
      Begin VB.TextBox txtProve 
         Height          =   285
         Index           =   1
         Left            =   1560
         TabIndex        =   18
         Text            =   "Text1"
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtDescProve 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   2640
         TabIndex        =   17
         Text            =   "Text2"
         Top             =   360
         Width           =   3495
      End
      Begin MSComctlLib.ListView lw2 
         Height          =   3135
         Left            =   240
         TabIndex        =   24
         Top             =   840
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   5530
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Cod.Tele."
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   7168
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Cod.artic."
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Fecha"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Precio"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label3 
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   26
         Top             =   4200
         Width           =   3735
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   3
         Left            =   360
         Picture         =   "frmTelematVarios.frx":1242
         ToolTipText     =   "Seleccionar todo"
         Top             =   4200
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   2
         Left            =   720
         Picture         =   "frmTelematVarios.frx":138C
         ToolTipText     =   "Quitar seleccion"
         Top             =   4200
         Width           =   240
      End
      Begin VB.Image imgPorv 
         Height          =   240
         Index           =   1
         Left            =   1200
         Picture         =   "frmTelematVarios.frx":14D6
         ToolTipText     =   "Buscar centro coste"
         Top             =   360
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Proveedor"
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
         Index           =   4
         Left            =   240
         TabIndex        =   19
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame FrameProgress 
      Height          =   1215
      Left            =   6600
      TabIndex        =   14
      Top             =   2760
      Visible         =   0   'False
      Width           =   4815
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   375
         Left            =   240
         TabIndex        =   22
         Top             =   720
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label2 
         Caption         =   "progreso"
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width           =   4215
      End
   End
   Begin VB.Frame FrameActprec 
      Height          =   8175
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   12615
      Begin VB.CheckBox chkCabel 
         Caption         =   "Proveedor CABEL"
         Height          =   195
         Index           =   3
         Left            =   10680
         TabIndex        =   64
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox txtDecimal 
         Height          =   285
         Index           =   0
         Left            =   5280
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   480
         Width           =   855
      End
      Begin VB.CheckBox chkImportes 
         Caption         =   "Actualiza precio compra"
         Height          =   255
         Left            =   1320
         TabIndex        =   23
         Top             =   7680
         Width           =   2175
      End
      Begin VB.CommandButton cmdActualizar 
         Caption         =   "Actualizar"
         Height          =   375
         Left            =   9720
         TabIndex        =   13
         Top             =   7680
         Visible         =   0   'False
         Width           =   1095
      End
      Begin MSComctlLib.ListView lw1 
         Height          =   6615
         Left            =   240
         TabIndex        =   12
         Top             =   960
         Width           =   12135
         _ExtentX        =   21405
         _ExtentY        =   11668
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Cod.Tele."
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   7168
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Cod.artic."
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Fecha"
            Object.Width           =   2028
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "P.V.P."
            Object.Width           =   2134
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "M"
            Object.Width           =   529
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Precio"
            Object.Width           =   2134
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Familia"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   0
         Left            =   11040
         TabIndex        =   11
         Top             =   7680
         Width           =   1215
      End
      Begin VB.CommandButton cmdVerArt 
         Height          =   495
         Left            =   10080
         Picture         =   "frmTelematVarios.frx":15D8
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Carga datos"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   1
         Left            =   8640
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   0
         Left            =   7200
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txtDescProve 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   1200
         TabIndex        =   6
         Text            =   "Text2"
         Top             =   480
         Width           =   3735
      End
      Begin VB.TextBox txtProve 
         Height          =   285
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Margen%"
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
         Index           =   7
         Left            =   5280
         TabIndex        =   61
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "En rojo articulos en promoción y/o precio especial.     M-> Supera el margen"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   3600
         TabIndex        =   41
         Top             =   7680
         Visible         =   0   'False
         Width           =   5655
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   1
         Left            =   600
         Picture         =   "frmTelematVarios.frx":1FDA
         ToolTipText     =   "Quitar seleccion"
         Top             =   7680
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   0
         Left            =   240
         Picture         =   "frmTelematVarios.frx":2124
         ToolTipText     =   "Seleccionar todo"
         Top             =   7680
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Fecha"
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
         Index           =   1
         Left            =   6240
         TabIndex        =   8
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Proveedor"
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
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   855
      End
      Begin VB.Image imgF 
         Height          =   240
         Index           =   1
         Left            =   9120
         Picture         =   "frmTelematVarios.frx":226E
         ToolTipText     =   "Buscar centro coste"
         Top             =   240
         Width           =   240
      End
      Begin VB.Image imgPorv 
         Height          =   240
         Index           =   0
         Left            =   1200
         Picture         =   "frmTelematVarios.frx":27F8
         ToolTipText     =   "Buscar centro coste"
         Top             =   240
         Width           =   240
      End
      Begin VB.Image imgF 
         Height          =   240
         Index           =   0
         Left            =   7680
         Picture         =   "frmTelematVarios.frx":28FA
         ToolTipText     =   "Buscar centro coste"
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Desde"
         Height          =   255
         Index           =   2
         Left            =   7200
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   3
         Left            =   8640
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmTelematVarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Opcion As Byte
        '0.- Actualizar precios proveedor
        '1.- Cruzar telamtel con sartic. Buscar referencias
        
        '2.- Imprimir
        '3.- Imprimir descuadre refencias
        
Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1

Dim SQL As String
Dim IT As ListItem
Dim N As Long


Private Sub cmdActualizar_Click()
Dim Sigue As Boolean
Dim FamiliasSeleccionadas As String

    If lw1.ListItems.Count = 0 Then Exit Sub
    
    SQL = ""
    For N = 1 To lw1.ListItems.Count
        If lw1.ListItems(N).Checked Then SQL = SQL & "X"
    Next
    
    If SQL = "" Then
        MsgBox "Seleccione algun artículo", vbExclamation
        Exit Sub
    End If
    
  
    
    
    
    pb1.Max = Len(SQL)
    pb1.Value = 0
    
    
    'Si en el desde / hasta hay fechanue NO puede seguir.Tiene
    'Voy a comprobar si hay articulos con fechanue
    SQL = ""
    For N = 1 To lw1.ListItems.Count
        If lw1.ListItems(N).Checked Then SQL = SQL & ", " & DBSet(lw1.ListItems(N).SubItems(2), "T")
    Next
    SQL = Mid(SQL, 2)
    SQL = "(" & SQL & ")"
    SQL = " NOT fechanue is null and codartic IN " & SQL & " AND codlista "
    SQL = DevuelveDesdeBD(conAri, "count(*)", "slista", SQL, vParamAplic.CodTarifa)
    If SQL = "" Then SQL = "0"
    If Val(SQL) > 0 Then
        MsgBox "Hay articulos en lista precio venta con fecha de cambio sin actualizar el precio", vbExclamation
        Exit Sub
    End If
    
    'Si marca actualizar en PROVE comprobaremos tb que no existan valores con fechanue <>null
    If chkImportes.Value = 1 Then
           SQL = ""
           
           For N = 1 To lw1.ListItems.Count
               If lw1.ListItems(N).Checked Then SQL = SQL & ", " & DBSet(lw1.ListItems(N).SubItems(2), "T")
           Next
           SQL = Mid(SQL, 2)
           SQL = "(" & SQL & ")"
           
           
           If Me.chkCabel(3).Value = 0 Then
                'Para todos los proveedores
                SQL = " NOT fechanue is null and codartic IN " & SQL & " AND codprove "
                SQL = DevuelveDesdeBD(conAri, "count(*)", "slispr", SQL, Me.txtProve(0).Text)
                If SQL = "" Then SQL = "0"
                If Val(SQL) > 0 Then
                    MsgBox "hay articulos en lista precio compra con fecha de cambio sin actualizar el precio", vbExclamation
                    Exit Sub
                End If
            
        
            Else
                'PARA CABEL
                SQL = " AND NOT fechanue is null and slispr.codartic IN " & SQL
                Set miRsAux = New ADODB.Recordset
                SQL = "Select slispr.*,sartic.codprove proveEnSartic,nomartic from slispr,sartic WHERE slispr.codartic=sartic.codartic" & SQL
                miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                SQL = ""
                While Not miRsAux.EOF
                    If miRsAux!proveEnSartic = miRsAux!Codprove Then
                        'ESTE ES el articulo/proveedor, con lo cual, tienen que actualizar
                        SQL = SQL & "    -" & miRsAux!codArtic & vbCrLf & "       " & LCase(miRsAux!NomArtic) & "     PROV: " & miRsAux!Codprove & vbCrLf
                    End If
                    miRsAux.MoveNext
                Wend
                miRsAux.Close
                Set miRsAux = Nothing
                If SQL <> "" Then
                    SQL = "Falta actualizar precios compra." & vbCrLf & String(42, "=") & vbCrLf & SQL
                    MsgBox SQL, vbExclamation
                    Exit Sub
                End If
                
            End If
    End If
    FamiliasSeleccionadas = ""
    
    

    
    'Modificacion MARZO 2015
    'Pedira las familias que quiere actualizar
    SQL = ","
    NumRegElim = 0
    For N = 1 To lw1.ListItems.Count
        If lw1.ListItems(N).Checked Then
            If InStr(1, SQL, "," & Trim(lw1.ListItems(N).SubItems(7)) & ",") = 0 Then
                SQL = SQL & Trim(lw1.ListItems(N).SubItems(7)) & ","
                NumRegElim = NumRegElim + 1
            End If
        End If
    Next
     
    'SQL no puede ser EOF
    FamiliasSeleccionadas = ""
    CadenaDesdeOtroForm = ""
    If NumRegElim > 1 Then
        'HAY MAS DE UNA familia. Sacamos previo
        frmListado5.OpcionListado = 7
        'quitamos las comas
        SQL = Mid(SQL, 2)
        SQL = Mid(SQL, 1, Len(SQL) - 1)
        frmListado5.OtrosDatos = SQL
        frmListado5.Show vbModal
        If CadenaDesdeOtroForm = "" Then Exit Sub 'Ha cancelado
        FamiliasSeleccionadas = Replace(CadenaDesdeOtroForm, ",", "|")
    Else
        'Solo hay una familia.. Preguntamos
        FamiliasSeleccionadas = Replace(SQL, ",", "|")
        SQL = "Va a realizar la actualización de los artículo seleccionados" & vbCrLf & "¿Continuar?"
        If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then Exit Sub
    End If
    
    
    
    
    
    
    
    Me.FrameProgress.visible = True
    DoEvents
    
    Screen.MousePointer = vbHourglass
    
    Sigue = True
    If vParamAplic.ActualizaPrecioEspecial Then
        If Not BloqueoManual("ACTPRE", "1") Then Sigue = False
    End If
    
    If Sigue Then
        For N = lw1.ListItems.Count To 1 Step -1
            If lw1.ListItems(N).Checked Then
                
                IncrementaPG lw1.ListItems(N).SubItems(2), 1
                If (pb1.Value Mod 10) = 0 Then
                    Me.Refresh
                    DoEvents
                End If

                
                
                'NUEVO
                'MARZO 2015
                'Si no es de las familias seleccionadas NO hacemos nada
                SQL = "|" & lw1.ListItems(N).SubItems(7) & "|"
                
                If InStr(1, FamiliasSeleccionadas, SQL) = 0 Then
                    'Esa familia NO la ha seleccionado
                    'Stop
                Else
                
                    If ActualizarPrecio2 Then lw1.ListItems.Remove N
                End If
            End If
        Next
    End If
    
    If vParamAplic.ActualizaPrecioEspecial Then DesBloqueoManual "ACTPRE"
    
    Me.FrameProgress.visible = False
    Screen.MousePointer = vbDefault
     
End Sub


Private Sub IncrementaPG(texto As String, Inc As Integer)
On Error Resume Next
    Me.Label2.Caption = texto
    Me.Label2.Refresh
    pb1.Value = pb1.Value + Inc
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub cmdBuscaArticulos_Click()
    Screen.MousePointer = vbHourglass
    CruzarReferenciasProveedor
    Screen.MousePointer = vbDefault
End Sub

Private Sub CruzarReferenciasProveedor()
    
    If chkCabel(2).Value = 0 Then
        If txtProve(1).Text = "" Or Me.txtDescProve(1).Text = "" Then
            MsgBox "Falta proveedor", vbExclamation
            Exit Sub
        End If
    Else
        txtProve(1).Text = ""
        Me.txtDescProve(1).Text = ""
    End If
    
    cmdBuscaArticulos.Tag = txtProve(1).Text 'para que no me cambien el proveedor
    'Buscamos los articulos del proveedor
    Label3.Caption = "Leyendo referencias sin asignar"
    Label3.Refresh
    Set miRsAux = New ADODB.Recordset
    lw2.ListItems.Clear
    If chkCabel(2).Value = 0 Then
        SQL = "Select * from stelem where codprove = " & txtProve(1).Text & " AND codartic IS NULL "
    Else
        SQL = "Select * from stelem where codprove is null AND codartic IS NULL "
    End If
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set IT = lw2.ListItems.Add()
        IT.Text = miRsAux!referprov
        IT.SubItems(1) = miRsAux!Nombre
        IT.SubItems(2) = "  "
        IT.SubItems(3) = Format(miRsAux!FechaCambio, "dd/mm/yyyy")
        IT.SubItems(4) = Format(miRsAux!Precio, FormatoPrecio)
        IT.Tag = miRsAux!codtelem
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If lw2.ListItems.Count > 0 Then
        Label3.Caption = "Cruzando datos articulos"
        Label3.Refresh
        DoEvents
        Espera 0.5
        'Cargo los articulos de ese proveedor
        If chkCabel(2).Value = 0 Then
            SQL = "Select codartic,referprov from sartic where codprove = " & Me.txtProve(1).Text
        Else
            SQL = "Select codartic,referprov from sartic,sfamia where sartic.codfamia=sfamia.codfamia and marcapropia = 1"
        End If
        miRsAux.Open SQL, conn, adOpenKeyset, adLockPessimistic, adCmdText
        
        'Ahora cojo y en desde el final voy intentando conseguir
        For N = lw2.ListItems.Count To 1 Step -1
            SQL = "referprov = " & DBSet(lw2.ListItems(N).Text, "T")
            miRsAux.Find SQL, , adSearchForward, 1
            If miRsAux.EOF Then
                'NO existe. Borro
                lw2.ListItems.Remove N
            Else
                lw2.ListItems(N).SubItems(2) = miRsAux!codArtic
            End If
    
        Next N
        miRsAux.Close
    End If
    Label3.Caption = ""
    Set miRsAux = Nothing
End Sub

Private Sub cmdCancel_Click(Index As Integer)
    Unload Me
End Sub







Private Sub cmdCruzar_Click()
Dim AUx As String

    If Me.lw2.ListItems.Count = 0 Then Exit Sub
    SQL = "Seleccione algun articulo para actualizar la referencia"
    For N = 1 To lw2.ListItems.Count
        If lw2.ListItems(N).Checked Then
            SQL = ""
            Exit For
        End If
    Next
    
    If SQL = "" Then
        If MsgBox("Desea actualizar las referencias?", vbQuestion + vbYesNo) = vbNo Then SQL = "SAL"
    Else
        MsgBox SQL, vbInformation
    End If
        
    If SQL <> "" Then Exit Sub
    
    If cmdBuscaArticulos.Tag <> txtProve(1).Text Then
        MsgBox "Has vuelto a cambiar el proveedor", vbExclamation
        Exit Sub
    End If
    
    
    For N = 1 To lw2.ListItems.Count
        If lw2.ListItems(N).Checked Then
            SQL = "UPDATE stelem set codartic=" & DBSet(lw2.ListItems(N).SubItems(2), "T")
            SQL = SQL & " WHERE codtelem = " & lw2.ListItems(N).Tag
            conn.Execute SQL
            
            
            SQL = DevuelveDesdeBD(conAri, "codean", "stelem", "codtelem", lw2.ListItems(N).Tag, "N")
            If SQL <> "" Then
                AUx = DevuelveDesdeBD(conAri, "max(numlinea)", "sarti3", "codartic", DevNombreSQL(lw2.ListItems(N).SubItems(2)), "T")
                If AUx = "" Then AUx = "0"
                AUx = CStr(Val(AUx) + 1)
        
                SQL = "INSERT INTO sarti3(codartic,numlinea,codigoea) VALUES (" & DBSet(lw2.ListItems(N).SubItems(2), "T") & "," & AUx & "," & DBSet(SQL, "T") & ")"

                ejecutar SQL, False
            End If
        End If
    Next
    Unload Me
End Sub

Private Sub cmdImprimir_Click()
Dim AUx As String

    If Not FijarSqlListadoTelematel(False) Then Exit Sub


    
    If Not HayRegParaInforme("stelem left join sartic on stelem.codartic=sartic.codartic", SQL) Then Exit Sub
    
    FijarSqlListadoTelematel True
    
    If CadenaDesdeOtroForm <> "" Then CadenaDesdeOtroForm = "pDesde=""" & CadenaDesdeOtroForm & """|"
    With frmImprimir
        .FormulaSeleccion = SQL
        .NumeroParametros = 1
        If CadenaDesdeOtroForm <> "" Then .NumeroParametros = 2
        
        CadenaDesdeOtroForm = "|pNomEmpre=""" & vEmpresa.nomempre & """|" & CadenaDesdeOtroForm
        .OtrosParametros = CadenaDesdeOtroForm
        .SoloImprimir = False
        .EnvioEMail = False
        .Opcion = 5
        .Titulo = "Listado artículos telematel"
        .NombreRPT = "rTelematel.rpt"  'Nombre fichero .rpt a Imprimir
        .Show vbModal
    End With

End Sub

Private Function FijarSqlListadoTelematel(ParaReport As Boolean) As Boolean
Dim Familia As String

    FijarSqlListadoTelematel = False


'    '===================================================
'    '============ PARAMETROS ===========================
    SQL = ""
    CadenaDesdeOtroForm = ""
    
    If chkCabel(1).Value = 1 Then
        ''CABEL
        If txtProve(2).Text <> "" Or txtProve(3).Text <> "" Then
            MsgBox "No debe indicar valores para povoeedor si marca la opción CABEL", vbExclamation
            Exit Function
        End If
        
        CadenaDesdeOtroForm = "Proveedor CABEL"
        If ParaReport Then
            SQL = " isnull({stelem.codprove}) "
        Else
            SQL = " (stelem.codprove)  IS null "
        End If
        
        
        'FAMILIA 2017
         'FAMILIA 2017
        Familia = ""
        If txtFamia(2).Text <> "" Then Familia = Familia & "Desde " & Trim(txtFamia(2).Text & "  " & Me.txtDescFamia(2).Text)
        If txtFamia(3).Text <> "" Then Familia = Trim(Familia & "    hasta " & txtFamia(3).Text & "  " & Me.txtDescFamia(3).Text)
         
        If Familia <> "" Then
            Familia = "Familia " & Familia
            If CadenaDesdeOtroForm <> "" Then CadenaDesdeOtroForm = CadenaDesdeOtroForm & """ + chr(13) + """
            CadenaDesdeOtroForm = CadenaDesdeOtroForm & Familia
            
            Familia = CadenaDesdeHastaBD(txtFamia(2).Text, txtFamia(3).Text, "(sartic.codfamia)", "N")
            If Familia <> "" Then
                'quito los aprentesis
                Familia = Mid(Familia, 2)
                Familia = Mid(Familia, 1, Len(Familia) - 1)
            End If
                        
            If SQL <> "" Then SQL = SQL & " AND "
            SQL = SQL & Familia
        End If
        
        
        
        
    Else
        If txtProve(2).Text <> "" Or txtProve(3).Text <> "" Then
            If txtProve(2).Text <> "" Then CadenaDesdeOtroForm = CadenaDesdeOtroForm & "Desde " & Trim(txtProve(2).Text & "  " & Me.txtDescProve(2).Text)
            If txtProve(3).Text <> "" Then CadenaDesdeOtroForm = Trim(CadenaDesdeOtroForm & "    hasta " & txtProve(3).Text & "  " & Me.txtDescProve(3).Text)
            If CadenaDesdeOtroForm <> "" Then CadenaDesdeOtroForm = "Proveedor " & CadenaDesdeOtroForm
            SQL = CadenaDesdeHastaBD(txtProve(2).Text, txtProve(3).Text, "(stelem.codprove)", "N")
            If SQL <> "" Then
                'quito los aprentesis
                SQL = Mid(SQL, 2)
                SQL = Mid(SQL, 1, Len(SQL) - 1)
            End If
            
             
         End If
         If txtFamia(2).Text <> "" Or txtFamia(3).Text <> "" Then
            'FAMILIA 2017
            Familia = ""
            If txtFamia(2).Text <> "" Then Familia = Familia & "Desde " & Trim(txtFamia(2).Text & "  " & Me.txtDescFamia(2).Text)
            If txtFamia(3).Text <> "" Then Familia = Trim(Familia & "    hasta " & txtFamia(3).Text & "  " & Me.txtDescFamia(3).Text)
             
            If Familia <> "" Then
                Familia = "Familia " & Familia
                If CadenaDesdeOtroForm <> "" Then CadenaDesdeOtroForm = CadenaDesdeOtroForm & """ + chr(13) + """
                CadenaDesdeOtroForm = CadenaDesdeOtroForm & Familia
                
                Familia = CadenaDesdeHastaBD(txtFamia(2).Text, txtFamia(3).Text, "(sartic.codfamia)", "N")
                If Familia <> "" Then
                    'quito los aprentesis
                    Familia = Mid(Familia, 2)
                    Familia = Mid(Familia, 1, Len(Familia) - 1)
                End If
                            
                If SQL <> "" Then SQL = SQL & " AND "
                SQL = SQL & Familia
            End If
            
           
            
            
            
        End If
        
        If ParaReport Then
            SQL = Replace(SQL, "(", "{")
            SQL = Replace(SQL, ")", "}")
        End If
        
    End If
    
    
    If Combo1.ListIndex > 0 Then
        'Solo capturadas  o pendientes
        If Combo1.ListIndex = 1 Then
            CadenaDesdeOtroForm = CadenaDesdeOtroForm & "      Capturadas "
            If SQL <> "" Then SQL = SQL & " AND "
            If ParaReport Then
                SQL = SQL & " {stelem.codartic} <>"""""
            Else
                SQL = SQL & " (stelem.codartic) <>"""""
            End If
            
            
        Else
            CadenaDesdeOtroForm = CadenaDesdeOtroForm & "      Pendientes "
            SQL = SQL & " AND "
            If ParaReport Then
                SQL = SQL & " isnull({stelem.codartic})"
                
            Else
                SQL = SQL & "(stelem.codartic is null) "
            End If
        End If
        CadenaDesdeOtroForm = Trim(CadenaDesdeOtroForm)
    End If
    FijarSqlListadoTelematel = True

End Function



Private Sub cmdListadoSinCruadrar_Click()
Dim B As Boolean
    
    Screen.MousePointer = vbHourglass
    Label6.Caption = "Prepara datos"
    Label6.Refresh
    SQL = "DELETE FROM tmpinformes where codusu = " & vUsu.codigo
    conn.Execute SQL
    
    If chkCabel(0).Value = 0 Then
        B = CargaTablaReferenciasSinCruzar
    Else
        B = CargaTablaReferenciasSinCruzarCABEL
    End If
    
    Label6.Caption = ""
    Screen.MousePointer = vbDefault
    
    If B Then
        CadenaDesdeOtroForm = ""
        SQL = ""
        If Me.txtProve(4).Text <> "" Then SQL = "desde " & txtProve(4).Text & " - " & txtDescProve(4).Text
        If Me.txtProve(5).Text <> "" Then SQL = SQL & "hesde " & txtProve(5).Text & " - " & txtDescProve(5).Text
        If SQL <> "" Then SQL = "Proveedor " & SQL
        CadenaDesdeOtroForm = "pAnyo=""" & SQL & """|"
    
        SQL = ""
        If Me.txtFamia(0).Text <> "" Then SQL = "desde " & txtFamia(0).Text & " - " & txtDescFamia(0).Text
        If Me.txtFamia(1).Text <> "" Then SQL = SQL & "hesde " & txtFamia(1).Text & " - " & txtDescFamia(1).Text
        If SQL <> "" Then SQL = "Familia: " & SQL
        CadenaDesdeOtroForm = CadenaDesdeOtroForm & "pDHCliente=""" & SQL & """|"
    
        With frmImprimir
            .FormulaSeleccion = "{tmpinformes.codusu}=" & vUsu.codigo
            
            'D/H
            .NumeroParametros = 3
            .OtrosParametros = CadenaDesdeOtroForm
            .SoloImprimir = False
            .EnvioEMail = False
            .Opcion = 5
            .Titulo = "Referencias sin cruzar telematel"
            .NombreRPT = "rTelematelSinX.rpt"
            .Show vbModal
         End With
         CadenaDesdeOtroForm = ""
    End If
    
    
End Sub

Private Sub cmdVerArt_Click()
Dim RT As ADODB.Recordset
Dim Importe As Currency
Dim K As Long

    Label5.visible = False
    
    SQL = ""
    If chkCabel(3).Value = 0 Then
        If txtProve(0).Text = "" Then SQL = "Falta proveedor"
    Else
        If txtProve(0).Text <> "" Then SQL = "No debe indicar proveedor "
    End If
    If SQL <> "" Then
        MsgBox SQL, vbExclamation
        Exit Sub
    End If
    
    'Si habia datos...
    Set RT = New ADODB.Recordset
    'Articulos en promocion
    If chkCabel(3).Value = 0 Then
        SQL = "select distinct(spromo.codartic) codartic from spromo,sartic where spromo.codartic=sartic.codartic and codprove= " & txtProve(0).Text
        SQL = SQL & " Union select distinct(sprees.codartic) codartic from sprees,sartic where sprees.codartic=sartic.codartic and codprove= " & txtProve(0).Text
        RT.Open SQL, conn, adOpenKeyset, adLockPessimistic, adCmdText
    End If
    
    'Limpiamos
    lw1.ListItems.Clear
    
    SQL = "Select  stelem.*,preciove,codfamia from stelem "
    'Marzo 2015
    'SQL = SQL & " left join sartic on stelem.codartic=sartic.codartic"
    SQL = SQL & " inner join sartic on stelem.codartic=sartic.codartic"
    
    SQL = SQL & " Where stelem.codArtic <> """" AND stelem.codprove "
    If chkCabel(3).Value = 0 Then
        SQL = SQL & " = " & txtProve(0).Text
    Else
        SQL = SQL & " IS NULL"
    End If
    
    
    'Si tiene fecha
    If txtFecha(0).Text <> "" Then SQL = SQL & " AND fechacambio >=" & DBSet(txtFecha(0).Text, "F")
    If txtFecha(1).Text <> "" Then SQL = SQL & " AND fechacambio <=" & DBSet(txtFecha(1).Text, "F")
    
    Set miRsAux = New ADODB.Recordset
    
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set IT = lw1.ListItems.Add()
        IT.Text = miRsAux!codtelem
        IT.SubItems(1) = miRsAux!Nombre
        IT.SubItems(2) = miRsAux!codArtic
        IT.SubItems(3) = Format(miRsAux!FechaCambio, "dd/mm/yyyy")
        
        IT.SubItems(5) = " "
        If IsNull(miRsAux!PrecioVe) Then
            IT.SubItems(4) = " "
        Else
            IT.SubItems(4) = Format(miRsAux!PrecioVe, FormatoPrecio)
            Importe = ImporteFormateado(Me.txtDecimal(0).Text)
            Importe = (Importe + 100) / 100
            Importe = miRsAux!PrecioVe * Importe
            If miRsAux!Precio > Importe Then IT.SubItems(5) = "*"
        End If
        
        'IT.SubItems(4) = Format(miRsAux!Precio, FormatoPrecio)
        IT.SubItems(6) = Format(miRsAux!Precio, FormatoPrecio)
        
        IT.SubItems(7) = DBLet(miRsAux!Codfamia, "T") 'La familia del articulo
        
        
        
        
        If Me.chkCabel(3).Value = 0 Then
            RT.Find "codartic=" & DBSet(miRsAux!codArtic, "T"), , adSearchForward, 1
            
        Else
            If lw1.ListItems.Count > 1 Then RT.Close
            SQL = "select spromo.codartic codartic from spromo WHERE codartic = " & DBSet(miRsAux!codArtic, "T")
            SQL = SQL & " Union select sprees.codartic codartic from sprees where sprees.codartic=" & DBSet(miRsAux!codArtic, "T")
            RT.Open SQL, conn, adOpenKeyset, adLockPessimistic, adCmdText
        End If
        If Not RT.EOF Then
            Label5.visible = True
            IT.ForeColor = vbRed
            IT.ListSubItems(1).ForeColor = vbRed
            IT.ListSubItems(2).ForeColor = vbRed
            IT.ListSubItems(3).ForeColor = vbRed
            IT.ListSubItems(4).ForeColor = vbRed
            IT.ListSubItems(5).ForeColor = vbRed
            IT.ListSubItems(6).ForeColor = vbRed
        End If
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    RT.Close
    Set miRsAux = Nothing
    Set RT = Nothing
    
    cmdActualizar.visible = lw1.ListItems.Count > 0
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KEYpressGnral KeyAscii, 3, False
End Sub

Private Sub Form_Load()
    Me.Icon = frmPpal.Icon
    FrameActprec.visible = False
    Me.FrameCruzar.visible = False
    FrameImprimirTel.visible = False
    FrameDescuadreRefencias.visible = False
    limpiar Me
    Select Case Opcion
    Case 0
        PonerFrameVisible Me.FrameActprec
        Caption = "Actualizar precios"
        LeerGuardarMargen True
    Case 1
        PonerFrameVisible Me.FrameCruzar
        Caption = "Buscar referencias"
    Case 2
        PonerFrameVisible Me.FrameImprimirTel
        Caption = "Imprimir"
        Me.Combo1.ListIndex = 0
    Case 3
        Label6.Caption = "" 'indicador
        PonerFrameVisible Me.FrameDescuadreRefencias
        Caption = "Imprimir"
        
    End Select
    
    Me.cmdCancel(Opcion).Cancel = True
End Sub

Private Sub PonerFrameVisible(Fr As Frame)

    Fr.visible = True
    Fr.Top = 0
    Fr.Left = 120
    
    Me.Height = Fr.Height + 540
    Me.Width = Fr.Width + 340
    
    
    
    'El frame del prgogress
    N = CInt(Fr.Height - FrameProgress.Height) \ 2
    FrameProgress.Top = N
    N = CInt(Fr.Width - FrameProgress.Width) \ 2
    FrameProgress.Left = N
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Opcion = 0 Then LeerGuardarMargen False
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
    SQL = CadenaDevuelta
End Sub

Private Sub frmC_Selec(vFecha As Date)
    SQL = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub imgCheck_Click(Index As Integer)
    If Index < 2 Then
        For N = 1 To lw1.ListItems.Count
            lw1.ListItems(N).Checked = Index = 0
        Next
    Else
        For N = 1 To lw2.ListItems.Count
            lw2.ListItems(N).Checked = Index = 3
        Next
    End If
End Sub

Private Sub imgF_Click(Index As Integer)
    SQL = ""
    Set frmC = New frmCal
    frmC.Fecha = Now
    If txtFecha(Index).Text <> "" Then frmC.Fecha = CDate(txtFecha(Index).Text)
    frmC.Show vbModal
    Set frmC = Nothing
    If SQL <> "" Then txtFecha(Index).Text = SQL
    SQL = ""
End Sub

Private Sub imgFamilia_Click(Index As Integer)
    SQL = ""
    Screen.MousePointer = vbHourglass
    Set frmB = New frmBuscaGrid
    frmB.vTitulo = "Familia"
    frmB.vCampos = "Codigo|sfamia|Codfamia|N||20·descripcion|sfamia|nomfamia|T||45·"
    frmB.vCargaFrame = False
    frmB.vDevuelve = "0|1|"
    frmB.vselElem = 1
    frmB.vConexionGrid = 1  'ODBC Ariges
    frmB.vTabla = "sfamia"
    frmB.vSQL = ""
    frmB.Show vbModal
    Set frmB = Nothing
    Screen.MousePointer = vbDefault
    If SQL <> "" Then
        
        Me.txtFamia(Index).Text = RecuperaValor(SQL, 1)
        Me.txtDescFamia(Index).Text = RecuperaValor(SQL, 2)
        PonerFoco txtFamia(Index)
        SQL = ""
    End If
End Sub

Private Sub imgPorv_Click(Index As Integer)
    lanzaBusqueda 0
    If SQL <> "" Then
        txtProve(Index).Text = RecuperaValor(SQL, 1)
        txtDescProve(Index).Text = RecuperaValor(SQL, 2)
        SQL = ""
    End If
End Sub

Private Sub txtFamia_GotFocus(Index As Integer)
    ConseguirFoco txtFamia(Index), 3
End Sub

Private Sub txtFamia_KeyPress(Index As Integer, KeyAscii As Integer)
     KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtFamia_LostFocus(Index As Integer)
    txtFamia(Index).Text = Trim(txtFamia(Index).Text)
    SQL = ""
    If txtFamia(Index).Text <> "" Then
        If IsNumeric(txtFamia(Index).Text) Then
            SQL = DevuelveDesdeBD(conAri, "nomfamia", "sfamia", "codfamia", txtFamia(Index).Text, "N")
            If SQL = "" Then MsgBox "El codigo no pertence a ningun familia", vbExclamation
        Else
            MsgBox "Campo numerico", vbExclamation
            txtFamia(Index).Text = ""
            PonerFoco txtFamia(Index)
        End If
    End If
     
    Me.txtDescFamia(Index).Text = SQL
    
End Sub


Private Sub txtFecha_GotFocus(Index As Integer)
     ConseguirFoco txtFecha(Index), 3
End Sub

Private Sub txtFecha_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 3, False
End Sub



Private Sub txtFecha_LostFocus(Index As Integer)
    PonerFormatoFecha txtFecha(Index)
End Sub

Private Sub txtProve_GotFocus(Index As Integer)
     ConseguirFoco txtProve(Index), 3
End Sub

Private Sub txtProve_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 3, False
End Sub

Private Sub txtProve_LostFocus(Index As Integer)
    SQL = ""
    If txtProve(Index).Text <> "" Then
        If Not PonerFormatoEntero(txtProve(Index)) Then
            txtProve(Index).Text = ""
        Else
            SQL = PonerNombreDeCod(txtProve(Index), conAri, "sprove", "nomprove", "codprove")
            
        End If
    End If
    Me.txtDescProve(Index).Text = SQL
    If SQL = "" Then
        'Segun sea la opcin pondre un lw u otra a blanco
        
        If Opcion = 0 Then lw1.ListItems.Clear
    End If
End Sub


Private Sub txtDecimal_GotFocus(Index As Integer)
    ConseguirFoco txtDecimal(Index), 3
End Sub

Private Sub txtDecimal_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtDecimal_LostFocus(Index As Integer)
Dim B As Boolean
    txtDecimal(Index).Text = Trim(txtDecimal(Index).Text)
    If txtDecimal(Index).Text <> "" Then
       ' If Index = 0 Then
            B = PonerFormatoDecimal(txtDecimal(Index), 4)
       ' Else
       '     B = PonerFormatoDecimal(txtDecimal(Index), 3)
       ' End If
        If B Then

        Else
            txtDecimal(Index).Text = ""
        End If
    End If
End Sub





Private Sub lanzaBusqueda(Cual As Byte)

            Set frmB = New frmBuscaGrid
            SQL = ""
            SQL = SQL & "Código|sprove|codprove|N|000000|18·"
            SQL = SQL & "Nombre|sprove|nomprove|T||40·"
            SQL = SQL & "Nom.Comer.|sprove|nomcomer|T||40·"
        
            frmB.vTabla = "sprove"
            frmB.vTitulo = "Proveedores"
        
            frmB.vCampos = SQL
            
           
            frmB.vSQL = ""
            '###A mano
            frmB.vDevuelve = "0|1|"
            frmB.vselElem = 1
            frmB.vConexionGrid = conAri
            frmB.vCargaFrame = False
            '#
            SQL = ""
            frmB.Show vbModal
            Set frmB = Nothing


End Sub



Private Function ActualizarPrecio2() As Boolean
Dim PVP As Currency
Dim MenError As String

    On Error GoTo EAC
    
    
    'Voy a actualizar el pvp
        
   
    

    PVP = ImporteFormateado(lw1.ListItems(N).SubItems(6))
    
  
                

  
        
         

    SQL = "UPDATE slista SET  fechanue = " & DBSet(lw1.ListItems(N).SubItems(3), "F")
    SQL = SQL & " ,precionu = " & TransformaComasPuntos(CStr(PVP))
    SQL = SQL & " WHERE codartic = " & DBSet(lw1.ListItems(N).SubItems(2), "T") & " AND codlista = " & vParamAplic.CodTarifa
    conn.Execute SQL
         
      
    If vParamAplic.ActualizaPrecioEspecial Then ActualizarPrecioEspecialGenerico lw1.ListItems(N).SubItems(2), PVP, False
        






    '------------------------------------------------------------------------------------------------------------------
    '------------------------------------------------------------------------------------------------------------------
    'PRECIO COMPRA
    '------------------------------------------------------------------------------------------------------------------
    If Me.chkImportes.Value = 1 Then
                                
            
           If Me.chkCabel(3).Value = 1 Then
                SQL = DevuelveDesdeBD(conAri, "codprove", "sartic", "codartic", DBSet(lw1.ListItems(N).SubItems(2), "T"))
                If SQL = "" Then
                    MsgBox "Error obteniendo proveedor para el articulo: " & lw1.ListItems(N).SubItems(2), vbExclamation
                    SQL = "-1"
                End If
           Else
                SQL = txtProve(0).Text
           End If
           SQL = " WHERE codartic = " & DBSet(lw1.ListItems(N).SubItems(2), "T") & " AND codprove = " & SQL
           SQL = ", precionu = " & TransformaComasPuntos(CStr(PVP)) & SQL
           SQL = "UPDATE slispr SET fechanue = " & DBSet(lw1.ListItems(N).SubItems(3), "F") & SQL
           conn.Execute SQL

    End If   'del check



    ActualizarPrecio2 = True
          
EAC:
    
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description & vbCrLf & SQL
    Set miRsAux = Nothing
End Function
















Private Function CargaTablaReferenciasSinCruzar() As Boolean
Dim CP As Collection
Dim AUx As String

    On Error GoTo eCargaTablaReferenciasSinCruzar
    
    Label6.Caption = "Prepara datos"
    Label6.Refresh
    SQL = "DELETE FROM tmpinformes where codusu = " & vUsu.codigo
    conn.Execute SQL
    Set miRsAux = New ADODB.Recordset
    Set CP = New Collection
    

    SQL = "Select codprove from sartic WHERE artvario=0 and referprov<> '' "
    AUx = CadenaDesdeHastaBD(txtProve(4).Text, txtProve(5).Text, "(sartic.codprove)", "N")

    If AUx <> "" Then SQL = SQL & " AND " & AUx
    AUx = CadenaDesdeHastaBD(txtFamia(0).Text, txtFamia(1).Text, "(sartic.codfamia)", "N")
    If AUx <> "" Then SQL = SQL & " AND " & AUx
    SQL = SQL & " GROUP BY 1"
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        CP.Add CStr(miRsAux!Codprove)
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    NumRegElim = 0
    
    
    'Para cada proveedor
    For N = 1 To CP.Count
            
        SQL = "select codartic,nomartic,codtelem,referprov,codfamia from sartic"
        SQL = SQL & " where codprove=" & CP.item(N) & " and codfamia<>9998 and artvario=0 "
        AUx = CadenaDesdeHastaBD(txtFamia(0).Text, txtFamia(1).Text, "(sartic.codfamia)", "N")
        If AUx <> "" Then SQL = SQL & " AND " & AUx
        SQL = SQL & " AND referprov not in (select referprov from stelem where codprove=" & CP.item(N) & ")"

        
        Me.Label6.Caption = CP.item(N) & " (" & N & "/" & CP.Count & ")"
        Label6.Refresh
        
        miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        SQL = ""
        'insert into `tmpinformes` (`codusu`,`codigo1`,`campo1`,`campo2`,`nombre1`,`nombre2`,`nombre3`)
        'values ( '1','1','1','2','codartic','nomartic','refprove')
        
        While Not miRsAux.EOF
            NumRegElim = NumRegElim + 1
            SQL = SQL & ", (" & vUsu.codigo & "," & NumRegElim & "," & CP.item(N) & "," & miRsAux!Codfamia & ","
            SQL = SQL & DBSet(miRsAux!codArtic, "T") & "," & DBSet(miRsAux!NomArtic, "T") & "," & DBSet(miRsAux!referprov, "T") & ")"
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        If SQL <> "" Then
            SQL = Mid(SQL, 2)
            AUx = "insert into `tmpinformes` (`codusu`,`codigo1`,`campo1`,`campo2`,`nombre1`,`nombre2`,`nombre3`) VALUES " & SQL
            conn.Execute AUx
        End If
    Next
    If NumRegElim > 0 Then CargaTablaReferenciasSinCruzar = True
eCargaTablaReferenciasSinCruzar:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set miRsAux = Nothing

End Function


Private Function CargaTablaReferenciasSinCruzarCABEL() As Boolean
Dim AUx As String

    On Error GoTo eCargaTablaReferenciasSinCruzarC
    
    
    Set miRsAux = New ADODB.Recordset

    

    
        
            
        SQL = "select codartic,nomartic,codtelem,referprov,sartic.codfamia,sartic.codprove from sartic,sfamia"
        SQL = SQL & " WHERE sartic.codfamia=sfamia.codfamia AND sfamia.marcapropia=1"
        SQL = SQL & " AND sartic.codfamia<>9998 and artvario=0 "
        AUx = CadenaDesdeHastaBD(txtFamia(0).Text, txtFamia(1).Text, "(sartic.codfamia)", "N")
        If AUx <> "" Then SQL = SQL & " AND " & AUx
        SQL = SQL & " AND referprov not in (select referprov from stelem where codprove is null )"

        
        Me.Label6.Caption = "Procesando"
        Label6.Refresh
        
        miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        SQL = ""
        'insert into `tmpinformes` (`codusu`,`codigo1`,`campo1`,`campo2`,`nombre1`,`nombre2`,`nombre3`)
        'values ( '1','1','1','2','codartic','nomartic','refprove')
        
        While Not miRsAux.EOF
            NumRegElim = NumRegElim + 1
            SQL = SQL & ", (" & vUsu.codigo & "," & NumRegElim & "," & miRsAux!Codfamia & "," & miRsAux!Codfamia & ","
            SQL = SQL & DBSet(miRsAux!codArtic, "T") & "," & DBSet(miRsAux!NomArtic, "T") & "," & DBSet(miRsAux!referprov, "T") & ")"
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        If SQL <> "" Then
            NumRegElim = NumRegElim + 1
            SQL = Mid(SQL, 2)
            AUx = "insert into `tmpinformes` (`codusu`,`codigo1`,`campo1`,`campo2`,`nombre1`,`nombre2`,`nombre3`) VALUES " & SQL
            conn.Execute AUx
        End If
    
    If NumRegElim > 0 Then CargaTablaReferenciasSinCruzarCABEL = True
eCargaTablaReferenciasSinCruzarC:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set miRsAux = Nothing

End Function




Private Sub LeerGuardarMargen(Leer As Boolean)
    On Error GoTo ELeerGuardarImportes
    
    SQL = App.Path & "\telmar.xdf"
    N = FreeFile
    If Leer Then
        
        If Dir(SQL, vbArchive) <> "" Then
            Open SQL For Input As #N
            Line Input #N, SQL
            Close #N
            
        Else
            SQL = ""
        End If
        If SQL = "" Then SQL = "30,00"
        txtDecimal(0).Text = SQL
        txtDecimal(0).Tag = SQL
        
    Else
        If txtDecimal(0).Tag <> txtDecimal(0).Text Then
            Open SQL For Output As #N
            Print #N, txtDecimal(0).Text
            Close #N
        End If
    End If
    
ELeerGuardarImportes:
    If Err.Number <> 0 Then Err.Clear
    
End Sub

