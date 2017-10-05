VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmListado4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado"
   ClientHeight    =   10170
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10170
   ScaleWidth      =   13710
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame FrameUpdateaSoloProveedor 
      Height          =   3135
      Left            =   5040
      TabIndex        =   110
      Top             =   840
      Visible         =   0   'False
      Width           =   6615
      Begin VB.CommandButton cmdActualizaSoloProveedor 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3960
         TabIndex        =   112
         Top             =   2640
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   15
         Left            =   5280
         TabIndex        =   111
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Label Label2 
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   121
         Top             =   2640
         Width           =   3495
      End
      Begin VB.Label Label2 
         Height          =   255
         Index           =   9
         Left            =   240
         TabIndex        =   120
         Top             =   2040
         Width           =   3255
      End
      Begin VB.Label Label2 
         Height          =   255
         Index           =   8
         Left            =   2640
         TabIndex        =   119
         Top             =   1560
         Width           =   3255
      End
      Begin VB.Label Label2 
         Caption         =   "Destino:"
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
         Left            =   360
         TabIndex        =   118
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label2 
         Height          =   255
         Index           =   6
         Left            =   1440
         TabIndex        =   117
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label2 
         Height          =   255
         Index           =   5
         Left            =   2640
         TabIndex        =   116
         Top             =   1080
         Width           =   3255
      End
      Begin VB.Label Label2 
         Caption         =   "Origen:"
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
         Left            =   360
         TabIndex        =   115
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "0001"
         Height          =   255
         Index           =   3
         Left            =   1440
         TabIndex        =   114
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Modificar proveedor articulo - familia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   360
         Index           =   20
         Left            =   480
         TabIndex        =   113
         Top             =   240
         Width           =   5235
      End
   End
   Begin VB.Frame frameMultiAlbaranes 
      Height          =   7335
      Left            =   360
      TabIndex        =   104
      Top             =   120
      Visible         =   0   'False
      Width           =   11055
      Begin VB.ComboBox cboTipoIVA 
         Height          =   315
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   108
         Top             =   6940
         Width           =   2655
      End
      Begin VB.CommandButton cmdMultialbaran 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   8280
         TabIndex        =   107
         Top             =   6840
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   14
         Left            =   9600
         TabIndex        =   106
         Top             =   6840
         Width           =   1095
      End
      Begin MSComctlLib.ListView lw 
         Height          =   6495
         Index           =   10
         Left            =   120
         TabIndex        =   105
         Top             =   240
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   11456
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Albaran"
            Object.Width           =   2381
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Fecha"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Cliente"
            Object.Width           =   1588
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Nombre"
            Object.Width           =   7039
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Dtos"
            Object.Width           =   1501
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Importe"
            Object.Width           =   2187
         EndProperty
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo IVA"
         Height          =   255
         Left            =   1440
         TabIndex        =   109
         Top             =   6960
         Width           =   855
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   11
         Left            =   600
         Picture         =   "frmListado4.frx":0000
         Top             =   6960
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   10
         Left            =   240
         Picture         =   "frmListado4.frx":014A
         Top             =   6960
         Width           =   240
      End
   End
   Begin VB.Frame FrameTipoprecFac 
      Height          =   9015
      Left            =   120
      TabIndex        =   24
      Top             =   0
      Visible         =   0   'False
      Width           =   13575
      Begin VB.TextBox txtDecimal 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   3720
         TabIndex        =   30
         Text            =   "Text2"
         Top             =   8640
         Width           =   855
      End
      Begin VB.ComboBox cboTipoPrecio2 
         Height          =   315
         ItemData        =   "frmListado4.frx":0294
         Left            =   1680
         List            =   "frmListado4.frx":02A1
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   8640
         Width           =   1815
      End
      Begin VB.CommandButton cmdFraTipoPrecio 
         Caption         =   "&Cambiar"
         Height          =   375
         Left            =   4920
         TabIndex        =   31
         Top             =   8520
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Salir"
         Height          =   375
         Index           =   6
         Left            =   12000
         TabIndex        =   25
         Top             =   8520
         Width           =   1095
      End
      Begin MSComctlLib.ListView lw 
         Height          =   7695
         Index           =   2
         Left            =   120
         TabIndex        =   26
         Top             =   720
         Width           =   13215
         _ExtentX        =   23310
         _ExtentY        =   13573
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
         NumItems        =   10
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Albaran"
            Object.Width           =   2734
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Articulo"
            Object.Width           =   2681
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Descripcion"
            Object.Width           =   7039
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Cantidad"
            Object.Width           =   1501
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "precio"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Dto1"
            Object.Width           =   1060
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Dto2"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Importe"
            Object.Width           =   1834
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "Comi."
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   9
            Text            =   "ECO"
            Object.Width           =   1058
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "%"
         Height          =   255
         Left            =   4680
         TabIndex        =   103
         Top             =   8640
         Width           =   135
      End
      Begin VB.Label lblTitulo 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   3
         Left            =   6960
         TabIndex        =   28
         Top             =   240
         Width           =   5235
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Modificar comision linea facturas "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   360
         Index           =   2
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   7635
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   5
         Left            =   480
         Picture         =   "frmListado4.frx":02BC
         Top             =   8640
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   4
         Left            =   120
         Picture         =   "frmListado4.frx":0406
         Top             =   8640
         Width           =   240
      End
      Begin VB.Image imgAyuda 
         Height          =   255
         Index           =   1
         Left            =   960
         ToolTipText     =   "Factura. Precio normal/eco"
         Top             =   8640
         Width           =   255
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   360
      Top             =   9600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameCambioFamMarca 
      Height          =   9735
      Left            =   600
      TabIndex        =   81
      Top             =   0
      Visible         =   0   'False
      Width           =   9975
      Begin VB.CommandButton cmdUpdatearFamiliaMarca 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   7200
         TabIndex        =   85
         Top             =   9240
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   12
         Left            =   8520
         TabIndex        =   82
         Top             =   9240
         Width           =   1095
      End
      Begin MSComctlLib.ListView lw 
         Height          =   8295
         Index           =   8
         Left            =   120
         TabIndex        =   83
         Top             =   720
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   14631
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Codigo"
            Object.Width           =   3200
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripcion"
            Object.Width           =   7805
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Marca"
            Object.Width           =   1305
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Desc. marca"
            Object.Width           =   3572
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Familia"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "codprove"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label lblTitulo 
         Caption         =   "T"
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
         Height          =   360
         Index           =   19
         Left            =   1320
         TabIndex        =   102
         Top             =   9240
         Width           =   5865
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Cambiar familia/marca/proveedor del artículo"
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
         Index           =   14
         Left            =   120
         TabIndex        =   84
         Top             =   240
         Width           =   7665
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   9
         Left            =   720
         Picture         =   "frmListado4.frx":0550
         Top             =   9360
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   8
         Left            =   240
         Picture         =   "frmListado4.frx":069A
         Top             =   9360
         Width           =   240
      End
   End
   Begin VB.Frame FrameAgua 
      Height          =   6975
      Left            =   120
      TabIndex        =   86
      Top             =   0
      Visible         =   0   'False
      Width           =   12255
      Begin VB.Frame Frame1 
         Height          =   735
         Left            =   120
         TabIndex        =   95
         Top             =   6120
         Width           =   9495
         Begin VB.TextBox txtDecimal 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   8040
            TabIndex        =   99
            Text            =   "Text2"
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   4080
            TabIndex        =   98
            Text            =   "Text1"
            Top             =   240
            Width           =   3975
         End
         Begin VB.OptionButton optAgua2 
            Caption         =   "Quitar asignacion"
            Height          =   255
            Index           =   1
            Left            =   1560
            TabIndex        =   97
            Top             =   240
            Width           =   1695
         End
         Begin VB.OptionButton optAgua2 
            Caption         =   "Asignar cuota"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   96
            Top             =   240
            Value           =   -1  'True
            Width           =   1695
         End
         Begin VB.Label lblTitulo 
            Caption         =   "Descr:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   17
            Left            =   3360
            TabIndex        =   101
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label lblTitulo 
            Caption         =   "€"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   16
            Left            =   9240
            TabIndex        =   100
            Top             =   240
            Width           =   255
         End
      End
      Begin VB.OptionButton optAgua 
         Caption         =   "Nombre"
         Height          =   255
         Index           =   2
         Left            =   11040
         TabIndex        =   93
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdAgua 
         Caption         =   "Actualizar"
         Height          =   375
         Left            =   9960
         TabIndex        =   92
         Top             =   6240
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   13
         Left            =   11040
         TabIndex        =   91
         Top             =   6240
         Width           =   975
      End
      Begin VB.OptionButton optAgua 
         Caption         =   "Cod. cliente"
         Height          =   255
         Index           =   1
         Left            =   9480
         TabIndex        =   90
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton optAgua 
         Caption         =   "Contador"
         Height          =   255
         Index           =   0
         Left            =   8160
         TabIndex        =   89
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
      Begin MSComctlLib.ListView lw 
         Height          =   5295
         Index           =   9
         Left            =   120
         TabIndex        =   87
         Top             =   720
         Width           =   11895
         _ExtentX        =   20981
         _ExtentY        =   9340
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Contador"
            Object.Width           =   2187
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Cod"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Nombre"
            Object.Width           =   6068
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Cuota"
            Object.Width           =   4304
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Importe"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "Fact"
            Object.Width           =   1058
         EndProperty
      End
      Begin VB.Label lblAzira 
         Caption         =   "Ordenado :"
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
         Index           =   2
         Left            =   6960
         TabIndex        =   94
         Top             =   240
         Width           =   1005
      End
      Begin VB.Image Imga 
         Height          =   240
         Index           =   5
         Left            =   2520
         Picture         =   "frmListado4.frx":07E4
         ToolTipText     =   "Eliminar telefono"
         Top             =   360
         Width           =   240
      End
      Begin VB.Image Imga 
         Height          =   240
         Index           =   4
         Left            =   2160
         Picture         =   "frmListado4.frx":11E6
         ToolTipText     =   "Añadir telefono"
         Top             =   360
         Width           =   240
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Contador / cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   15
         Left            =   120
         TabIndex        =   88
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.Frame FrameAjuteCuotasTfnia 
      Height          =   9495
      Left            =   120
      TabIndex        =   57
      Top             =   0
      Width           =   12375
      Begin VB.ComboBox cboOperadora 
         Height          =   315
         Left            =   8760
         Style           =   2  'Dropdown List
         TabIndex        =   71
         Top             =   4440
         Width           =   2055
      End
      Begin VB.CommandButton cmdCutoasMasivas 
         Caption         =   "Generar"
         Height          =   375
         Left            =   9840
         TabIndex        =   70
         Top             =   8760
         Width           =   975
      End
      Begin VB.OptionButton optTfnia 
         Caption         =   "Operadora"
         Height          =   255
         Index           =   3
         Left            =   10800
         TabIndex        =   69
         Top             =   3840
         Width           =   1335
      End
      Begin VB.OptionButton optTfnia 
         Caption         =   "Nombre"
         Height          =   255
         Index           =   2
         Left            =   9600
         TabIndex        =   68
         Top             =   3840
         Width           =   1335
      End
      Begin VB.OptionButton optTfnia 
         Caption         =   "Cod.Cliente"
         Height          =   255
         Index           =   1
         Left            =   8280
         TabIndex        =   67
         Top             =   3840
         Width           =   1335
      End
      Begin VB.OptionButton optTfnia 
         Caption         =   "Teléfono"
         Height          =   255
         Index           =   0
         Left            =   6960
         TabIndex        =   66
         Top             =   3840
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.Frame FrameOpcionesModCutoas 
         Height          =   735
         Left            =   6960
         TabIndex        =   63
         Top             =   2520
         Width           =   5175
         Begin VB.CommandButton cmdTfnia 
            Height          =   375
            Index           =   2
            Left            =   1200
            Picture         =   "frmListado4.frx":1BE8
            Style           =   1  'Graphical
            TabIndex        =   74
            ToolTipText     =   "Imprimir listado telefonos"
            Top             =   240
            Width           =   375
         End
         Begin VB.CommandButton cmdTfnia 
            Height          =   375
            Index           =   1
            Left            =   600
            Picture         =   "frmListado4.frx":25EA
            Style           =   1  'Graphical
            TabIndex        =   65
            ToolTipText     =   "Guardar datos telefonos"
            Top             =   240
            Width           =   375
         End
         Begin VB.CommandButton cmdTfnia 
            Height          =   375
            Index           =   0
            Left            =   120
            Picture         =   "frmListado4.frx":2FEC
            Style           =   1  'Graphical
            TabIndex        =   64
            ToolTipText     =   "Leer datos guardados"
            Top             =   240
            Width           =   375
         End
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   10
         Left            =   11040
         TabIndex        =   58
         Top             =   8760
         Width           =   1095
      End
      Begin MSComctlLib.ListView lw 
         Height          =   8655
         Index           =   5
         Left            =   120
         TabIndex        =   59
         Top             =   600
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   15266
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Telefono"
            Object.Width           =   2187
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Codigo"
            Object.Width           =   1516
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Nombre"
            Object.Width           =   5009
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Operadora"
            Object.Width           =   1940
         EndProperty
      End
      Begin MSComctlLib.ListView lw 
         Height          =   1575
         Index           =   6
         Left            =   6960
         TabIndex        =   60
         Top             =   600
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   2778
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Cuota"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripcion"
            Object.Width           =   4304
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Importe"
            Object.Width           =   1764
         EndProperty
      End
      Begin VB.Label lblAzira 
         Caption         =   "l"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   6960
         TabIndex        =   73
         Top             =   8280
         Width           =   5295
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Forzar operadora"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   10
         Left            =   6960
         TabIndex        =   72
         Top             =   4440
         Width           =   1710
      End
      Begin VB.Image Imga 
         Height          =   240
         Index           =   3
         Left            =   9120
         Picture         =   "frmListado4.frx":39EE
         ToolTipText     =   "Eliminar cuota"
         Top             =   240
         Width           =   240
      End
      Begin VB.Image Imga 
         Height          =   240
         Index           =   2
         Left            =   8760
         Picture         =   "frmListado4.frx":43F0
         ToolTipText     =   "Añadir cuota"
         Top             =   240
         Width           =   240
      End
      Begin VB.Image Imga 
         Height          =   240
         Index           =   0
         Left            =   2040
         Picture         =   "frmListado4.frx":4DF2
         ToolTipText     =   "Añadir telefono"
         Top             =   240
         Width           =   240
      End
      Begin VB.Image Imga 
         Height          =   240
         Index           =   1
         Left            =   2400
         Picture         =   "frmListado4.frx":57F4
         ToolTipText     =   "Eliminar telefono"
         Top             =   240
         Width           =   240
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Cuotas a aplicar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   9
         Left            =   6960
         TabIndex        =   62
         Top             =   240
         Width           =   1710
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Teléfono / socio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   8
         Left            =   120
         TabIndex        =   61
         Top             =   240
         Width           =   1710
      End
   End
   Begin VB.Frame FramePedidoArticulos 
      Height          =   6615
      Left            =   120
      TabIndex        =   20
      Top             =   0
      Visible         =   0   'False
      Width           =   10095
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Regresar"
         Height          =   375
         Index           =   5
         Left            =   8760
         TabIndex        =   23
         Top             =   6120
         Width           =   1095
      End
      Begin MSComctlLib.ListView lw 
         Height          =   5175
         Index           =   1
         Left            =   120
         TabIndex        =   22
         Top             =   720
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   9128
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Articulo"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripcion"
            Object.Width           =   6527
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Uds"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "precio"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Dto1"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Dto2"
            Object.Width           =   1587
         EndProperty
      End
      Begin VB.Image imgAyuda 
         Height          =   255
         Index           =   0
         Left            =   1080
         ToolTipText     =   "Precio normal/eco"
         Top             =   6240
         Width           =   255
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   3
         Left            =   120
         Picture         =   "frmListado4.frx":61F6
         Top             =   6240
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   2
         Left            =   600
         Picture         =   "frmListado4.frx":6340
         Top             =   6240
         Width           =   240
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Articulos varios. Asignar tipo precio (Norma/Eco)"
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
         Index           =   1
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   7635
      End
   End
   Begin VB.Frame FrameCambioPassword 
      Height          =   3615
      Left            =   960
      TabIndex        =   8
      Top             =   3120
      Visible         =   0   'False
      Width           =   5895
      Begin VB.CommandButton cmdCambiarPasswd 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3360
         TabIndex        =   13
         Top             =   3000
         Width           =   975
      End
      Begin VB.CheckBox chkPass 
         Alignment       =   1  'Right Justify
         Caption         =   "Mostrar"
         Height          =   255
         Left            =   3960
         TabIndex        =   19
         Top             =   2280
         Width           =   1455
      End
      Begin VB.TextBox txtPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   3960
         PasswordChar    =   "*"
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   1920
         Width           =   1455
      End
      Begin VB.TextBox txtPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   3960
         PasswordChar    =   "*"
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   1560
         Width           =   1455
      End
      Begin VB.TextBox txtPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   960
         PasswordChar    =   "*"
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   1560
         Width           =   1455
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Salir"
         Height          =   375
         Index           =   4
         Left            =   4560
         TabIndex        =   14
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Confirmar"
         Height          =   255
         Index           =   2
         Left            =   2760
         TabIndex        =   18
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Nuevo"
         Height          =   255
         Index           =   1
         Left            =   2760
         TabIndex        =   17
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Antiguo"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   16
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label lblNombreUsu 
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   1080
         Width           =   5175
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Cambio password"
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
         Left            =   1320
         TabIndex        =   9
         Top             =   360
         Width           =   3075
      End
   End
   Begin VB.Frame FrameEscandallo 
      Height          =   6495
      Left            =   1200
      TabIndex        =   4
      Top             =   0
      Width           =   9375
      Begin VB.TextBox txtEscandallo 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5175
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Text            =   "frmListado4.frx":648A
         Top             =   600
         Width           =   8775
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Salir"
         Height          =   375
         Index           =   1
         Left            =   7920
         TabIndex        =   5
         Top             =   6000
         Width           =   1095
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Pertenece a los conjuntos:"
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
         Index           =   18
         Left            =   240
         TabIndex        =   7
         Top             =   120
         Width           =   8715
      End
   End
   Begin VB.Frame FramePMP 
      Height          =   6615
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   10215
      Begin VB.CommandButton cmdActualizaPMP 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   7560
         TabIndex        =   3
         Top             =   6000
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   0
         Left            =   8880
         TabIndex        =   1
         Top             =   6000
         Width           =   1095
      End
      Begin MSComctlLib.ListView lw 
         Height          =   5535
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   9763
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Articulo"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripcion"
            Object.Width           =   6527
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Prov"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Fam"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Precio"
            Object.Width           =   1941
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Calculado"
            Object.Width           =   1941
         EndProperty
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   1
         Left            =   600
         Picture         =   "frmListado4.frx":6490
         Top             =   6000
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   0
         Left            =   240
         Picture         =   "frmListado4.frx":65DA
         Top             =   6000
         Width           =   240
      End
   End
   Begin VB.Frame FrameListadoAlzira 
      Height          =   7575
      Left            =   120
      TabIndex        =   32
      Top             =   0
      Visible         =   0   'False
      Width           =   9375
      Begin VB.ComboBox cboEjercicio 
         Height          =   315
         Left            =   6960
         Style           =   2  'Dropdown List
         TabIndex        =   42
         Top             =   240
         Width           =   2175
      End
      Begin VB.CommandButton cmdtreeview1 
         Height          =   375
         Index           =   3
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Eliminar familia"
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdtreeview1 
         Height          =   375
         Index           =   2
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Eliminar familia"
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdtreeview1 
         Height          =   375
         Index           =   1
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Añadir familia"
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdtreeview1 
         Height          =   375
         Index           =   0
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Cambiar texto"
         Top             =   240
         Width           =   375
      End
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   6135
         Left            =   240
         TabIndex        =   37
         Top             =   720
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   10821
         _Version        =   393217
         Style           =   7
         Appearance      =   1
      End
      Begin VB.CommandButton cmdVtentasAlzira 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   6960
         TabIndex        =   35
         Top             =   7080
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Salir"
         Height          =   375
         Index           =   7
         Left            =   8160
         TabIndex        =   34
         Top             =   7080
         Width           =   1095
      End
      Begin VB.Label lblTitulo 
         Alignment       =   1  'Right Justify
         Caption         =   "Ej"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   5
         Left            =   5760
         TabIndex        =   43
         Top             =   240
         Width           =   1050
      End
      Begin VB.Label lblAzira 
         Caption         =   "l"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   36
         Top             =   7200
         Width           =   5295
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Datos ventas"
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
         Index           =   4
         Left            =   240
         TabIndex        =   33
         Top             =   240
         Width           =   2265
      End
   End
   Begin VB.Frame FrameFAS 
      Height          =   9255
      Left            =   120
      TabIndex        =   52
      Top             =   0
      Visible         =   0   'False
      Width           =   12255
      Begin VB.CommandButton cmdGenerarFAS 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   9720
         TabIndex        =   56
         Top             =   8640
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   9
         Left            =   10920
         TabIndex        =   54
         Top             =   8640
         Width           =   1095
      End
      Begin MSComctlLib.ListView lw 
         Height          =   7695
         Index           =   4
         Left            =   240
         TabIndex        =   53
         Top             =   480
         Width           =   11775
         _ExtentX        =   20770
         _ExtentY        =   13573
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
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nº Presu"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Fecha"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "FechaOCULTA"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Codigo"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Nombre"
            Object.Width           =   5909
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Bruto"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "% IVA"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Imp. IVA"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "Total"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   7
         Left            =   600
         Picture         =   "frmListado4.frx":6724
         Top             =   8280
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   6
         Left            =   240
         Picture         =   "frmListado4.frx":686E
         Top             =   8280
         Width           =   240
      End
      Begin VB.Label lblTitulo 
         Alignment       =   1  'Right Justify
         Caption         =   "Leyendo B.D."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   7
         Left            =   8640
         TabIndex        =   55
         Top             =   8280
         Width           =   3330
      End
   End
   Begin VB.Frame FrameControlDirEnv 
      Height          =   9615
      Left            =   -3600
      TabIndex        =   75
      Top             =   0
      Visible         =   0   'False
      Width           =   15735
      Begin VB.CommandButton cmdCancelar 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   375
         Index           =   11
         Left            =   14400
         TabIndex        =   76
         Top             =   9120
         Width           =   1095
      End
      Begin MSComctlLib.ListView lw 
         Height          =   8055
         Index           =   7
         Left            =   120
         TabIndex        =   77
         Top             =   840
         Width           =   15375
         _ExtentX        =   27120
         _ExtentY        =   14208
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   10
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "TipoAlb"
            Object.Width           =   1323
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "NºAlb"
            Object.Width           =   1773
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "F.Alb."
            Object.Width           =   1852
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Cli."
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Nombre cliente"
            Object.Width           =   6085
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Dir."
            Object.Width           =   1535
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Direccion de envio"
            Object.Width           =   5927
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Tipo"
            Object.Width           =   1453
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "NºFactura"
            Object.Width           =   2143
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   9
            Text            =   "Fecha"
            Object.Width           =   1852
         EndProperty
      End
      Begin VB.Label lblTitulo 
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
         Index           =   11
         Left            =   120
         TabIndex        =   78
         Top             =   360
         Width           =   7635
      End
   End
   Begin VB.Frame FrameTelefono 
      Height          =   8535
      Left            =   120
      TabIndex        =   44
      Top             =   120
      Visible         =   0   'False
      Width           =   11415
      Begin VB.ComboBox cboTelefono 
         Height          =   315
         Left            =   5880
         Style           =   2  'Dropdown List
         TabIndex        =   50
         Top             =   480
         Width           =   4215
      End
      Begin VB.OptionButton optTelefono 
         Caption         =   "Varios"
         Height          =   255
         Index           =   2
         Left            =   3480
         TabIndex        =   49
         Top             =   480
         Width           =   855
      End
      Begin VB.OptionButton optTelefono 
         Caption         =   "Cuotas"
         Height          =   255
         Index           =   1
         Left            =   2160
         TabIndex        =   48
         Top             =   480
         Width           =   1455
      End
      Begin VB.OptionButton optTelefono 
         Caption         =   "Detalle llamada"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   47
         Top             =   480
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Salir"
         Height          =   375
         Index           =   8
         Left            =   10320
         TabIndex        =   45
         Top             =   480
         Width           =   855
      End
      Begin MSComctlLib.ListView lw 
         Height          =   6855
         Index           =   3
         Left            =   240
         TabIndex        =   46
         Top             =   1080
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   12091
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Telefono"
            Object.Width           =   1852
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Tipo"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Descripcion"
            Object.Width           =   5981
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Fecha"
            Object.Width           =   1325
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Hora"
            Object.Width           =   1340
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Uds"
            Object.Width           =   1060
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Cantidad"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Importe"
            Object.Width           =   1834
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "Libre"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lblTitulo 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   13
         Left            =   6570
         TabIndex        =   80
         Top             =   8040
         Width           =   2850
      End
      Begin VB.Label lblTitulo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   12
         Left            =   240
         TabIndex        =   79
         Top             =   8040
         Width           =   2850
      End
      Begin VB.Label lblTitulo 
         Alignment       =   1  'Right Justify
         Caption         =   "Concepto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   6
         Left            =   4560
         TabIndex        =   51
         Top             =   480
         Width           =   1170
      End
   End
End
Attribute VB_Name = "frmListado4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Opcion As Byte
    '   0.- Precio medio ponderado
    '   1.- Articulo es escandallo de otros
    '   2.- Precio ultima compra (RECALCULO)

    '   4.- Cambio password
    '   5.- Articulos de varios de pedido y a albaran. Cuales son precio normal o ECO
    '   6.- Tipo precio factura (para poder modificarlo)
    
    '   7.- Listado ventas ALZIRA
    '   8.- Telefonia.  Dado un fichero, ver conceptos de BD telefonos
    
    '   9.- Pase de presu a FAS  HERBELCA
    '   10.- Modificacion masiva cuotas telefonia
    '
    '   11.- Cavevinum. Control de albaranes
    '   12.- UPDATEAR familia marca
    
    '   13.- Cuotas varias AGUA-
    
    '   14.- Multimpresion albaranes
    
    '   15.- Updatear proveedor. Es como el 12 pero updatea tooooooooodas las familias marcas del proveedor
    
    
Public vCadena As String

Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1

Dim PrimVez As Boolean
Dim SQL As String



Dim CadParam As String
Dim numParam As Integer

Dim IT As ListItem

Private Sub cboTelefono_Click()
    If Me.cboTelefono.ListIndex < 0 Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    CargarDatosTelefonia
    Screen.MousePointer = vbDefault
End Sub

Private Sub cboTipoPrecio_Change()
    
End Sub

Private Sub cboTipoPrecio2_Click()
    If PrimVez Then Exit Sub
    
    If cboTipoPrecio2.ListIndex < 0 Then Exit Sub
    
    
    
    txtDecimal(1).Text = Format(Val(Me.cboTipoPrecio2.ItemData(cboTipoPrecio2.ListIndex)) / 100, FormatoCantidad)
    
End Sub

Private Sub chkPass_Click()
    If Me.chkPass.Value = 0 Then
        txtPassword(2).PasswordChar = "*"
    Else
        txtPassword(2).PasswordChar = ""
    End If
End Sub

Private Sub cmdActualizaPMP_Click()
    SQL = ""
    For NumRegElim = 1 To Me.lw(0).ListItems.Count
        If Me.lw(0).ListItems(NumRegElim).Checked Then SQL = SQL & "X"
    Next NumRegElim
    
    
    If SQL = "" Then
        MsgBox "Seleccione algún articulo para actualizar", vbExclamation
        Exit Sub
    End If
    
    
    SQL = "Va a actualizar " & Len(SQL) & " referencia(s)"
    SQL = SQL & vbCrLf & vbCrLf & "¿Continuar?"
    If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    
    Screen.MousePointer = vbHourglass
    ActualizarReferencias
    Screen.MousePointer = vbDefault
    If SQL = "" Then
        CadenaDesdeOtroForm = "OK"
        Unload Me  'ha ido bien
    End If
    
End Sub

Private Sub cmdActualizaSoloProveedor_Click()

    On Error GoTo ecmdActualizaSoloProveedor
    CadenaDesdeOtroForm = "Origen    " & Me.Label2(3).Caption & " " & Me.Label2(5).Caption & vbCrLf
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & "Destino   " & Me.Label2(6).Caption & " " & Me.Label2(8).Caption
    SQL = "Va a cambiar en la BD el proveedor:"
    SQL = SQL & vbCrLf & CadenaDesdeOtroForm
    SQL = SQL & vbCrLf & "¿Continuar?"
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    
    Set miRsAux = New ADODB.Recordset
   
    Label2(10).Caption = "Articulo(1/5)"
    Label2(10).Refresh
    
    CadParam = "UPDATE sartic set codprove =" & Label2(6).Caption
    CadParam = CadParam & " WHERE codprove=" & Label2(3).Caption
    conn.Execute CadParam
    Espera 0.8
    
    conn.Execute "SET FOREIGN_KEY_CHECKS=0;"
    Label2(10).Caption = "Precios(2/5)"
    Label2(10).Refresh
    CadParam = Replace(CadParam, "sartic", "slispr")
    conn.Execute CadParam
    Label2(10).Caption = "Hco precios(3/5)"
    Label2(10).Refresh
    CadParam = Replace(CadParam, "slispr", "slisp1")
    conn.Execute CadParam
    
    Label2(10).Caption = "Dto proveedor(41/5)"
    Label2(10).Refresh
    CadParam = Replace(CadParam, "slisp1", "sdtomp")
    conn.Execute CadParam
    
    Label2(10).Caption = "Familias(5/5)"
    Label2(10).Refresh
    CadParam = Replace(CadParam, "sdtomp", "sfamia")
    conn.Execute CadParam
    
    
    conn.Execute "SET FOREIGN_KEY_CHECKS=1;"

    
    
    
    Set LOG = New cLOG
    
     Label2(10).Caption = ""
    LOG.Insertar 32, vUsu, CadenaDesdeOtroForm
    
    Screen.MousePointer = vbDefault
    CadenaDesdeOtroForm = ""
    Unload Me
    
    
ecmdActualizaSoloProveedor:
    If Err.Number <> 0 Then
        MuestraError Err.Number, , Err.Description
         Label2(10).Caption = ""
    End If
    Set miRsAux = Nothing
    Set LOG = Nothing
   
End Sub

Private Sub cmdAgua_Click()
    If Me.lw(9).ListItems.Count < 1 Then Exit Sub
    
    If Me.optAgua2(0).Value Then
        If Text1.Text = "" Or txtDecimal(0).Text = "" Then
            MsgBox "Campos obligados", vbExclamation
            Exit Sub
        End If
    End If
    
    If Me.optAgua2(1).Value Then
        SQL = "quitar la facturacion de la cuota de varios."
    Else
        SQL = "añadir a la facturacion la cuota:"
        SQL = SQL & vbCrLf & "Cuota: " & Text1.Text
        SQL = SQL & vbCrLf & "Importe: " & Me.txtDecimal(0).Text
        
    End If
    SQL = "Va a " & SQL & vbCrLf & "Contadores: " & Me.lw(9).ListItems.Count & vbCrLf & vbCrLf
    SQL = SQL & "¿Continuar?"
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    
    
    
    For NumRegElim = 1 To lw(9).ListItems.Count
        
        If Me.optAgua2(1).Value Then
            'QUITAR FACTURACION
            SQL = "0"
        Else
            SQL = "1, descripcion=" & DBSet(Text1.Text, "T") & ", importeconcepto=" & DBSet(txtDecimal(0).Text, "N")
        End If
        SQL = "UPDATE aguacontadoresconce set facturar=" & SQL & " WHERE aguacontadoresconce.codconceAg= 7 "
        SQL = SQL & " and contador=" & DBSet(lw(9).ListItems(NumRegElim).Text, "T")
        conn.Execute SQL
    Next
    
End Sub

Private Sub cmdCambiarPasswd_Click()
    SQL = ""
    For NumRegElim = 0 To 2
        txtPassword(NumRegElim).Text = Trim(txtPassword(NumRegElim).Text)
        If txtPassword(NumRegElim).Text = "" Then
            SQL = "1"
            Exit For
        End If
    Next
    If SQL <> "" Then
        MsgBox "Campos obligatorios", vbExclamation
        PonerFoco txtPassword(NumRegElim)
        Exit Sub
    End If
    
    
    If txtPassword(1).Text <> txtPassword(2).Text Then
        MsgBox "No coincide el nuevo password", vbExclamation
        Exit Sub
    End If
    
    If Me.txtPassword(0).Text <> vUsu.PasswdPROPIO Then
        MsgBox "Error en el password actual", vbExclamation
        Exit Sub
    End If
    
    
    If MsgBox("Desea cambiar el password para las aplicaciones ARIADNA SOFTWARE?", vbQuestion + vbYesNo) = vbYes Then
        NumRegElim = (vUsu.codigo Mod 1000)
        SQL = "UPDATE usuarios.usuarios SET passwordpropio=" & DBSet(txtPassword(2).Text, "T")
        SQL = SQL & " WHERE codusu = " & NumRegElim
        conn.Execute SQL
        
        vUsu.PasswdPROPIO = Me.txtPassword(2).Text
        
        Unload Me
        
    End If
    
    
End Sub

Private Sub cmdCancelar_Click(Index As Integer)
    If Index = 0 Then CadenaDesdeOtroForm = ""
    If Index = 5 Then AsignarValoresLineasPedidoPrecioECO
    
    Unload Me
End Sub



Private Sub cmdCutoasMasivas_Click()
    SQL = ""
    If lw(5).ListItems.Count = 0 Then SQL = "-Telefonos"
    If SQL <> "" Then
        MsgBox "Debe insertar: " & vbCrLf & SQL, vbExclamation
        Exit Sub
    End If
    
    
    If lw(6).ListItems.Count = 0 Then
        SQL = "Va a ELIMINAR(borrar) las cuotas para los telefonos seleccionados"
    
    Else
    
        SQL = "Va a generar las cutoas selecciondas para los " & Me.lw(5).ListItems.Count & " telefono(s)"
        If Me.cboOperadora.ListIndex > 0 Then SQL = SQL & vbCrLf & vbCrLf & "***  Va a forzar la operadora a : " & Me.cboOperadora.List(Me.cboOperadora.ListIndex)
    
    End If
    
    SQL = SQL & vbCrLf & vbCrLf & "¿Continuar?"
    If MsgBox(SQL, vbQuestion + vbYesNo) <> vbYes Then Exit Sub
    
    
    Screen.MousePointer = vbHourglass
    
    
    hacerProcesoCuotas lw(6).ListItems.Count = 0
    
    Me.lblAzira(0).Caption = "" 'indicador de proceso
    Screen.MousePointer = vbDefault
    
    
End Sub

Private Sub cmdFraTipoPrecio_Click()
    
    If cboTipoPrecio2.ListIndex < 0 Then Exit Sub
    If Me.txtDecimal(1).Text = "" Then Exit Sub
    
    
    SQL = ""
    For NumRegElim = 1 To Me.lw(2).ListItems.Count
        If Me.lw(2).ListItems(NumRegElim).Checked Then SQL = SQL & "X"
    Next NumRegElim
    
    
    If SQL = "" Then Exit Sub
    SQL = Len(SQL) & " vencimiento(s)"
    SQL = "Va a modificar la comision para " & SQL & vbCrLf & "¿Continuar?"
    If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
        
    Screen.MousePointer = vbHourglass
    lblTitulo(3).Caption = "Actualizar BD....."
    lblTitulo(3).Refresh
    
    
    
    For NumRegElim = 1 To Me.lw(2).ListItems.Count
        '(codtipom,numfactu,fecfactu,codtipoa,numalbar,numlinea)
        If lw(2).ListItems(NumRegElim).Checked Then
        
            SQL = "UPDATE slifac set pvpInferior =" & cboTipoPrecio2.ListIndex
            SQL = SQL & ", comisionagente =" & DBSet(txtDecimal(1).Text, "N", "N")
            SQL = SQL & " Where " & Me.lw(2).Tag & " AND "
            SQL = SQL & "(codtipoa,numalbar,numlinea) IN (" & lw(2).ListItems(NumRegElim).Tag & ")"
            
            conn.Execute SQL
        
            lw(2).ListItems(NumRegElim).SubItems(8) = txtDecimal(1).Text
            If cboTipoPrecio2.ListIndex = 0 Then
                lw(2).ListItems(NumRegElim).SubItems(9) = " "
                lw(2).ListItems(NumRegElim).ListSubItems(9).Bold = False
                lw(2).ListItems(NumRegElim).ListSubItems(9).ForeColor = vbBlack
            ElseIf cboTipoPrecio2.ListIndex = 2 Then
                lw(2).ListItems(NumRegElim).SubItems(9) = "S"
                lw(2).ListItems(NumRegElim).ListSubItems(9).Bold = True
                lw(2).ListItems(NumRegElim).ListSubItems(9).ForeColor = vbRed
            Else
                lw(2).ListItems(NumRegElim).SubItems(9) = "e"
                lw(2).ListItems(NumRegElim).ListSubItems(9).Bold = False
                lw(2).ListItems(NumRegElim).ListSubItems(9).ForeColor = vbBlue
            End If
        
    
        End If
        
    Next
    
    
    
    lblTitulo(3).Caption = ""
    
    Screen.MousePointer = vbDefault
            
  
    
End Sub



Private Sub cmdGenerarFAS_Click()
Dim cTipo As CTiposMov
Dim vCli As New CCliente

    If lblTitulo(7).Tag = 0 Then Exit Sub
       
    numParam = 0
    For NumRegElim = 1 To Me.lw(4).ListItems.Count
        If Me.lw(4).ListItems(NumRegElim).Checked Then numParam = numParam + 1
    Next
    If numParam = 0 Then
        
    Else
        Set vCli = New CCliente
        CadParam = RecuperaValor(Me.vCadena, 1)
        SQL = ""
        If Not vCli.LeerDatos(CadParam) Then
            SQL = "N"
        Else
            If vCli.ClienteBloqueado(False) Then SQL = "N"
        End If
        If SQL <> "" Then
            Set vCli = Nothing
            Exit Sub
        End If
        Set cTipo = New CTiposMov
        If cTipo.Leer("FAS") Then
            If cTipo.ConseguirContador("FAS") Then
                CadParam = "Va a crear " & numParam & " FAS por un importe total de : " & lblTitulo(7).Caption & vbCrLf & "¿Continuar?"
                If MsgBox(CadParam, vbQuestion + vbYesNo) = vbYes Then
                    
                                    
                    Screen.MousePointer = vbHourglass
                    For NumRegElim = Me.lw(4).ListItems.Count To 1 Step -1
                        If Me.lw(4).ListItems(NumRegElim).Checked Then
                            lblTitulo(7).Caption = Me.lw(4).ListItems(NumRegElim).Text & " - " & Me.lw(4).ListItems(NumRegElim).SubItems(1)
                            lblTitulo(7).Refresh
                            
                            If CambiarFactura_A_FAS(cTipo.Contador, vCli) Then
                                
                                cTipo.IncrementarContador cTipo.TipoMovimiento
                                lw(4).ListItems.Remove lw(4).ListItems(NumRegElim).Index
                            Else
                                lw(4).ListItems(NumRegElim).ForeColor = vbRed
                                lw(4).ListItems(NumRegElim).Checked = False
                            End If
                            
                        End If
                    Next
                    ejecutar "SET FOREIGN_KEY_CHECKS=1;", True
                    
                    Screen.MousePointer = vbDefault
                    lblTitulo(7).Caption = "0,00"
                    lblTitulo(7).Tag = CCur("0,0")
                    
                End If
        
            End If
        End If
        Set cTipo = Nothing
        Set vCli = Nothing
        
        
    End If
            
    
    
    


End Sub

Private Sub cmdMultialbaran_Click()
Dim indRPT  As Byte
Dim J As Integer
Dim CadParam As String

Dim EsAlbaranDeRuta As Boolean

    pPdfRpt = ""
    For J = 1 To Me.lw(10).ListItems.Count
        If lw(10).ListItems(J).Checked Then pPdfRpt = pPdfRpt & "X"
    Next
    If pPdfRpt = "" Then Exit Sub


    CadParam = ""   'CADPARAM
    SQL = ""        'nomDocu
    CadenaDesdeOtroForm = ""  'cadformula
    numParam = 0
    
    
     EsAlbaranDeRuta = True
    
    
    
    
    If Not EsAlbaranDeRuta Then
                            'ALBARANES
                '        If hcoCodTipoM = "ALZ" Then
                '            indRPT = 29   'Albaranes B
                '        ElseIf hcoCodTipoM = "ALR" Then
                '            indRPT = 36
                '        ElseIf hcoCodTipoM = "ALS" Then
                '            indRPT = 39
                '        ElseIf hcoCodTipoM = "ALI" Then
                '            indRPT = 56
                '        Else
                '            If EsHistorico Then
                '                indRPT = 11 'Hist. Albaranes clientes
                '            Else
                                indRPT = 10 'Albaran Clientes
                '            End If
                '        End If
    Else
        'Albaranes de ruta
        indRPT = 49
       
    End If
    If Not PonerParamRPT2(indRPT, CadParam, CByte(numParam), SQL, pImprimeDirecto, pPdfRpt, pRptvMultiInforme) Then Exit Sub
    frmImprimir.NombrePDF = SQL
    frmImprimir.NombreRPT = SQL
    
    'Añadir el codigo de usuario como parametro para link con tabla Temporal (tmptiposiva) en el Report
    'tabla temporal para el calculo del bruto total para cada tipo de IVA
    CadParam = CadParam & "pCodUsu=" & vUsu.codigo & "|"
    numParam = numParam + 1
    
    'PORTES
    CadParam = CadParam & "vPortes=""" & vParamAplic.ArtPortesN & """|"
    numParam = numParam + 1
    
    'PUNTO VERDE
    CadParam = CadParam & "PuntoVerde=""" & vParamAplic.ArtReciclado & """|"
    numParam = numParam + 1
    
    'FALTA###
    'Si se imprimen importes y/o
    'devuelve = DevuelveDesdeBD(conAri, "albarcon", "sclien", "codclien", Text1(4).Text, "N")
    pPdfRpt = ""
    If pPdfRpt = "" Then pPdfRpt = "0"
    ' 0 "Todo"
    ' 1 "Cantidad y Precio"
    ' 2 "Cantidad"
    CadParam = CadParam & "Albarcon=" & pPdfRpt & "|"
    numParam = numParam + 1
    
    
    'Nombre fichero .rpt a Imprimir
    frmImprimir.SeleccionaRPTCodigo = pRptvMultiInforme

        
        
    '===================================================
    '================= FORMULA =========================
    'Cadena para seleccion Nº de Albaran
    '---------------------------------------------------
    CadenaDesdeOtroForm = ""
    
    
    pPdfRpt = ""
    For J = 1 To Me.lw(10).ListItems.Count
        If lw(10).ListItems(J).Checked Then pPdfRpt = pPdfRpt & ", " & lw(10).ListItems(J).Text
    Next
    pPdfRpt = Mid(pPdfRpt, 2)
    
    pPdfRpt = "{scaalb.codtipom}='ALV' AND ({scaalb.numalbar} IN [" & pPdfRpt & "])"
    If Not AnyadirAFormula(CadenaDesdeOtroForm, pPdfRpt) Then Exit Sub
    
    
    If Not EsAlbaranDeRuta Then
        'Aqui imprimiria los albaranes como si fuera uno a uno
        'De momento no lo ha pedido nadie, pero podria servir
        frmImprimir.NumeroCopias = vParamAplic.NumCop_AlbaranNormal
    Else
    
        'En pPdfRpt tengo el select del rpt
        'lo transformo a MYSQL
        pPdfRpt = Replace(pPdfRpt, "{", "(")
        pPdfRpt = Replace(pPdfRpt, "}", ")")
        pPdfRpt = Replace(pPdfRpt, "[", "(")
        pPdfRpt = Replace(pPdfRpt, "]", ")")
    
        frmImprimir.NumeroCopias = vParamAplic.NumCop_AlbaranRuta
        
        'Impresion modo albaranes ruta
        If Not CargarDatosImprimeAlbaranConTransporte Then Exit Sub
        
    End If
    davidCodtipom = "0"
    CadParam = CadParam & "pTipoIVA=" & davidCodtipom & "|"
    numParam = numParam + 1
    
    davidCodtipom = ""
    
    With frmImprimir
        'Febrero 2010
       
        
        .FormulaSeleccion = CadenaDesdeOtroForm
        .OtrosParametros = CadParam
        .NumeroParametros = numParam
        .SoloImprimir = False
        .EnvioEMail = False
        .Opcion = 45
        
        .Titulo = "Albaran de Cliente(Seleccion)"
        
        .ConSubInforme = True
        .Show vbModal
        
         
    End With
    

    If EsAlbaranDeRuta Then
        If HaPulsadoElBotonDeImprimir Then
            'UPDATEAMOS scaalb para que no reimpimrpima los albaranes
            
            SQL = "UPDATE scaalb SET albImpreso = 1 WHERE " & pPdfRpt
            ejecutar SQL, False
        End If
    End If






End Sub





Private Function CargarDatosImprimeAlbaranConTransporte() As Boolean
    CargarDatosImprimeAlbaranConTransporte = False
    
   
        
    'Para cada albaran pendiente de reeimprimir habra que ver si tiene resto de pedido
    'Si lo tiene cargaremos la tabla
    SQL = "DELETE FROM tmpsliped WHERE codusu = " & vUsu.codigo
    conn.Execute SQL
    
    
    'Para tener un temporal por si se va la luz
    SQL = "DELETE FROM tmpnseries WHERE codusu =" & vUsu.codigo
    conn.Execute SQL
    
    SQL = "INSERT INTO tmpnseries (codusu ,numlinealb,numserie) "
    SQL = SQL & " select " & vUsu.codigo & " ,numalbar,fechaalb from scaalb where " & pPdfRpt
    conn.Execute SQL
    
    

    
    
    
    '
    '**** linkamos POR codzona--> CODDIREN.  pARA NO CREAR MAS CAMPOS EN TMPSLIPED.. En codlamac llevare el coddiren
    '
    SQL = "Select " & vUsu.codigo & ",scaped.numpedcl,numlinea,codartic,nomartic,cantidad,coddiren,codclien FROM scaped,sliped where scaped.numpedcl =sliped.numpedcl"
    SQL = SQL & " AND (scaped.numpedcl,fecpedcl) in "
    SQL = SQL & "( select numpedcl,fecpedcl from scaalb where " & pPdfRpt & ")"
    SQL = "INSERT INTO tmpsliped(codusu, numpedcl, numlinea, codartic, nomartic, cantidad,codzona,codclien) " & SQL
    If ejecutar(SQL, False) Then
        'Pondre a cero la codzona pq si no el rpt no enlaza bien
        SQL = "UPDATE tmpsliped SET codzona = 0 where codusu = " & vUsu.codigo & " AND codzona is null"
        ejecutar SQL, False
        CargarDatosImprimeAlbaranConTransporte = True
    End If
    
End Function
Private Sub cmdTfnia_Click(Index As Integer)
    If Index < 2 Then
        'LEER /guardar
        If Index = 1 Then
            If lw(5).ListItems.Count = 0 Then
                MsgBox "Ningun dato para guardar", vbExclamation
                Exit Sub
            End If
            GuadarFichero
            
        Else
            If lw(5).ListItems.Count > 0 Then
                If MsgBox("Ya existen datos. Eliminar?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
                Me.lw(5).ListItems.Clear
            End If
            LeerFichero
        End If
    Else
        'Imprimir
        If lw(5).ListItems.Count = 0 Then Exit Sub
        
        conn.Execute "Delete from tmpinformes where codusu =" & vUsu.codigo
        SQL = "INSERT INTO tmpinformes(codusu,codigo1,nombre1,campo1,nombre2,nombre3) VALUES "
        CadParam = ""
        For numParam = 1 To lw(5).ListItems.Count
            'tmpinformes(codusu,codigo1,nombre1,nombre2,nombre3)
            CadParam = CadParam & ", (" & vUsu.codigo & "," & numParam & ",'" & lw(5).ListItems(numParam).Text & "',"
            CadParam = CadParam & DBSet(lw(5).ListItems(numParam).SubItems(1), "T") & "," & DBSet(lw(5).ListItems(numParam).SubItems(2), "T") & "," & DBSet(lw(5).ListItems(numParam).SubItems(3), "T") & ") "
        Next
        CadParam = Mid(CadParam, 2)
        SQL = SQL & CadParam
        conn.Execute SQL
        
        InicializarVbles True
        SQL = lw(5).ColumnHeaders(lw(5).SortKey + 1)
        CadParam = CadParam & "Valores=""Orden: " & SQL & """|"
        numParam = numParam + 1
        LlamarImprimir "Mod. cuotas telefonia", "{tmpinformes.codusu}=" & vUsu.codigo, "rTelefModCuota.rpt"
        
    End If
End Sub

Private Sub cmdtreeview1_Click(Index As Integer)

    If Me.TreeView1.SelectedItem Is Nothing Then Exit Sub

    If Index = 0 Then
        'Modificar label identificativo
        If Mid(Me.TreeView1.SelectedItem.Key, 1, 1) = "F" Then
            
        Else
            SQL = InputBox("Etiqueta del grupo", "", TreeView1.SelectedItem.Text)
            If SQL <> "" Then
                If SQL <> TreeView1.SelectedItem.Text Then
                    'OK, a actualizar
                
                    CadParam = "UPDATE sventasalzira SET textocolumn =" & DBSet(SQL, "T") & " WHERE"
                    CadParam = CadParam & " Grupo = " & Mid(TreeView1.SelectedItem.Key, 2, 2)
                    CadParam = CadParam & " AND columna = "
                    If TreeView1.SelectedItem.Parent Is Nothing Then
                        CadParam = CadParam & "0"
                    Else
                        CadParam = CadParam & Mid(TreeView1.SelectedItem.Key, 4, 4)
                    End If
                    If ejecutar(CadParam, False) Then TreeView1.SelectedItem.Text = SQL
                End If
            End If
        End If
        
    Else
        
        If Index = 1 Then
            If Mid(Me.TreeView1.SelectedItem.Key, 1, 1) <> "N" Then
                MsgBox "Seleccione el nodo donde insertar", vbExclamation
            Else
                'aÑadir,modificar familias
                CadenaDesdeOtroForm = ""
                frmVarios3.Opcion = 0
                frmVarios3.Show vbModal
                If CadenaDesdeOtroForm <> "" Then
                    'Pequeña comprobacion
                    
                
                    'Insertamos
                    SQL = "insert into sventasalzira(Grupo,Columna,Familia,Interna,TextoColumn) VALUES ("
                    SQL = SQL & Mid(TreeView1.SelectedItem.Key, 2, 2) '2 el grupo
                    SQL = SQL & "," & Mid(TreeView1.SelectedItem.Key, 4, 4) '4 la columna
                    SQL = SQL & "," & RecuperaValor(CadenaDesdeOtroForm, 1) & "," & RecuperaValor(CadenaDesdeOtroForm, 3) & ",'')"
                    If ejecutar(SQL, False) Then
                        TreeView1.Nodes.Add TreeView1.SelectedItem.Key, tvwChild, "F" & TreeView1.Nodes.Count + 10, RecuperaValor(CadenaDesdeOtroForm, 2)
                        TreeView1.Nodes(TreeView1.Nodes.Count).EnsureVisible
                    End If
                End If
            End If
        Else
            'borrar modificar
            If Mid(Me.TreeView1.SelectedItem.Key, 1, 1) <> "F" Then
            
                MsgBox "No se puede eliminar-modificar datos que no sean familias", vbExclamation
        
            Else
                If Index = 3 Then
                    'eliminar
                    SQL = "Va a eliminar la familia: " & TreeView1.SelectedItem.FullPath & vbCrLf & "¿Continuar?"
                    If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
                        'parent.key: N010002     Grupo:1 columna:2
                        SQL = "Grupo = " & Mid(TreeView1.SelectedItem.Parent.Key, 2, 2)
                        SQL = SQL & " AND columna = " & Mid(TreeView1.SelectedItem.Parent.Key, 4, 4)
                        SQL = SQL & " AND familia = " & Mid(TreeView1.SelectedItem.Text, 1, 4)
                        SQL = "DELETE FROM sventasalzira WHERE " & SQL
                        If ejecutar(SQL, False) Then TreeView1.Nodes.Remove TreeView1.SelectedItem.Index
                    End If
                Else
                    'modificar
                    CadenaDesdeOtroForm = Replace(TreeView1.SelectedItem, " - ", "|")
                    CadenaDesdeOtroForm = CadenaDesdeOtroForm & "|0|"
                    frmVarios3.Opcion = 0
                    frmVarios3.Show vbModal
                    If CadenaDesdeOtroForm <> "" Then numParam = 1
                End If
            End If
        End If
        
        
    End If
    CadParam = "" 'por si vuelve a apretar aceptar
        
End Sub

Private Sub cmdUpdatearFamiliaMarca_Click()
Dim Aux As String

    SQL = ""
    Aux = ""
    For NumRegElim = 1 To Me.lw(8).ListItems.Count
        If Not Me.lw(8).ListItems(NumRegElim).Checked Then
            SQL = SQL & "X"
            lw(8).ListItems(NumRegElim).Checked = True
        End If
    Next NumRegElim

    If SQL <> "" Then MsgBox "Todos los articulos serán seleccionados.", vbInformation
    
    numParam = NumRegElim
    
    SQL = ""
    If numParam > 1 Then SQL = "s"
    
    SQL = "Va a actualizar " & numParam & " articulo" & SQL & " seleccionado" & SQL & "." & vbCrLf
    SQL = SQL & vbCrLf & CadenaDesdeOtroForm
    SQL = SQL & vbCrLf & "¿Continuar?"
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    
    
    

    
    
    
    Screen.MousePointer = vbHourglass
    
    '  LOG de acciones
    Set LOG = New cLOG
    
    
    
    
    numParam = 0
    Do
        
        SQL = CadenaDesdeOtroForm & vbCrLf & "Total: " & lw(8).ListItems.Count & " Actualizar: " & numParam & vbCrLf & "Artic:"
        If Len(CadenaDesdeOtroForm) > 200 Then
            If numParam = 0 Then
                LOG.Insertar 27, vUsu, CadenaDesdeOtroForm & " Sigue secuencia"
                Espera 1
                
            End If
            numParam = numParam + 1
            SQL = "Secuencia:" & numParam & vbCrLf
        Else
            SQL = CadenaDesdeOtroForm
        End If
        
        NumRegElim = Len(Aux) + Len(SQL)
        
        If NumRegElim > 252 Then
            
            NumRegElim = 252 - Len(SQL)
            
            
            
            If Len(Aux) > NumRegElim Then
                SQL = SQL & Mid(Aux, 1, NumRegElim) & "..."
                Aux = Mid(Aux, NumRegElim + 1)
            Else
                SQL = SQL & Aux
                Aux = ""
            End If
        Else
            SQL = SQL & Aux
            Aux = ""
        End If
    
        LOG.Insertar 27, vUsu, SQL
        Espera 1
    
    Loop Until Aux = ""
    'Lo que updateamos
    CadenaDesdeOtroForm = ""
    SQL = RecuperaValor(vCadena, 2)
    If SQL <> "" Then CadenaDesdeOtroForm = CadenaDesdeOtroForm & ", codfamia =" & SQL
    SQL = RecuperaValor(vCadena, 3)
    If SQL <> "" Then CadenaDesdeOtroForm = CadenaDesdeOtroForm & ", codmarca =" & SQL
    SQL = RecuperaValor(vCadena, 4)
    If Val(SQL) > 0 Then CadenaDesdeOtroForm = CadenaDesdeOtroForm & ", codprove=" & SQL
    CadenaDesdeOtroForm = Mid(CadenaDesdeOtroForm, 2) 'quitamos la primera coma
    'montamos el sql
    
    CadenaDesdeOtroForm = "UPDATE sartic set " & CadenaDesdeOtroForm
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & " WHERE codartic = "
    numParam = 0
    SQL = ""
    For NumRegElim = 1 To Me.lw(8).ListItems.Count
        If Me.lw(8).ListItems(NumRegElim).Checked Then ActualizaFamiliaMarca
        
    Next NumRegElim
    
    
    CadParam = RecuperaValor(vCadena, 4)
    If CadParam = "-1" Then CadParam = lw(8).ListItems(1).SubItems(5)
    CadParam = "UPDATE sdtomp SET codprove=" & CadParam
    If RecuperaValor(vCadena, 2) <> "" Then CadParam = CadParam & ", codfamia= " & RecuperaValor(vCadena, 2)
    'HEREBLCA. LA marca no la trato
    'If RecuperaValor(vCadena, 3) <> "" Then CadParam = CadParam & ", codmarca= " & RecuperaValor(vCadena, 3)
    CadParam = CadParam & " WHERE codprove = " & lw(8).ListItems(1).SubItems(5)
    CadParam = CadParam & " AND codfamia = " & lw(8).ListItems(1).SubItems(4)
    
    ejecutar CadParam, False
        
        
    Set LOG = Nothing
    Screen.MousePointer = vbDefault
    Unload Me
    
    
    
End Sub

Private Sub ActualizaFamiliaMarca()
Dim YProveedor As String

        'Numregelim llevo el indice al lw(8) para coger el articulo
        
        'Si queremos insertar LOG
        YProveedor = ""
        SQL = RecuperaValor(vCadena, 4)
        If SQL = "-1" Then SQL = ""
        If SQL <> "" Then
        
            'Esta actualizando el proveedor. Vemos el del articulo
            CadParam = lw(8).ListItems(NumRegElim).SubItems(5)
            If CadParam <> SQL Then YProveedor = CadParam
            
        Else
            YProveedor = lw(8).ListItems(NumRegElim).SubItems(5)
            SQL = YProveedor
        End If
       
        CadParam = CadenaDesdeOtroForm & DBSet(lw(8).ListItems(NumRegElim).Text, "T")
        conn.Execute CadParam
        
        
            conn.Execute "SET FOREIGN_KEY_CHECKS=0;"
            CadParam = " WHERE codprove = " & YProveedor
            CadParam = "UPDATE slispr SET codprove=" & SQL & CadParam
            CadParam = CadParam & " AND codartic = " & DBSet(lw(8).ListItems(NumRegElim).Text, "T")
                
            
            ejecutar CadParam, False
                
            CadParam = Replace(CadParam, "slispr", "slisp1")
            ejecutar CadParam, False
         
            
            conn.Execute "SET FOREIGN_KEY_CHECKS=1;"


        SQL = ""

End Sub



Private Sub cmdVtentasAlzira_Click()
Dim NO As Node
    If CadParam <> "" Then
        If MsgBox("El proceso puede costar mucho tiempo. Continuar?", vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    
    If GenerarDatosVentasAlzira Then
        InicializarVbles True
            
        NumRegElim = 0
        
        SQL = "XX"
        While SQL <> ""
            If Len(SQL) = 2 Then
                'Es el primer bloque
                Set NO = TreeView1.Nodes(1).Child
                SQL = "X"
            Else
                Set NO = TreeView1.Nodes(1)
                Set NO = NO.Next
                Set NO = NO.Child
                SQL = ""
            End If
            
        
            While Not NO Is Nothing
                NumRegElim = NumRegElim + 1
                If NumRegElim <= 9 Then
                    CadParam = CadParam & "C" & NumRegElim & "= """ & NO.Text & """|"
                    numParam = numParam + 1
                End If
                Set NO = NO.Next
            Wend
        
        Wend
        
        NumRegElim = Me.cboEjercicio.ItemData(cboEjercicio.ListIndex)
        If Year(vEmpresa.FechaIni) = Year(vEmpresa.FechaFin) Then
            'Mismo añoa ejercicios
            SQL = NumRegElim & "|" & NumRegElim - 1 & "|"
        Else
            SQL = NumRegElim & "/" & (NumRegElim + 1) - 2000 & "|"
            SQL = SQL & NumRegElim - 1 & "/" & (NumRegElim) - 2000 & "|"
        End If
        CadParam = CadParam & "TextoActual= """ & RecuperaValor(SQL, 1) & """|"
        CadParam = CadParam & "TextoAnterior= """ & RecuperaValor(SQL, 2) & """|"
        numParam = numParam + 2
         
        LlamarImprimir "Ventas familia agrupado", "{tmpinformes.codusu}=" & vUsu.codigo, "rVtasAgrupaFamiliaAlz.rpt"
    End If
    
    Screen.MousePointer = vbDefault
    
End Sub



Private Sub Form_Activate()
    If PrimVez Then
        PrimVez = False
        If Opcion = 0 Or Opcion = 2 Then CargaLwPrecioMP
        If Opcion = 1 Then CarArticulos
        
        If Opcion = 5 Then CargarLineasPedidoVarios
        If Opcion = 6 Then CargarLineasFacturaTipoPrecio
        If Opcion = 7 Then CargarTreevieVentasAlzira
        If Opcion = 8 Then
            optTelefono_Click 0
            Set miRsAux = New ADODB.Recordset
        End If
        
        If Opcion = 9 Then CargaDatosFAZ
        If Opcion = 11 Then CargaControlDireccionesEnvio
        If Opcion = 12 Then CargarArticulosCambio
        If Opcion = 14 Then CargaAlbaranes True
        
        If Opcion = 15 Then PonerCamposCambioProveedor
    End If
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Form_Load()
Dim H As Integer
Dim W As Integer
Dim indice As Integer


    CargaIconosAyuda
    
    PrimVez = True
    Me.Icon = frmPpal.Icon
    FramePMP.visible = False
    FrameEscandallo.visible = False
    FrameCambioPassword.visible = False
    FramePedidoArticulos.visible = False
    FrameTipoprecFac.visible = False
    FrameListadoAlzira.visible = False
    FrameFAS.visible = False
    FrameAjuteCuotasTfnia.visible = False
    FrameControlDirEnv.visible = False
    FrameCambioFamMarca.visible = False
    FrameAgua.visible = False
    FrameUpdateaSoloProveedor.visible = False
    indice = Opcion
    If Opcion >= 4 Then limpiar Me
    Select Case Opcion
    Case 0, 2
        'Ajuste precio
        '   0: PrecioMP
        '   2: PrecioMP
        
        PonerFrameVisible FramePMP, H, W
        If Opcion = 0 Then
            Caption = "Actualizar precio medio ponderado"
            
        Else
            Caption = "Actualizar último precio compra"
            
        End If
        indice = 0 'el cmdcancelar
        CadenaDesdeOtroForm = ""
    Case 1
        
        PonerFrameVisible FrameEscandallo, H, W
        Caption = "Ver articulos"
    Case 4
        PonerFrameVisible FrameCambioPassword, H, W
        Caption = "Ariadna Software"
        lblNombreUsu.Caption = vUsu.Nombre
        
    Case 5
        PonerFrameVisible FramePedidoArticulos, H, W
        Caption = "Pase de pedido..."
        
    Case 6
        PonerFrameVisible FrameTipoprecFac, H, W
        Caption = "Factura"
        lblTitulo(3).Caption = RecuperaValor(Me.vCadena, 1)
        CargaComboTipoComision
'        cboTipoPrecio2.ListIndex = 0
        
    Case 7
        'Ventas alzira
        PonerFrameVisible FrameListadoAlzira, H, W
        
        'lblAzira(0).Caption = "Mostrar los datos de ventas anuales(comparativos)" & vbCrLf
        'lblAzira(0).Caption = lblAzira(0).Caption & "Por ventas o internas y agrupado por familias"
        lblAzira(1).Caption = ""
        
        cmdtreeview1(0).Picture = frmPpal.imgListComun.ListImages(12).Picture
        cmdtreeview1(1).Picture = frmPpal.imgListComun.ListImages(3).Picture
        cmdtreeview1(2).Picture = frmPpal.imgListComun.ListImages(4).Picture
        cmdtreeview1(3).Picture = frmPpal.imgListComun.ListImages(43).Picture
        
        CargaComboEjercicio
       
    Case 8
        Set miRsAux = New ADODB.Recordset
        PonerFrameVisible FrameTelefono, H, W
       
       
    Case 9
        PonerFrameVisible FrameFAS, H, W
       
    Case 10
        PonerFrameVisible FrameAjuteCuotasTfnia, H, W
        '                                                                       quitamos vodafone
        CargarCombo_Tabla cboOperadora, "stfnooperador", "codoperador", "nombre", "codoperador<4", True
        Me.lblAzira(0).Caption = "" 'indicador de proceso
        
    Case 11
        PonerFrameVisible FrameControlDirEnv, H, W
    Case 12
    
        PonerFrameVisible FrameCambioFamMarca, H, W
        lblTitulo(19).visible = False
        lblTitulo(19).Tag = 0
    Case 13
        PonerFrameVisible FrameAgua, H, W
        Frame1.BorderStyle = 0
        Text1.Text = "Varios"
        
    Case 14
        PonerFrameVisible frameMultiAlbaranes, H, W
        Me.cboTipoIVA.Clear
        Caption = "Impresión albaranes"
    Case 15
    
        PonerFrameVisible FrameUpdateaSoloProveedor, H, W
        Caption = "Actualizar proveedor"
    End Select
    
    Height = H + 150
    Me.Width = W
    cmdCancelar(CInt(indice)).Cancel = True
    
End Sub


'Dado un FRAME lo pone a true y lo situa en x:120 y:0 y devuelve lo que debe medir el form
Private Sub PonerFrameVisible(ByRef F As Frame, ByRef CH As Integer, CW As Integer)
    F.Top = 0
    F.Left = 120
    F.visible = True
    CH = F.Height + 420
    CW = F.Width + 240
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Opcion = 8 Then Set miRsAux = Nothing
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
    SQL = CadenaDevuelta
End Sub

Private Sub Imga_Click(Index As Integer)
    Select Case Index
    Case 0
        LanzaBuscaGrid 0
    
    Case 1, 3, 5
        numParam = 5
        If Index = 3 Then numParam = 6
        If Index = 5 Then numParam = 9
            
        If lw(numParam).SelectedItem Is Nothing Then
            MsgBox "Seleccione un dato  a borrar", vbExclamation
            Exit Sub
        End If
        
        SQL = lw(numParam).SelectedItem.Text & " " & lw(numParam).SelectedItem.SubItems(1)
        SQL = "Desea eliminar el elemento seleccionado: " & SQL & "?"
        If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then Exit Sub
        
        
        lw(numParam).ListItems.Remove lw(numParam).SelectedItem.Index
        
        
    Case 2

        CadenaDesdeOtroForm = ""
        frmVarios3.Opcion = 4
        frmVarios3.Show vbModal
        If CadenaDesdeOtroForm <> "" Then AnyadeNodoTelefonia False

    Case 3
    
    
    '-------------------------------------- AGUA
    Case 4
        LanzaBuscaGrid 1
    
    
    
    
    End Select
End Sub

Private Sub imgayuda_Click(Index As Integer)
Dim Ayuda As String

    
    'Sera las ayuda. Tampoco queiero la biblia, pero,
    'si un "pelin" de ayuda no me vendria mal a mi, imaginemos a el cliente final
    Select Case Index
    Case 0
        
        Ayuda = vbCrLf & "Establecerá para las lineas de artículos varios "
        Ayuda = Ayuda & "si el precio es normal o eco"
        Ayuda = Ayuda & vbCrLf & " CHECKED=  ECO"
        
    Case 1
        
        Ayuda = vbCrLf & "Establecerá para los artículos facturados que seleccionemos: " & vbCrLf
        Ayuda = Ayuda & "Tipo comision: Normal / Eco / Supereco "
        Ayuda = Ayuda & vbCrLf & "%Comision:  Que llevara la linea"
    End Select
    
    Ayuda = imgayuda(Index).ToolTipText & vbCrLf & String(45, "=") & vbCrLf & Ayuda
    MsgBox Ayuda, vbInformation



End Sub

Private Sub imgCheck_Click(Index As Integer)
Dim Cual As Byte
Dim Chec As Boolean
Dim Importe As Currency

    If Index < 2 Then
        'trabajadores....
        Cual = 0
        Chec = True
        If (Index Mod 2) = 0 Then Chec = False
    ElseIf Index < 4 Then
        Cual = 1
        Chec = True
        If (Index Mod 2) = 1 Then Chec = False
    ElseIf Index < 6 Then
        Cual = 2
        Chec = True
        If (Index Mod 2) = 0 Then Chec = False
        vCadena = "OK" 'para saber que han cambiado cosas
    ElseIf Index < 8 Then
        '6 7
        Cual = 4
        Chec = (Index Mod 2) = 1
    ElseIf Index < 10 Then
        '8 9
        Cual = 8
        Chec = (Index Mod 2) = 1
    Else
        '10   11
        Cual = 10
        Chec = (Index Mod 2) = 1
    End If
         
    For NumRegElim = 1 To lw(Cual).ListItems.Count
        lw(Cual).ListItems(NumRegElim).Checked = Chec
    Next

    If Cual = 4 Then
        If Not Chec Then
            Importe = 0
        Else
            For NumRegElim = 1 To lw(Cual).ListItems.Count
                Importe = Importe + ImporteFormateado(lw(Cual).ListItems(NumRegElim).SubItems(8))
            Next
        End If
        lblTitulo(7).Tag = Importe
        lblTitulo(7).Caption = Format(lblTitulo(7).Tag, FormatoImporte)
    End If
End Sub


Private Sub CargaLwPrecioMP()
    Set miRsAux = New ADODB.Recordset
    Me.lw(0).ListItems.Clear
    
    SQL = "Select * from tmpinformes where codusu = " & vUsu.codigo & " ORDER BY campo1,nombre1"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set IT = lw(0).ListItems.Add()
        IT.Text = miRsAux!nombre1  'codartic
        IT.SubItems(1) = miRsAux!nombre2 'nomartic
        IT.SubItems(2) = miRsAux!campo1
        IT.SubItems(3) = miRsAux!campo2
        IT.SubItems(4) = Format(miRsAux!importeb1, FormatoPrecio)
        IT.SubItems(5) = Format(miRsAux!importeb2, FormatoPrecio)
        
        
        IT.Checked = False
        
        If miRsAux!importeb1 <> 0 And miRsAux!importeb2 <> 0 Then IT.Checked = True
            
            
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
End Sub



Private Sub lw_ColumnClick(Index As Integer, ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If Index = 5 Then
        'Telefonos para las acciones masivas
        lw(5).SortKey = ColumnHeader.Index - 1
        
        
    ElseIf Index = 7 Then
    
        'For numParam = 1 To lw(7).ColumnHeaders.Count: Debug.Print numParam & " " & (7); lw(7).ColumnHeaders(numParam).Width: Next numParam
    
        If lw(7).SortKey = ColumnHeader.Index - 1 Then
            If lw(7).SortOrder = lvwAscending Then
                lw(7).SortOrder = lvwDescending
            Else
                lw(7).SortOrder = lvwAscending
            End If
        Else
            lw(7).SortKey = ColumnHeader.Index - 1
        End If
    End If
End Sub

Private Sub lw_DblClick(Index As Integer)
    If Me.lw(Index).SelectedItem Is Nothing Then Exit Sub

    If Index = 0 Then
        
        
        frmAlmArticulos.DeConsulta = True
        frmAlmArticulos.DatosADevolverBusqueda = "::" & Me.lw(0).SelectedItem.Text
        frmAlmArticulos.Show vbModal
        
        
    Else
        Screen.MousePointer = vbHourglass
        If Index = 3 Then
        
                SQL = Trim(lw(Index).SelectedItem.Text)
                SQL = "telefono = '" & SQL & "' AND fichero "
                
                SQL = DevuelveDesdeBD(conAri, "concat(serie,'|',ano,'|',numfact,'|')", "tel_cab_factura", SQL, vCadena, "T")
                If Len(SQL) > 3 Then
                    SQL = "serie='" & RecuperaValor(SQL, 1) & "' AND ano =" & RecuperaValor(SQL, 2) & " AND numfact=" & RecuperaValor(SQL, 3)
                    SQL = vCadena & "|" & SQL & "|"
                    frmTelefonoVerFra.TieneAlbaranes = False
                    frmTelefonoVerFra.Where2 = SQL
                    frmTelefonoVerFra.Show vbModal
                    
                    'Vuelvo a poner el new
                    Set miRsAux = Nothing
                    Set miRsAux = New ADODB.Recordset
                End If
                
        ElseIf Index = 2 Then
        
           '
        
        Else
        
            AbrirFormularioDireccionEnvio
        
        End If
        
        Screen.MousePointer = vbDefault
        
    
    End If
End Sub

Private Sub ActualizarReferencias()
Dim HayError As Boolean
    
        
        
    vCadena = ""
    HayError = False
    For NumRegElim = lw(0).ListItems.Count To 1 Step -1
        If lw(0).ListItems(NumRegElim).Checked Then
            
            If Opcion = 0 Then
                SQL = "preciomp"
            Else
                SQL = "preciouc"
            End If
            
            SQL = "UPDATE sartic set " & SQL & " = " & DBSet(lw(0).ListItems(NumRegElim).SubItems(5), "N")
            SQL = SQL & " WHERE codartic = " & DBSet(lw(0).ListItems(NumRegElim).Text, "T")
            If Not ejecutar(SQL, False) Then
                HayError = True
                NumRegElim = Me.lw(0).ListItems.Count + 1
            Else
                vCadena = vCadena & "  ·  " & DBSet(lw(0).ListItems(NumRegElim).Text, "T")
                lw(0).ListItems.Remove lw(0).ListItems(NumRegElim).Index
                
                If Len(vCadena) > 230 Then InsertaLog  'y pone vcdena a ""
                    
            End If
        End If
    Next NumRegElim
    
    If vCadena <> "" Then InsertaLog 'y pone vcdena a ""
    
    'Si llega aqui... tutto benne
    If Not HayError Then SQL = ""
        
End Sub

Private Sub InsertaLog()
        
        '------------------------------------------------------------------------------
        '  LOG de acciones
        Set LOG = New cLOG
        vCadena = Mid(vCadena, 6) 'quitamos el primer separador
        If Opcion = 0 Then
            vCadena = "PMP: " & vCadena
        Else
            vCadena = "UPC: " & vCadena
        End If
        vCadena = Replace(vCadena, "'", "")
        LOG.Insertar 19, vUsu, vCadena
        vCadena = ""
        Set LOG = Nothing
        '-----------------------------------------------------------------------------
        Espera 0.6
End Sub


Private Sub CarArticulos()
Dim Aux As String

    Set miRsAux = New ADODB.Recordset
    txtEscandallo.Text = ""
    While vCadena <> ""
        NumRegElim = InStr(1, vCadena, "|")
        If NumRegElim = 0 Then
            vCadena = ""
        Else
            SQL = Mid(vCadena, 1, NumRegElim - 1)
            vCadena = Mid(vCadena, NumRegElim + 1)
            'Pongo el nombre
            Aux = "Select nomartic from sartic where codartic =" & DBSet(SQL, "T")
            miRsAux.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            Aux = Mid(SQL & Space(16), 1, 16) & " "
            If Not miRsAux.EOF Then Aux = Aux & miRsAux!NomArtic
            miRsAux.Close
            txtEscandallo.Text = txtEscandallo.Text & Aux & vbCrLf
            'los aticulos de los cuales es componente
            Aux = "select sartic.codartic,nomartic from sarti1,sartic where"
            Aux = Aux & " sarti1.codartic=sartic.codartic and codarti1=" & DBSet(SQL, "T") & " ORDER BY 2"
            miRsAux.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not miRsAux.EOF
                Aux = miRsAux!codArtic
                Aux = Space(10) & "- " & Mid(Aux & Space(16), 1, 16) & " " & miRsAux!NomArtic
                txtEscandallo.Text = txtEscandallo.Text & Aux & vbCrLf
                miRsAux.MoveNext
            Wend
            miRsAux.Close
            'Aun quedan mas articulos
            If vCadena <> "" Then txtEscandallo.Text = txtEscandallo.Text & vbCrLf
        End If
    Wend
    Set miRsAux = Nothing
End Sub




Private Sub lw_ItemCheck(Index As Integer, ByVal item As MSComctlLib.ListItem)
Dim Importe As Currency
    If Index = 2 Then
        vCadena = "Cambiado"
    
    ElseIf Index = 4 Then
        'PRESU a FAS
        Importe = ImporteFormateado(item.SubItems(8))
        If Not item.Checked Then Importe = Importe * -1
        
        lblTitulo(7).Tag = lblTitulo(7).Tag + Importe
        lblTitulo(7).Caption = Format(lblTitulo(7).Tag, FormatoImporte)
        
        
    End If
End Sub

Private Sub optAgua_Click(Index As Integer)
    
    lw(9).SortKey = Index
    
End Sub

Private Sub optAgua2_Click(Index As Integer)
    lblTitulo(16).visible = Index = 0
    lblTitulo(17).visible = Index = 0
    Text1.visible = Index = 0
    Me.txtDecimal(0).visible = Index = 0
    
End Sub

Private Sub optTelefono_Click(Index As Integer)
    Me.lw(3).ListItems.Clear
    cboTelefono.Clear

    CargaComboTelef
    
    lblTitulo(12).Caption = ""
    lblTitulo(13).Caption = ""
    
End Sub

Private Sub optTfnia_Click(Index As Integer)
    lw(5).SortKey = Index
End Sub

Private Sub txtDecimal_GotFocus(Index As Integer)
    ConseguirFoco txtDecimal(Index), 3
End Sub

Private Sub txtDecimal_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtDecimal_LostFocus(Index As Integer)
Dim b As Boolean
    txtDecimal(Index).Text = Trim(txtDecimal(Index).Text)
    If txtDecimal(Index).Text <> "" Then
       ' If Index = 0 Or Index = 9 Then
       '     B = PonerFormatoDecimal(txtDecimal(Index), 2)
       ' Else
            b = PonerFormatoDecimal(txtDecimal(Index), 5)
       ' End If
        If b Then

        Else
            txtDecimal(Index).Text = ""
        End If
    End If
End Sub


Private Sub txtPassword_GotFocus(Index As Integer)
    ConseguirFoco txtPassword(Index), 3
End Sub

Private Sub txtPassword_KeyPress(Index As Integer, KeyAscii As Integer)
     KEYpressGnral KeyAscii, 2, True
End Sub


Private Sub CargarLineasPedidoVarios()
    
    Set miRsAux = New ADODB.Recordset
    ' vCadena  TRAE el sql desde los pedidos
    miRsAux.Open vCadena, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        
        Set IT = lw(1).ListItems.Add(, "C" & Format(miRsAux!numlinea, "000000"))
        IT.Text = miRsAux!codArtic
        IT.SubItems(1) = miRsAux!NomArtic 'nomartic
        IT.SubItems(2) = Format(miRsAux.Fields(2), FormatoPorcen)  'cantidad o servidas
        IT.SubItems(3) = Format(miRsAux!precioar, FormatoPrecio)
        IT.SubItems(4) = Format(miRsAux!dtoline1, FormatoPorcen)
        IT.SubItems(5) = Format(miRsAux!dtoline2, FormatoPorcen)
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    
    
    
End Sub


Private Sub AsignarValoresLineasPedidoPrecioECO()
Dim J As Integer

    CadenaDesdeOtroForm = ""
    For J = 1 To Me.lw(1).ListItems.Count
        If Me.lw(1).ListItems(J).Checked Then CadenaDesdeOtroForm = CadenaDesdeOtroForm & ", " & Mid(Me.lw(1).ListItems(J).Key, 2)
    Next
End Sub

Private Sub CargaIconosAyuda()
Dim Ima As Image
    On Error Resume Next 'mejor que no diera errores, pero bien, tampoco vamos a enfadarnos
    For Each Ima In Me.imgayuda
        Ima.Picture = frmPpal.imgListComun.ListImages(46).Picture
    Next
    Err.Clear
End Sub




Private Sub CargarLineasFacturaTipoPrecio()
Dim MasDeUnAlbaran As Boolean
    Set miRsAux = New ADODB.Recordset
    ' vCadena  TRAE la factura
    
    
    'Me guardo el SQL
    Me.lw(2).Tag = RecuperaValor(vCadena, 2)
    SQL = "select * from slifac where " & Me.lw(2).Tag
    SQL = SQL & " order by codtipom,numalbar,numlinea"
    
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    SQL = ""
    MasDeUnAlbaran = False
    While Not miRsAux.EOF
        Set IT = lw(2).ListItems.Add(, "C" & Mid(DBLet(miRsAux!codtipoa, "T"), 1, 3) & Format(miRsAux!NumAlbar, "0000000") & Format(miRsAux!numlinea, "000000"))
        
        IT.Text = DBLet(miRsAux!codtipoa, "T") & Format(miRsAux!NumAlbar, "0000000")
        IT.SubItems(1) = miRsAux!codArtic 'nomartic
        IT.SubItems(2) = miRsAux!NomArtic 'nomartic
        IT.SubItems(3) = Format(miRsAux!cantidad, FormatoPorcen)  'cantidad o servidas
        IT.SubItems(4) = Format(miRsAux!precioar, FormatoPrecio)
        IT.SubItems(5) = Format(miRsAux!dtoline1, FormatoPorcen)
        IT.SubItems(6) = Format(miRsAux!dtoline2, FormatoPorcen)
        IT.SubItems(7) = Format(miRsAux!ImporteL, FormatoImporte)
        
        '(codtipom,numfactu,fecfactu,codtipoa,numalbar,numlinea)
        IT.Tag = "(" & DBSet(miRsAux!codtipoa, "T", "N") & "," & miRsAux!NumAlbar & "," & miRsAux!numlinea & ")"
        
        
        'Diciembre 2014
        IT.SubItems(8) = Format(miRsAux!comisionagente, FormatoPorcen)

        If DBLet(miRsAux!PVPInferior, "N") = 0 Then
            IT.SubItems(9) = " "
        ElseIf miRsAux!PVPInferior = 2 Then
            IT.SubItems(9) = "S"
            IT.ListSubItems(9).Bold = True
            IT.ListSubItems(9).ForeColor = vbRed
        Else
           
            IT.SubItems(9) = "e"
            IT.ListSubItems(9).ForeColor = vbBlue
        End If
            
        'IT.Checked = DBLet(miRsAux!PVPInferior, "N") = 1
        If SQL <> Format(miRsAux!NumAlbar, "000000") Then
            If SQL <> "" Then
                MasDeUnAlbaran = True
                IT.ForeColor = vbBlue
                IT.Bold = True
            End If
            SQL = Format(miRsAux!NumAlbar, "000000")
        End If
        
                    
                    
                    
                    
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    
    If MasDeUnAlbaran Then
        Me.lw(2).ListItems(1).ForeColor = vbBlue
        Me.lw(2).ListItems(1).Bold = True
    End If
    vCadena = "" 'Si la modifica es que ha cambiado valores
End Sub





Private Sub InicializarVbles(AñadireElDeEmpresa As Boolean)
    'cadFormula = ""
    'cadSelect = ""
    CadParam = "|"
    numParam = 0
    
    
    If AñadireElDeEmpresa Then
        CadParam = CadParam & "pEmpresa=""" & vEmpresa.nomempre & """|"
        numParam = numParam + 1
    End If
    
End Sub

Private Sub LlamarImprimir(cadTitulo As String, cadFormula As String, NombreRPT As String)
    With frmImprimir
        .FormulaSeleccion = cadFormula
        .OtrosParametros = CadParam
        .NumeroParametros = numParam

        .SoloImprimir = False
        .EnvioEMail = False
        .Titulo = cadTitulo
        .Opcion = 3000   'VAN TODOS EN ESTE SACO
        .NombrePDF = ""
        .NombrePDF = NombreRPT
        .NombreRPT = NombreRPT
        .ConSubInforme = False 'conSubRPT
        .MostrarTreeDesdeFuera = False ' vMostrarTree
        .Show vbModal
    End With
End Sub


Private Sub CargarTreevieVentasAlzira()
Dim N As Node

    TreeView1.Nodes.Clear

    Set miRsAux = New ADODB.Recordset
    SQL = "Select grupo,TextoColumn from sventasalzira WHERE columna=0 ORDER BY grupo "  'siempre serán 2
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set N = TreeView1.Nodes.Add(, , "G" & miRsAux!Grupo, miRsAux!TextoColumn)
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    SQL = "Select * from sventasalzira WHERE columna>0 ORDER BY grupo,columna,familia"
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    numParam = 0
    While Not miRsAux.EOF
        'El indicador de columna
         If miRsAux!Familia < 0 Then
            'Metemos la columna
            CadParam = "N" & Format(miRsAux!Grupo, "00") & Format(miRsAux!columna, "0000")
            Set N = TreeView1.Nodes.Add("G" & miRsAux!Grupo, tvwChild, CadParam, miRsAux!TextoColumn)
             N.EnsureVisible
        Else
        
            'Metemos la familia
            CadParam = "N" & Format(miRsAux!Grupo, "00") & Format(miRsAux!columna, "0000")
            numParam = numParam + 1
            Set N = TreeView1.Nodes.Add(CadParam, tvwChild, "F" & Format(numParam, "0000"), miRsAux!Familia)
          
        End If
         
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    
    
    'Cargaremos la familias
    SQL = "Select codfamia,nomfamia FROM sfamia"
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    
    SQL = ""
    For numParam = 1 To TreeView1.Nodes.Count
        If Not TreeView1.Nodes(numParam).Parent Is Nothing Then
            If Mid(TreeView1.Nodes(numParam).Key, 1, 1) <> "N" Then
                'OK es una familia
                SQL = TreeView1.Nodes(numParam).Text
                SQL = "codfamia = " & SQL
                miRsAux.Find SQL, , adSearchForward, 1
                If miRsAux.EOF Then
                    SQL = "0000 - No encontrada"
                Else
                    SQL = Format(miRsAux!Codfamia, "0000") & " - " & miRsAux!nomfamia
                End If
                TreeView1.Nodes(numParam).Text = SQL
            End If
        End If
    Next numParam
    
    miRsAux.Close
    Set miRsAux = Nothing
    CadParam = "" 'por si vuelve a apretar aceptar
End Sub


Private Function GenerarDatosVentasAlzira() As Boolean
Dim NO As Node
Dim F As Date
    GenerarDatosVentasAlzira = False
    
    lblAzira(1).Caption = "Preparando datos"
    lblAzira(1).Refresh
    
    SQL = "DELETE FROM tmpinformes where codusu = " & vUsu.codigo
    conn.Execute SQL
    
    
    'Cargaremos en tmpinformes  12 columnas por ejercicio
    ' mas la del anterior
    'Primero: Cargamos a cero todos los valores
    
    SQL = "codusu,codigo1,campo1,campo2,nombre1,importe1,importe2,importe3,importe4,importe5,importeb1,importeb2,importeb3,importeb4"
    For numParam = 0 To 1
        
        CadParam = Format(vEmpresa.FechaIni, "dd/mm/") & Me.cboEjercicio.ItemData(cboEjercicio.ListIndex)
        F = CDate(CadParam)
        If numParam = 1 Then F = DateAdd("yyyy", -1, F)         'anterior
        CadParam = ""
        For NumRegElim = 1 To 12
            CadParam = CadParam & ", (" & vUsu.codigo & "," & (numParam * 100) + Month(F) & "," & Format(F, "yyyymm") & ","
            CadParam = CadParam & numParam & "," & DBSet(MonthName(Month(F)), "T") & ",0,0,0,0,0,0,0,0,0)"
            F = DateAdd("m", 1, F)
        Next
        CadParam = Mid(CadParam, 2)
        CadParam = "INSERT INTO tmpinformes(" & SQL & ") VALUES " & CadParam
        conn.Execute CadParam
    Next
    
    
    Set miRsAux = New ADODB.Recordset
    
    'Ya tenemos los valores a cero. Vamos con los datos de las ventas
    'Hacemos primero el G1 que son ventas desde TPV y demas
    Set NO = TreeView1.Nodes(1).Child
    If NO Is Nothing Then Exit Function
    HacerDatosVentasAlzira NO
    
    
    'Hacemos despues G2
    Set NO = TreeView1.Nodes(1)
    Set NO = NO.Next
    If Not NO Is Nothing Then
        Set NO = NO.Child
        If Not NO Is Nothing Then HacerDatosVentasAlzira NO
    End If
    
    
    'select nombre1,sum(codigo1)+1000,2,sum(if(campo2=1,-importe1,importe1))
    'from tmpinformes WHERE codusu = 22000  group by 1 order by campo1,codigo1
    lblAzira(1).Caption = "Calculando diferencias"
    lblAzira(1).Refresh
    SQL = ""
    For NumRegElim = 1 To 5
        CadParam = "importe" & NumRegElim
        SQL = SQL & ", sum(if( campo2=1,-" & CadParam & "," & CadParam & "))"
        
        If NumRegElim <> 5 Then
            CadParam = "importeb" & NumRegElim
            SQL = SQL & ", sum(if( campo2=1,-" & CadParam & "," & CadParam & "))"
        End If
    
    Next NumRegElim
    CadParam = "INSERT INTO tmpinformes(codusu,codigo1,campo1,campo2,nombre1"
    For NumRegElim = 1 To 5
        CadParam = CadParam & ",importe" & NumRegElim
        If NumRegElim <> 5 Then CadParam = CadParam & ",importeb" & NumRegElim
    Next NumRegElim
    
    SQL = "select " & vUsu.codigo & ",min(codigo1)+2000,max(campo1)+1000,2,nombre1" & SQL
    SQL = SQL & " FROM tmpinformes WHERE codusu =" & vUsu.codigo & " group by nombre1 order by 3 "
    SQL = CadParam & ") " & SQL
    conn.Execute SQL
    
    
    'Si  algun valor, todas las coluimnas son cero, en el apartado de difernecias pondemos un cero
    
    lblAzira(1).Caption = "Valores a cero"
    lblAzira(1).Refresh
    SQL = "select * FROM tmpinformes WHERE codusu =" & vUsu.codigo & " and campo2<2 "
    SQL = SQL & " and importe1=0 and importeb1=0 and importe2=0 and importeb2=0 and importe3=0 "
    SQL = SQL & " and importeb3=0 and importe4=0 and importeb4=0 and importe5=0 order by campo1"
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        SQL = "UPDATE tmpinformes set importe1=0 , importeb1=0 , importe2=0 , importeb2=0 , importe3=0"
        SQL = SQL & " ,importeb3=0 , importe4=0 , importeb4=0 , importe5=0  WHERE codusu =" & vUsu.codigo
        SQL = SQL & " AND campo2=2 and nombre1=" & DBSet(miRsAux!nombre1, "T")
        conn.Execute SQL
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    GenerarDatosVentasAlzira = True
    
    
    
    lblAzira(1).Caption = ""
End Function



Private Sub HacerDatosVentasAlzira(ByRef N As Node)
Dim N2 As Node
Dim Aux As String
Dim EsPrimerGrupo As Boolean
Dim F1 As Date
    
    'El grupo 1, veremos las familias vinculadas
    'Enviamos el nodo
    SQL = Mid(N.Key, 2, 2)
    EsPrimerGrupo = SQL = "01"
    
    While Not N Is Nothing
        numParam = Mid(N.Key, 4)
        lblAzira(1).Caption = "Leyendo datos " & N.Text
        lblAzira(1).Refresh
        'OK
        Set N2 = N.Child
        SQL = ""
        While Not N2 Is Nothing
            SQL = SQL & "," & Mid(N2.Text, 1, 5) 'los 5 primeros son la cod familia
            Set N2 = N2.Next
        Wend
        
            
        'Select
        If SQL <> "" Then
            SQL = "(" & Mid(SQL, 2) & ")"
            
            
            Aux = Format(vEmpresa.FechaIni, "dd/mm/") & Me.cboEjercicio.ItemData(cboEjercicio.ListIndex) - 1
            F1 = CDate(Aux)
            
            'El select
            Aux = "Select year(slifac.fecfactu)*100+month(slifac.fecfactu),sum(importel) FROM slifac,scafac1,sartic WHERE"
            Aux = Aux & " slifac.codartic=sartic.codartic AND "
            Aux = Aux & " slifac.codtipom=scafac1.codtipom and slifac.numfactu=scafac1.numfactu and slifac.fecfactu=scafac1.fecfactu and"
            Aux = Aux & " slifac.codtipoa=scafac1.codtipoa and slifac.numalbar=scafac1.numalbar and"
            
            'Aux = Aux & " slifac.fecfactu between " & DBSet(DateAdd("yyyy", -1, vEmpresa.FechaIni), "F") & " AND "
            Aux = Aux & " slifac.fecfactu between " & DBSet(F1, "F") & " AND "
            F1 = DateAdd("yyyy", 2, F1) '+ 2 años
            F1 = DateAdd("d", -1, F1) ' menos un dia
            'Antes
            'Aux = Aux & DBSet(vEmpresa.FechaFin, "F")
            Aux = Aux & DBSet(F1, "F")
            Aux = Aux & " AND codfamia in " & SQL
            
            'Esto es para alzira
            If EsPrimerGrupo Then
                'Ventas TPV directa
                Aux = Aux & " AND (not referenc like 'Parte%' or referenc is null)"
            Else
              '  Aux = Aux & " AND referenc like 'Parte%'"
            End If
            
            
            
            Aux = Aux & " GROUP BY 1"
            miRsAux.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not miRsAux.EOF
                lblAzira(1).Caption = N.Text & " " & miRsAux.Fields(0)
                lblAzira(1).Refresh
                DoEvents
                If EsPrimerGrupo Then
                    Aux = RecuperaValor("importe1|importe2|importe3|", numParam)
                Else
                    Aux = RecuperaValor("importe4|importe5|importeb1|importeb2|importeb3|importeb4|", numParam)
                End If
                If DBLet(miRsAux.Fields(1), "N") <> 0 Then
                    Aux = "UPDATE tmpinformes set " & Aux & " = " & DBSet(miRsAux.Fields(1), "N")
                    Aux = Aux & " WHERE codusu = " & vUsu.codigo & " AND campo1=" & miRsAux.Fields(0)
                    conn.Execute Aux
                End If
                miRsAux.MoveNext
            Wend
            miRsAux.Close
            
            
        End If
            

        Set N = N.Next
    Wend
    
End Sub




Private Sub CargaComboEjercicio()
Dim F1 As Date
    
    cboEjercicio.Clear
    SQL = "Select min(fecfactu),max(fecfactu) FROM scafac"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    F1 = vEmpresa.FechaIni
    If Not miRsAux.EOF Then
        F1 = miRsAux.Fields(0)
        NumRegElim = Val(Format(miRsAux.Fields(1), "yyyymm"))
    End If
    miRsAux.Close
    Set miRsAux = Nothing
    
    

    If Year(vEmpresa.FechaIni) = Year(vEmpresa.FechaFin) Then
        F1 = CDate("01/01/" & Year(F1))
        lblTitulo(5).Caption = "Año"
    Else
        CadParam = Format(F1, "mmdd")
        F1 = CDate(Format(F1, "dd/mm/") & Year(F1))
        If CadParam < Format(vEmpresa.FechaIni, "mmdd") Then F1 = DateAdd("yyyy", -1, F1)
        lblTitulo(5).Caption = "Ejercicio"
    End If
            
            
    Do
        If Year(vEmpresa.FechaIni) = Year(vEmpresa.FechaFin) Then
            'Años naturales
            SQL = Year(F1)
            CadParam = "31/12/" & Year(F1)
        Else
            'Años partidos
            SQL = Year(F1) & " / " & Year(F1) + 1
            CadParam = Format(vEmpresa.FechaFin, "dd/mm/") & Year(F1) + 1
        End If
        cboEjercicio.AddItem SQL
        cboEjercicio.ItemData(cboEjercicio.NewIndex) = Year(F1)
        F1 = DateAdd("yyyy", 1, F1)
        
        If Val(Format(CDate(CadParam), "yyyymm")) > NumRegElim Then SQL = ""
    
    Loop Until SQL = ""
    cboEjercicio.ListIndex = cboEjercicio.ListCount - 1

End Sub



Private Sub CargaComboTelef()
    Me.cboTelefono.AddItem "Todos"
    
    
    If Me.optTelefono(0).Value Then
        CadParam = "select codigo_de_trafico,tipo_de_trafico from telefono.detalle_de_llamadas "
    ElseIf Me.optTelefono(1).Value Then
        CadParam = "select Codigo_de_cuota ,Descripcion_de_cuota from telefono.cuotas "
    Else
        CadParam = "select Codigo_de_vario , Descripcion_de_vario  from telefono.varios "
    End If
    CadParam = CadParam & " where fichero ='" & vCadena & "'"
    
    If Me.optTelefono(2).Value Then CadParam = CadParam & " AND Codigo_de_vario<>'' "
    CadParam = CadParam & "  group by 1 ORDER BY 1"
    
    miRsAux.Open CadParam, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        CadParam = Trim(miRsAux.Fields(1)) & "  [" & Trim(miRsAux.Fields(0)) & "]"
        cboTelefono.AddItem CadParam
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
        
    
    
End Sub


Private Sub CargarDatosTelefonia()
    Me.lw(3).ListItems.Clear
    CargarDatosDetalle
End Sub

Private Sub CargarDatosDetalle()
Dim ImporteT As Currency

    If Me.optTelefono(0).Value Then
        CadParam = "select Numero_de_telefono,codigo_de_trafico,Tipo_de_trafico,Fecha,Hora_inicio,Unidad_de_medida,"
        CadParam = CadParam & " Cantidad_medida_originada,Importe,Libre FROM telefono.detalle_de_llamadas "
        CadParam = CadParam & " where fichero ='" & vCadena & "' AND fecha <> '0000'"
        
    ElseIf Me.optTelefono(1).Value Then
        CadParam = "SELECT Numero_de_telefono,Codigo_de_cuota,Descripcion_de_cuota, '0000' fecha"
        CadParam = CadParam & " ,'' Hora_inicio,'-' Unidad_de_medida,1 Cantidad_medida_originada,Importe,'' Libre"
        CadParam = CadParam & " From telefono.cuotas WHERE fichero ='" & vCadena & "' AND Numero_de_telefono<>'0'"
    Else
        'VARIOS
        CadParam = "SELECT Numero_de_telefono,Codigo_de_vario,Descripcion_de_vario, '0000' fecha"
        CadParam = CadParam & " ,'' Hora_inicio,'-' Unidad_de_medida,1 Cantidad_medida_originada,Importe,'' Libre"
        CadParam = CadParam & " From telefono.varios WHERE fichero ='" & vCadena & "' AND Numero_de_telefono<>'0'  AND Codigo_de_vario<>''"
    End If
    SQL = ""
    If Me.cboTelefono.ListIndex > 0 Then
        numParam = InStr(1, Me.cboTelefono.Text, "[")
        If numParam > 0 Then
            SQL = Mid(Me.cboTelefono, numParam + 1)
            numParam = InStr(1, SQL, "]")
            If numParam = 0 Then
                SQL = ""
            Else
                SQL = Mid(SQL, 1, numParam - 1)
            End If
        End If
        If SQL <> "" Then
            SQL = " AND  @@@ ='" & SQL & "'"
            If Me.optTelefono(0).Value Then
                SQL = Replace(SQL, "@@@", "codigo_de_trafico")
            ElseIf Me.optTelefono(1).Value Then
                SQL = Replace(SQL, "@@@", "Codigo_de_cuota")
            
            Else
                SQL = Replace(SQL, "@@@", "Codigo_de_vario")
            End If
            
        End If
    End If
    CadParam = CadParam & SQL & " ORDER BY "
    
    If Me.optTelefono(2).Value Then
        
        CadParam = CadParam & " Numero_de_telefono,Codigo_de_vario"
    Else
        CadParam = CadParam & " Fecha , Hora_inicio, Numero_de_telefono"
    End If
    miRsAux.Open CadParam, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not miRsAux.EOF
        Set IT = lw(3).ListItems.Add()
        
        'Numero_de_telefono,codigo_de_trafico,Tipo_de_trafico,Fecha,Hora_inicio,Unidad_de_medida,
        'Cantidad_medida_originada,Importe,Libre FROM telefono.detalle_de_llamadas
        IT.Text = miRsAux!Numero_de_telefono & " "
        IT.SubItems(1) = miRsAux.Fields(1) ' CStr(miRsAux!codigo_de_trafico)
        IT.SubItems(2) = miRsAux.Fields(2) ' miRsAux!Tipo_de_trafico
        If miRsAux!Fecha = "0000" Then
            IT.SubItems(3) = "-"
            IT.SubItems(4) = "-"
        Else
            IT.SubItems(3) = Mid(miRsAux!Fecha, 3) & "/" & Mid(miRsAux!Fecha, 1, 2)
            
            IT.SubItems(4) = Mid(miRsAux!Fecha, 1, 2) & ":" & Mid(miRsAux!Fecha, 3)
        End If
        IT.SubItems(5) = Mid(miRsAux!Unidad_de_medida, 1, 3)
        IT.SubItems(6) = Format(miRsAux!Cantidad_medida_originada, FormatoCantidad)
        IT.SubItems(7) = Format(miRsAux!Importe, FormatoPrecio)
        ImporteT = miRsAux!Importe + ImporteT
        If DBLet(miRsAux!Libre, "T") = "" Then
            IT.SubItems(8) = " "
        Else
            IT.SubItems(8) = miRsAux!Libre
        End If
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    If lw(3).ListItems.Count = 0 Then
        lblTitulo(12).Caption = ""
        lblTitulo(13).Caption = ""
    Else
        lblTitulo(12).Caption = "Reg: " & Me.lw(3).ListItems.Count
        lblTitulo(13).Caption = Format(ImporteT, FormatoPrecio)
        

    End If
End Sub


'------------------------------------------------------------------------------------

Private Sub CargaDatosFAZ()
Dim Importe As Currency
Dim Aux2 As Currency

    Screen.MousePointer = vbHourglass
    Set miRsAux = New ADODB.Recordset
    Me.Refresh
    lw(4).ListItems.Clear
    
    'LOS IVAS
    CadParam = ""
    SQL = "Select * from tiposiva"
    miRsAux.Open SQL, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        CadParam = CadParam & Format(miRsAux!codigiva, "0000") & "#" & Right(Space(5) & miRsAux!PorceIVA, 5) & "|"
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    
    NumRegElim = 0
    SQL = "select scafac.numfactu,scafac.fecfactu,codigiva,scafac.codclien,scafac.nomclien,sum(importel) suma"
    SQL = SQL & " from scafac,slifac,sartic,sclien where scafac.codclien =sclien.codclien AND "
    SQL = SQL & " scafac.codtipom = slifac.codtipom And scafac.Numfactu = slifac.Numfactu And scafac.FecFactu = "
    SQL = SQL & " slifac.FecFactu and sartic.codartic=slifac.codartic "
    SQL = SQL & " AND slifac.codartic<>'01005000'"
    'cadenadesdeotroform= SELECT desde frmlistado3
    SQL = SQL & " AND " & RecuperaValor(vCadena, 4)
    SQL = SQL & " group by scafac.numfactu,scafac.fecfactu,codigiva"

    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = -1
    While Not miRsAux.EOF
    
        numParam = InStr(1, CadParam, Format(miRsAux!codigiva, "0000") & "#")
        If numParam = 0 Then
            SQL = "0"
        Else
            SQL = Trim(Mid(CadParam, numParam + 5, 5))
        End If
    
        If NumRegElim = miRsAux!Numfactu Then
            
            
            'Fra con dos IVAS
            IT.SubItems(6) = IT.SubItems(6) & " - " & SQL & "(" & miRsAux!codigiva & ")"
            'BRUTO
            Importe = ImporteFormateado(IT.SubItems(5)) + miRsAux!Suma
            IT.SubItems(5) = Format(Importe, FormatoImporte)
            
            Importe = (CCur(SQL) / 100)
            Importe = Round2(Importe * miRsAux!Suma, 2)
            
            Aux2 = ImporteFormateado(IT.SubItems(7)) + Importe
            IT.SubItems(7) = Format(Aux2, FormatoImporte)
            
            Importe = ImporteFormateado(IT.SubItems(7)) + miRsAux!Suma + Importe
            IT.SubItems(8) = Format(Importe + miRsAux!Suma, FormatoImporte)
        Else
            'EL IVA
            
                Set IT = lw(4).ListItems.Add()
                
                IT.SubItems(6) = SQL
                Importe = (CCur(SQL) / 100)
            
                
                IT.Text = Format(miRsAux!Numfactu, "000000")
                IT.SubItems(1) = Format(miRsAux!FecFactu, "dd/mm/yyyy")
                IT.SubItems(2) = Format(miRsAux!FecFactu, "yyyymmdd") 'para la ordenacion
                IT.SubItems(3) = Format(miRsAux!codClien, "0000")
                IT.SubItems(4) = miRsAux!NomClien
                IT.SubItems(5) = Format(miRsAux!Suma, FormatoImporte)
                
                'It.SubItems(6) = Format(Importe, FormatoImporte)
                
                Importe = Round2(Importe * miRsAux!Suma, 2)
                IT.SubItems(7) = Format(Importe, FormatoImporte)
                
                IT.SubItems(8) = Format(Importe + miRsAux!Suma, FormatoImporte)
                If (lw(4).ListItems.Count Mod 200) = 0 Then
                    lblTitulo(7).Caption = "Leyendo ... " & lw(4).ListItems.Count
                    Me.Refresh
                    DoEvents
                End If
                
        End If
        NumRegElim = miRsAux!Numfactu
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    lblTitulo(7).Caption = "0,00"
    lblTitulo(7).Tag = CCur("0,0")
End Sub



'Numregelim tendra el subitem correspondiente
Private Function CambiarFactura_A_FAS(ContadorFAS As Long, ByRef vCl As CCliente) As Boolean
Dim vFac As CFactura

    
    On Error GoTo eCambi
    
    CadParam = "slifac.codtipom = 'FAZ' AND slifac.numfactu=" & lw(4).ListItems(NumRegElim).Text & " AND slifac.fecfactu = " & DBSet(Me.lw(4).ListItems(NumRegElim).SubItems(1), "F")
    
    ''LA smoval
    SQL = "Select codartic,scafac1.numalbar,scafac1.fechaalb from slifac,scafac1 WHERE "
    SQL = SQL & " scafac1.codtipom=slifac.codtipom AND scafac1.numfactu=slifac.numfactu AND scafac1.fecfactu=slifac.fecfactu AND "
    SQL = SQL & CadParam
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    SQL = ""
    'NO Puede ser eof
    Do
        SQL = "UPDATE smoval SET detamovi='ALS' WHERE codartic=" & DBSet(miRsAux!codArtic, "T")
        SQL = SQL & " AND detamovi='ALZ' AND document='" & Format(miRsAux!NumAlbar, "0000000") & "'"
        SQL = SQL & " AND fechamov = " & DBSet(miRsAux!FechaAlb, "F")
        miRsAux.MoveNext
        ejecutar SQL, False
    Loop Until miRsAux.EOF
    miRsAux.Close
    
    
    
    conn.BeginTrans
    CadParam = "codtipom = 'FAZ' AND numfactu=" & lw(4).ListItems(NumRegElim).Text & " AND fecfactu = " & DBSet(Me.lw(4).ListItems(NumRegElim).SubItems(1), "F")
    'Vamos a cambiar... scafac,scafac1,slifac,svenci
    'codtipom numfactu fecfactu
    
    SQL = RecuperaValor(vCadena, 3)
    CadenaDesdeOtroForm = "numfactu = " & ContadorFAS + 1 & ", codtipom='FAS', fecfactu='" & Format(SQL, FormatoFecha) & "'"
    
    
    conn.Execute "SET FOREIGN_KEY_CHECKS=0;"
    
    'El vto lo borro
    SQL = "DELETE FROM svenci  WHERE " & CadParam
    conn.Execute SQL
    
    SQL = "UPDATE slifac SET " & CadenaDesdeOtroForm & ",codtipoa='ALS' WHERE " & CadParam
    conn.Execute SQL
    
    SQL = "UPDATE scafac1 SET " & CadenaDesdeOtroForm & ",codtipoa='ALS' WHERE " & CadParam
    conn.Execute SQL
    
    SQL = "UPDATE scafac SET " & CadenaDesdeOtroForm & ",codagent=" & vCl.Agente & " WHERE " & CadParam
    conn.Execute SQL
    
    
        
    
    'Insertamos el cobro
    
    
    
    conn.Execute "SET FOREIGN_KEY_CHECKS=1;"

        
    'AHora recalculamos los importes de la factura
    
    Espera 0.25
    Set vFac = New CFactura
    lblTitulo(7).Caption = "Recalculando importes"
    lblTitulo(7).Refresh
    
    SQL = RecuperaValor(vCadena, 3)
    If vFac.LeerDatos("FAS", ContadorFAS + 1, SQL) Then
        Debug.Print vFac.Agente
        
        SQL = "codtipom='" & vFac.codtipom & "' AND fecfactu=" & DBSet(vFac.FecFactu, "F") & " AND numfactu=" & vFac.Numfactu
        
        vFac.CuentaPrev = DevuelveDesdeBD(conAri, "codagent", "scafac", SQL & " AND 1", "1")
        If Val(vFac.CuentaPrev) = 0 Then vFac.CuentaPrev = DevuelveDesdeBD(conAri, "codagent", "sagent", "codagent>0 AND 1", "1")
        vFac.Agente = Val(CInt(vFac.CuentaPrev))
        vFac.CuentaPrev = ""
        
        If vFac.CalcularDatosFactura(SQL, "scafac", "slifac", False) Then
            
            SQL = "UPDATE scafac set "
            'baseimp1 codigiv1 porciva1 imporiv1  NO LLEVA REA
            SQL = SQL & " baseimp1=" & DBSet(vFac.BaseIVA1, "N", "N") & " , codigiv1= " & DBSet(vFac.TipoIVA1, "N", "N")
            SQL = SQL & ", porciva1=" & DBSet(vFac.PorceIVA1, "N", "N") & " , imporiv1= " & DBSet(vFac.ImpIVA1, "N", "N")
            
            'La base 2  NO LLEVA REA
            SQL = SQL & ", baseimp2=" & DBSet(vFac.BaseIVA2, "N", "S") & " , codigiv2= " & DBSet(vFac.TipoIVA2, "N", "S")
            SQL = SQL & ", porciva2=" & DBSet(vFac.PorceIVA2, "N", "S") & " , imporiv2= " & DBSet(vFac.ImpIVA2, "N", "S")
            
            'La base 3
            'baseimp1 codigiv1 porciva1 imporiv1  NO LLEVA REA
            SQL = SQL & ", baseimp3=" & DBSet(vFac.BaseIVA3, "N", "S") & " , codigiv3= " & DBSet(vFac.TipoIVA3, "N", "S")
            SQL = SQL & ", porciva3=" & DBSet(vFac.PorceIVA3, "N", "S") & " , imporiv3= " & DBSet(vFac.ImpIVA3, "N", "S")
            'Total
            SQL = SQL & ", TotalFac=" & DBSet(vFac.TotalFac, "N", "N")
            SQL = SQL & ", codbanco=0,codsucur=0,digcontr=null,cuentaba=null,coddirec=null,nomdirec=null"
            
            
         
            
            vFac.Cliente = vCl.codigo
            vFac.BancoPr = RecuperaValor(Me.vCadena, 2)
            'domclien codpobla pobclien proclien nifclien telclien   codforpa
            SQL = SQL & ", codclien=" & vCl.codigo
            SQL = SQL & ", nomclien=" & DBSet(vCl.Nombre, "T")
            SQL = SQL & ", domclien=" & DBSet(vCl.Domicilio, "T")
            SQL = SQL & ", codpobla=" & DBSet(vCl.CPostal, "T")
            SQL = SQL & ", pobclien=" & DBSet(vCl.Poblacion, "T")
            SQL = SQL & ", proclien=" & DBSet(vCl.Provincia, "T")
            SQL = SQL & ", nifclien=" & DBSet(vCl.NIF, "T")
            SQL = SQL & ", telclien=" & DBSet(vCl.TfnoClien, "T")
            
            SQL = SQL & " WHERE codtipom='" & vFac.codtipom & "' AND fecfactu=" & DBSet(vFac.FecFactu, "F") & " AND numfactu=" & vFac.Numfactu
            
            conn.Execute SQL
            
            'Diciembre 2014. Vuelven a querer el vto en tesoreria
            vFac.CuentaPrev = DevuelveDesdeBD(conAri, "codmacta", "sbanpr", "codbanpr", vFac.BancoPr)
            'vFac.CuentaPrev = ""
            If Not vFac.InsertarEnTesoreria(vFac.codtipom & vFac.Numfactu & "||", SQL) Then MsgBox SQL, vbExclamation
            
        End If
    End If
    Set vFac = Nothing
    
    conn.CommitTrans
    CambiarFactura_A_FAS = True
    
    Exit Function
eCambi:
    MuestraError Err.Number, "Actualizando registro"
    conn.RollbackTrans

End Function


Private Sub LanzaBuscaGrid(KOpcion As Byte)

    Screen.MousePointer = vbHourglass
    Set frmB = New frmBuscaGrid
    
    Select Case KOpcion
    Case 0
    
        SQL = "Código|sclien|codclien|T||10·Nombre|sclien|nomclien|T||50·"
        SQL = SQL & "Telefono|sclientfno|IdTelefono|T||20·Nombre|stfnooperador|nombre|T||20·"
        frmB.vCampos = SQL
        frmB.vTabla = "sclien,sclientfno,stfnooperador"
        frmB.vSQL = "sclien.codclien=sclientfno.codclien and stfnooperador.codoperador=sclientfno.operador"
        
        '###A mano
        frmB.vDevuelve = "0|1|2|3|"
        frmB.vTitulo = "Telefonos"
        frmB.vselElem = 2
        frmB.vConexionGrid = conAri
    
    
    Case 1
        'aguacontadores ,sclien where aguacontadores.codclien=sclien.codclien
        SQL = "Código|sclien|codclien|T||10·Nombre||nomclien|T||50·"
        SQL = SQL & "Contador|aguacontadores|contador|T||15·"   'Nombre|aguacontadores|nombre|T||20·"
        frmB.vCampos = SQL
        frmB.vTabla = "aguacontadores ,sclien "
        frmB.vSQL = "aguacontadores.codclien=sclien.codclien"
        
        
        
         'aguacontadores ,sclien where aguacontadores.codclien=sclien.codclien
        SQL = "Código|sclien|codclien|T||10·Nombre||nomclien|T||50·"
        SQL = SQL & "Contador|aguacontadores|contador|T||15·Facturar||if(coalesce(aguacontadoresconce.Facturar,0)=1,""Si"","""")|T||20·"
        frmB.vCampos = SQL
        SQL = "sclien inner join aguacontadores on aguacontadores.codclien=sclien.codclien left join aguacontadoresconce ON  aguacontadores.Contador = aguacontadoresconce.Contador"
        frmB.vTabla = SQL
        frmB.vSQL = "aguacontadoresconce.codconceAg = 7 "
        
        
        
        
        
        frmB.vDevuelve = "2|"
        frmB.vTitulo = "Contadores"
        frmB.vselElem = 2
        frmB.vConexionGrid = conAri
    
    
    End Select
    SQL = ""
    frmB.vCargaFrame = False
    frmB.Show vbModal
    Set frmB = Nothing
    If SQL <> "" Then
       Select Case KOpcion
       Case 0
            'Añadimos el nodo. si formateamos el KEY del nodo como 00000999999999 codclien+tfno, si ya existe da error
            AnyadeNodoTelefonia True
            'Else
             'AnyadeNodoTelefonia False
             
       Case 1
            CadenaDesdeOtroForm = SQL
            SQL = "select *    from aguacontadores  left join aguacontadoresconce ON "
            
            SQL = "select aguacontadores.contador,aguacontadores.codclien,sclien.codclien,nomclien,descripcion,importeconcepto,aguacontadoresconce.Facturar"
            SQL = SQL & " from sclien inner join aguacontadores on aguacontadores.codclien=sclien.codclien "
            SQL = SQL & " left join aguacontadoresconce ON  aguacontadores.Contador = aguacontadoresconce.Contador "
            SQL = SQL & " WHERE aguacontadoresconce.codconceAg = 7 "
            SQL = SQL & " and aguacontadores.contador=" & DBSet(RecuperaValor(CadenaDesdeOtroForm, 1), "T")
            Set miRsAux = New ADODB.Recordset
            miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            SQL = ""
            If miRsAux.EOF Then
                'por el motivo que fuese(no viene al caso) no tenia la cuota 7 de varios
                MsgBox "Error leyendo caontador", vbExclamation
                
            Else
                SQL = miRsAux!Contador & "|" & miRsAux!NomClien & "|"
                SQL = SQL & DBLet(miRsAux!Descripcion, "T") & "|"
                SQL = SQL & Format(DBLet(miRsAux!importeconcepto, "N"), FormatoImporte)
                SQL = SQL & "|" & miRsAux!codClien & "|" & miRsAux!facturar & "|"
                
            End If
            CadenaDesdeOtroForm = SQL
            miRsAux.Close
            Set miRsAux = Nothing
            If SQL <> "" Then AnyadeNodoAguaContadores
       End Select
    End If
End Sub

'En sql iran codclien, nomclien, tfno operador
Private Sub AnyadeNodoTelefonia(ElDeTelefono As Boolean)
On Error GoTo eAnyadeNodoTfno

    If ElDeTelefono Then
        CadParam = Format(RecuperaValor(SQL, 1), "000000") & Trim(RecuperaValor(SQL, 3))
        Set IT = lw(5).ListItems.Add(, "K" & CadParam)
        IT.Text = RecuperaValor(SQL, 3)
        IT.SubItems(1) = Format(RecuperaValor(SQL, 1), "000000")
        IT.SubItems(2) = Trim(RecuperaValor(SQL, 2))
        IT.SubItems(3) = RecuperaValor(SQL, 4)
        IT.EnsureVisible
        IT.Selected = True
    
    Else
        SQL = CadenaDesdeOtroForm
        CadParam = RecuperaValor(CadenaDesdeOtroForm, 1)
        Set IT = lw(6).ListItems.Add(, "K" & CadParam)
        IT.Text = RecuperaValor(CadenaDesdeOtroForm, 1)
        IT.SubItems(1) = RecuperaValor(CadenaDesdeOtroForm, 2)
        IT.SubItems(2) = RecuperaValor(CadenaDesdeOtroForm, 3)
    End If
    
    Exit Sub
eAnyadeNodoTfno:
    MuestraError Err.Number, Err.Description, SQL
End Sub


Private Sub AnyadeNodoAguaContadores()
On Error GoTo eAnyadeNodoTfno
    'SQL = CadenaDesdeOtroForm lo hace antes de llamar a la funcion
    CadParam = RecuperaValor(CadenaDesdeOtroForm, 1)
    Set IT = lw(9).ListItems.Add(, "K" & CadParam)
    IT.Text = RecuperaValor(CadenaDesdeOtroForm, 1)
    IT.SubItems(1) = RecuperaValor(CadenaDesdeOtroForm, 5)
    IT.SubItems(2) = RecuperaValor(CadenaDesdeOtroForm, 2)
    IT.SubItems(3) = RecuperaValor(CadenaDesdeOtroForm, 3)
    IT.SubItems(4) = RecuperaValor(CadenaDesdeOtroForm, 4)
    If RecuperaValor(CadenaDesdeOtroForm, 5) = 1 Then
        IT.SubItems(5) = "Si"
    Else
        IT.SubItems(5) = "NO"
    End If
    Exit Sub
eAnyadeNodoTfno:
    MuestraError Err.Number, Err.Description, SQL

End Sub

Private Sub LeerFichero()
Dim OK As Boolean
Dim ColLineas As Collection
        
    On Error GoTo ELeer
    
    numParam = -1
    cd1.FileName = ""
    cd1.InitDir = "c:\"
    cd1.CancelError = False
    cd1.ShowOpen
    If cd1.FileName = "" Then Exit Sub
    CadParam = cd1.FileName
    
    
    
    
    
    numParam = FreeFile
    Open CadParam For Input As #numParam
    'Primera linea
    Line Input #numParam, CadParam
    'Si empieza por #001 y acaba por 002# esta bien
    OK = False
    If Mid(CadParam, 1, 4) = "#001" Then
        If Right(CadParam, 4) = "002#" Then OK = True
    End If
    If Not OK Then
        SQL = "Encabezado del fichero incorrecto"
        
    Else
        Set ColLineas = New Collection
        'SQL--> Errores
        While Not EOF(numParam)
            Line Input #numParam, CadParam
            'Iremos leyendo linea a linea hasta que la ultima empieza por #format(numreg,"0000")#
            If Mid(CadParam, 1, 1) = "#" Then
                If Right(CadParam, 1) <> "#" Then
                    SQL = SQL & vbCrLf & "Fin fichero incorrecto (#)"
                Else
                    vCadena = Mid(CadParam, 2)
                    vCadena = Mid(vCadena, 1, Len(vCadena) - 1)
                    If Not IsNumeric(vCadena) Then
                        vCadena = "No numerica"
                    Else
                        If Val(vCadena) <> ColLineas.Count Then
                            vCadena = "Nº registros incorrectos"
                        Else
                            vCadena = ""
                            SQL = ""
                        End If
                    End If
                    If vCadena <> "" Then SQL = SQL & vbCrLf & "Fin fichero incorrecto"
                End If
            Else
                
                'Como sabemos que la linea correcta.
                'La codclien + numerotelefono --> CadenaTextoMod97
                NumRegElim = InStr(1, CadParam, "@")
                If NumRegElim = 0 Then
                    SQL = SQL & vbCrLf & "Linea incorrecta (@)"
                Else
                    vCadena = Mid(CadParam, 1, NumRegElim - 1)
                    CadenaDesdeOtroForm = CadenaTextoMod97(vCadena)
                    vCadena = Mid(CadParam, NumRegElim + 1)
                    If CadenaDesdeOtroForm <> vCadena Then
                        SQL = SQL & vbCrLf & "Linea incorrecta (Mod97): " & CadenaDesdeOtroForm
                    Else
                        ColLineas.Add Mid(CadParam, 1, NumRegElim - 1)
                    End If
                End If
            End If
        Wend
    End If
    Close #numParam
    numParam = -1
    vCadena = ""
    CadenaDesdeOtroForm = ""
    If SQL <> "" Then
        MsgBox SQL, vbExclamation
        Set ColLineas = Nothing
        Exit Sub
    End If
    
    Set miRsAux = New ADODB.Recordset
    For NumRegElim = 1 To ColLineas.Count
        SQL = "SELECT sclien.codclien, sclien.nomclien, idtelefono, nombre FROM sclien,sclientfno,stfnooperador WHERE sclien.codclien=sclientfno.codclien"
        SQL = SQL & " and stfnooperador.codoperador=sclientfno.operador and sclientfno.codclien="
        CadParam = Mid(ColLineas.item(NumRegElim), 1, 9)
        vCadena = Mid(ColLineas.item(NumRegElim), 10)
        SQL = SQL & vCadena & " and idtelefono='" & CadParam & "'"
        miRsAux.Open SQL, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
        If Not miRsAux.EOF Then
            SQL = miRsAux!codClien & "|" & miRsAux!NomClien & "|" & miRsAux!idtelefono & "|" & miRsAux!Nombre & "|"
            AnyadeNodoTelefonia True
        Else
            SQL = Mid(SQL, InStr(1, SQL, " WHERE ") + 12)
            MsgBox "No existe telefono: " & SQL, vbExclamation
        End If
        miRsAux.Close
    Next
    
    
ELeer:
    If Err.Number <> 0 Then MuestraError Err.Number
    If numParam >= 0 Then CerrarFichero
    Set ColLineas = Nothing
    Set miRsAux = Nothing
    vCadena = ""
End Sub


Private Sub GuadarFichero()

    On Error GoTo eGuadarFichero
    numParam = -1
    cd1.FileName = ""
    cd1.InitDir = "c:\"
    cd1.CancelError = False
    cd1.ShowSave
    If cd1.FileName = "" Then Exit Sub
    CadParam = cd1.FileName

    If Dir(CadParam, vbArchive) <> "" Then
        If MsgBox("El fichero ya existe. Sobreescribir?", vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    End If
    
    numParam = FreeFile
    Open CadParam For Output As #numParam
    'Primera linea
    CadParam = "#001 Ariges. Generado en Ariges(Ariadna Software) por " & vUsu.Nombre & "002#"
    Print #numParam, CadParam
    
    For NumRegElim = 1 To lw(5).ListItems.Count
        SQL = lw(5).ListItems(NumRegElim).Text & Me.lw(5).ListItems(NumRegElim).SubItems(1)
        vCadena = SQL
        CadParam = CadenaTextoMod97(SQL)
        
        Print #numParam, vCadena & "@" & CadParam
    Next
    Print #numParam, "#" & Format(NumRegElim - 1, "0000000") & "#"
    Close #numParam
        
    MsgBox "Fichero guardado: " & cd1.FileName, vbInformation
    Exit Sub
eGuadarFichero:
    MuestraError Err.Number, Err.Description
    If numParam > 0 Then CerrarFichero
End Sub


Private Sub CerrarFichero()
    On Error Resume Next
    Close #numParam
    Err.Clear
End Sub


Private Sub hacerProcesoCuotas(SoloBorrar As Boolean)

    'Borramos todas las cuotas para los telefonos seleccionados
    Me.lblAzira(0).Caption = "Borre datos anteriores"
    Me.lblAzira(0).Refresh
    CargaTelefonosParaElUpdate
    SQL = "delete from sclientfnocuotas where IdTelefono in (" & SQL & ")"
    conn.Execute SQL
    
    
    
    
    If Not SoloBorrar Then
        
    
        'Insertamos las cuotas para los telefonos
        For numParam = 1 To lw(6).ListItems.Count
            Me.lblAzira(0).Caption = "Insertar cuota " & lw(6).ListItems(numParam).SubItems(1)
            Me.lblAzira(0).Refresh
            CargaTelefonosParaElUpdate
            
            'Insertamios  en  sclientfnocuotas IdTelefono numlinea descripcion precio
            SQL = " FROM sclientfno WHERE idtelefono in ( " & SQL & ")"
            CadParam = "SELECT idtelefono," & lw(6).ListItems(numParam) & "," & DBSet(lw(6).ListItems(numParam).SubItems(1), "T")
            CadParam = CadParam & "," & TransformaComasPuntos(lw(6).ListItems(numParam).SubItems(2))
            SQL = CadParam & SQL
            SQL = "INSERT INTO  sclientfnocuotas(IdTelefono ,numlinea ,descripcion ,precio) " & SQL
            conn.Execute SQL
            
        Next
        
        
        
        If Me.cboOperadora.ListIndex > 0 Then
            Me.lblAzira(0).Caption = "Actualizar operadora"
            Me.lblAzira(0).Refresh
            'Ha forzado operadora
            CargaTelefonosParaElUpdate
            CadParam = "UPDATE sclientfno set operador = " & Me.cboOperadora.ItemData(cboOperadora.ListIndex)
            SQL = CadParam & " WHERE  idtelefono in ( " & SQL & ")"
            conn.Execute SQL
        End If
        
    End If
    
    Set LOG = New cLOG
    
        
        CadParam = ""
        For NumRegElim = 1 To Me.lw(6).ListItems.Count
            CadParam = CadParam & "   " & lw(6).ListItems(NumRegElim).SubItems(1) & "(" & Me.lw(6).ListItems(NumRegElim).Text & ")"
        Next
        
        If Not SoloBorrar Then
            If Me.cboOperadora.ListIndex > 0 Then CadParam = CadParam & "   Operadora: " & cboOperadora.List(cboOperadora.ListIndex)
        Else
            CadParam = "Solo BORRAR"
        End If
        CadParam = CadParam & vbCrLf
        numParam = Len(CadParam)
            
        
        
        SQL = ""
        For NumRegElim = 1 To Me.lw(5).ListItems.Count
            
            SQL = SQL & "    " & lw(5).ListItems(NumRegElim).Text
            If Len(SQL) + numParam > 230 Then
                SQL = CadParam & SQL
                LOG.Insertar 24, vUsu, SQL
                SQL = ""
                Espera 1
            End If
        Next
        If SQL <> "" Then
            SQL = CadParam & SQL
            LOG.Insertar 24, vUsu, SQL
        End If
    Set LOG = Nothing
            
    
    
        
    
End Sub

Private Sub CargaTelefonosParaElUpdate()
    SQL = ""
    For NumRegElim = 1 To Me.lw(5).ListItems.Count
        SQL = SQL & ", '" & lw(5).ListItems(NumRegElim).Text & "'"
    Next
    SQL = Mid(SQL, 2)  'no puede ser """
End Sub

Private Function PonerDesdeHastaControlDir(Albaranes As Boolean) As String
    PonerDesdeHastaControlDir = ""
    'CLIENTE
    CadParam = RecuperaValor(vCadena, 1)
    If CadParam <> "" Then
        If Albaranes Then
            CadParam = "scaalb.codclien >=" & CadParam
        Else
            CadParam = "scafac.codclien >=" & CadParam
        End If
        PonerDesdeHastaControlDir = PonerDesdeHastaControlDir & " AND " & CadParam
    End If
    
    CadParam = RecuperaValor(vCadena, 2)
    If CadParam <> "" Then
        If Albaranes Then
            CadParam = "scaalb.codclien <=" & CadParam
        Else
            CadParam = "scafac.codclien <=" & CadParam
        End If
        PonerDesdeHastaControlDir = PonerDesdeHastaControlDir & " AND " & CadParam
    End If
    
    
    If Albaranes Then Exit Function
    
    CadParam = RecuperaValor(vCadena, 3)
    If CadParam <> "" Then PonerDesdeHastaControlDir = PonerDesdeHastaControlDir & " AND scafac.fecfactu >= '" & CadParam & "'"
    
    CadParam = RecuperaValor(vCadena, 4)
    If CadParam <> "" Then PonerDesdeHastaControlDir = PonerDesdeHastaControlDir & " AND scafac.fecfactu <= '" & CadParam & "'"
    
    
    
End Function

Private Sub CargaControlDireccionesEnvio()

    'Comprobar direcciones de envio
    Me.lblTitulo(11).Caption = "Leyendo desde BD..."
    Me.lblTitulo(11).Refresh
    DoEvents
    Me.lw(7).ListItems.Clear
    lw(7).SortKey = 1
    lw(7).Sorted = True
    
    SQL = "SELECT codtipom as codtipoa,numalbar ,scaalb.fechaalb ,scaalb.codclien ,nomclien,"
    SQL = SQL & "scaalb.coddiren ,nomdiren ,'' codtipom_ ,null AS numfactu,null AS fecfactu"
    SQL = SQL & " From scaalb, sdirenvio WHERE scaalb.codclien=sdirenvio.codclien AND"
    SQL = SQL & " scaalb.coddiren = sdirenvio.coddiren" & PonerDesdeHastaControlDir(True)
    'Si tiene WHERE
    SQL = SQL & " Union"
    SQL = SQL & " SELECT scafac1.codtipoa,scafac1.numalbar ,scafac1.fechaalb ,scafac.codclien ,scafac.nomclien,"
    SQL = SQL & " scafac1.coddiren ,nomdiren ,scafac.codtipom codtipom_,scafac.numfactu ,scafac.fecfactu"
    SQL = SQL & " From scafac, scafac1, sdirenvio"
    SQL = SQL & " WHERE scafac.codtipom=scafac1.codtipom AND"
    SQL = SQL & " scafac.numfactu=scafac1.numfactu AND scafac.fecfactu=scafac1.fecfactu AND"
    SQL = SQL & " scafac.codclien=sdirenvio.codclien AND scafac1.coddiren = sdirenvio.coddiren"
    SQL = SQL & PonerDesdeHastaControlDir(False)
    
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set IT = lw(7).ListItems.Add()
        CargaItemDireccionEnvio IT
        miRsAux.MoveNext
        
        
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    '
    Me.lblTitulo(11).Caption = "Comprobar direcciones de envio"
End Sub

Private Sub CargaItemDireccionEnvio(ByRef elit As ListItem)
        elit.Text = miRsAux!codtipoa
        elit.SubItems(1) = Format(miRsAux!NumAlbar, "000000")
        elit.SubItems(2) = Format(miRsAux!FechaAlb, FormatoFecha)
        elit.SubItems(3) = Format(miRsAux!codClien, "00000")
        elit.SubItems(4) = miRsAux!NomClien
        elit.SubItems(5) = Format(miRsAux!coddiren, "0000")
        elit.SubItems(6) = miRsAux!nomdiren
        elit.SubItems(7) = miRsAux!codtipom_
        SQL = " "
        If Not IsNull(miRsAux!Numfactu) Then SQL = Format(miRsAux!Numfactu, "00000")
        elit.SubItems(8) = SQL
        SQL = " "
        If Not IsNull(miRsAux!FecFactu) Then SQL = Format(miRsAux!FecFactu, FormatoFecha)
        elit.SubItems(9) = SQL
        
End Sub
Private Sub AbrirFormularioDireccionEnvio()
    'InsertadoAlbaran = variable global que me da el ultimo albaran insertado
    InsertadoAlbaran = 0

    If Trim(lw(7).SelectedItem.SubItems(8)) = "" Then
        'ALBARAN
        numParam = 0
        If vParamAplic.TipoFormularioClientes = 0 Then
             With frmFacEntAlbaranes2
                .hcoCodMovim = lw(7).SelectedItem.SubItems(1)
                .hcoCodTipoM = lw(7).SelectedItem.Text
                .Show vbModal
            End With
            
        Else
            'FORMULARIO SAIL
             With frmFacEntAlbSAIL
                .hcoCodMovim = lw(7).SelectedItem.SubItems(1)
                .hcoCodTipoM = lw(7).SelectedItem.Text
                .Show vbModal
            End With
        End If
    Else
        numParam = 1
        With frmFacHcoFacturas2
            .DesdeFichaCliente = True 'Si ponemos TRUE busca directamente por numero de factura
            .hcoCodMovim = lw(7).SelectedItem.SubItems(8)
            .hcoCodTipoM = lw(7).SelectedItem.SubItems(7)
            .hcoFechaMov = lw(7).SelectedItem.SubItems(9)
            
            .Show vbModal
        End With
    End If
    
    
    'Volvemos a leer los datos de ese ITEM,
    'InsertadoAlbaran = variable global que me da el ultimo albaran insertado
    If InsertadoAlbaran = 0 Then
    
        If numParam = 0 Then
            'DESDE ALBARANES
            SQL = "SELECT codtipom as codtipoa,numalbar ,scaalb.fechaalb ,scaalb.codclien ,nomclien,"
            SQL = SQL & "scaalb.coddiren ,nomdiren ,'' codtipom_ ,null AS numfactu,null AS fecfactu"
            SQL = SQL & " From scaalb, sdirenvio WHERE scaalb.codclien=sdirenvio.codclien AND"
            SQL = SQL & " scaalb.coddiren = sdirenvio.coddiren AND scaalb.codtipom= '" & lw(7).SelectedItem.Text
            SQL = SQL & "' AND scaalb.numalbar = " & lw(7).SelectedItem.SubItems(1)
        Else
            'Desde factura
            SQL = " SELECT scafac1.codtipoa,scafac1.numalbar ,scafac1.fechaalb ,scafac.codclien ,scafac.nomclien,"
            SQL = SQL & " scafac1.coddiren ,nomdiren ,scafac.codtipom codtipom_,scafac.numfactu ,scafac.fecfactu"
            SQL = SQL & " From scafac, scafac1, sdirenvio"
            SQL = SQL & " WHERE scafac.codtipom=scafac1.codtipom AND"
            SQL = SQL & " scafac.numfactu=scafac1.numfactu AND scafac.fecfactu=scafac1.fecfactu AND"
            SQL = SQL & " scafac.codclien=sdirenvio.codclien AND scafac1.coddiren = sdirenvio.coddiren"
            SQL = SQL & " AND scafac.codtipom= '" & lw(7).SelectedItem.SubItems(7) & "' AND scafac.numfactu=" & lw(7).SelectedItem.SubItems(8)
            SQL = SQL & " AND scafac.fecfactu = '" & lw(7).SelectedItem.SubItems(9) & "'"
        End If
        Set miRsAux = New ADODB.Recordset
        miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If miRsAux.EOF Then
            NumRegElim = 0
            
        Else
            NumRegElim = 1
            CargaItemDireccionEnvio Me.lw(7).SelectedItem
        End If
        miRsAux.Close
        Set miRsAux = Nothing
        
        If NumRegElim = 0 Then
            'Ha dado error situando datos, vuelvo a cargar
            Screen.MousePointer = vbHourglass
            CargaControlDireccionesEnvio
            Screen.MousePointer = vbDefault
        End If
        

    Else
        Screen.MousePointer = vbHourglass
        CargaControlDireccionesEnvio
        
        For NumRegElim = 1 To lw(7).ListItems.Count
            If Trim(lw(7).ListItems(NumRegElim).SubItems(8)) = "" Then
                If Val(Me.lw(7).ListItems(NumRegElim).SubItems(1)) = InsertadoAlbaran Then
                    'OK: Es este. Vamos a ver
                    lw(7).ListItems(NumRegElim).Selected = True
                    lw(7).ListItems(NumRegElim).EnsureVisible
                    Set lw(7).SelectedItem = lw(7).ListItems(NumRegElim)
                    Exit For
                End If
            End If
        Next NumRegElim
        Screen.MousePointer = vbDefault
    End If
End Sub




Private Sub CargarArticulosCambio()
    
    
    SQL = RecuperaValor(vCadena, 4)
    If Val(SQL) > 0 Then
        Me.lblTitulo(19).Tag = Val(SQL)
        Me.lblTitulo(19).Caption = "->" & RecuperaValor(vCadena, 5)
        Me.lblTitulo(19).visible = True
    End If
    
    Me.lw(8).ListItems.Clear
    
    SQL = " SELECT codartic,nomartic,sartic.codmarca,nommarca,codfamia,codprove"
    SQL = SQL & " From sartic,smarca"
    SQL = SQL & " WHERE sartic.codmarca=smarca.codmarca AND " & RecuperaValor(vCadena, 1)
    
    
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set IT = lw(8).ListItems.Add()
        IT.Text = miRsAux!codArtic
        IT.SubItems(1) = miRsAux!NomArtic
        IT.SubItems(2) = miRsAux!codmarca
        IT.SubItems(3) = miRsAux!nommarca
        IT.SubItems(4) = miRsAux!Codfamia
        IT.SubItems(5) = miRsAux!Codprove
        IT.Checked = True
        
        miRsAux.MoveNext
        
        
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
  

End Sub


Private Sub CargaComboTipoComision()
Dim Age As cAgente

Set Age = New cAgente
    cboTipoPrecio2.Clear
    
    Set miRsAux = New ADODB.Recordset
    SQL = RecuperaValor(vCadena, 2)
    SQL = "Select codagent from scafac where " & SQL
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then Age.LeerDatos CStr(miRsAux!CodAgent)
    miRsAux.Close
    
    cboTipoPrecio2.AddItem "Normal"
    cboTipoPrecio2.ItemData(cboTipoPrecio2.NewIndex) = Age.ComsionNormal * 100
    
    cboTipoPrecio2.AddItem "Eco"
    cboTipoPrecio2.ItemData(cboTipoPrecio2.NewIndex) = Age.ComsionEco * 100
    
    cboTipoPrecio2.AddItem "Supereco"
    cboTipoPrecio2.ItemData(cboTipoPrecio2.NewIndex) = Age.ComsionPVPMin * 100
    
        
        
    
End Sub



Private Sub CargaAlbaranes(CargaLosTiposIVA As Boolean)
    Set miRsAux = New ADODB.Recordset
    
    
    
    
    
    'Primeros los tipos de IVA
    If CargaLosTiposIVA Then
        Me.cboTipoIVA.Clear
        cboTipoIVA.visible = False
        Label3.visible = False
        
        SQL = "select codclien from scaalb where codtipom='ALV' AND codclien =" & vCadena
        
        SQL = "select distinct(tipoiva) from sclien where codclien in (" & SQL & ") ORDER BY 1"
        miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            
            CargarComboTipoIVA miRsAux!TipoIVA
            miRsAux.MoveNext
        Wend
        miRsAux.Close
    
    
        If cboTipoIVA.ListCount > 1 Then
            cboTipoIVA.visible = True
            Label3.visible = True
        End If
        
        cboTipoIVA.ListIndex = 0
    End If
    
    SQL = "select scaalb.codtipom,scaalb.numalbar,scaalb.codclien,scaalb.nomclien,"
    SQL = SQL & " if (scaalb.dtoppago+scaalb.dtognral>0,"" * "","" "") TieneDto"
    SQL = SQL & " ,count(*) Lineas,sum(importel) Base,fechaalb"
    SQL = SQL & " from scaalb,slialb,sclien where scaalb.codtipom=slialb.codtipom and "
    SQL = SQL & " scaalb.numalbar=slialb.numalbar and scaalb.codclien=sclien.codclien "
    SQL = SQL & " AND slialb.codtipom  ='ALV'"
    SQL = SQL & " AND scaalb.codclien =" & vCadena
    SQL = SQL & " AND tipoiva =  " & Me.cboTipoIVA.ItemData(cboTipoIVA.ListIndex)
    SQL = SQL & " group by 1,2 order by fechaalb,numalbar"
    
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
                
        Set IT = lw(10).ListItems.Add()
        IT.Text = miRsAux!NumAlbar
        IT.SubItems(1) = Format(miRsAux!FechaAlb, "dd/mm/yyyy")
        IT.SubItems(2) = miRsAux!codClien
        IT.SubItems(3) = miRsAux!NomClien
        IT.SubItems(4) = miRsAux!TieneDto
        IT.SubItems(5) = Format(miRsAux!Base, FormatoImporte)
        IT.Checked = True
        miRsAux.MoveNext
        
        
        
        
    Wend
    miRsAux.Close
    
    Set miRsAux = Nothing
        
        
End Sub



Private Sub CargarComboTipoIVA(Tipo As Integer)
'Cogido de fac clientes
' -> Quitamos el cliear y añadimos segun tipo
'0-Normal, 1-Con Recargo de Equivalencia, 2-Exento de IVA

    Select Case Tipo
    
    Case 1
    
        cboTipoIVA.AddItem "Recargo Equivalencia"

    Case 2
        cboTipoIVA.AddItem "Exento de IVA"
        
    Case 3
        cboTipoIVA.AddItem "Intracomunitario"
        
    Case 4
        'Junio 2012 Reducido
        cboTipoIVA.AddItem "Reducido"
        
    Case Else 'CERO
        cboTipoIVA.AddItem "Normal"
        
    End Select

    cboTipoIVA.ItemData(cboTipoIVA.NewIndex) = Tipo
End Sub


Private Sub PonerCamposCambioProveedor()
    SQL = RecuperaValor(vCadena, 1)
    numParam = InStr(1, SQL, "=")
    SQL = Trim(Mid(SQL, numParam + 1))
    
    Label2(3).Caption = Right("00000" & SQL, 5)
    Label2(5).Caption = DevuelveDesdeBD(conAri, "nomprove", "sprove", "codprove", SQL)
    SQL = Right("00000" & RecuperaValor(vCadena, 4), 5)
    Label2(6).Caption = SQL
    Label2(8).Caption = RecuperaValor(vCadena, 5)
    Label2(10).Caption = ""
    
    Me.cmdActualizaSoloProveedor.Enabled = vUsu.Nivel <= 1
    
End Sub
