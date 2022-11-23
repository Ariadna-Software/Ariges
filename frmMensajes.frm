VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMensajes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mensajes"
   ClientHeight    =   11745
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   18240
   Icon            =   "frmMensajes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11745
   ScaleWidth      =   18240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame FramemarcarAlbaranesFacturar 
      Height          =   9975
      Left            =   600
      TabIndex        =   167
      Top             =   720
      Visible         =   0   'False
      Width           =   13575
      Begin VB.CommandButton cmdMarcarFacturar 
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   1
         Left            =   12000
         TabIndex        =   172
         Top             =   9600
         Width           =   1185
      End
      Begin VB.CommandButton cmdMarcarFacturar 
         Caption         =   "Validar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   0
         Left            =   10320
         TabIndex        =   168
         ToolTipText     =   "Validar cambios realizados"
         Top             =   9600
         Width           =   1305
      End
      Begin MSComctlLib.ListView ListView12 
         Height          =   8625
         Left            =   240
         TabIndex        =   171
         Top             =   720
         Width           =   12960
         _ExtentX        =   22860
         _ExtentY        =   15214
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
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
         NumItems        =   10
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Estado-Tipo"
            Object.Width           =   2470
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Albarán"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Fecha"
            Object.Width           =   2382
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Nombre"
            Object.Width           =   8872
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "F.Pago"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Term."
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Bases"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Orden1"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Orden2"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Orden3"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   11
         Left            =   12840
         Picture         =   "frmMensajes.frx":000C
         Top             =   360
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   10
         Left            =   12360
         Picture         =   "frmMensajes.frx":0156
         Top             =   360
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "FF"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   17
         Left            =   240
         TabIndex        =   170
         Top             =   9600
         Width           =   5895
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Albaranes"
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
         Index           =   5
         Left            =   240
         TabIndex        =   169
         Top             =   240
         Width           =   3105
      End
   End
   Begin VB.Frame FrameImportarPedidosFonteas 
      Height          =   7695
      Left            =   2040
      TabIndex        =   155
      Top             =   720
      Visible         =   0   'False
      Width           =   15015
      Begin VB.CommandButton cmdElimPedidoXLS 
         Height          =   375
         Left            =   3840
         Picture         =   "frmMensajes.frx":02A0
         Style           =   1  'Graphical
         TabIndex        =   166
         ToolTipText     =   "Eliminar pedido XLS"
         Top             =   840
         Width           =   375
      End
      Begin VB.CommandButton cmdpededioDeseExcel 
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
         Height          =   495
         Index           =   0
         Left            =   13200
         TabIndex        =   164
         Top             =   6960
         Width           =   1545
      End
      Begin VB.CommandButton cmdpededioDeseExcel 
         Caption         =   "Traer pedido"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   11400
         TabIndex        =   163
         Top             =   6960
         Width           =   1665
      End
      Begin MSComctlLib.ListView lwPedidosFontenas 
         Height          =   2655
         Index           =   0
         Left            =   240
         TabIndex        =   156
         Top             =   1320
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   4683
         View            =   3
         LabelEdit       =   1
         SortOrder       =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Pedido"
            Object.Width           =   2141
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Fecha"
            Object.Width           =   3387
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Lineas"
            Object.Width           =   2133
         EndProperty
      End
      Begin MSComctlLib.ListView lwPedidosFontenas 
         Height          =   2175
         Index           =   1
         Left            =   240
         TabIndex        =   158
         Top             =   4560
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   3836
         View            =   3
         LabelEdit       =   1
         SortOrder       =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Cliente"
            Object.Width           =   8361
         EndProperty
      End
      Begin MSComctlLib.ListView lwPedidosFontenas 
         Height          =   5415
         Index           =   2
         Left            =   5520
         TabIndex        =   159
         Top             =   1320
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   9551
         View            =   3
         LabelEdit       =   1
         SortOrder       =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Alm."
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Articulo"
            Object.Width           =   3492
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Descripcion"
            Object.Width           =   6772
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Cant."
            Object.Width           =   1923
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Lote"
            Object.Width           =   2221
         EndProperty
      End
      Begin VB.Label Label2 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   16
         Left            =   120
         TabIndex        =   165
         Top             =   7200
         Width           =   8175
      End
      Begin VB.Label Label2 
         Caption         =   "FICHEROS pendientes de procesar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   15
         Left            =   240
         TabIndex        =   162
         Top             =   4200
         Width           =   4695
      End
      Begin VB.Label Label2 
         Caption         =   "Articulos del pedido"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Index           =   14
         Left            =   5520
         TabIndex        =   161
         Top             =   960
         Width           =   2415
      End
      Begin VB.Label Label2 
         Caption         =   "Pedidos pendientes importar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Index           =   13
         Left            =   240
         TabIndex        =   160
         Top             =   960
         Width           =   3375
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Importación pedidos desde EXCEL"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   480
         Index           =   4
         Left            =   240
         TabIndex        =   157
         Top             =   240
         Width           =   7185
      End
   End
   Begin VB.Frame FrameTaxco 
      Height          =   8175
      Left            =   360
      TabIndex        =   90
      Top             =   0
      Visible         =   0   'False
      Width           =   17775
      Begin VB.CommandButton cmdTaxco 
         Height          =   495
         Index           =   2
         Left            =   16080
         Picture         =   "frmMensajes.frx":0CA2
         Style           =   1  'Graphical
         TabIndex        =   99
         ToolTipText     =   "Imprimir factura extendida"
         Top             =   180
         Width           =   495
      End
      Begin VB.CommandButton cmdTaxco 
         Height          =   495
         Index           =   3
         Left            =   14880
         Picture         =   "frmMensajes.frx":1D24
         Style           =   1  'Graphical
         TabIndex        =   98
         ToolTipText     =   "Establecer kilometros"
         Top             =   180
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton cmdTaxco 
         Height          =   495
         Index           =   1
         Left            =   15480
         Picture         =   "frmMensajes.frx":3796
         Style           =   1  'Graphical
         TabIndex        =   97
         ToolTipText     =   "Imprimir listado"
         Top             =   180
         Width           =   495
      End
      Begin VB.CommandButton cmdBusMatr 
         Height          =   495
         Left            =   6240
         Picture         =   "frmMensajes.frx":4198
         Style           =   1  'Graphical
         TabIndex        =   92
         ToolTipText     =   "Buscar"
         Top             =   120
         Width           =   495
      End
      Begin VB.TextBox txtMatr 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         TabIndex        =   91
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdTaxco 
         Height          =   495
         Index           =   0
         Left            =   17040
         Picture         =   "frmMensajes.frx":4B9A
         Style           =   1  'Graphical
         TabIndex        =   93
         ToolTipText     =   "SALIR"
         Top             =   180
         Width           =   495
      End
      Begin MSComctlLib.ListView lwTaxco 
         Height          =   7215
         Left            =   240
         TabIndex        =   94
         Top             =   720
         Width           =   17295
         _ExtentX        =   30506
         _ExtentY        =   12726
         SortKey         =   14
         View            =   3
         LabelEdit       =   1
         SortOrder       =   -1  'True
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   15
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Cliente"
            Object.Width           =   1773
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "NºFact."
            Object.Width           =   1773
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Ser"
            Object.Width           =   901
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Fecha"
            Object.Width           =   2453
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Articulo"
            Object.Width           =   3476
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Descripcion"
            Object.Width           =   6817
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Cant."
            Object.Width           =   1923
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Importe"
            Object.Width           =   2222
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Matricula"
            Object.Width           =   2170
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   9
            Text            =   "Kms"
            Object.Width           =   2134
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Observa"
            Object.Width           =   4179
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "ordenCli"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "ordenNumfac"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "ordenserie"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "order fecfactu"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Matricula"
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
         Index           =   1
         Left            =   3360
         TabIndex        =   96
         Top             =   300
         Width           =   810
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Datos reparaciones"
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
         Left            =   240
         TabIndex        =   95
         Top             =   240
         Width           =   3105
      End
   End
   Begin VB.Frame FrameAnticiposprov 
      Height          =   7455
      Left            =   -1800
      TabIndex        =   148
      Top             =   3360
      Visible         =   0   'False
      Width           =   8895
      Begin VB.CommandButton cmdAnticipoProv 
         Caption         =   "Cancelar"
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
         Index           =   0
         Left            =   7440
         TabIndex        =   150
         Top             =   6840
         Width           =   1065
      End
      Begin VB.CommandButton cmdAnticipoProv 
         Caption         =   "Aceptar"
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
         Index           =   1
         Left            =   6120
         TabIndex        =   149
         Top             =   6840
         Width           =   1065
      End
      Begin MSComctlLib.ListView ListView11 
         Height          =   5175
         Left            =   240
         TabIndex        =   151
         Top             =   960
         Width           =   8445
         _ExtentX        =   14896
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Codigo"
            Object.Width           =   2188
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Documento"
            Object.Width           =   4304
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Fecha"
            Object.Width           =   3246
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Anticipo"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label2 
         Caption         =   "Nomprove"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   12
         Left            =   240
         TabIndex        =   153
         Top             =   480
         Width           =   8175
      End
      Begin VB.Label Label2 
         Caption         =   "Marque,si lo desea, (el)los anticipos que quiera descontar en la factura "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   10
         Left            =   240
         TabIndex        =   152
         Top             =   6240
         Width           =   8175
      End
   End
   Begin VB.Frame FramePuntosCaducados 
      Height          =   6735
      Left            =   840
      TabIndex        =   142
      Top             =   240
      Width           =   14415
      Begin VB.CommandButton cmdPuntosCaducados 
         Caption         =   "Imprimir"
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
         Index           =   2
         Left            =   9840
         TabIndex        =   154
         Top             =   6120
         Width           =   1185
      End
      Begin VB.CommandButton cmdPuntosCaducados 
         Caption         =   "Caducar"
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
         Index           =   0
         Left            =   11520
         TabIndex        =   146
         Top             =   6120
         Width           =   1185
      End
      Begin VB.CommandButton cmdPuntosCaducados 
         Caption         =   "Salir"
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
         Index           =   1
         Left            =   12840
         TabIndex        =   145
         Top             =   6120
         Width           =   1185
      End
      Begin MSComctlLib.ListView ListView10 
         Height          =   5175
         Left            =   240
         TabIndex        =   143
         Top             =   720
         Width           =   13845
         _ExtentX        =   24421
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
            Text            =   "Codigo"
            Object.Width           =   2188
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Cliente"
            Object.Width           =   7832
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Puntos"
            Object.Width           =   3246
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "P.calculado"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "P.caducan"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Ult. canje"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Ult. caduca"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   9
         Left            =   13680
         Picture         =   "frmMensajes.frx":5C1C
         Top             =   360
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   8
         Left            =   13320
         Picture         =   "frmMensajes.frx":5D66
         Top             =   360
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "lbl2 9"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   9
         Left            =   240
         TabIndex        =   147
         Top             =   6120
         Width           =   5655
      End
      Begin VB.Label Label2 
         Caption         =   "Puntos caducados"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   8
         Left            =   240
         TabIndex        =   144
         Top             =   360
         Width           =   2820
      End
   End
   Begin VB.Frame FrameBloqueoEmpresas 
      Height          =   7455
      Left            =   0
      TabIndex        =   110
      Top             =   0
      Visible         =   0   'False
      Width           =   11415
      Begin VB.CommandButton cmdBlEmp 
         Caption         =   "<<"
         Height          =   375
         Index           =   3
         Left            =   5520
         TabIndex        =   116
         Top             =   3480
         Width           =   375
      End
      Begin VB.CommandButton cmdBlEmp 
         Caption         =   ">>"
         Height          =   375
         Index           =   2
         Left            =   5520
         TabIndex        =   115
         Top             =   3000
         Width           =   375
      End
      Begin VB.CommandButton cmdBlEmp 
         Caption         =   "<"
         Height          =   375
         Index           =   1
         Left            =   5520
         TabIndex        =   114
         Top             =   1920
         Width           =   375
      End
      Begin VB.CommandButton cmdBlEmp 
         Caption         =   ">"
         Height          =   375
         Index           =   0
         Left            =   5520
         TabIndex        =   113
         Top             =   1440
         Width           =   375
      End
      Begin VB.CommandButton cmdBloqEmpre 
         Caption         =   "Cancelar"
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
         Index           =   1
         Left            =   9840
         TabIndex        =   112
         Top             =   6840
         Width           =   1335
      End
      Begin VB.CommandButton cmdBloqEmpre 
         Caption         =   "Aceptar"
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
         Index           =   0
         Left            =   8400
         TabIndex        =   111
         Top             =   6840
         Width           =   1335
      End
      Begin MSComctlLib.ListView ListView99 
         Height          =   5775
         Index           =   0
         Left            =   210
         TabIndex        =   117
         Top             =   840
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   10186
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Codigo"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Empresa"
            Object.Width           =   5644
         EndProperty
      End
      Begin MSComctlLib.ListView ListView99 
         Height          =   5775
         Index           =   1
         Left            =   6240
         TabIndex        =   118
         Top             =   840
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   10186
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Codigo"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Empresa"
            Object.Width           =   5644
         EndProperty
      End
      Begin VB.Label Label41 
         Alignment       =   1  'Right Justify
         Caption         =   "Bloqueadas"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   10050
         TabIndex        =   121
         Top             =   600
         Width           =   1125
      End
      Begin VB.Label Label41 
         Caption         =   "Permitidas"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   120
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Bloqueo de empresas por usuario"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   495
         Index           =   4
         Left            =   2880
         TabIndex        =   119
         Top             =   240
         Width           =   5295
      End
   End
   Begin VB.Frame FrameArticulosProv 
      Height          =   8775
      Left            =   0
      TabIndex        =   105
      Top             =   0
      Visible         =   0   'False
      Width           =   9255
      Begin MSComctlLib.ListView ListView5 
         Height          =   6795
         Left            =   360
         TabIndex        =   108
         Top             =   900
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   11986
         View            =   3
         LabelEdit       =   1
         SortOrder       =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Codigo"
            Object.Width           =   3007
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripcion"
            Object.Width           =   8891
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Precio"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.CommandButton cmdAcepArticPro 
         Caption         =   "Aceptar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6240
         TabIndex        =   107
         Top             =   7950
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelArticPro 
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7680
         TabIndex        =   106
         Top             =   7950
         Width           =   1215
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Artículos Proveedor"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   480
         Index           =   3
         Left            =   360
         TabIndex        =   109
         Top             =   285
         Width           =   5145
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   6
         Left            =   8055
         Picture         =   "frmMensajes.frx":5EB0
         Top             =   495
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   7
         Left            =   8520
         Picture         =   "frmMensajes.frx":5FFA
         Top             =   495
         Width           =   240
      End
   End
   Begin VB.Frame FrameEtiqEstant 
      Height          =   7455
      Left            =   0
      TabIndex        =   31
      Top             =   -120
      Width           =   10535
      Begin VB.CommandButton cmdEtiqEstan 
         Caption         =   "Imprimir"
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
         Index           =   1
         Left            =   7590
         TabIndex        =   34
         Top             =   6960
         Width           =   1215
      End
      Begin VB.CommandButton cmdEtiqEstan 
         Caption         =   "Cancelar"
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
         Index           =   0
         Left            =   9030
         TabIndex        =   33
         Top             =   6960
         Width           =   1215
      End
      Begin MSComctlLib.ListView ListView3 
         Height          =   6495
         Left            =   240
         TabIndex        =   32
         Top             =   360
         Width           =   10055
         _ExtentX        =   17727
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
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Descripcion"
            Object.Width           =   8467
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Precio"
            Object.Width           =   2470
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Familia"
            Object.Width           =   4057
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Cant."
            Object.Width           =   2118
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Codalmac"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   63
         Top             =   6960
         Width           =   4095
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   1
         Left            =   600
         Picture         =   "frmMensajes.frx":6144
         Top             =   6960
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   0
         Left            =   240
         Picture         =   "frmMensajes.frx":628E
         Top             =   6960
         Width           =   240
      End
   End
   Begin VB.Frame FrameAcvitivad 
      Height          =   8775
      Left            =   3360
      TabIndex        =   100
      Top             =   120
      Visible         =   0   'False
      Width           =   9255
      Begin VB.CommandButton cmdSelActividad 
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   7680
         TabIndex        =   104
         Top             =   8040
         Width           =   1215
      End
      Begin VB.CommandButton cmdSelActividad 
         Caption         =   "Aceptar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   6240
         TabIndex        =   103
         Top             =   8040
         Width           =   1215
      End
      Begin MSComctlLib.ListView lwActividad 
         Height          =   6975
         Left            =   360
         TabIndex        =   101
         Top             =   960
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   12303
         View            =   3
         LabelEdit       =   1
         SortOrder       =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Codigo"
            Object.Width           =   1949
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripcion"
            Object.Width           =   9596
         EndProperty
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   5
         Left            =   8520
         Picture         =   "frmMensajes.frx":63D8
         Top             =   720
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   4
         Left            =   8040
         Picture         =   "frmMensajes.frx":6522
         Top             =   720
         Width           =   240
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Actividades"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   480
         Index           =   2
         Left            =   360
         TabIndex        =   102
         Top             =   240
         Width           =   5145
      End
   End
   Begin VB.Frame FrameAcercaDe 
      BorderStyle     =   0  'None
      Height          =   3855
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   6375
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cambios version: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   240
         Left            =   120
         TabIndex        =   89
         Top             =   2040
         Width           =   1875
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Haga click en este enlace  para ver los cambios"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   1920
         MousePointer    =   3  'I-Beam
         TabIndex        =   88
         ToolTipText     =   "Haga click para seguir enlace"
         Top             =   2040
         Width           =   4410
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pasaje Ventura Feliú, 13 entlo.izquierdo 2ª"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   195
         Left            =   2400
         TabIndex        =   10
         Top             =   3000
         Width           =   3855
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "46007 - VALENCIA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   195
         Left            =   3000
         TabIndex        =   9
         Top             =   3285
         Width           =   2535
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tfno:  902 88 88 78  -  96 380 55 79"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   195
         Left            =   2685
         TabIndex        =   8
         Top             =   3555
         Width           =   3165
      End
      Begin VB.Label Label27 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   195
         Left            =   4215
         TabIndex        =   7
         Top             =   3480
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Image Image2 
         Height          =   735
         Left            =   0
         Picture         =   "frmMensajes.frx":666C
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   1920
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Version"
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
         Left            =   -120
         TabIndex        =   6
         Top             =   1260
         Width           =   4155
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   81.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1725
         Left            =   4260
         TabIndex        =   5
         Top             =   0
         Width           =   1350
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "ARIGES"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   915
         Left            =   1080
         TabIndex        =   4
         Top             =   300
         Width           =   3495
      End
   End
   Begin VB.Frame FrameArticulosAgrupados 
      Height          =   6015
      Left            =   0
      TabIndex        =   68
      Top             =   1080
      Visible         =   0   'False
      Width           =   9375
      Begin VB.Frame FrameSelecArtAgrupado 
         Height          =   3255
         Left            =   1440
         TabIndex        =   75
         Top             =   1320
         Width           =   6495
         Begin VB.TextBox txtNoEditable 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   375
            Index           =   5
            Left            =   4560
            Locked          =   -1  'True
            TabIndex        =   81
            Text            =   "6"
            Top             =   2400
            Width           =   1335
         End
         Begin VB.TextBox txtNoEditable 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   4
            Left            =   2280
            Locked          =   -1  'True
            TabIndex        =   80
            Text            =   "5"
            Top             =   2400
            Width           =   1335
         End
         Begin VB.TextBox txtNoEditable 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   480
            Locked          =   -1  'True
            TabIndex        =   79
            Text            =   "4"
            Top             =   2400
            Width           =   1335
         End
         Begin VB.TextBox txtNoEditable 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   480
            Locked          =   -1  'True
            TabIndex        =   78
            Text            =   "3"
            Top             =   1560
            Width           =   5295
         End
         Begin VB.TextBox txtNoEditable 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   2280
            Locked          =   -1  'True
            TabIndex        =   77
            Text            =   "2"
            Top             =   720
            Width           =   1695
         End
         Begin VB.TextBox txtNoEditable 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   480
            Locked          =   -1  'True
            TabIndex        =   76
            Text            =   "1"
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label Label8 
            Caption         =   "TOTAL"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   4680
            TabIndex        =   87
            Top             =   2040
            Width           =   855
         End
         Begin VB.Label Label8 
            Caption         =   "PVP"
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
            Index           =   4
            Left            =   2400
            TabIndex        =   86
            Top             =   2040
            Width           =   735
         End
         Begin VB.Label Label8 
            Caption         =   "Uds"
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
            Index           =   3
            Left            =   480
            TabIndex        =   85
            Top             =   2040
            Width           =   735
         End
         Begin VB.Label Label8 
            Caption         =   "Descripción"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   2
            Left            =   480
            TabIndex        =   84
            Top             =   1200
            Width           =   1245
         End
         Begin VB.Label Label8 
            Caption         =   "Referencia"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   1
            Left            =   2280
            TabIndex        =   83
            Top             =   360
            Width           =   1170
         End
         Begin VB.Label Label8 
            Caption         =   "Id"
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
            Index           =   0
            Left            =   480
            TabIndex        =   82
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.TextBox txtArtAgrupado 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   70
         Text            =   "1"
         Top             =   5400
         Width           =   735
      End
      Begin VB.CommandButton cmdArtAgrupado 
         Caption         =   "&Salir"
         Height          =   375
         Index           =   1
         Left            =   8160
         TabIndex        =   72
         Top             =   5520
         Width           =   975
      End
      Begin VB.CommandButton cmdArtAgrupado 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   0
         Left            =   7080
         TabIndex        =   71
         Top             =   5520
         Width           =   975
      End
      Begin MSComctlLib.ListView lwArticulosAgrupados 
         Height          =   4455
         Left            =   480
         TabIndex        =   69
         Top             =   840
         Width           =   8625
         _ExtentX        =   15214
         _ExtentY        =   7858
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Caja"
            Object.Width           =   2187
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripcion"
            Object.Width           =   7408
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Referencia"
            Object.Width           =   2187
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "PVP"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label7 
         Caption         =   "CAJAS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   495
         Index           =   3
         Left            =   480
         TabIndex        =   74
         Top             =   5400
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Artículos agrupados"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   495
         Index           =   2
         Left            =   480
         TabIndex        =   73
         Top             =   240
         Width           =   5295
      End
   End
   Begin VB.Frame FrameNSeries 
      Height          =   5000
      Left            =   480
      TabIndex        =   11
      Top             =   120
      Width           =   6975
      Begin VB.CommandButton cmdSelTodos 
         Caption         =   "&Todos"
         Height          =   315
         Left            =   720
         TabIndex        =   22
         Top             =   4320
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.CommandButton cmdDeselTodos 
         Caption         =   "&Ninguno"
         Height          =   315
         Left            =   1680
         TabIndex        =   21
         Top             =   4320
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5040
         TabIndex        =   20
         Top             =   4320
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptarNSeries 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3960
         TabIndex        =   12
         Top             =   4320
         Width           =   975
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   3015
         Left            =   720
         TabIndex        =   13
         Top             =   840
         Width           =   3945
         _ExtentX        =   6959
         _ExtentY        =   5318
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label7 
         Caption         =   "Empresas en el sistema"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   495
         Index           =   1
         Left            =   720
         TabIndex        =   30
         Top             =   240
         Visible         =   0   'False
         Width           =   5295
      End
   End
   Begin VB.Frame FrameEmail 
      Height          =   6975
      Left            =   3600
      TabIndex        =   50
      Top             =   0
      Width           =   8055
      Begin VB.TextBox txmemail 
         Height          =   315
         Index           =   4
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   61
         Text            =   "Text2"
         Top             =   2160
         Width           =   2535
      End
      Begin VB.CommandButton cmdEmail 
         Caption         =   "Salir"
         Height          =   375
         Left            =   6720
         TabIndex        =   60
         Top             =   6480
         Width           =   975
      End
      Begin VB.TextBox txmemail 
         Height          =   3555
         Index           =   3
         Left            =   360
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   58
         Text            =   "frmMensajes.frx":6A8A
         Top             =   2760
         Width           =   7335
      End
      Begin VB.TextBox txmemail 
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   56
         Text            =   "Text2"
         Top             =   2160
         Width           =   4815
      End
      Begin VB.TextBox txmemail 
         Height          =   315
         Index           =   1
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   52
         Text            =   "Text2"
         Top             =   1560
         Width           =   7335
      End
      Begin VB.TextBox txmemail 
         Height          =   315
         Index           =   0
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   51
         Text            =   "Text2"
         Top             =   960
         Width           =   7335
      End
      Begin VB.Label Label5 
         Caption         =   "Fecha"
         Height          =   255
         Index           =   4
         Left            =   5280
         TabIndex        =   62
         Top             =   1920
         Width           =   2535
      End
      Begin VB.Label Label5 
         Caption         =   "Mensaje"
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   59
         Top             =   2520
         Width           =   2535
      End
      Begin VB.Label Label5 
         Caption         =   "Adjuntos"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   57
         Top             =   1920
         Width           =   2535
      End
      Begin VB.Label Label5 
         Caption         =   "Asunto"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   55
         Top             =   1320
         Width           =   2535
      End
      Begin VB.Label Label5 
         Caption         =   "Label5"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   54
         Top             =   720
         Width           =   2535
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Email CRM"
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
         Height          =   375
         Index           =   15
         Left            =   960
         TabIndex        =   53
         Top             =   240
         Width           =   5415
      End
   End
   Begin VB.Frame FrameErrores 
      Height          =   5535
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   8415
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "&Cerrar"
         Height          =   375
         Left            =   7080
         TabIndex        =   29
         Top             =   4920
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   4335
         Left            =   360
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   28
         Text            =   "frmMensajes.frx":6A90
         Top             =   360
         Width           =   7695
      End
   End
   Begin VB.Frame FrameComponentes 
      Height          =   3975
      Left            =   0
      TabIndex        =   14
      Top             =   120
      Width           =   5295
      Begin VB.CommandButton cmdAceptarComp 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3240
         TabIndex        =   19
         Top             =   3240
         Width           =   975
      End
      Begin VB.Frame FrameComponentes2 
         Caption         =   "Mostrar Equipos del :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   2175
         Left            =   960
         TabIndex        =   15
         Top             =   720
         Width           =   3255
         Begin VB.OptionButton OptCompXClien 
            Caption         =   "Cliente"
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
            Left            =   360
            TabIndex        =   18
            Top             =   1440
            Width           =   2535
         End
         Begin VB.OptionButton OptCompXDpto 
            Caption         =   "Departamento"
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
            Left            =   360
            TabIndex        =   17
            Top             =   960
            Width           =   2535
         End
         Begin VB.OptionButton OptCompXMant 
            Caption         =   "Mantenimiento"
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
            Left            =   360
            TabIndex        =   16
            Top             =   480
            Width           =   2055
         End
      End
   End
   Begin VB.Frame FrameCobrosPtes 
      Height          =   4935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8655
      Begin VB.CommandButton cmdCancelarCobros 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   7080
         TabIndex        =   25
         Top             =   4440
         Width           =   975
      End
      Begin VB.TextBox txtParam 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   1095
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   23
         Text            =   "frmMensajes.frx":6A96
         Top             =   120
         Width           =   6615
      End
      Begin VB.CommandButton cmdAceptarCobros 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   5400
         TabIndex        =   1
         Top             =   4440
         Width           =   975
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3135
         Left            =   120
         TabIndex        =   2
         Top             =   1200
         Width           =   8100
         _ExtentX        =   14288
         _ExtentY        =   5530
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "¿Desea continuar?"
         Height          =   375
         Index           =   1
         Left            =   480
         TabIndex        =   26
         Top             =   4440
         Width           =   7215
      End
      Begin VB.Label Label1 
         Caption         =   "Departamento:"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   24
         Top             =   240
         Width           =   7215
      End
   End
   Begin VB.Frame FrameErrorCC 
      Height          =   6135
      Left            =   6000
      TabIndex        =   64
      Top             =   960
      Width           =   6495
      Begin VB.TextBox txtCCError 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4335
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   67
         Text            =   "frmMensajes.frx":6A9C
         Top             =   840
         Width           =   5895
      End
      Begin VB.CommandButton cmdSalirCC 
         Caption         =   "&Salir"
         Height          =   375
         Left            =   5040
         TabIndex        =   66
         Top             =   5520
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Errores centro de coste"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   495
         Index           =   0
         Left            =   240
         TabIndex        =   65
         Top             =   360
         Width           =   5295
      End
   End
   Begin VB.Frame FramePrevisualizar 
      Height          =   11190
      Left            =   0
      TabIndex        =   122
      Top             =   0
      Width           =   16170
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   13770
         TabIndex        =   134
         Top             =   180
         Width           =   1845
      End
      Begin VB.CommandButton cmdAcepPrev 
         Caption         =   "Continuar"
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
         Left            =   13545
         TabIndex        =   133
         Top             =   10710
         Width           =   1155
      End
      Begin VB.CommandButton cmdCanPrev 
         Caption         =   "Salir"
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
         Left            =   14805
         TabIndex        =   123
         Top             =   10710
         Width           =   1065
      End
      Begin MSComctlLib.ListView ListView6 
         Height          =   4815
         Left            =   120
         TabIndex        =   124
         Top             =   585
         Width           =   7845
         _ExtentX        =   13838
         _ExtentY        =   8493
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Codigo"
            Object.Width           =   2188
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Cliente"
            Object.Width           =   7831
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Importe"
            Object.Width           =   3247
         EndProperty
      End
      Begin MSComctlLib.ListView ListView7 
         Height          =   4815
         Left            =   8055
         TabIndex        =   127
         Top             =   585
         Width           =   7845
         _ExtentX        =   13838
         _ExtentY        =   8493
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Codigo"
            Object.Width           =   2187
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Forma de Pago"
            Object.Width           =   7832
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Importe"
            Object.Width           =   3246
         EndProperty
      End
      Begin MSComctlLib.ListView ListView8 
         Height          =   4815
         Left            =   90
         TabIndex        =   128
         Top             =   5805
         Width           =   7845
         _ExtentX        =   13838
         _ExtentY        =   8493
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Codigo"
            Object.Width           =   2188
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Agente"
            Object.Width           =   7832
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Importe"
            Object.Width           =   3246
         EndProperty
      End
      Begin MSComctlLib.ListView ListView9 
         Height          =   4815
         Left            =   8055
         TabIndex        =   129
         Top             =   5805
         Width           =   7845
         _ExtentX        =   13838
         _ExtentY        =   8493
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Codigo"
            Object.Width           =   2188
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Actividad"
            Object.Width           =   7832
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Importe"
            Object.Width           =   3247
         EndProperty
      End
      Begin VB.Label Label2 
         Caption         =   "TOTAL GLOBAL:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   6
         Left            =   11970
         TabIndex        =   135
         Top             =   225
         Width           =   1740
      End
      Begin VB.Label Label2 
         Caption         =   "Totales por Actividad"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   5
         Left            =   8055
         TabIndex        =   132
         Top             =   5490
         Width           =   7725
      End
      Begin VB.Label Label2 
         Caption         =   "Totales por Agente"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   4
         Left            =   90
         TabIndex        =   131
         Top             =   5490
         Width           =   7725
      End
      Begin VB.Label Label2 
         Caption         =   "Totales por Forma de Pago"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   2
         Left            =   8055
         TabIndex        =   130
         Top             =   270
         Width           =   7725
      End
      Begin VB.Label Label2 
         Caption         =   "Totales por Cliente"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   126
         Top             =   240
         Width           =   7725
      End
      Begin VB.Label Label11 
         Caption         =   "Label3"
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
         Left            =   135
         TabIndex        =   125
         Top             =   10665
         Width           =   5055
      End
   End
   Begin VB.Frame FramePWD 
      Height          =   2640
      Left            =   0
      TabIndex        =   136
      Top             =   0
      Width           =   6720
      Begin VB.CommandButton cmdCanPWD 
         Caption         =   "Salir"
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
         Left            =   5040
         TabIndex        =   139
         Top             =   1800
         Width           =   1065
      End
      Begin VB.CommandButton cmdAcepPWD 
         Caption         =   "Aceptar"
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
         Left            =   3780
         TabIndex        =   138
         Top             =   1800
         Width           =   1065
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1620
         PasswordChar    =   "*"
         TabIndex        =   137
         Top             =   1080
         Width           =   1845
      End
      Begin VB.Label Label2 
         Caption         =   "Introduzca el pasword para continuar con el proceso"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   11
         Left            =   210
         TabIndex        =   141
         Top             =   375
         Width           =   5655
      End
      Begin VB.Label Label2 
         Caption         =   "PASSWORD:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   7
         Left            =   225
         TabIndex        =   140
         Top             =   1125
         Width           =   1740
      End
   End
   Begin VB.Frame FrameTraspasoMante 
      Height          =   3135
      Left            =   0
      TabIndex        =   43
      Top             =   0
      Width           =   5415
      Begin VB.TextBox txtMante 
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
         Left            =   2640
         TabIndex        =   48
         Top             =   1080
         Width           =   900
      End
      Begin VB.CheckBox chkMante 
         Caption         =   "Copiar importes en siguiente"
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
         Left            =   360
         TabIndex        =   46
         Top             =   1800
         Value           =   1  'Checked
         Width           =   3330
      End
      Begin VB.CommandButton cmdMante 
         Caption         =   "Cancelar"
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
         Index           =   1
         Left            =   3960
         TabIndex        =   45
         Top             =   2520
         Width           =   1065
      End
      Begin VB.CommandButton cmdMante 
         Caption         =   "Aceptar"
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
         Index           =   0
         Left            =   2745
         TabIndex        =   44
         Top             =   2520
         Width           =   1065
      End
      Begin VB.Label Label4 
         Caption         =   "Año a traspasar"
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
         Left            =   1005
         TabIndex        =   49
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Pasar importes mantenimiento a histórico."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   47
         Top             =   360
         Width           =   4695
      End
   End
   Begin VB.Frame FrameCorreccionPrecios 
      Height          =   9930
      Left            =   0
      TabIndex        =   35
      Top             =   600
      Width           =   15675
      Begin VB.ComboBox cmbActualizarTar 
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
         ItemData        =   "frmMensajes.frx":6AA2
         Left            =   10500
         List            =   "frmMensajes.frx":6AA4
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Top             =   9330
         Width           =   2175
      End
      Begin VB.CommandButton cmdCorrecotrPrecios 
         Caption         =   "Salir"
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
         Index           =   1
         Left            =   14415
         TabIndex        =   38
         Top             =   9255
         Width           =   1065
      End
      Begin VB.CommandButton cmdCorrecotrPrecios 
         Caption         =   "Aceptar"
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
         Index           =   0
         Left            =   13215
         TabIndex        =   37
         Top             =   9255
         Width           =   1065
      End
      Begin MSComctlLib.ListView ListView4 
         Height          =   8505
         Left            =   120
         TabIndex        =   36
         Top             =   600
         Width           =   15360
         _ExtentX        =   27093
         _ExtentY        =   15002
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
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
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Codigo"
            Object.Width           =   3952
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Denominación"
            Object.Width           =   7126
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "U.P.Compra"
            Object.Width           =   2541
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "% M"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "PVP"
            Object.Width           =   2364
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "%T"
            Object.Width           =   1234
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "P.Tarifa"
            Object.Width           =   2364
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "PVP Correcto"
            Object.Width           =   2717
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "Tarifa correc."
            Object.Width           =   2717
         EndProperty
      End
      Begin VB.Label Label2 
         Caption         =   "Actualizar:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   1
         Left            =   9180
         TabIndex        =   42
         Top             =   9375
         Width           =   1215
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   2
         Left            =   14640
         Picture         =   "frmMensajes.frx":6AA6
         Top             =   240
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   3
         Left            =   15240
         Picture         =   "frmMensajes.frx":6BF0
         Top             =   240
         Width           =   240
      End
      Begin VB.Label lblIndicadorCorregir 
         Caption         =   "Label3"
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
         TabIndex        =   40
         Top             =   9255
         Width           =   5055
      End
      Begin VB.Label Label2 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   39
         Top             =   240
         Width           =   8175
      End
   End
End
Attribute VB_Name = "frmMensajes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'====================== VBLES PUBLICAS ================================

Public Event DatoSeleccionado(CadenaSeleccion As String)

Public OpcionMensaje As Byte
'======================================
'==== FACTURACION =====================
' 1 .- Mensaje de Cobros Pendientes
' 2 .- Mensaje de No hay suficiente Stock para pasar de Pedido a Albaran
' 3 .- Mensaje Acerca de...
' 4 .- Listado de los Nº de Serie de un Articulo
' 5 .- Seleccionar tipo de Componente a Mostrar en Mant. de Nº de Series
' 6 .- Mostrar Prefacturacion de Albaranes
' 7 .- Mostrar Prefacturacion Mantenimientos
' 8 .- Mostrar lista clientes para seleccionar los que queremos imprimir (Etiquetas)
' 9 .- Mostrar lista Proveedores para seleccionar los que queremos imprimir (Etiquetas)
'10 .- Mostrar lista de Errores de las facturas NO contabilizadas
'11 .- Mostrar lista lineas de factura a Rectificar para seleccionar las q queremos traer al Albaran de FAct. Rectificativa
'12 .- Mostrar Albaranes del Rango que no se van a Facturar. (Facturar Albaranes Venta)

'13 .- Mostrar Errores
'14 .- Mostrar Empresas existentes en el sistema



'15 .- Mostrar lista de articulos para imprimir etiquetas estanteria
'16 .- Lista de articulos para corregir importes
'17 .- Etiquetas clientes. LO MISMO QUE EL 8 pero hecho por david
'18 .- Mantenimientos. paso ejercicio siguiente a actual

'20 .- IGual que el 16. Pero los importes son de los articulos que tienen componentes

'21 .- Ver un mensaje enlazado desde el outlook para el CRM

'22 .-  Muestra clientes potenciales

'23 .- Igual que 15. Listado PVP con IVA  (para los TPVs)

'24 .- Lineas de factura sib centro de coste

'25 .- Articulos agruopados en ventas TPV

'26 .-  Taxco Dado un cliente, mostrar datos exendidos de reparaciones

'27 .- Seleccion de activiadad

'28 .- Precios de articulos a modificar



'29 .- Previsualizacion de prefacturacion
'30 .- Introduccion del PWD

'31 .- Puntos Caducados    'AUN NO ESTA REALIZADA !!!!!!!
'32 .- Anticipos proveedores

'33.- Importar pedidos EXEL Fontenas
'34.- Marcar albaranes para facturar



'99 .- Bloqueo por empresas





Public cadWhere As String 'Cadena para pasarle la WHERE de la SELECT de los cobros pendientes o de Pedido(para comp. stock)
                          'o CodArtic para seleccionar los Nº Series
                          'para cargar el ListView
                          
Public cadWHERE2 As String

Public vCampos As String 'Articulo y cantidad Empipados para Nº de Series
                         'Tambien para pasar el nombre de la tabla de lineas (sliped, slirep,...)
                         'Dependiendo desde donde llamemos, de Pedidos o Reparaciones

Public Parametros As String


'====================== VBLES LOCALES ================================

Dim PulsadoSalir As Boolean 'Solo salir con el boton de Salir no con aspa del form
Dim PrimeraVez As Boolean

'Para los Nº de Serie
Dim TotalArray As Integer
Dim codArtic() As String
Dim cantidad() As Integer


Dim OK As Integer
Dim NE As Integer
Dim Sql As String
Dim Rs As Recordset
Dim ItmX As ListItem
Dim Errores As String
Dim Importe As Currency


Private Sub cmdAcepArticPro_Click()
Dim i As Integer
Dim Seleccionados As Integer
Dim Cad As String

    Seleccionados = 0
    Cad = ""
    
    For i = 1 To Me.ListView5.ListItems.Count
        If ListView5.ListItems(i).Checked Then
            Seleccionados = Seleccionados + 1
            Cad = Cad & ",'" & ListView5.ListItems(i).Text & "'"
        End If
    Next i
        
    RaiseEvent DatoSeleccionado(Cad)
    PulsadoSalir = False
    Unload Me
        
End Sub

Private Sub cmdAcepPrev_Click()
    PulsadoSalir = True
    Unload Me
End Sub

Private Sub cmdAcepPWD_Click()
    PulsadoSalir = True
    Unload Me
End Sub

Private Sub cmdAceptarCobros_Click()
    If OpcionMensaje = 12 Then vCampos = "1"
    Unload Me
End Sub


Private Sub cmdAceptarComp_Click()
'Boton Aceptar de Componentes del Mant. de Nº de Series en Reparaciones
Dim H As Integer, W As Integer

    ponerFrameComponentesVisible False, H, W
    PonerFrameCobrosPtesVisible True, H, W
    Me.Height = H + 350
    Me.Width = W + 70
    
    If Me.OptCompXMant.Value Then
        'Mostrar Resumen de los Nº de Serie del Mantenimiento
        Me.Caption = "Equipos del Mantenimiento"
        CargarListaComponentes (1)
    ElseIf Me.OptCompXDpto.Value Then
        'Mostrar Resumen de los Nº de Serie del Departamento
        Me.Caption = "Equipos del Departamento"
        CargarListaComponentes (2)
    ElseIf Me.OptCompXClien.Value Then
        'Mostrar Resumen de los Nº de Serie del Cliente
        Me.Caption = "Equipos del Cliente"
        CargarListaComponentes (3)
    End If
    PonerFocoBtn Me.cmdAceptarCobros
End Sub


Private Sub cmdAceptarNSeries_Click()
Dim i As Integer, J As Byte
Dim Seleccionados As Integer
Dim Cad As String, Sql As String
Dim Articulo As String
Dim Rs As ADODB.Recordset
Dim C1 As String * 10, C2 As String * 10, c3 As String * 10


    If OpcionMensaje = 4 Then
        'Comprobar que se han seleccionado el nº correcto de  Nº de Serie para cada Articulo
        Seleccionados = 0
        Articulo = ""
      
        'Si se ha seleccionado la cantidad correcta de Nº de series, empiparlos y
        'devolverlos al form de Albaranes(facturacion)
        Cad = ""
        For J = 0 To TotalArray
            Articulo = codArtic(J)
            Cad = Cad & Articulo & "|"
            For i = 1 To ListView2.ListItems.Count
                If ListView2.ListItems(i).Checked Then
                    If Articulo = ListView2.ListItems(i).ListSubItems(1).Text Then
                        If Seleccionados < Abs(cantidad(J)) Then
                            Seleccionados = Seleccionados + 1
                            Cad = Cad & ListView2.ListItems(i).Text & "|"
                        End If
                   'cad = cad & Data1.Recordset.Fields(1) & "|"
                    End If
                End If
            Next i
            If Seleccionados < Abs(cantidad(J)) Then
                'Comprobar que si tiene Nºs de serie de ese articulos cargados seleccione los
                'que corresponden
                Sql = "SELECT count(sserie.numserie)"
                Sql = Sql & " FROM sserie " 'INNER JOIN sartic ON sserie.codartic=sartic.codartic "
                Sql = Sql & " WHERE sserie.codartic=" & DBSet(Articulo, "T")
                Sql = Sql & " AND (isnull(sserie.numfactu) or sserie.numfactu='') and (isnull(sserie.numalbar) or sserie.numalbar='') "
                Sql = Sql & " ORDER BY sserie.codartic, numserie "
                Set Rs = New ADODB.Recordset
                Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                
                If Rs.Fields(0).Value >= Abs(cantidad(J)) - Seleccionados Then
                    MsgBox "Debe seleccionar " & cantidad(J) & " Nº Series para el articulo " & codArtic(J), vbExclamation
                    Exit Sub
                Else
                    'No hay Nº Serie y Pedirlos
                End If
                Rs.Close
                Set Rs = Nothing
            
            End If
            Cad = Cad & "·"
            Seleccionados = 0
        Next J
      
    ElseIf OpcionMensaje = 8 Or OpcionMensaje = 9 Or OpcionMensaje = 17 Or OpcionMensaje = 22 Then
        'concatenar todos los clientes seleccionados para imprimir etiquetas
        If OpcionMensaje = 17 Then
            
            '----------------------------------------------------------------
            '                                                      pongo numlinea cone l contador de registro como clave
            Cad = "insert into tmpnlotes (codusu,numalbar,fechaalb,numlinea,codalmac,codprove) values ("
            ' ---- [30/10/2009] (LAURA) : agrupar por cliente y departamento
'            cad = cad & vUsu.Codigo & ",1,'2005-04-12',1,"
            Cad = Cad & vUsu.Codigo & ",1,'2005-04-12',"
            
            
            For i = 1 To ListView2.ListItems.Count
                If ListView2.ListItems(i).Checked Then
                    ' ---- [30/10/2009] (LAURA) : agrupar por cliente y departamento
'                    conn.Execute cad & (ListView2.ListItems(I).Text) & ")"

                                                    
                    conn.Execute Cad & NumRegElim & "," & DBSet(ListView2.ListItems(i).ListSubItems(3).Text, "N", "S") & "," & (ListView2.ListItems(i).Text) & ")"
                    
                    NumRegElim = NumRegElim + 1
                End If
            Next i
            
            
            '----------------------------------------------------------------
            '
            ' 29/11/2010
            '
            'A partir de los datos vamos a meter en la tmpinfomres los valore
            If Not CargaDatosEtiquetas Then Exit Sub
            
        Else
            Cad = ""
            NumRegElim = 0
            For i = 1 To ListView2.ListItems.Count
                If ListView2.ListItems(i).Checked Then
                    NumRegElim = NumRegElim + 1
                    Cad = Cad & Val(ListView2.ListItems(i).Text) & ","
                     'cad = cad & Data1.Recordset.Fields(1) & "|"
                End If
            Next i
            If NumRegElim > 1000 Then
                MsgBox "Maximo número de etiquetas: 1000 (" & NumRegElim & ")", vbExclamation
                NumRegElim = 0
                Cad = ""
                Exit Sub
            End If
            NumRegElim = 0
            If Cad <> "" Then Cad = Mid(Cad, 1, Len(Cad) - 1)
        End If
    ElseIf OpcionMensaje = 11 Then
    'Lineas Factura a rectificar
        'cad = "(" & cadWHERE & ")"
        Cad = ""
        C1 = ""
        C2 = ""
        c3 = ""
        Sql = ""
        For i = 1 To ListView2.ListItems.Count
            If ListView2.ListItems(i).Checked Then
                If Sql = "" Then
                    C1 = DBSet(ListView2.ListItems(i), "T", "N")
                    C2 = ListView2.ListItems(i).ListSubItems(1)
'                    c3 = ListView2.ListItems(i).ListSubItems(2)
                    Cad = "(codtipoa=" & Trim(C1) & " and numalbar=" & Val(C2) & " and numlinea IN (" & ListView2.ListItems(i).ListSubItems(2)

                Else
                    If Trim(DBSet(ListView2.ListItems(i), "T", "N")) = Trim(C1) And Trim(ListView2.ListItems(i).ListSubItems(1)) = Trim(C2) Then
                    'es el mismo albaran y concatenamos lineas
                        Cad = "," & ListView2.ListItems(i).ListSubItems(2)

                    Else
                        If Cad <> "" Then Sql = Sql & ")) "
                        C1 = DBSet(ListView2.ListItems(i), "T", "N")
                        C2 = ListView2.ListItems(i).ListSubItems(1)
'                    c3 = ListView2.ListItems(i).ListSubItems(2)
                        Cad = " or (codtipoa=" & Trim(C1) & " and numalbar=" & Val(C2) & " and numlinea IN (" & ListView2.ListItems(i).ListSubItems(2)
                        
'                       cad=cad &
                    End If
                End If
                Sql = Sql & Cad
'                If cad <> "" Then cad = cad & " OR "
'                cad = cad & "(codtipoa=" & DBSet(ListView2.ListItems(i), "T", "N") & " and numalbar=" & Val(ListView2.ListItems(i).ListSubItems(1)) & " and numlinea=" & ListView2.ListItems(i).ListSubItems(2) & ")"
            Else
'                cad = ""
            End If
        Next i
        If Cad <> "" Then
            Sql = Sql & "))"
            Cad = "(" & cadWhere & ") AND (" & Sql & ")"
        End If
'        If cad <> "" Then cad = "(" & cadWHERE & ") AND (" & cad & ")"
    ElseIf OpcionMensaje = 14 Then
        Cad = RegresarCargaEmpresas
    End If
    
    
    
     'Actualizar la tabla sseries asignando los valores correspondientes a los
      'campos: codclien, coddirec, tieneman, codtipom, numalbar, fechavta, numline1
      'y Salir (Volver a Mto Albaranes Clientes (Facturacion)
      PulsadoSalir = True
      'RaiseEvent CargarNumSeries
      RaiseEvent DatoSeleccionado(Cad)
      Unload Me
End Sub


Private Sub cmdAnticipoProv_Click(Index As Integer)
    CadenaDesdeOtroForm = ""
    If Index = 0 Then
        If MsgBox("Desea cancelar proceso generacion factura proveedor?", vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
        CadenaDesdeOtroForm = "CANCEL"
    Else
        OK = 0
        Importe = 0
        For NE = 1 To ListView11.ListItems.Count
            If ListView11.ListItems(NE).Checked Then
                OK = OK + 1
                Importe = Importe + ImporteFormateado(ListView11.ListItems(NE).SubItems(3))
                CadenaDesdeOtroForm = CadenaDesdeOtroForm & ", " & ListView11.ListItems(NE).Text
            End If
        Next
        If OK = 0 Then
            Sql = "No compensar ningún anticipo"
        Else
            Sql = ""
           
                
            Sql = "Compensar factura sobre: " & vbCrLf & "Anticipos: " & OK & vbCrLf & "Importe : " & Space(11) & Format(Importe, FormatoImporte)
            Sql = Sql & vbCrLf & "Importe factura: " & Parametros
            
             If Importe <> CCur(Parametros) Then Sql = Sql & vbCrLf & vbCrLf & "*** Importe anticipos distinto factura ***"
                
            
        End If
        Sql = Sql & vbCrLf & vbCrLf & "¿Continuar?"
        OK = MsgBox(Sql, vbQuestion + vbYesNoCancel)
        If OK <> vbYes Then
            CadenaDesdeOtroForm = ""
            Exit Sub
        End If
    End If
    PulsadoSalir = True
    Unload Me
End Sub

Private Sub cmdArtAgrupado_Click(Index As Integer)
Dim Impor As Currency
    CadenaDesdeOtroForm = ""
    If FrameSelecArtAgrupado.visible Then
        If Index = 0 Then
            'OK este es el lote y las uds que quiere
            CadenaDesdeOtroForm = lwArticulosAgrupados.SelectedItem.Text & "|" & txtArtAgrupado.Text & "|"  'lote y uds
            Unload Me
        Else
            ponerframeTotaAgrupadoVisible False
        End If
    Else
        If Index = 0 Then
            If txtArtAgrupado.Text = "" Then txtArtAgrupado.Text = "1"
            If Me.lwArticulosAgrupados.SelectedItem Is Nothing Then Exit Sub
            
            With lwArticulosAgrupados.SelectedItem
                Me.txtNoEditable(0).Text = .Text
                Me.txtNoEditable(1).Text = .SubItems(2)
                Me.txtNoEditable(2).Text = .SubItems(1)
                Me.txtNoEditable(3).Text = txtArtAgrupado.Text
                Me.txtNoEditable(4).Text = .SubItems(3)
                Impor = ImporteFormateado(.SubItems(3))
                Impor = Impor * CInt(Me.txtArtAgrupado.Text)
                Me.txtNoEditable(5).Text = Format(Impor, FormatoImporte)
            End With
            
            ponerframeTotaAgrupadoVisible True
            
            PonerFocoBtn cmdArtAgrupado(0)
        Else
            Unload Me
        End If
    End If
End Sub

Private Sub ponerframeTotaAgrupadoVisible(visible As Boolean)
    FrameSelecArtAgrupado.visible = visible
    Me.lwArticulosAgrupados.Enabled = Not visible
    Me.txtArtAgrupado.Enabled = Not visible
End Sub

Private Sub cmdBlEmp_Click(Index As Integer)
Dim i As Integer


    Select Case Index
    Case 0, 1
        'Index Me dira que listview
        For OK = ListView99(Index).ListItems.Count To 1 Step -1
            If ListView99(Index).ListItems(OK).Selected Then
                i = ListView99(Index).ListItems(OK).Index
                PasarUnaEmpresaBloqueada Index = 0, i
            End If
        Next OK
    Case Else
        If Index = 2 Then
            OK = 0
        Else
            OK = 1
        End If
        For NumRegElim = ListView99(OK).ListItems.Count To 1 Step -1
            PasarUnaEmpresaBloqueada OK = 0, ListView99(OK).ListItems(NumRegElim).Index
        Next NumRegElim
        OK = 0
    End Select
End Sub

Private Sub PasarUnaEmpresaBloqueada(ABLoquedas As Boolean, Indice As Integer)
Dim Origen As Integer
Dim Destino As Integer
Dim IT
Dim Sql As String

    If ABLoquedas Then
        Origen = 0
        Destino = 1
        NE = 2
    Else
        Origen = 1
        Destino = 0
        NE = 1 'icono
    End If
    
    Sql = ListView99(Origen).ListItems(Indice).Key
    Set IT = ListView99(Destino).ListItems.Add(, Sql)
    IT.SmallIcon = NE
    IT.Text = ListView99(Origen).ListItems(Indice).Text
    IT.SubItems(1) = ListView99(Origen).ListItems(Indice).SubItems(1)

    'Borramos en origen
    ListView99(Origen).ListItems.Remove Indice
End Sub

Private Sub cmdBloqEmpre_Click(Index As Integer)
Dim Sql As String
Dim i As Integer

    If Index = 0 Then
        Sql = "DELETE FROM usuarios.usuarioempresasariges WHERE codusu =" & Parametros
        conn.Execute Sql
        Sql = ""
        For i = 1 To ListView99(1).ListItems.Count
            Sql = Sql & ", (" & Parametros & "," & Val(Mid(ListView99(1).ListItems(i).Key, 2)) & ")"
        Next i
        If Sql <> "" Then
            'Quitmos la primera coma
            Sql = Mid(Sql, 2)
            Sql = "INSERT INTO usuarios.usuarioempresasariges(codusu,codempre) VALUES " & Sql
            If Not ejecutar(Sql, False) Then MsgBox "Se han producido errores insertando datos", vbExclamation
        End If
    End If
    Unload Me
End Sub

Private Sub cmdBusMatr_Click()
    Screen.MousePointer = vbHourglass
    lblTitulo(0).Caption = "Leyendo datos"
    lblTitulo(0).Refresh
    CargarDatosReparaciones
    lblTitulo(0).Caption = "Historico reparaciones"
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdCancelar_Click()
    If OpcionMensaje = 4 Then
        MsgBox "Debe introducir los nº de serie necesarios para el Albaran.", vbInformation
        Exit Sub
    End If
    PulsadoSalir = True
    RaiseEvent DatoSeleccionado("")
    Unload Me
End Sub

Private Sub cmdCancelarCobros_Click()
    vCampos = "0"
    Unload Me
End Sub

Private Sub cmdCancelArticPro_Click()
    RaiseEvent DatoSeleccionado("NO")
    Unload Me
End Sub

Private Sub cmdCanPrev_Click()
    RaiseEvent DatoSeleccionado("")
    Unload Me
End Sub

Private Sub cmdCanPWD_Click()
    RaiseEvent DatoSeleccionado("")
    Unload Me
End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub cmdCorrecotrPrecios_Click(Index As Integer)
    
    If Index = 0 Then
        
        If Not ActualizarPrecios Then Exit Sub
        
    End If
    Unload Me
End Sub

Private Function ActualizarPrecios() As Boolean
Dim Sql As String
    
    
    
        
        ActualizarPrecios = False
        'Compruebo si ha seleccionado algun articulo de los de precio ultima compra=0
        cadWHERE2 = ""
        Sql = ""
        For TotalArray = 1 To Me.ListView4.ListItems.Count
            If ListView4.ListItems(TotalArray).Checked Then
                If ListView4.ListItems(TotalArray).Tag = "" Then
                    Sql = Sql & "M"
                Else
                    cadWHERE2 = cadWHERE2 & "M"
                End If
            End If
        Next
    
        If Sql <> "" Then
            MsgBox "No puede actualizar los articulos cuyo precio ultima compra sea 0", vbExclamation
            Exit Function
        End If
    
        If cadWHERE2 = "" Then
            MsgBox "Seleccione algun articulo para actualizar", vbExclamation
            Exit Function
        End If
    
        'Llegado aqui todo correcto. Hacemos la pregunta de actualizar y a correr
        Sql = "artículo"
        If Len(cadWHERE2) > 1 Then Sql = Sql & "s"
        Sql = "Va a actualizar los precios de " & Len(cadWHERE2) & " " & Sql & vbCrLf & vbCrLf & "¿Desea continuar?"
        If MsgBox(Sql, vbQuestion + vbYesNo) <> vbYes Then Exit Function
        
        
        'Aqui esta el proceso de actualizacion de articulos
        Me.lblIndicadorCorregir.Caption = "Actualización precios"
        Me.Refresh
        Espera 0.5
        
       'Para el LOG
       Sql = cadWhere & vbCrLf
       For TotalArray = 1 To Me.ListView4.ListItems.Count
            If ListView4.ListItems(TotalArray).Checked Then
                If ListView4.ListItems(TotalArray).Tag <> "" Then Sql = Sql & ListView4.ListItems(TotalArray).Text & "|"
            End If
        Next
        Sql = Mid(Sql, 1, 237)
        
        '------------------------------------------------------------------------------
        '  LOG de acciones
        Set LOG = New cLOG
        LOG.Insertar 4, vUsu, "Correccion precios: " & vbCrLf & Sql
        Set LOG = Nothing
        '-----------------------------------------------------------------------------
        
        
        
        
        
        
        
        
        
        
        For TotalArray = 1 To Me.ListView4.ListItems.Count
            If ListView4.ListItems(TotalArray).Checked Then
                If ListView4.ListItems(TotalArray).Tag <> "" Then
                    
                    'lo metemos en transaccion. Si queremos vamos
                    Me.lblIndicadorCorregir.Caption = ListView4.ListItems(TotalArray).Text
                    Me.lblIndicadorCorregir.Refresh
                    
                                        
                    conn.BeginTrans
                    If ActualizaPrecios(TotalArray) Then
                        conn.CommitTrans
                    Else
                        conn.RollbackTrans
                    End If
                    
                    
                End If
            End If
        Next
    
    
        ActualizarPrecios = True
End Function


Private Function ActualizaPrecios(NumeroItem As Integer) As Boolean

On Error GoTo EActualizaPrecios
    ActualizaPrecios = False
    With ListView4.ListItems(NumeroItem)
        If OpcionMensaje = 16 Then
            'ACtualizador de precio normal
            If Me.cmbActualizarTar.ListIndex <> 2 Then
                cadWHERE2 = TransformaComasPuntos(CStr(ImporteFormateado(.SubItems(7))))
                cadWHERE2 = "UPDATE sartic set preciove=" & cadWHERE2 & " WHERE codartic = '" & ListView4.ListItems(NumeroItem).Tag & "'"
                conn.Execute cadWHERE2
            End If
            If Me.cmbActualizarTar.ListIndex <> 1 Then
                cadWHERE2 = TransformaComasPuntos(CStr(ImporteFormateado(.SubItems(8))))
                cadWHERE2 = "UPDATE slista set precioac=" & cadWHERE2 & " WHERE codartic = '" & ListView4.ListItems(NumeroItem).Tag & "' AND codlista =" & vCampos
                conn.Execute cadWHERE2
            End If
            
        Else
            'Precio articulos componentes
            '----------------------------
            vCampos = ""
            If Me.cmbActualizarTar.ListIndex <> 2 Then
                cadWHERE2 = TransformaComasPuntos(CStr(ImporteFormateado(.SubItems(7))))
                vCampos = " preciove = " & cadWHERE2
            End If
            If Me.cmbActualizarTar.ListIndex <> 1 Then
                cadWHERE2 = TransformaComasPuntos(CStr(ImporteFormateado(.SubItems(8))))
                If vCampos <> "" Then vCampos = vCampos & ","
                vCampos = vCampos & " preciouc = " & cadWHERE2
            End If
            cadWHERE2 = "UPDATE sartic set " & vCampos & " WHERE codartic = '" & ListView4.ListItems(NumeroItem).Tag & "'"
            conn.Execute cadWHERE2
            
            
                        

            
        End If
        
    End With
        
    ActualizaPrecios = True
    Exit Function
EActualizaPrecios:
    MuestraError Err.Number, ListView4.ListItems(NumeroItem).Text
End Function


Private Sub cmdDeselTodos_Click()
Dim i As Integer

    For i = 1 To ListView2.ListItems.Count
        ListView2.ListItems(i).Checked = False
    Next i
End Sub

Private Sub cmdElimPedidoXLS_Click()
    If Me.lwPedidosFontenas(0).ListItems.Count = 0 Then Exit Sub
    If Me.lwPedidosFontenas(0).SelectedItem Is Nothing Then Exit Sub
    
    
    Sql = DevuelveDesdeBD(conAri, "count(*)", "slipedxls", "codigo", Me.lwPedidosFontenas(0).SelectedItem.Text)
    'NO PUEDE SER 0
    Sql = "Va a eliminar el pedido " & Me.lwPedidosFontenas(0).SelectedItem.Text & vbCrLf & "Lineas : " & Sql
    If MsgBox(Sql, vbQuestion + vbYesNoCancel) = vbYes Then
        Sql = "DELETE FROM slipedxls WHERE codigo= " & lwPedidosFontenas(0).SelectedItem
        conn.Execute Sql
        Me.lwPedidosFontenas(2).ListItems.Clear
        cargarDatosImportarPedidosEXCEL
    End If
    
End Sub

Private Sub cmdEmail_Click()
    Unload Me
End Sub

Private Sub cmdEtiqEstan_Click(Index As Integer)
    Screen.MousePointer = vbHourglass
    If Index = 1 Then
        If OpcionMensaje = 23 Then
            'lISTADO PRECIOS tpv
            ImprimeListadoTPV
        Else
            GenerarEtiquetasEstanterias Me.ListView3, cadWhere
            
            
            
        End If
    Else
        If TotalArray > 0 Then
            TotalArray = -1
            Exit Sub
        End If
        NumRegElim = 0
    End If
    
    Screen.MousePointer = vbDefault
    Unload Me
End Sub



Private Sub cmdMante_Click(Index As Integer)
Dim B As Boolean
    If Index = 0 Then
        
        
        If Val(txtMante(0).Text) = 0 Then
            MsgBox "El campo Año a traspasar debe ser numérico", vbExclamation
            Exit Sub
        End If
        
        
        If MsgBox("El proceso es irreversible. Continuar de igual modo?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        
        '-------------------------------------------
        Screen.MousePointer = vbHourglass
        Set miRsAux = New ADODB.Recordset
        conn.BeginTrans
        B = TraspasarMantenimientos
        Set miRsAux = Nothing
        Screen.MousePointer = vbDefault
        If B Then
            conn.CommitTrans
        Else
            conn.RollbackTrans
        End If
        
        
    End If
    Unload Me
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdMarcarFacturar_Click(Index As Integer)
    
    If Index = 0 Then
        Sql = ""
        For NE = 1 To Me.ListView12.ListItems.Count
            If ListView12.ListItems(NE).Tag <> Abs(ListView12.ListItems(NE).Checked) Then Sql = Sql & "Z"
        Next NE
        
        If Sql = "" Then
            MsgBox "Ningun cambio realizado", vbExclamation
            Exit Sub
        Else
            Sql = "Va a actualizar los albaranes modificados. (" & Len(Sql) & ")" & vbCrLf & "¿Continuar?"
            If MsgBox(Sql, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
            
            
            'Haremos dos pasadas. Poner a 1 los que estaban a 0
            'y poner a 0 los que estaban a 1
            '  SI ES QUE HAY !!!
           
            For OK = 1 To 2
                Screen.MousePointer = vbHourglass
                Sql = ""
                For NE = 1 To Me.ListView12.ListItems.Count

                    If ListView12.ListItems(NE).Tag <> Abs(ListView12.ListItems(NE).Checked) Then
                        'Se ha modificado
                        'Veamos los que tenemos que poner o a 0, o a 1 , seguna pasada bucle
                            
                        If OK = 1 Then
                            'Primera pasada. Poner marca a cero
                            If Not ListView12.ListItems(NE).Checked Then Sql = Sql & ", ('" & Trim(ListView12.ListItems(NE).Text) & "'," & Trim(ListView12.ListItems(NE).SubItems(1)) & ")"
                        Else
                            If ListView12.ListItems(NE).Checked Then Sql = Sql & ", ('" & Trim(ListView12.ListItems(NE).Text) & "'," & Trim(ListView12.ListItems(NE).SubItems(1)) & ")"
                        End If
                            
                        'ListView1.ListItems(NE).Checked
                                
                                
                                
                
                    End If
                    
                    NumRegElim = Len(Sql)
                    If NE >= ListView12.ListItems.Count Then NumRegElim = 2000
                    
                    If NumRegElim > 500 Then
                        If Sql <> "" Then
                            Screen.MousePointer = vbHourglass
                            Label2(17).Caption = "Marcar " & IIf(OK = 1, "No", "Si") & ".  Poscion: " & NE & " de " & Me.ListView12.ListItems.Count
                            Label2(17).Refresh
                            Espera 0.25
                            
                            Sql = Mid(Sql, 2)
                            vCampos = IIf(OK = 1, 0, 1)
                            vCampos = "UPDATE scaalb SET factursn =" & vCampos
                            vCampos = vCampos & " WHERE (codtipom,numalbar) IN (" & Sql & ")"
                            ejecutar vCampos, False
                            
                            Sql = ""
                        End If
                        Label2(17).Caption = ""
                        Label2(17).Refresh
                    End If
                    
                Next NE
            Next OK
         End If
         Screen.MousePointer = vbDefault
     End If
    Unload Me
End Sub

Private Sub cmdpededioDeseExcel_Click(Index As Integer)
    
    CadenaDesdeOtroForm = ""
    
    If Index = 2 Then
        
        If lwPedidosFontenas(2).ListItems.Count = 0 Then
            Sql = "Ningun dato a devolver"
        Else
            If lwPedidosFontenas(0).SelectedItem Is Nothing Then
                Sql = "Seleccione un pedido a insertar"
            Else
                If lwPedidosFontenas(0).SelectedItem Is Nothing Then
                    Sql = "Seleccione un pedido a insertar"
                Else
                    If lwPedidosFontenas(0).SelectedItem.Tag = 1 Then
                        Sql = "NO existe  el pedido: " & lwPedidosFontenas(0).SelectedItem.Text
                    Else
                        Sql = ""
                    End If
                End If
            End If
        End If
        
        If Sql = "" Then
            'Que todos los articulos del pedido XLS existen
            For NE = 1 To lwPedidosFontenas(2).ListItems.Count
                If lwPedidosFontenas(2).ListItems(NE).Tag = 1 Then
                    Sql = Sql & "X"
                    lwPedidosFontenas(2).ListItems(NE).ListSubItems(2).Bold = True
                End If
            Next
            
            If Sql <> "" Then
                Sql = "No existen articulos: " & Len(Sql) & vbCrLf & "Revise listado"
                lwPedidosFontenas(2).Refresh
            End If
        End If
        
        
        If Sql <> "" Then
        
            MsgBox Sql, vbExclamation
            Exit Sub
        End If
        
        Sql = DevuelveDesdeBD(conAri, "count(*)", "sliped", "numpedcl", Me.lwPedidosFontenas(0).SelectedItem.Text)
        NE = Val(Sql)
        
        Sql = "Va a insertar los datos importados desde la excel." & vbCrLf
        Sql = Sql & "Pedido : " & Me.lwPedidosFontenas(0).SelectedItem.Text & vbCrLf
        Sql = Sql & "Lineas: " & Me.lwPedidosFontenas(2).ListItems.Count & vbCrLf & vbCrLf
        If NE > 0 Then
            Sql = Sql & String(45, "*") & vbCrLf & "Pedido tiene " & NE & " linea(s).   Serán eliminadas" & vbCrLf & String(45, "*") & vbCrLf
        End If
        Sql = Sql & "¿ Continuar ?"
        NE = CInt(MsgBox(Sql, vbQuestion + vbYesNoCancel))
        If NE = vbCancel Then Exit Sub
        
        If NE = vbYes Then CadenaDesdeOtroForm = lwPedidosFontenas(0).SelectedItem.Text
        
        
        
        End If
    Unload Me
End Sub

Private Sub cmdPuntosCaducados_Click(Index As Integer)
    
    
    
    If Index = 2 Then
            'IMPRIMIR
            hazImprimirCaducidad
            
            Exit Sub
    End If
    
    
    
    If Index <> 1 Then
    
        NE = 0
        For NumRegElim = 1 To Me.ListView10.ListItems.Count
            If ListView10.ListItems(NumRegElim).Checked Then NE = NE + 1
        Next
        If NE = 0 Then
            MsgBox "Seleccione algún cliente ", vbExclamation
            Exit Sub
        End If
        
        Sql = "Va a caducar puntos a " & NE & " cliente" & IIf(NE > 1, "s", "") & vbCrLf & "   ¿Continuar?"
        If MsgBox(Sql, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
        
        'Caducar puntos
        Screen.MousePointer = vbHourglass
        CaducarPuntos
        Screen.MousePointer = vbDefault
    
    End If
    Unload Me
End Sub

Private Sub cmdSalirCC_Click()
    Unload Me
End Sub

Private Sub cmdSelActividad_Click(Index As Integer)
    CadenaDesdeOtroForm = ""
    If Index = 0 Then
        cadWHERE2 = "N"
        For NumRegElim = 1 To lwActividad.ListItems.Count
            
            If Not lwActividad.ListItems(NumRegElim).Checked Then
                CadenaDesdeOtroForm = CadenaDesdeOtroForm & ", 'A" & Val(lwActividad.ListItems(NumRegElim).Text) & "'" 'que quite el formato
            Else
                AnchoLogin = AnchoLogin & " - " & lwActividad.ListItems(NumRegElim).Text
                cadWHERE2 = ""
            End If
          
        Next NumRegElim
        If cadWHERE2 <> "" Then
            MsgBox "Seleccione alguna actividad", vbExclamation
            Exit Sub
        End If
        
        If CadenaDesdeOtroForm = "" Then CadenaDesdeOtroForm = "IG"  'igual  estan todas selcconads
        
    Else
        CadenaDesdeOtroForm = "NO"
    End If
    Unload Me
End Sub

Private Sub cmdSelTodos_Click()
    Dim i As Integer

    For i = 1 To ListView2.ListItems.Count
        ListView2.ListItems(i).Checked = True
    Next i
End Sub



Private Sub cmdTaxco_Click(Index As Integer)
    Select Case Index
    Case 0
        Unload Me
    Case 1
        ImprimeSeleccionReparaciones
        
    Case 2
        ImprimeFra
    Case Else
        If vUsu.Nivel = 0 Then CambiaKilometros
    End Select
End Sub

Private Sub Form_Activate()
Dim OK As Boolean
    
    
    
    Select Case OpcionMensaje
        Case 4 'Mostrar Nº Series
            If PrimeraVez Then
                PrimeraVez = False
                Me.Refresh
                Screen.MousePointer = vbHourglass
                OK = ObtenerTamanyosArray
                If OK Then OK = SeparaCampos
                If Not OK Then
                    'Error en SQL
                    'Salimos
                    Unload Me
                    Exit Sub
                End If
                CargarListaNSeries
            End If
            
        Case 8, 9, 17, 22 'Etiquetas de clientes/Proveedores
            CargarListaClientes
'        Case 10 'Errores al contabilizar facturas
'            CargarListaErrContab
        Case 11 'Lineas Factura a rectificar
            CargarListaLinFactu
            
        Case 14 'Mostrar Empresas del sistema
            CargarListaEmpresas
            
        Case 15
            'Etiquetas estanteria
            CargarArticulosEstanteria
            
        Case 16, 20
            'Articulos para corregir
            If OpcionMensaje = 16 Then
                CargarArticulosCorreccionPrecio
            Else
                CargaPVPPreciosArticulosConComponentes
            End If
            If Me.ListView4.ListItems.Count = 0 Then
                MsgBox "Ningún dato para mostrar", vbExclamation
                Unload Me
            End If
        Case 18
            PonerFoco txtMante(0)
        Case 21
            CargarEmail
        Case 23
            CargarPVPArticulos   'aqui aqui auqi
            
        Case 24
            txtCCError.Text = vCampos
            vCampos = ""
            
        Case 25
            CargaArticulosAgrupados
        Case 26
        
        Case 27
            CargaListActividades
                        
        Case 28
            CargaListArticulosProv
            
        Case 29 ' previsualizacion de facturacion
            cargaFactPrevisualizacion
                        
        Case 30 ' introduccion del password
            PonerFoco Text3
            
        Case 31
            If PrimeraVez Then
                PrimeraVez = False
                CargaPuntosCaducados
            End If
            
        Case 32
            CargaAnticipoProveedor
            
        Case 33
            Screen.MousePointer = vbHourglass
            Label2(16).Caption = "Leyendo datos .."
            Label2(16).Refresh
            cargarDatosImportarPedidosEXCEL
            
        Case 34
            
            Screen.MousePointer = vbHourglass
            Label2(16).Caption = "Leyendo datos .."
            Label2(16).Refresh
            cargarDatosAlbaranes
        
        Case 99
            cargaempresasbloquedas
                        
                        
    End Select
    
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
Dim H As Integer, W As Integer
Dim Cad As String
On Error Resume Next

    Me.FrameCobrosPtes.visible = False
    Me.FrameAcercaDe.visible = False
    Me.FrameNSeries.visible = False
    Me.FrameComponentes.visible = False
    Me.FrameComponentes2.visible = False
    Me.FrameErrores.visible = False
    FrameEtiqEstant.visible = False
    FrameCorreccionPrecios.visible = False
    FrameTraspasoMante.visible = False
    FrameEMail.visible = False
    FrameErrorCC.visible = False
    FrameArticulosAgrupados.visible = False
    Me.FrameTAXCO.visible = False
    FrameAcvitivad.visible = False
    FrameArticulosProv.visible = False
    FrameBloqueoEmpresas.visible = False
    FramePrevisualizar.visible = False
    FramePWD.visible = False
    FramePuntosCaducados.visible = False
    FrameAnticiposprov.visible = False
    FrameImportarPedidosFonteas.visible = False
    FramemarcarAlbaranesFacturar.visible = False
    
    PulsadoSalir = True
    PrimeraVez = True
    
    Select Case OpcionMensaje
        Case 1 'Mensaje de Cobros Pendientes
            PonerFrameCobrosPtesVisible True, H, W
            CargarListaCobrosPtes
            Me.Caption = "Cobros Pendientes"
            PonerFocoBtn Me.cmdAceptarCobros
            
        Case 2 'Mensaje de no hay suficiente Stock
            PonerFrameCobrosPtesVisible True, H, W
            CargarListaArtSinStock (vCampos)
            Me.Caption = "Artículos sin stock suficiente"
            PonerFocoBtn Me.cmdAceptarCobros
            
        Case 3 'Mensaje ACERCA DE
            CargaImagen
            Me.Caption = "Acerca de ....."
            PonerFrameAcercaDeVisible True, H, W
            vCampos = ""
            PonerFechaArchivo
            If vCampos = "" Then
                vCampos = "Versión:  "
            Else
                vCampos = vCampos & "         ver:"
            End If
            Me.lblVersion.Caption = vCampos & App.Major & "." & App.Minor & "." & App.Revision & " "
        
        Case 4 'Listado Nº Series Articulo
            PonerFrameNSeriesVisible True, H, W
            Me.Caption = "Nº Serie"
            Me.Label7(1).Caption = "Seleccione los Nº de serie para el Albaran."
            Me.Label7(1).FontSize = 12
            PulsadoSalir = False
            
        Case 5 'Seleccionar tipo de Componente que queremos mostrar en Resumen
                'En mant. de Nº Series de Reparacion
            ponerFrameComponentesVisible True, H, W
            Me.Caption = "Componentes"
            Me.OptCompXMant.Value = True
            PonerFocoBtn Me.cmdAceptarComp
        
        Case 6 'Mostrar Prefacturacion de Albaranes
            PonerFrameCobrosPtesVisible True, H, W
            CargarListaPreFacturar
            Me.Caption = "Prefacturación Albaranes"
            Cad = RecuperaValor(vCampos, 1)
            If Cad <> "" Then Cad = Mid(Cad, 1, Len(Cad) - 1)
            Me.txtParam.Text = Cad
            Cad = RecuperaValor(vCampos, 2)
            If Cad <> "" Then
                Cad = Mid(Cad, 1, Len(Cad) - 1)
                If Trim(Me.txtParam.Text) <> "" Then
                    txtParam.Text = Me.txtParam.Text & vbCrLf & Cad
                Else
                    txtParam.Text = Cad
                End If
            End If
            Cad = RecuperaValor(vCampos, 3)
            If Cad <> "" Then
                Cad = Mid(Cad, 1, Len(Cad) - 1)
                If Trim(Me.txtParam.Text) <> "" Then
                    txtParam.Text = Me.txtParam.Text & vbCrLf & Cad
                Else
                    txtParam.Text = Cad
                End If
            End If
            
            PonerFocoBtn Me.cmdAceptarComp
            
        Case 8, 17, 22 'Etiquetas de Clientes
            PonerFrameNSeriesVisible True, H, W
            Me.Caption = "Clientes"
            If OpcionMensaje = 22 Then Me.Caption = Me.Caption & " potenciales"
            
            
            PulsadoSalir = False
            Me.cmdSelTodos.visible = True
            Me.cmdDeselTodos.visible = True
            
        Case 9 'Etiquetas de Proveedores
            PonerFrameNSeriesVisible True, H, W
            Me.Caption = "Proveedores"
            PulsadoSalir = False
            Me.cmdSelTodos.visible = True
            Me.cmdDeselTodos.visible = True
        
        Case 10 'Errores al contabilizar facturas
            PonerFrameCobrosPtesVisible True, H, W
            CargarListaErrContab
            Me.Caption = "Facturas NO contabilizadas: "
            PonerFocoBtn Me.cmdAceptarCobros
        
        Case 11 'Lineas Factura a Rectificar
            PonerFrameNSeriesVisible True, H, W
            Me.Caption = "Lineas Factura a Rectificar"
            PulsadoSalir = False
            Me.cmdSelTodos.visible = True
            Me.cmdDeselTodos.visible = True
            Me.cmdAceptarNSeries.Left = Me.cmdAceptarNSeries.Left + 1000
            Me.cmdCancelar.Left = Me.cmdCancelar.Left + 1000
        
        Case 12 'Mensaje Albaranes que no se van a Facturar
            PonerFrameCobrosPtesVisible True, H, W
            CargarListaAlbaranes
            Me.Caption = "Facturación Albaranes"
            Me.Label1(0).Caption = "Existen Albaranes que NO se van a Facturar:"
            Me.Label1(0).Top = 260
            Me.Label1(0).Left = 480
            PonerFocoBtn Me.cmdAceptarCobros
            
        Case 13 'Muestra Errores
            H = 6000
            W = 8800
            PonerFrameVisible Me.FrameErrores, True, H, W
            Me.Text1.Text = vCampos
            Me.Caption = "Errores"
        
        Case 14 'Muestra Empresas del sistema
            PonerFrameNSeriesVisible True, H, W
            Me.Caption = "Selección"
            CargarListaEmpresas
        Case 15, 23
            'Etiquetas estanteria
            'PVP para TPVs
            H = FrameEtiqEstant.Height
            W = FrameEtiqEstant.Width
            PonerFrameVisible FrameEtiqEstant, True, H, W
            If OpcionMensaje = 23 Then
                ListView3.ColumnHeaders(3).Text = "Precio tarifa"
                ListView3.ColumnHeaders(3).Alignment = lvwColumnRight
            End If
            
            
        Case 16, 20
            
            
            Caption = "Corrección precios"
            H = FrameCorreccionPrecios.Height
            W = FrameCorreccionPrecios.Width
            PonerFrameVisible FrameCorreccionPrecios, True, H, W
            Me.cmdCorrecotrPrecios(1).Cancel = True
            lblIndicadorCorregir.Caption = ""
            CargaComboActualizarPrecios
            If OpcionMensaje = 20 Then
                ListView4.ColumnHeaders(9).Text = " PUC correc."
                Label2(0).Caption = " Corrección de precios de articulos con componentes"
            Else
                ListView4.ColumnHeaders(9).Text = "Tarifa correc."
                Label2(0).Caption = " Corrección de errores y actualización de tarifas"
            End If
            
        Case 18
            
            Caption = "Mantenimientos"
            H = FrameTraspasoMante.Height
            W = FrameTraspasoMante.Width
            PonerFrameVisible FrameTraspasoMante, True, H, W
            
        Case 21
            'Ver email
            limpiar Me
            H = FrameEMail.Height
            W = FrameEMail.Width
            PonerFrameVisible FrameEMail, True, H, W
            If cadWHERE2 = "0" Then
                Caption = "Enviados"
                Label5(0).Caption = "Para"
            Else
                Label5(0).Caption = "De"
                Caption = "Recibidos"
            End If
            cmdEmail.Cancel = True
            PonerFocoBtn Me.cmdEmail
            
        Case 24
            Caption = "Analítica"
            H = FrameErrorCC.Height
            W = FrameErrorCC.Width
            PonerFrameVisible FrameErrorCC, True, H, W
            PonerFocoBtn cmdSalirCC
        Case 25
            Caption = "LOTES"
            
            H = FrameArticulosAgrupados.Height
            W = FrameArticulosAgrupados.Width
            PonerFrameVisible FrameArticulosAgrupados, True, H, W
            FrameSelecArtAgrupado.visible = False
            cmdArtAgrupado(1).Cancel = True
        Case 26
            Caption = "Historico"
             
            H = FrameTAXCO.Height
            W = FrameTAXCO.Width
            PonerFrameVisible FrameTAXCO, True, H, W
            'cmdTaxco.Picture = frmPpal.imgListComun.ListImages(15).Picture
            cmdTaxco(0).Cancel = True
            
            cmdTaxco(3).visible = vUsu.Nivel = 0
            
        Case 27
            H = FrameAcvitivad.Height
            W = FrameAcvitivad.Width
            PonerFrameVisible FrameAcvitivad, True, H, W
            
        Case 28
            H = FrameArticulosProv.Height
            W = FrameArticulosProv.Width
            PonerFrameVisible FrameArticulosProv, True, H, W
        
        Case 29 ' previsualizacion de prefacturacion
            H = FramePrevisualizar.Height
            W = FramePrevisualizar.Width
            PonerFrameVisible FramePrevisualizar, True, H, W
            PulsadoSalir = False
            Caption = "Previsión facturación"
            
        Case 30 ' peticion de pwd
            H = FramePWD.Height
            W = FramePWD.Width
            PonerFrameVisible FramePWD, True, H, W
            PulsadoSalir = False
        
        
        Case 31 '
            H = FramePuntosCaducados.Height
            W = FramePuntosCaducados.Width
            PonerFrameVisible FramePuntosCaducados, True, H, W
            PulsadoSalir = True
            Label2(9).Caption = ""
            Me.Caption = "Puntos."
         
        Case 32
            PulsadoSalir = False
             H = FrameAnticiposprov.Height
            W = FrameAnticiposprov.Width
            PonerFrameVisible FrameAnticiposprov, True, H, W
            Label2(9).Caption = ""
            Me.Caption = "Anticipo proveedor"
            CadenaDesdeOtroForm = ""
        
        
        Case 33
            Me.Tag = "C:\PedidosAriadna"
            Me.lwPedidosFontenas(0).SmallIcons = frmPpal.ImgListPpal
            Me.lwPedidosFontenas(1).SmallIcons = frmPpal.ImgListPpal
             H = FrameImportarPedidosFonteas.Height
            W = FrameImportarPedidosFonteas.Width
            PonerFrameVisible FrameImportarPedidosFonteas, True, H, W
            Label2(16).Caption = ""
            Me.Caption = "Importador pedidos"
            CadenaDesdeOtroForm = ""
            
            
        Case 34
        
            Label2(17).Caption = "Marcar para facturar"
        
            W = Me.FramemarcarAlbaranesFacturar.Width
            H = Me.FramemarcarAlbaranesFacturar.Height + 300
        
            PonerFrameVisible FramemarcarAlbaranesFacturar, True, H, W
            Me.cmdMarcarFacturar(1).Cancel = True
        
        Case 99 ' bloqueo por empresa
            Me.FrameBloqueoEmpresas.visible = True
            Caption = "Bloqueo empresas"
            W = Me.FrameBloqueoEmpresas.Width
            H = Me.FrameBloqueoEmpresas.Height + 300
            'Como cuando venga por esta opcion, viene llamado desde el manteusu
            Me.ListView99(0).SmallIcons = frmMantenusu2.ImageList1
            Me.ListView99(1).SmallIcons = frmMantenusu2.ImageList1
            Me.cmdBloqEmpre(1).Cancel = True
        
    End Select
    'Me.cmdCancel(indFrame).Cancel = True
    Me.Height = H + 350
    Me.Width = W + 70
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub PonerFrameCobrosPtesVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Pone el Frame de Cobros Pendientes Visible y Ajustado al Formulario, y visualiza los controles
'necesario para el Informe

    H = 4600
        
    Select Case OpcionMensaje
        Case 1
            H = 5000
            W = 8600
            Me.Label1(0).Caption = "CLIENTE: " & vCampos
        Case 2
            W = 8800
            Me.cmdAceptarCobros.Top = 4000
            Me.cmdAceptarCobros.Left = 4200
        Case 5 'Componentes
            W = 6000
            H = 5000
            Me.cmdAceptarCobros.Left = 4000

        Case 6, 7 'Prefacturar Albaranes
            W = 7000
            H = 6000
            Me.cmdAceptarCobros.Top = 5400
            Me.cmdAceptarCobros.Left = 4600

        Case 10, 12 'Errores al contabilizar facturas
            H = 6000
            W = 8400
            Me.cmdAceptarCobros.Top = 5300
            Me.cmdAceptarCobros.Left = 4900
            If OpcionMensaje = 12 Then
                Me.cmdCancelarCobros.Top = 5300
                Me.cmdCancelarCobros.Left = 4600
                Me.cmdAceptarCobros.Left = 3300
                Me.Label1(1).Top = 4800
                Me.Label1(1).Left = 3400
                Me.cmdAceptarCobros.Caption = "&SI"
                Me.cmdCancelarCobros.Caption = "&NO"
            End If
    End Select
            
    PonerFrameVisible Me.FrameCobrosPtes, visible, H, W

    If visible = True Then
        Me.txtParam.visible = (OpcionMensaje = 6 Or OpcionMensaje = 7)
        Me.Label1(0).visible = (OpcionMensaje = 1) Or (OpcionMensaje = 5) Or (OpcionMensaje = 12)
        Me.cmdCancelarCobros.visible = (OpcionMensaje = 12)
        Me.Label1(1).visible = (OpcionMensaje = 12)
    End If
End Sub


Private Sub PonerFrameAcercaDeVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Pone el Frame ACERCA DE visible y Ajustado al Formulario

    Me.FrameAcercaDe.visible = visible
    If visible = True Then
        'Ajustar Tamaño del Frame para ajustar tamaño de Formulario al del Frame
        Me.FrameAcercaDe.Top = -90
        Me.FrameAcercaDe.Left = 0
        Me.FrameAcercaDe.Height = 4555
        Me.FrameAcercaDe.Width = 6600
        
        W = Me.FrameAcercaDe.Width
        H = Me.FrameAcercaDe.Height
    End If
End Sub


Private Sub PonerFrameNSeriesVisible(visible As Boolean, H As Integer, W As Integer)
'Pone el Frame de Nº Serie Visible y Ajustado al Formulario, y visualiza los controles
'necesario para el Informe

    H = 5000
   
    If OpcionMensaje = 11 Then 'Lineas Factura a Rectificar
        W = 10900
    ElseIf OpcionMensaje = 14 Then
        W = 6500
        Me.Label7(1).visible = True
    ElseIf OpcionMensaje = 17 Then
        W = 10500
        Me.Label7(1).visible = False
    Else
        W = 8500
        Me.Label7(1).visible = False
    End If
    PonerFrameVisible Me.FrameNSeries, visible, H, W
End Sub


Private Sub ponerFrameComponentesVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Pone el Frame de Componentes Visible y Ajustado al Formulario, y visualiza los controles
'necesario para el Informe

'    Me.FrameComponentes.visible = visible
    Me.FrameComponentes2.visible = visible
    
    H = 4000
    W = 5300
    PonerFrameVisible Me.FrameComponentes, visible, H, W
        
    'Ajustar Tamaño del Frame para ajustar tamaño de Formulario al del Frame
    If visible Then Me.OptCompXDpto.Caption = DevuelveTextoDepto(False)
        
    
End Sub


Private Sub CargarListaCobrosPtes()
'Muestra la lista Detallada de cobros en un ListView
'Carga los valores de la tabla scobro de la Contabilidad
Dim Rs As ADODB.Recordset
Dim ItmX As ListItem
Dim Sql As String
Dim Impor As Currency
Dim Borrame As Currency


    If vParamAplic.ContabilidadNueva Then
        Sql = "SELECT numserie,numfactu,fecfactu,fecvenci,impvenci,impcobro,gastos FROM "
        Sql = Sql & " cobros INNER JOIN formapago ON cobros.codforpa=formapago.codforpa "
    Else
        Sql = "SELECT numserie, codfaccl, fecfaccl, fecvenci, impvenci, impcobro ,gastos FROM "
        Sql = Sql & " scobro INNER JOIN sforpa ON scobro.codforpa=sforpa.codforpa "
    End If
    Sql = Sql & cadWhere

    Set Rs = New ADODB.Recordset
    Rs.Open Sql, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
     
    ListView1.Top = 900
    ListView1.Height = 3250
    ListView1.Width = 8100
    ListView1.Left = 160
    
    'Los encabezados
    ListView1.ColumnHeaders.Clear
    
    ListView1.ColumnHeaders.Add , , "Serie", 600
    ListView1.ColumnHeaders.Add , , "Nº Factura", 1000, 1
    ListView1.ColumnHeaders.Add , , "Fecha Factura", 1200, 2
    ListView1.ColumnHeaders.Add , , "Fecha Venci.", 1200, 2
    ListView1.ColumnHeaders.Add , , "Imp. Venci.(€)", 1250, 1
    ListView1.ColumnHeaders.Add , , "Imp. Cobro(€)", 1250, 1
    ListView1.ColumnHeaders.Add , , "Pte. Cobro(€)", 1250, 1
   ' Borrame = 0
    While Not Rs.EOF
        Set ItmX = ListView1.ListItems.Add
        ItmX.Text = Rs.Fields(0).Value 'Nº Serie
        ItmX.SubItems(1) = Rs.Fields(1).Value 'Nº Factura
        ItmX.SubItems(2) = Rs.Fields(2).Value 'Fecha Factura
        ItmX.SubItems(3) = Rs.Fields(3).Value 'Fecha Vencimiento
        ItmX.SubItems(4) = Rs.Fields(4).Value 'Importe Vencido
        ItmX.SubItems(5) = DBLet(Rs.Fields(5).Value, "N") 'Importe Cobrado
        'ItmX.SubItems(6) = RS.Fields(4).Value + DBLet(RS!gastos, "N") - DBLet(RS.Fields(5).Value, "N") 'Pendiente de cobro
        Impor = Rs.Fields(4).Value + DBLet(Rs!gastos, "N") - DBLet(Rs.Fields(5).Value, "N") 'Pendiente de cobro
        ItmX.SubItems(6) = Impor
        If ItmX.SubItems(6) > 0 Then
            ItmX.ListSubItems(6).ForeColor = vbRed
           ' Borrame = Borrame + Impor
        End If
        
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
End Sub


Private Sub CargarListaArtSinStock(NomTabla As String)
'Muestra la lista Detallada de Articulos que no tienen stock suficiente en un ListView
Dim Rs As ADODB.Recordset
Dim Sql As String
    
    Sql = "SELECT " & NomTabla & ".codalmac," & NomTabla & ".codartic, " & NomTabla & ".nomartic, salmac.canstock as canstock, SUM(cantidad) as cantidad, canstock-SUM(cantidad) as disp ,conjunto "
    Sql = Sql & " FROM ((" & NomTabla & " INNER JOIN sartic ON " & NomTabla & ".codartic=sartic.codartic) INNER JOIN sfamia ON sartic.codfamia=sfamia.codfamia) "
    Sql = Sql & " INNER JOIN salmac ON " & NomTabla & ".codalmac=salmac.codalmac and " & NomTabla & ".codartic=salmac.codartic "
    Sql = Sql & cadWhere 'Where numpedcl = 2 And sfamia.instalac = 0
    Sql = Sql & " GROUP by " & NomTabla & ".codalmac, " & NomTabla & ".codartic "
    

    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
     
    Me.ListView1.Top = 500
     
    'Los encabezados
    ListView1.Width = 8400
    ListView1.Height = 3150
    ListView1.ColumnHeaders.Clear
    
    ListView1.ColumnHeaders.Add , , "Alm.", 500
    ListView1.ColumnHeaders.Add , , "Articulo", 1800, 2
    ListView1.ColumnHeaders.Add , , "Dec. Artic", 3300
    ListView1.ColumnHeaders.Add , , "Stock", 950, 2
    ListView1.ColumnHeaders.Add , , "Cantidad", 900, 2
    ListView1.ColumnHeaders.Add , , "No Disp.", 900, 2
    
    While Not Rs.EOF
        CargaItemStock Rs, ""
        'Si no tiene produccion miraremos si es conjunto
        If Not vParamAplic.Produccion Then
            If Rs!Conjunto = 1 Then
                Sql = Rs!codAlmac & "|" & Rs!codArtic & "|" & Rs!cantidad & "|"
                CargaStockConjuntos Sql
            End If
        End If
        Rs.MoveNext
    Wend
    Rs.Close
    
    Set Rs = Nothing

    
    
End Sub
    
Private Sub CargaStockConjuntos(linea As String)
    
        
        Set miRsAux = New ADODB.Recordset
            'Deberiamos cargar los elementos que tiene subconjuntos
            cadWHERE2 = "SELECT " & RecuperaValor(linea, 1) & ",sarti1.codarti1,nomartic,"
            cadWHERE2 = cadWHERE2 & " sarti1.cantidad * " & TransformaComasPuntos(RecuperaValor(linea, 3)) & " as cantidad,"
            cadWHERE2 = cadWHERE2 & " salmac.canstock as canstock,  canstock-(sarti1.cantidad * " & TransformaComasPuntos(RecuperaValor(linea, 3))
            cadWHERE2 = cadWHERE2 & ") as disp From sarti1, salmac, sartic"
            cadWHERE2 = cadWHERE2 & " Where sarti1.codarti1 = salmac.codArtic And sarti1.codarti1 = sartic.codArtic"
            cadWHERE2 = cadWHERE2 & " and sarti1.codartic='" & DevNombreSQL(RecuperaValor(linea, 2)) & "'"
            
            miRsAux.Open cadWHERE2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not miRsAux.EOF
                CargaItemStock miRsAux, " * "
                miRsAux.MoveNext
            Wend
            miRsAux.Close
        
        cadWHERE2 = ""
    Set miRsAux = Nothing
End Sub
 
    
Private Sub CargaItemStock(ByRef R As ADODB.Recordset, ByRef TxtAñadido As String)
Dim ItmX As ListItem
     If R!disp < 0 Then
            Set ItmX = ListView1.ListItems.Add
            ItmX.Text = Format(R.Fields(0).Value, "000") 'Cod Almacen
            If TxtAñadido <> "" Then TxtAñadido = "[" & TxtAñadido & "]"
            ItmX.SubItems(1) = TxtAñadido & " " & R.Fields(1).Value 'Cod Artic
            ItmX.SubItems(2) = R.Fields(2).Value 'Nom Artic
            ItmX.SubItems(3) = R.Fields(3).Value 'Stock
            ItmX.SubItems(4) = R.Fields(4).Value 'Cantidad
            ItmX.SubItems(5) = R.Fields(5).Value 'No Disp
    End If
End Sub


Private Sub CargarListaNSeries()
'Carga las lista con todos los Nº de serie encontrados en la tabla:sserie
'para el articulo pasado como parametro en la cadwhere: "codartic='00012'"
'y que esten disponibles: numfactu y numalbar no tengan valor
Dim Rs As ADODB.Recordset
Dim ItmX As ListItem
Dim Sql As String
Dim cadLista As String
Dim Dif As Single

    On Error GoTo ECargarLista

    If cadWHERE2 = "" Then
        'Mostramos los nº serie libres para seleccionar la cantidad
        Sql = "SELECT sserie.numserie, sserie.codartic, sartic.nomartic "
        Sql = Sql & "FROM sserie INNER JOIN sartic ON sserie.codartic=sartic.codartic "
        Sql = Sql & cadWhere 'Where codartic='000012'
        'seleccionamos los que no esten asignados a ninguna factura ni albaran
        Sql = Sql & " AND ((isnull(sserie.numfactu) or sserie.numfactu='') and (isnull(sserie.numalbar) or sserie.numalbar='')) "
        Sql = Sql & " ORDER BY sserie.codartic, numserie "
        
    Else 'venimos de modificar la cantidad y seleccionamos los ya asignados
        If InStr(1, cadWHERE2, "|") > 0 Then
            Dif = CSng(RecuperaValor(cadWHERE2, 1))
            cadWHERE2 = RecuperaValor(cadWHERE2, 2)
        
            'seleccionamos nº serie del albaran que modificamos
            Sql = "SELECT sserie.numserie, sserie.codartic, sartic.nomartic "
            Sql = Sql & "FROM sserie INNER JOIN sartic ON sserie.codartic=sartic.codartic "
            Sql = Sql & cadWHERE2
                
            
            If Dif < 0 Then
                'Si la diferencia de cantidad es < 0, mostrar en la lista los nº serie que
                'tiene la linea de albaran asignado con todos marcados y desmarcar el que no queremos
                
            Else
                'si la diferencia de cantidad es > 0, mostrar en la lista los nº de serie que
                'ya tenia asignados la linea del albaran más los libres para seleccionar los que añadimos de mas
                cadLista = ""
                Set Rs = New ADODB.Recordset
                Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                While Not Rs.EOF
                    cadLista = cadLista & ", " & Rs!numSerie
                    Rs.MoveNext
                Wend
                Rs.Close
                Set Rs = Nothing
                
                'mostrar tambien los nº serie sin asignar
                Sql = Sql & " OR (" & Replace(cadWhere, "WHERE", "") & " and (numalbar=''or isnull(numalbar)))"
            End If
        Else
            'viene de una factura rectificativa, seleccionamos los nº de serie de
            'esa factura y marcamos los que queremos quitar
            Sql = cadWHERE2
        End If
    End If
    

    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    'Los encabezados
    ListView2.Width = 7400
    Me.ListView2.Height = 3100
    Me.ListView2.Left = 650
    ListView2.ColumnHeaders.Clear
    
    ListView2.ColumnHeaders.Add , , "Nº Serie", 1800
    ListView2.ColumnHeaders.Add , , "Articulo", 1800
    ListView2.ColumnHeaders.Add , , "Desc. Artic", 3650
        
    If Rs.EOF Then Unload Me
    
    While Not Rs.EOF
         Set ItmX = ListView2.ListItems.Add
         ItmX.Text = Rs.Fields(0).Value 'num serie
         If Dif < 0 Then
            ItmX.Checked = True
         ElseIf Dif > 0 Then
            If InStr(1, cadLista, CStr(Rs!numSerie)) > 0 Then
                ItmX.Checked = True
            Else
                ItmX.Checked = False
            End If
         Else
            ItmX.Checked = False
         End If
         ItmX.SubItems(1) = Rs.Fields(1).Value 'Desc Artic
         ItmX.SubItems(2) = Rs.Fields(2).Value 'Nom Artic
         Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    
    
ECargarLista:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargar Nº Series", Err.Description
End Sub


Private Sub CargarListaComponentes(opt As Byte)
'Muestra la lista Detallada de cobros en un ListView
'Carga los valores de la tabla scobro de la Contabilidad
Dim Rs As ADODB.Recordset
Dim ItmX As ListItem
Dim Sql As String
Dim Codigo As String, cadCodigo As String

    Select Case opt
        Case 1 'Mantenimiento
            Codigo = RecuperaValor(vCampos, 1)
            If Codigo = "" Then
                cadCodigo = " isnull(nummante) "
            Else
                cadCodigo = " nummante=" & DBSet(Codigo, "T")
            End If
            Sql = ObtenerSQLcomponentes(cadWhere & " and " & cadCodigo)
            Me.Label1(0).Caption = "Mantenimiento: " & Codigo
            
        Case 2 'Departamento
            Codigo = RecuperaValor(vCampos, 2)
            If Codigo = "" Then
                cadCodigo = "isnull(coddirec)"
            Else
                cadCodigo = " coddirec=" & Codigo
            End If
            Sql = ObtenerSQLcomponentes(cadWhere & " and " & cadCodigo)
            If vParamAplic.HayDeparNuevo = 1 Then
                Me.Caption = "Equipos del Departamento"
                Me.Label1(0).Caption = " Departamento: " & RecuperaValor(vCampos, 3)
            ElseIf vParamAplic.HayDeparNuevo = 0 Then
                Me.Caption = "Equipos de la Dirección"
                Me.Label1(0).Caption = " Dirección: " & Codigo & " " & RecuperaValor(vCampos, 3)
            Else
                Me.Caption = "Equipos de la obra"
                Me.Label1(0).Caption = " Obra: " & Codigo & " " & RecuperaValor(vCampos, 3)
            End If
        
        Case 3 'Cliente
            Sql = ObtenerSQLcomponentes(cadWhere)
            Me.Caption = "Equipos del Cliente"
            Me.Label1(0).Caption = "Cliente: " & RecuperaValor(vCampos, 4)
    End Select
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    'Los encabezados
    ListView1.Top = 800
    ListView1.Left = 280
    ListView1.Width = 4900
    ListView1.Height = 3250
    ListView1.ColumnHeaders.Clear
    
    ListView1.ColumnHeaders.Add , , "TA", 760
    ListView1.ColumnHeaders.Add , , "Tipo Articulo", 2800
    ListView1.ColumnHeaders.Add , , "Cantidad", 1280, 2
    
    If Not Rs.EOF Then
        While Not Rs.EOF
            Set ItmX = ListView1.ListItems.Add
            ItmX.Text = Rs.Fields(0).Value 'TA
            ItmX.SubItems(1) = Rs.Fields(1).Value 'Tipo Articulo
            ItmX.SubItems(2) = Rs.Fields(2).Value 'Cantidad
            Rs.MoveNext
        Wend
    End If
    Rs.Close
    Set Rs = Nothing
End Sub


Private Sub CargarListaPreFacturar()
'Muestra la lista Detallada de Albaranes a Factura en un ListView
'Carga los valores de la tabla scobro de la Contabilidad
Dim Rs As ADODB.Recordset
Dim ItmX As ListItem
Dim Sql As String

    On Error GoTo ECargarList
    
    Sql = "CREATE TEMPORARY TABLE tmp ( "
    Sql = Sql & "codforpa SMALLINT(3) UNSIGNED  DEFAULT '0' NOT NULL, "
    Sql = Sql & "numalbar MEDIUMINT(7) UNSIGNED  DEFAULT '0' NOT NULL, "
    Sql = Sql & "dtoppago DECIMAL(4,2) UNSIGNED  DEFAULT '0.0' NOT NULL, "
    Sql = Sql & "dtopgnral DECIMAL(4,2) UNSIGNED  DEFAULT '0.0' NOT NULL, "
    Sql = Sql & "importe DECIMAL(12,2) UNSIGNED  DEFAULT '0.0' NOT NULL, "
    Sql = Sql & "bruto DECIMAL(12,2) UNSIGNED  DEFAULT '0.0' NOT NULL) "
    conn.Execute Sql
     
'     SQL = "LOCK TABLES scaalb READ, slialb READ;"
'     Conn.Execute SQL
     
    Sql = "SELECT scaalb.codforpa, scaalb.numalbar, dtoppago, dtognral, round(sum(importel),2) as importe, round(sum(importel),2) - round(((round(sum(importel),2)*dtoppago)/100),2) - round(((round(sum(importel),2)*dtognral)/100),2) as bruto "
    Sql = Sql & " FROM (scaalb INNER JOIN slialb ON scaalb.codtipom=slialb.codtipom and scaalb.numalbar=slialb.numalbar) "
    Sql = Sql & " WHERE " & cadWhere
    Sql = Sql & " GROUP BY scaalb.numalbar "
    Sql = Sql & " ORDER BY scaalb.codforpa, scaalb.numalbar "

    Sql = " INSERT INTO tmp " & Sql
    conn.Execute Sql
     
    Sql = " SELECT tmp.codforpa, sforpa.nomforpa, sum(tmp.bruto) as bruto"
    Sql = Sql & " FROM tmp, sforpa WHERE tmp.codforpa=sforpa.codforpa "
    Sql = Sql & " GROUP BY tmp.codforpa "
    
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        ListView1.Height = 3850
        ListView1.Width = 5400
        ListView1.Left = 500
        ListView1.Top = 1200
    '    ListView1.GridLines = False
    
        'Los encabezados
        ListView1.ColumnHeaders.Clear
        
        ListView1.ColumnHeaders.Add , , " Forma de Pago", 3300
        ListView1.ColumnHeaders.Add , , "Base Imp.(€)", 2020, 1
     
        While Not Rs.EOF
            Set ItmX = ListView1.ListItems.Add
            ItmX.Text = Format(Rs!codforpa.Value, "000") & "  " & Rs!nomforpa.Value
            
            ItmX.SubItems(1) = Rs!bruto
            Rs.MoveNext
        Wend
    End If
    Rs.Close
    Set Rs = Nothing
    
    'Borrar la tabla temporal
    Sql = " DROP TABLE IF EXISTS tmp;"
    conn.Execute Sql

ECargarList:
    If Err.Number <> 0 Then
         'Borrar la tabla temporal
        Sql = " DROP TABLE IF EXISTS tmp;"
        conn.Execute Sql
'        SQL = "UNLOCK TABLES "
'        Conn.Execute SQL
    End If
End Sub


Private Sub CargarListaClientes()
'Carga las lista con todos los clientes seleccionados en la tabla:sclien
'para imprimir etiquetas, pasando como parametro la cadwhere
Dim Rs As ADODB.Recordset
Dim ItmX As ListItem
Dim Sql As String, Men As String

    On Error GoTo ECargarLista

    Select Case OpcionMensaje
    Case 8
        'CLIENTES
        Sql = "SELECT codclien,nomclien,nifclien "
        Sql = Sql & "FROM sclien "
        If cadWhere <> "" Then Sql = Sql & " WHERE " & cadWhere
        Sql = Sql & " ORDER BY codclien "
        Men = "Cliente"
    Case 9
        'PROVEEDORES
        Sql = "SELECT codprove,nomprove,nifprove "
        Sql = Sql & "FROM sprove "
        If cadWhere <> "" Then Sql = Sql & " WHERE " & cadWhere
        Sql = Sql & " ORDER BY codprove "
        Men = "Proveedor"
    Case 17
        'CLIENTES MANTENIMIENTO
        Sql = cadWhere
        Men = "Cliente"
                
    Case 22
        'CLIENTES POTENCIALES
        Sql = "SELECT codclien,nomclien,nifclien "
        Sql = Sql & "FROM sclipot "
        If cadWhere <> "" Then Sql = Sql & " WHERE " & cadWhere
        Sql = Sql & " ORDER BY codclien "
        Men = "Cli. potenciales"
    End Select
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
   
        'Los encabezados
        ListView2.Width = 9400
        ListView2.Top = 500
        ListView2.Height = 3620
        ListView2.ColumnHeaders.Clear
        
        ListView2.ColumnHeaders.Add , , Men, 1050
        ListView2.ColumnHeaders.Add , , "Nombre", 4000
        ListView2.ColumnHeaders.Add , , "NIF", 1050
       
        
        If OpcionMensaje = 17 Then
            ListView2.Width = 9400
            ListView2.Left = 500
            ListView2.ColumnHeaders.Add , , "Dpto", 550
            If vParamAplic.HayDeparNuevo = 1 Then
                ListView2.ColumnHeaders.Add , , "Departamento", 2400
            ElseIf vParamAplic.HayDeparNuevo = 0 Then
                ListView2.ColumnHeaders.Add , , "Direccion", 2400
            Else
                ListView2.ColumnHeaders.Add , , "Obra", 2400
            End If
        Else
             ListView2.Width = 7000
        End If
     If Not Rs.EOF Then
        While Not Rs.EOF
             Set ItmX = ListView2.ListItems.Add
             ItmX.Text = Format(Rs.Fields(0).Value, "000000") 'cod clien/prove
             ItmX.Checked = False
             ItmX.SubItems(1) = Rs.Fields(1).Value 'Nom clien/prove
             ItmX.SubItems(2) = DBLet(Rs.Fields(2).Value, "T") 'NIF clien/prove
             
             If OpcionMensaje = 17 Then
                ItmX.SubItems(3) = DBLet(Rs.Fields(3).Value, "T") 'cod dpto
                ItmX.SubItems(4) = DBLet(Rs.Fields(4).Value, "T") 'nom dpto
             End If
            
             Rs.MoveNext
        Wend
    End If
    Rs.Close
    Set Rs = Nothing
    
ECargarLista:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Cargar " & Men, Err.Description
        PulsadoSalir = True
        Unload Me
    End If
End Sub



Private Sub CargarListaErrContab()
'Muestra la lista Detallada de Facturas que dieron error al contabilizar
'en un ListView
Dim Rs As ADODB.Recordset
Dim ItmX As ListItem
Dim Sql As String

    On Error GoTo ECargarList

    Sql = " SELECT  * "
    Sql = Sql & " FROM tmpErrFac "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        ListView1.Height = 4500
        ListView1.Width = 7400
        ListView1.Left = 500
        ListView1.Top = 500

        'Los encabezados
        ListView1.ColumnHeaders.Clear

        If Rs.Fields(0).Name = "codprove" Then
            'Facturas de Compra
             ListView1.ColumnHeaders.Add , , "Prove.", 700
        Else 'Facturas de Venta
            ListView1.ColumnHeaders.Add , , "Tipo", 600
        End If
        ListView1.ColumnHeaders.Add , , "Factura", 1000, 1
        ListView1.ColumnHeaders.Add , , "Fecha", 1100, 1
        ListView1.ColumnHeaders.Add , , "Error", 4620
    
        While Not Rs.EOF
            Set ItmX = ListView1.ListItems.Add
            'El primer campo será codtipom si llamamos desde Ventas
            ' y será codprove si llamamos desde Compras
            ItmX.Text = Rs.Fields(0).Value
            ItmX.SubItems(1) = Format(Rs!Numfactu, "0000000")
            ItmX.SubItems(2) = Rs!FecFactu
            ItmX.SubItems(3) = Rs!Error
            Rs.MoveNext
        Wend
    End If
    Rs.Close
    Set Rs = Nothing

ECargarList:
    If Err.Number <> 0 Then
        Err.Clear
    End If
End Sub


Private Sub CargarListaLinFactu()
'Carga las lista con todas las lineas de la factura que estamos rectificando
'seleccionamos las que nos queremos llevar al Albaran de rectificacion
Dim Rs As ADODB.Recordset
Dim ItmX As ListItem
Dim Sql As String

    On Error GoTo ECargarLista

    Sql = "SELECT codtipom,numfactu,fecfactu,codtipoa,numalbar,numlinea,codalmac,codartic,nomartic,ampliaci,cantidad,precioar,dtoline1,dtoline2,importel,origpre"
    Sql = Sql & " FROM slifac "
    If cadWhere <> "" Then Sql = Sql & " WHERE " & cadWhere
        
    'HERBELCA NO DEJA traer varios para GANDIA - CASTELLONS
    If vParamAplic.NumeroInstalacion = 2 Then
        'Si el almacen es gandia y castellon NO sale si el stock es cero
        If vUsu.AlmacenPorDefecto2 = 2 Or vUsu.AlmacenPorDefecto2 = 3 Then Sql = Sql & " AND NOT codartic IN (Select codartic from sartic where artvario=1)"
    End If
    
    
    
    
    Sql = Sql & " ORDER BY codtipom,numfactu,fecfactu,codtipoa,numalbar,numlinea "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        
        ListView2.Top = 500
        ListView2.Left = 380
        ListView2.Width = 10100
        ListView2.Height = 3620
        
        'Los encabezados
        ListView2.ColumnHeaders.Clear
    
        ListView2.ColumnHeaders.Add , , "T.Alb", 660
        ListView2.ColumnHeaders.Add , , "Nº Alb", 840
        ListView2.ColumnHeaders.Add , , "Lin.", 450
         ListView2.ColumnHeaders.Item(3).Alignment = lvwColumnRight
        ListView2.ColumnHeaders.Add , , "Alm", 460
        ListView2.ColumnHeaders.Add , , "Artic", 1380
        ListView2.ColumnHeaders.Add , , "Desc. Artic.", 2500
        ListView2.ColumnHeaders.Add , , "Cant.", 600
        ListView2.ColumnHeaders.Item(7).Alignment = lvwColumnRight
        ListView2.ColumnHeaders.Add , , "Precio", 960
        ListView2.ColumnHeaders.Item(8).Alignment = lvwColumnRight
        ListView2.ColumnHeaders.Add , , "Dto 1", 600
        ListView2.ColumnHeaders.Item(9).Alignment = lvwColumnRight
        ListView2.ColumnHeaders.Add , , "Dto 2", 600
        ListView2.ColumnHeaders.Item(10).Alignment = lvwColumnRight
        ListView2.ColumnHeaders.Add , , "Importe", 950
        ListView2.ColumnHeaders.Item(11).Alignment = lvwColumnRight
    
        While Not Rs.EOF
             Set ItmX = ListView2.ListItems.Add
             ItmX.Text = Rs!Codtipoa 'cod tipo alb
             ItmX.Checked = False
             ItmX.SubItems(1) = Format(Rs!Numalbar, "0000000") 'Nº Albaran
             ItmX.SubItems(2) = Rs!numlinea 'linea Albaran
             ItmX.SubItems(3) = Format(Rs!codAlmac, "000") 'cod almacen
             ItmX.SubItems(4) = Rs!codArtic 'Cod Articulo
             ItmX.SubItems(5) = Rs!NomArtic 'Nombre del Articulo
             ItmX.SubItems(6) = Rs!cantidad
             ItmX.SubItems(7) = Format(Rs!precioar, FormatoPrecio)
             ItmX.SubItems(8) = Rs!dtoline1
             ItmX.SubItems(9) = Rs!dtoline2
             ItmX.SubItems(10) = Format(Rs!ImporteL, FormatoImporte)
             Rs.MoveNext
        Wend
    End If
    Rs.Close
    Set Rs = Nothing
    
    'si aparece la barra de desplazamiento ajustar el ancho
    If Me.ListView2.ListItems.Count > 11 Then
        Me.ListView2.ColumnHeaders(5).Width = 1200 'codartic
        Me.ListView2.ColumnHeaders(8).Width = 920  'precio
    End If
   
    If ListView2.ListItems.Count = 0 Then
        MsgBox "Ninguna linea disponible para rectificar", vbExclamation
        PulsadoSalir = True
        Unload Me
    End If
    
ECargarLista:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Cargar Lineas Factura", Err.Description
        PulsadoSalir = True
        Unload Me
    End If
End Sub




Private Sub CargarListaAlbaranes()
'Muestra la lista Detallada de Facturas que dieron error al contabilizar
'en un ListView
Dim Rs As ADODB.Recordset
Dim ItmX As ListItem
Dim Sql As String

    On Error GoTo ECargarList

    Sql = cadWhere 'cadwhere ya le pasamos toda la SQL
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        ListView1.Height = 3900
        ListView1.Width = 7200
        ListView1.Left = 500
        ListView1.Top = 700

        'Los encabezados
        ListView1.ColumnHeaders.Clear

        ListView1.ColumnHeaders.Add , , "Tipo", 700
        ListView1.ColumnHeaders.Add , , "Nº Albaran", 1000, 1
        ListView1.ColumnHeaders.Add , , "Fecha", 1100, 1
        ListView1.ColumnHeaders.Item(3).Alignment = lvwColumnCenter
        ListView1.ColumnHeaders.Add , , "Cod. Cli.", 900
        ListView1.ColumnHeaders.Add , , "Cliente", 3400
    
        While Not Rs.EOF
            Set ItmX = ListView1.ListItems.Add
            ItmX.Text = Rs.Fields(0).Value
            ItmX.SubItems(1) = Format(Rs!Numalbar, "0000000")
            ItmX.SubItems(2) = Rs!FechaAlb
            ItmX.SubItems(3) = Format(Rs!codClien, "000000")
            ItmX.SubItems(4) = Rs!NomClien
            Rs.MoveNext
        Wend
    End If
    Rs.Close
    Set Rs = Nothing

ECargarList:
    If Err.Number <> 0 Then
        Err.Clear
    End If
End Sub



Private Sub CargarListaEmpresas()
'Carga las lista con todas las empresas que hay en el sistema
Dim Rs As ADODB.Recordset
Dim ItmX As ListItem
Dim Sql As String
Dim i As Integer

Dim Prohibidas As String

    On Error GoTo ECargarLista

    VerEmresasProhibidas Prohibidas
    
    Sql = "Select * from usuarios.empresasariges order by codempre"
    Set ListView2.SmallIcons = frmPpal.ImageListB
    ListView2.Width = 5000
    ListView2.ColumnHeaders.Clear
    ListView2.ColumnHeaders.Add , , "Empresa", 4900
    ListView2.HideColumnHeaders = True
    ListView2.GridLines = False
    ListView2.ListItems.Clear
    
    Set Rs = New ADODB.Recordset
    i = -1
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not Rs.EOF
        Sql = "|" & Rs!codempre & "|"
        If InStr(1, Prohibidas, Sql) = 0 Then
            Set ItmX = ListView2.ListItems.Add(, , Rs!nomempre, , 5)
            ItmX.Tag = Rs!codempre
            If ItmX.Tag = vEmpresa.codempre Then
                ItmX.Checked = True
                i = ItmX.Index
            End If
            ItmX.ToolTipText = Rs!AriGes
        End If
        Rs.MoveNext
    Wend
    Rs.Close
    If i > 0 Then Set ListView2.SelectedItem = ListView2.ListItems(i)

    
ECargarLista:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Cargando datos empresas", Err.Description
        PulsadoSalir = True
        Unload Me
    End If
End Sub


Private Sub VerEmresasProhibidas(ByRef VarProhibidas As String)
Dim Sql As String
Dim Rs As ADODB.Recordset

On Error GoTo EVerEmresasProhibidas
    VarProhibidas = "|"
    Sql = "Select codempre from usuarios.usuarioempresasariges WHERE codusu = " & (vUsu.Codigo Mod 1000)
    Sql = Sql & " order by codempre"
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not Rs.EOF
          VarProhibidas = VarProhibidas & Rs!codempre & "|"
          Rs.MoveNext
    Wend
    Rs.Close
    Exit Sub
EVerEmresasProhibidas:
    MuestraError Err.Number, Err.Description & vbCrLf & " Consulte soporte técnico"
    Set Rs = Nothing
End Sub



Private Sub CargaImagen()
On Error Resume Next
   ' Image2.Picture = LoadPicture(App.Path & "\minilogo.bmp")
    'Image1.Picture = LoadPicture(App.path & "\fondon.gif")
    Err.Clear
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If OpcionMensaje = 28 Then
        If PulsadoSalir = True Then
            RaiseEvent DatoSeleccionado("NO")
        Else
            PulsadoSalir = True
        End If
    End If
    
    ' previsualizacion
    If OpcionMensaje = 29 Then
        If PulsadoSalir = True Then
            RaiseEvent DatoSeleccionado("OK")
        Else
            PulsadoSalir = True
        End If
    End If
    
    ' introduccion de password
    If OpcionMensaje = 30 Then
        If PulsadoSalir = True Then
            RaiseEvent DatoSeleccionado(Text3.Text)
        Else
            PulsadoSalir = True
        End If
    End If

    

    If OpcionMensaje = 32 Then
        
        'If PulsadoSalir = False Then Cancel = 1
    End If

    If PulsadoSalir = False Then Cancel = 1
    
End Sub



Private Function ObtenerTamanyosArray() As Boolean
'Para el frame de los Nº de Serie de los Articulos
'En cada indice pone en CodArtic(i) el codigo del articulo
'y en Cantidad(i) la cantidad solicitada de cada codartic
Dim i As Integer, J As Integer

    ObtenerTamanyosArray = False
    'Primero a los campos de la tabla
    TotalArray = -1
    J = 0
    Do
        i = J + 1
        J = InStr(i, vCampos, "·")
        If J > 0 Then TotalArray = TotalArray + 1
    Loop Until J = 0
    
    If TotalArray < 0 Then Exit Function
    
    'Las redimensionaremos
    ReDim codArtic(TotalArray)
    ReDim cantidad(TotalArray)
    
    ObtenerTamanyosArray = True
End Function


Private Function SeparaCampos() As Boolean
'Para el frame de los Nº de Serie de los Articulos
Dim Grupo As String
Dim i As Integer
Dim J As Integer
Dim C As Integer 'Contador dentro del array

    SeparaCampos = False
    i = 0
    C = 0
    Do
        J = i + 1
        i = InStr(J, vCampos, "·")
        If i > 0 Then
            Grupo = Mid(vCampos, J, i - J)
            'Y en la martriz
            InsertaGrupo Grupo, C
            C = C + 1
        End If
    Loop Until i = 0
    SeparaCampos = True
End Function


Private Sub InsertaGrupo(Grupo As String, Contador As Integer)
Dim J As Integer
Dim Cad As String

    J = 0
    Cad = ""
    
    'Cod Artic
    J = InStr(1, Grupo, "|")
    If J > 0 Then
        Cad = Mid(Grupo, 1, J - 1)
        Grupo = Mid(Grupo, J + 1)
        J = 1
    End If
    codArtic(Contador) = Cad
    
    'Cantidad
    J = InStr(1, Grupo, "|")
    If J > 0 Then
        Cad = Mid(Grupo, 1, J - 1)
        Grupo = Mid(Grupo, J + 1)
    Else
        Cad = Grupo
        Grupo = ""
    End If
    cantidad(Contador) = Cad
End Sub





Private Sub imgCheck_Click(Index As Integer)
Dim B As Boolean
    If Index < 2 Then
        'En el listview3
        B = Index = 1
        For TotalArray = 1 To ListView3.ListItems.Count
            ListView3.ListItems(TotalArray).Checked = B
            If (TotalArray Mod 50) = 0 Then DoEvents
        Next TotalArray
        
    ElseIf Index < 4 Then
        'En el listview4
        B = Index = 3
        For TotalArray = 1 To ListView4.ListItems.Count
            If ListView4.ListItems(TotalArray).Tag <> "" Then
                ListView4.ListItems(TotalArray).Checked = B
            Else
                ListView4.ListItems(TotalArray).Checked = False
            End If
            If (TotalArray Mod 50) = 0 Then DoEvents
        Next TotalArray
        
        
    ElseIf Index < 6 Then
        'lwActividad
        B = Index = 5
        For TotalArray = 1 To lwActividad.ListItems.Count
            lwActividad.ListItems(TotalArray).Checked = B
            If (TotalArray Mod 50) = 0 Then DoEvents
        Next TotalArray
        
        
    ElseIf Index < 8 Then
        'Lista de articulos de proveedor
        'ListView5
        B = Index = 7
        For TotalArray = 1 To ListView5.ListItems.Count
            ListView5.ListItems(TotalArray).Checked = B
            If (TotalArray Mod 50) = 0 Then DoEvents
        Next TotalArray
        
    ElseIf Index < 10 Then
        B = Index = 9
        For TotalArray = 1 To ListView10.ListItems.Count
            ListView10.ListItems(TotalArray).Checked = B
            If (TotalArray Mod 50) = 0 Then DoEvents
        Next TotalArray
    Else
        B = Index = 11
        For TotalArray = 1 To ListView12.ListItems.Count
            ListView12.ListItems(TotalArray).Checked = B
            If (TotalArray Mod 50) = 0 Then DoEvents
        Next TotalArray
    End If
End Sub



Private Sub Label9_Click()
        LanzaVisorMimeDocumento Me.hwnd, "http://help-ariges.ariadnasw.com/Versiones.html"
End Sub

Private Sub ListView12_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Dim kCol As Integer
    
    
    Select Case ColumnHeader.Index
    Case 1
        kCol = 7
    Case 7
        kCol = 8
    Case 8
        kCol = 9
    Case Else
        kCol = ColumnHeader.Index - 1
    End Select
    If ListView12.SortKey = kCol Then
        ListView12.SortOrder = IIf(ListView12.SortOrder = lvwDescending, lvwAscending, lvwDescending)
    Else
        ListView12.SortKey = kCol
        ListView12.SortOrder = lvwAscending
    End If
    
End Sub

Private Sub ListView3_DblClick()
    If ListView3.ListItems.Count = 0 Then Exit Sub
    If ListView3.SelectedItem Is Nothing Then Exit Sub
    vCampos = InputBox("Art: " & ListView3.SelectedItem.Text, "Cambiar cantidad", Val(ListView3.SelectedItem.SubItems(3)))
    If vCampos <> "" Then
        If Val(vCampos) > 0 Then ListView3.SelectedItem.SubItems(3) = Val(vCampos)
    End If
End Sub

Private Sub ListView3_KeyPress(KeyAscii As Integer)
    ListView3_DblClick
End Sub

Private Sub lwArticulosAgrupados_DblClick()
    cmdArtAgrupado_Click 0
End Sub

Private Sub lwArticulosAgrupados_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub lwPedidosFontenas_DblClick(Index As Integer)
    If Index = 1 Then
        If lwPedidosFontenas(1).SelectedItem Is Nothing Then Exit Sub
                    
        If lwPedidosFontenas(1).SelectedItem.SmallIcon = 6 Then
            LanzaProcesoImportacionXLS
            
        Else
            ImportarFontenasCSV True, lwPedidosFontenas(1).SelectedItem.Text
        End If
        cargarDatosImportarPedidosEXCEL
    End If
End Sub

Private Sub lwPedidosFontenas_ItemClick(Index As Integer, ByVal Item As MSComctlLib.ListItem)
    If Index = 0 Then PonerDatosPedidosFontenas
    
End Sub

Private Sub lwTaxco_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Dim queColum As Integer
    
    queColum = ColumnHeader.Index
    If queColum = 1 Then queColum = 11
    If queColum = 2 Then queColum = 12
    If queColum = 3 Then queColum = 13
    If queColum = 4 Then queColum = 14
    If queColum < 11 Then queColum = queColum - 1
    
'    For NumRegElim = 1 To lwTaxco.ColumnHeaders.Count
'        Debug.Print lwTaxco.ColumnHeaders(NumRegElim).Text & ": " & lwTaxco.ColumnHeaders(NumRegElim).Width
'    Next

    If queColum = lwTaxco.SortKey Then
        If lwTaxco.SortOrder = lvwAscending Then
            lwTaxco.SortOrder = lvwDescending
        Else
            lwTaxco.SortOrder = lvwAscending
        End If
    Else
        If queColum = 14 Then
            lwTaxco.SortOrder = lvwDescending
        Else
            lwTaxco.SortOrder = lvwAscending
        End If
        lwTaxco.SortKey = queColum
    End If
    
    For queColum = 1 To lwTaxco.ColumnHeaders.Count
        If lwTaxco.ColumnHeaders(queColum).Width > 1 Then
            lwTaxco.ColumnHeaders.Item(queColum).Text = Replace(lwTaxco.ColumnHeaders.Item(queColum).Text, ">", "")
            lwTaxco.ColumnHeaders.Item(queColum).Text = Replace(lwTaxco.ColumnHeaders.Item(queColum).Text, "<", "")
        End If
    Next
    ColumnHeader.Text = Replace(ColumnHeader.Text, ">", "")
    ColumnHeader.Text = Trim(Replace(ColumnHeader.Text, "<", ""))
    ColumnHeader.Text = ColumnHeader.Text & " " & IIf(lwTaxco.SortOrder = lvwAscending, ">", "<")
End Sub

Private Sub lwTaxco_DblClick()
    If Me.lwTaxco.ListItems.Count = 0 Then Exit Sub
    If Me.lwTaxco.SelectedItem Is Nothing Then Exit Sub
    
    With lwTaxco.SelectedItem
        cadWHERE2 = "Cliente " & .Text & " " & .ToolTipText & vbCrLf
        cadWHERE2 = cadWHERE2 & "Fact: " & .SubItems(2) & .SubItems(1) & "   " & .SubItems(3) & vbCrLf
        cadWHERE2 = cadWHERE2 & "Matr: " & .SubItems(8) & "   " & .ListSubItems(8).ToolTipText & vbCrLf
        If Trim(.SubItems(9)) <> "" Then cadWHERE2 = cadWHERE2 & "Km: " & .SubItems(9)
           
        cadWHERE2 = cadWHERE2 & vbCrLf & "Obs: " & .SubItems(10)
    
    End With
    MsgBox cadWHERE2, vbInformation
End Sub

Private Sub OptCompXClien_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub OptCompXDpto_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub OptCompXMant_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 2, cerrar
    If cerrar Then Unload Me
End Sub



Private Function RegresarCargaEmpresas() As String
Dim Sql As String
Dim Parametros As String
Dim i As Integer

    CadenaDesdeOtroForm = ""
    
        Sql = ""
        Parametros = ""
        For i = 1 To ListView2.ListItems.Count
            If Me.ListView2.ListItems(i).Checked Then
                Sql = Sql & Me.ListView2.ListItems(i).Text & "|"
                Parametros = Parametros & "1" 'Contador
            End If
        Next i
        CadenaDesdeOtroForm = Len(Parametros) & "|" & Sql
        'Vemos las conta
        Sql = ""
        For i = 1 To ListView2.ListItems.Count
            If Me.ListView2.ListItems(i).Checked Then
                Sql = Sql & Me.ListView2.ListItems(i).Tag & "|"
            End If
        Next i
        CadenaDesdeOtroForm = CadenaDesdeOtroForm & Sql
    
    
        RegresarCargaEmpresas = CadenaDesdeOtroForm

End Function



Private Sub CargarArticulosEstanteria()
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim IT As ListItem
Dim RBarras As ADODB.Recordset
Dim J As Integer
Dim Aux As String

    
    Set RBarras = New ADODB.Recordset
    Label6.Caption = "Cargando"
    Label6.Refresh
    Sql = "Select * from sarti3 order by codartic,numlinea desc"
    RBarras.Open Sql, conn, adOpenKeyset, adLockPessimistic, adCmdText
    
    
    Sql = "select sartic.codartic,nomartic,preciove,codigiva,nomfamia from sartic,sfamia where sartic.codfamia=sfamia.codfamia"
    If cadWhere <> "" Then Sql = Sql & " AND " & cadWhere
    If vCampos <> "" Then Sql = Sql & " AND codartic in (Select codartic from salmac WHERE codalmac= " & vCampos & " AND  stockmin >0)"
    
    
    
    Sql = Sql & " ORDER BY sartic.codartic "
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    TotalArray = 0
    While Not Rs.EOF
        Set IT = ListView3.ListItems.Add
        Label6.Caption = Rs!codArtic
        Label6.Refresh
        
        RBarras.Find "codartic = " & DBSet(Rs!codArtic, "T"), , adSearchForward, 1
        If RBarras.EOF Then
            Sql = ""
            If vParamAplic.NumeroInstalacion = vbTaxco Then Sql = Rs!codArtic
        Else
            Sql = RBarras!codigoea
        End If
        'Ponemos el codigo de articulo y el TIPO de IVA
        IT.Tag = "'" & DevNombreSQL(Rs!codArtic) & "'," & Rs!Codigiva & ",'" & Sql & "'"
        IT.Text = Rs!NomArtic
        IT.SubItems(1) = Format(Rs!PrecioVe, cadWHERE2)
        IT.SubItems(2) = Rs!nomfamia
        
        IT.SubItems(3) = 1
        IT.SubItems(4) = 1 'cdoalmac
        
        
        If InStr(1, cadWhere, "slialp") > 0 Then
            J = InStr(cadWhere, "1  and codartic in (")
            If J > 0 Then
                Sql = Mid(cadWhere, J + 20)    'len("1  and codartic in (")  = 20
                Sql = Mid(Sql, 1, Len(Sql) - 1)
                J = InStr(1, Sql, " WHERE ")
                If J > 0 Then
                    Sql = Mid(Sql, J + 7) & " AND codartic "
                    Aux = "codalmac"
                    Sql = DevuelveDesdeBD(conAri, "cantidad", "slialp", Sql, Rs!codArtic, "T", Aux)
                    If Sql <> "" Then
                        IT.SubItems(3) = Val(Sql)
                        IT.SubItems(4) = Aux
                    End If
                End If
            End If
        End If
        IT.Checked = True
        Rs.MoveNext
        TotalArray = TotalArray + 1
        If TotalArray > 300 Then
            DoEvents
            If TotalArray < 0 Then
                'Han pulsado cancelar
                
                While Not Rs.EOF
                    Rs.MoveNext
                Wend
                
            End If
            TotalArray = 0
        End If
    Wend
    Rs.Close
    RBarras.Close
    
    
    'Febrero 2013
    'Opcion imprimir etiqetas articulo de un almacen determinado y que tengan stock minimo
    'Para ello se ha llamado al form poniendo en vCampos el codlamac
    
    
    Set RBarras = Nothing
    Set Rs = Nothing
    TotalArray = 0
    Label6.Caption = ""
    
        
End Sub




Private Sub CargarArticulosCorreccionPrecio()
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim IT As ListItem
Dim margen As Currency
Dim MargenT As Currency
Dim ImpPVP As Currency
Dim ImpTar As Currency
Dim Aux As Currency
Dim decimales As Long
Dim precioUC As Currency
Dim SoloImporteMenor As Boolean
Dim SobreUPC As Boolean

    'El amrgen a aplicar
    'Si la tarifa es sobre el PVP es el articulo
    'si es sobre UPC entonces es sobre el de la tarifa

    lblIndicadorCorregir = "LEYENDO BD"
    lblIndicadorCorregir.Refresh
    

    
    'Si NUMREGELIM=1 entonces esta marcada la opcion(check) de solo importes menores
    If NumRegElim = 1 Then SoloImporteMenor = True
    
    
    
    'Comprobamos la tarifa donde se aplica, si sobre PVP o sobre ultima compra (%tarifa)
    Sql = DevuelveDesdeBD(conAri, "opcionINC", "starif", "codlista", vCampos)
    SobreUPC = Val(Sql) = 1
            
    
    TotalArray = InStr(1, cadWHERE2, ",")
    Sql = Mid(cadWHERE2, TotalArray + 1)
    decimales = Len(Sql)
    'Formato
    cadWHERE2 = "#,##0." & Mid(cadWHERE2, TotalArray + 1)
    
    'Sql
    Sql = " SELECT sartic.nomartic,slista.codartic,sartic.preciove,sartic.preciouc,"
    Sql = Sql & "slista.precioac, slista.codlista, starif.nomlista,"
    Sql = Sql & "sartic.margecom as margenArt,starif.margecom as margetar"
    Sql = Sql & " FROM   (slista INNER JOIN sartic ON slista.codartic=sartic.codartic)"
    Sql = Sql & " INNER JOIN starif  ON slista.codlista=starif.codlista WHERE "

    Sql = Sql & cadWhere '& " AND "
    ''SQL = SQL & " sartic.preciove <> sartic.preciouc + round((sartic.preciouc * if(isnull(sartic.margecom),0,sartic.margecom))/100," & Decimales & ")"
    
    Sql = Sql & " ORDER BY slista.codartic"
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    

    '
  
    TotalArray = 0
    
    While Not Rs.EOF
        'Calculo los importes
        lblIndicadorCorregir.Caption = Rs!codArtic
        lblIndicadorCorregir.Refresh
        
        margen = DBLet(Rs!margenart, "N") / 100
        MargenT = DBLet(Rs!margetar, "N") / 100
        precioUC = DBLet(Rs!precioUC, "N")
        
        Aux = margen * precioUC
        ImpPVP = Round2(precioUC + Aux, decimales)
        
        'El de la tarifa
        If SobreUPC Then
            Aux = MargenT * precioUC
            ImpTar = Round2(precioUC + Aux, CLng(decimales))
        Else
        
            Aux = MargenT * ImpPVP
            ImpTar = Round2(ImpPVP + Aux, CLng(decimales))
        End If
        Aux = Round2(Rs!PrecioVe, decimales)
        
        Sql = ""
        

        If SoloImporteMenor Then
            If Aux >= ImpPVP Then
                'El primero esta bien
                'Veamos el segundo. En la tarifa
                Aux = Round2(Rs!precioac, decimales)
                If Aux < ImpTar Then Sql = "M"
            Else
                Sql = "M"
            End If
        
        
        Else
            If Aux = ImpPVP Then
                'El primero esta bien
                'Veamos el segundo. En la tarifa
                Aux = Round2(Rs!precioac, decimales)
                If Aux <> ImpTar Then Sql = "M"
            Else
                Sql = "M"
            End If
        End If
        
        If Sql <> "" Then
            Set IT = ListView4.ListItems.Add
            IT.Tag = DevNombreSQL(Rs!codArtic)
            IT.ToolTipText = IT.Tag
            IT.Text = IT.Tag
            IT.SubItems(1) = Rs!NomArtic
            '++
            IT.ListSubItems(1).ToolTipText = Rs!NomArtic
            
            Aux = Round2(precioUC, decimales)
            IT.SubItems(2) = Format(Aux, cadWHERE2)
            
            IT.SubItems(3) = Format(margen * 100, FormatoPorcen)
            Aux = Round2(Rs!PrecioVe, decimales)
            IT.SubItems(4) = Format(Aux, cadWHERE2)
            
            IT.SubItems(5) = Format(MargenT * 100, FormatoPorcen)
            Aux = Round2(Rs!precioac, decimales)
            IT.SubItems(6) = Format(Aux, cadWHERE2)
            

            IT.SubItems(7) = Format(ImpPVP, cadWHERE2)
            IT.SubItems(8) = Format(ImpTar, cadWHERE2)
            
            
            
            If precioUC = 0 Then
                'Precio ultima compra =0
                'NOOOOO se puede actualizar la tarifa
                IT.Tag = "" 'para no actualizar
                IT.Checked = False
                IT.Bold = True
                IT.ForeColor = vbRed
            Else
                
            End If
            IT.Checked = False
        End If
        Rs.MoveNext
        TotalArray = TotalArray + 1
        If TotalArray > 300 Then
            TotalArray = 0
            Me.Refresh
            DoEvents
        End If
    Wend
    Rs.Close
    cmbActualizarTar.ListIndex = 0
    lblIndicadorCorregir.Caption = ""
End Sub




Private Function TraspasarMantenimientos() As Boolean
    
    On Error GoTo ETraspasarMantenimientos
    TraspasarMantenimientos = False

    

    cadWhere = "Select count(*) from sliman where anomante =" & txtMante(0).Text
    miRsAux.Open cadWhere, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    If Not miRsAux.EOF Then NumRegElim = DBLet(miRsAux.Fields(0), "N")
    miRsAux.Close
    
    
    If NumRegElim > 0 Then
        MsgBox "Ya existen datos para el año " & txtMante(0).Text, vbExclamation
        Exit Function
    End If
    
    
    
    'Se divide en 4 pasos
    '1.- Introducir una linea en la sliman con los datos para el año
        cadWhere = "insert into sliman (anomante,codclien,nummante,mes01man,mes02man,mes03man,mes04man,mes05man,mes06man,mes07man,mes08man,mes09man,mes10man,mes11man,mes12man)"
        cadWhere = cadWhere & " SELECT " & txtMante(0).Text & ",codclien,nummante,mes01act,mes02act,mes03act,mes04act,mes05act,mes06act,mes07act,mes08act,mes09act,mes10act,mes11act,mes12act FROM scaman"
        conn.Execute cadWhere
    '2.- Updatear los campos de actual con siguiente
        cadWhere = ""
        For TotalArray = 1 To 12
            cadWhere = cadWhere & ", mes" & Format(TotalArray, "00") & "act = mes" & Format(TotalArray, "00") & "sig"
        Next TotalArray
        cadWhere = Mid(cadWhere, 2) 'Para quitar la primera coma
        cadWhere = "UPDATE scaman SET " & cadWhere
        conn.Execute cadWhere
        
    '3.- Si no han marcado la opcion copiar datos tengo que resetear a 0
        If chkMante.Value = 0 Then
            'NO SE COPIA, luego hay que resetear
            cadWhere = ""
            For TotalArray = 1 To 12
                cadWhere = cadWhere & ", mes" & Format(TotalArray, "00") & "sig = 0 "
            Next TotalArray
            cadWhere = Mid(cadWhere, 2) 'Para quitar la primera coma
            cadWhere = "UPDATE scaman SET " & cadWhere
            conn.Execute cadWhere
        End If
        
    '4.- Ultimo mes facturado pasa a ser  cero
        conn.Execute "UPDATE scaman SET ulmesfac=0"
        
    TraspasarMantenimientos = True
    
    Exit Function
ETraspasarMantenimientos:
    MuestraError Err.Number
End Function


Private Sub Text3_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub txtArtAgrupado_GotFocus()
    ConseguirFoco txtArtAgrupado, 3
End Sub


Private Sub txtArtAgrupado_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub txtArtAgrupado_LostFocus()
    If Not PonerFormatoEntero(txtArtAgrupado) Then
        txtArtAgrupado.Text = "1"
        PonerFoco txtArtAgrupado
    End If
End Sub


Private Sub txtMante_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub CargaPVPPreciosArticulosConComponentes()
Dim decimales As Byte
Dim Sql As String
Dim Impor As Currency
Dim IA As Currency
Dim PC As Currency
Dim PCC As Currency

    Set miRsAux = New ADODB.Recordset
    
    lblIndicadorCorregir = "LEYENDO BD"
    lblIndicadorCorregir.Refresh
    
    
    'Fomato importe
    TotalArray = InStr(1, cadWHERE2, ",")
    Sql = Mid(cadWHERE2, TotalArray + 1)
    decimales = Len(Sql)
    'Formato
    cadWHERE2 = "#,##0." & Mid(cadWHERE2, TotalArray + 1)
    
    
    'Tres columna svamos a ponerlas a tamaño 0
    ListView4.ColumnHeaders(6).Width = 0
    ListView4.ColumnHeaders(7).Width = 0
    
    Sql = "select sarti1.*,s1.nomartic,s1.preciove pre2,s1.margecom,s1.preciouc,"
    Sql = Sql & " sarti1.cantidad,s2.preciove, s2.preciouc coste"
    Sql = Sql & " from sarti1,sartic as s1,sartic as s2 where sarti1.codartic=s1.codartic and sarti1.codarti1=s2.codartic"
    'Si lleva WHERE
    If cadWhere <> "" Then
        vCampos = Replace(cadWhere, "sartic.", "s1.")
        Sql = Sql & " AND " & vCampos
        vCampos = ""
    End If
    
    Sql = Sql & " ORDER BY sarti1.codartic"
    miRsAux.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    Sql = ""

    While Not miRsAux.EOF
        If Sql <> miRsAux!codArtic Then
            'Nuevo articulo
            lblIndicadorCorregir = miRsAux!codArtic
            lblIndicadorCorregir.Refresh
            If Sql <> "" Then
                'Si precioventa distionto   o pcompra distionto
                If IA <> Impor Or PC <> PCC Then
                    vCampos = vCampos & Format(IA, cadWHERE2) & "|" & Format(Impor, cadWHERE2) & "|"
                    vCampos = vCampos & Format(PC, cadWHERE2) & "|" & Format(PCC, cadWHERE2) & "|"
                    InsertarItemARticuloConjunto vCampos
                End If
                    
                
            End If
            Sql = miRsAux!codArtic
            vCampos = miRsAux!codArtic & "|" & miRsAux!NomArtic & "|"
            PC = DBLet(miRsAux!precioUC, "N")
            vCampos = vCampos & Format(PC, cadWHERE2)
            vCampos = vCampos & "|" & Format(DBLet(miRsAux!margecom, "N"), FormatoPorcen) & "|"
            
            IA = miRsAux!pre2
            PCC = 0 'precio compra calculado
            Impor = 0
        End If
        Impor = Impor + Round2((miRsAux!cantidad * miRsAux!PrecioVe), CLng(decimales))
        PCC = PCC + Round2((miRsAux!cantidad * DBLet(miRsAux!coste, "N")), CLng(decimales))
        miRsAux.MoveNext
    Wend
    If Sql <> "" Then
        If IA <> Impor Or PC <> PCC Then
            vCampos = vCampos & Format(IA, cadWHERE2) & "|" & Format(Impor, cadWHERE2) & "|"
            vCampos = vCampos & Format(PC, cadWHERE2) & "|" & Format(PCC, cadWHERE2) & "|"
            InsertarItemARticuloConjunto vCampos
        End If
    End If
    miRsAux.Close
    lblIndicadorCorregir = ""
End Sub



Private Sub InsertarItemARticuloConjunto(Datos As String)
Dim IT As ListItem

        Set IT = ListView4.ListItems.Add
        IT.Tag = RecuperaValor(Datos, 1)
        IT.ToolTipText = IT.Tag
        IT.Text = IT.Tag
        IT.SubItems(1) = RecuperaValor(Datos, 2)  'nomartic
    
        IT.SubItems(2) = RecuperaValor(Datos, 3)  'precio UC del articulo
        IT.SubItems(3) = RecuperaValor(Datos, 4)  ' Margen
        
        IT.SubItems(4) = RecuperaValor(Datos, 5)  'PVP articulo
        IT.SubItems(7) = RecuperaValor(Datos, 6)  'PVP calculado
        IT.SubItems(8) = RecuperaValor(Datos, 8)  'PUC calculado
        
            
End Sub

Private Sub CargaComboActualizarPrecios()
    cmbActualizarTar.Clear
    
    If OpcionMensaje = 16 Then
        'ART Y TARIFAS
        cmbActualizarTar.Tag = "Artículos y tarifas|Solo artículo|Solo tarifas|"
    Else
        cmbActualizarTar.Tag = "PVP y PUC|Solo PVP|Solo PUC|"
    End If
    cmbActualizarTar.AddItem RecuperaValor(cmbActualizarTar.Tag, 1)
    cmbActualizarTar.AddItem RecuperaValor(cmbActualizarTar.Tag, 2)
    cmbActualizarTar.AddItem RecuperaValor(cmbActualizarTar.Tag, 3)
    cmbActualizarTar.Tag = ""
    cmbActualizarTar.ListIndex = 0
End Sub



Private Sub CargarEmail()
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open "Select * from scrmmail WHERE " & cadWhere, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        Me.txmemail(0).Text = miRsAux!email
        
        Me.txmemail(4).Text = miRsAux!FechaHora
        Me.txmemail(1).Text = DBLet(miRsAux!asunto, "T")
        Me.txmemail(2).Text = DBLet(miRsAux!adjuntos, "T")
        Me.txmemail(3).Text = DBLet(miRsAux!cuerpo, "T")
    
    
    End If
    miRsAux.Close
    Set miRsAux = Nothing
End Sub


Private Function CargaDatosEtiquetas() As Boolean

    On Error GoTo ECargaDatosEtiquetas
    CargaDatosEtiquetas = False
    
    
    cadWhere = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute cadWhere
    
    Set miRsAux = New ADODB.Recordset
    cadWhere = "Select codprove,codalmac from tmpnlotes where codusu = " & vUsu.Codigo & " ORDER by 1,2"
    miRsAux.Open cadWhere, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    cadWhere = ""
    vCampos = "" 'Para la etiqueta
    While Not miRsAux.EOF
        'para cada cliente departamento vere el campo attetiqu
        cadWhere = "attetiqu<>"""" and coddirec "
        If IsNull(miRsAux!codAlmac) Then
            cadWhere = cadWhere & " is null"
        Else
            cadWhere = cadWhere & " = " & miRsAux!codAlmac
        End If
        cadWhere = cadWhere & " AND codclien"
        cadWhere = DevuelveDesdeBD(conAri, "attetiqu", "scaman", cadWhere, CStr(miRsAux!Codprove), "N")
        


        cadWHERE2 = "insert into `tmpinformes` (`codusu`,`codigo1`,`campo1`,`nombre1`) VALUES (" & vUsu.Codigo & ","
        cadWHERE2 = cadWHERE2 & miRsAux!Codprove & ","
        If IsNull(miRsAux!codAlmac) Then
            cadWHERE2 = cadWHERE2 & "NULL"
        Else
            cadWHERE2 = cadWHERE2 & miRsAux!codAlmac
        End If
        cadWHERE2 = cadWHERE2 & "," & DBSet(cadWhere, "T") & ")"
        NumRegElim = NumRegElim + 1
        conn.Execute cadWHERE2
            
        miRsAux.MoveNext
        
    Wend
    miRsAux.Close
        
    If NumRegElim > 0 Then CargaDatosEtiquetas = True
    
    
 
ECargaDatosEtiquetas:
    If Err.Number <> 0 Then MuestraError Err.Number, "CargaDatosEtiquetas"
    Set miRsAux = Nothing
End Function



Private Sub CargarPVPArticulos()
Dim Sql As String
Dim IT As ListItem
Dim RIVA As ADODB.Recordset
Dim Precio As Currency
Dim ImpIva As Currency

    Set RIVA = New ADODB.Recordset
    Label6.Caption = "Cargando"
    Label6.Refresh
    Sql = "Select * from tiposiva order by codigiva"
    RIVA.Open Sql, ConnConta, adOpenKeyset, adLockPessimistic, adCmdText
    
      
    Sql = "select sartic.codartic,nomartic,preciove,codigiva,nomfamia,slista.precioac,fechanue,precionu from sfamia inner join sartic"
    Sql = Sql & " on sartic.codfamia=sfamia.codfamia left join slista on slista.codartic=sartic.codartic and codlista=" & vParamAplic.CodTarifa
    Sql = Sql & " WHERE 1=1 "
    If cadWhere <> "" Then Sql = Sql & " AND " & cadWhere
    
    '
    'If vCampos <> "" Then SQL = SQL & " AND codartic in (Select codartic from salmac WHERE codalmac= " & vCampos & " AND  stockmin >0)"
    
    
    
    Sql = Sql & " ORDER BY sartic.codartic "
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    TotalArray = 0
    While Not miRsAux.EOF
        Set IT = ListView3.ListItems.Add
        Label6.Caption = miRsAux!codArtic
        Label6.Refresh
        
        
        
        '`codartic`,`numlinea`,numserie,`numlinealb`,nummante
        
        RIVA.Find "codigiva = " & miRsAux!Codigiva, , adSearchForward, 1
        ImpIva = 0
        If Not RIVA.EOF Then ImpIva = DBLet(RIVA!PorceIVA)
                
        
        
        
        
        
        cadWHERE2 = " "
        If Not IsNull(miRsAux!precioac) Then
            Precio = miRsAux!precioac
            If Not IsNull(miRsAux!fechanue) Then
                If Now >= miRsAux!fechanue Then Precio = DBLet(miRsAux!precionu, "N")
            End If
            Precio = Round(((ImpIva * Precio) / 100) + Precio, 2)
            cadWHERE2 = Format(Precio, FormatoImporte)
        End If
        
        'PVP + IVA
        Precio = Round(((ImpIva * miRsAux!PrecioVe) / 100) + miRsAux!PrecioVe, 2)
        
        
        IT.Text = miRsAux!NomArtic
        IT.SubItems(1) = Format(Precio, FormatoImporte)
        IT.SubItems(2) = cadWHERE2
        IT.Checked = True
        
        '`codartic`,`numlinea`,numserie,`numlinealb`,nummante
        Sql = "'" & DevNombreSQL(miRsAux!codArtic) & "'," & miRsAux!Codigiva & ",'" & IT.SubItems(1) & "'," & TotalArray & ",'" & IT.SubItems(2) & "')"
        IT.Tag = Sql
        
        
        miRsAux.MoveNext
        TotalArray = TotalArray + 1
        If TotalArray > 300 Then
            DoEvents
            If TotalArray < 0 Then
                'Han pulsado cancelar
                
                While Not miRsAux.EOF
                    miRsAux.MoveNext
                Wend
                
            End If
            TotalArray = 0
        End If
    Wend
    miRsAux.Close
    RIVA.Close
    
    
    'Febrero 2013
    'Opcion imprimir etiqetas articulo de un almacen determinado y que tengan stock minimo
    'Para ello se ha llamado al form poniendo en vCampos el codlamac
    
    
    Set RIVA = Nothing
    Set miRsAux = Nothing
    TotalArray = 0
    Label6.Caption = ""
    
        
End Sub




Private Sub ImprimeListadoTPV()
        
            vCampos = ""
            For NumRegElim = 1 To Me.ListView3.ListItems.Count
                '                                                En el tag YA esta grabado
                If ListView3.ListItems(NumRegElim).Checked Then
                    vCampos = vCampos & ", (" & vUsu.Codigo & "," & ListView3.ListItems(NumRegElim).Tag
                    If (NumRegElim Mod 25) = 0 Then
                        conn.Execute "insert into `tmpnseries` (`codusu`,`codartic`,`numlinea`,numserie,`numlinealb`,nummante) VALUES " & Mid(vCampos, 2) & ";"
                        vCampos = ""
                        DoEvents
                    End If
                End If
            Next NumRegElim
            If vCampos <> "" Then conn.Execute "insert into `tmpnseries` (`codusu`,`codartic`,`numlinea`,numserie,`numlinealb`,nummante) VALUES " & Mid(vCampos, 2) & ";"


End Sub

Private Sub CargaArticulosAgrupados()
Dim ItmX As ListItem
    Set miRsAux = New ADODB.Recordset
    lwArticulosAgrupados.ListItems.Clear
    cadWHERE2 = "select * from sarticAgrupado ORDER BY  idcaja"
    miRsAux.Open cadWHERE2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set ItmX = lwArticulosAgrupados.ListItems.Add()
        ItmX.Text = miRsAux.Fields(0).Value 'Nº Serie
        ItmX.SubItems(1) = miRsAux.Fields(1).Value 'Nº Factura
        ItmX.SubItems(2) = miRsAux.Fields(2).Value 'Fecha Factura
        ItmX.SubItems(3) = Format(miRsAux!totalmostrar, FormatoImporte) 'Fecha Vencimiento
        
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    cadWHERE2 = ""
End Sub



Private Sub PonerFechaArchivo()
    On Error GoTo ePonerFechaArchivo
    
    vCampos = App.Path & "\Ariges4.exe"
    If Dir(vCampos, vbArchive) = "" Then
        vCampos = App.Path & "\" & App.EXEName & ".exe"
        If Dir(vCampos, vbArchive) = "" Then vCampos = ""
    End If
    If vCampos <> "" Then vCampos = FileDateTime(vCampos)
        
    
    
ePonerFechaArchivo:
    If Err.Number <> 0 Then
        Err.Clear
        vCampos = ""
    End If
End Sub








Private Sub CargarDatosReparaciones()
Dim ItmX As ListItem

On Error GoTo eCargarDatosReparaciones
    cadWHERE2 = "select f.codclien ,nomclien, f.numfactu, f.codtipom , f.fecfactu, codartic,nomartic,cantidad,importel,"
    cadWHERE2 = cadWHERE2 & "  bombamarca,motormodelo,numrepar,m.numalbar,m.observaciones"
    cadWHERE2 = cadWHERE2 & "  from scafac f ,scafac1  c left join scafac_eu m"
    cadWHERE2 = cadWHERE2 & " on c.codtipom=m.codtipom and c.numfactu=m.numfactu and c.fecfactu=m.fecfactu and"
    cadWHERE2 = cadWHERE2 & " C.Codtipoa = m.Codtipoa And C.Numalbar = m.Numalbar"
    cadWHERE2 = cadWHERE2 & " ,slifac l WHERE f.codtipom=c.codtipom and"
    cadWHERE2 = cadWHERE2 & " f.numfactu=c.numfactu and f.fecfactu=c.fecfactu and c.codtipom=l.codtipom and"
    cadWHERE2 = cadWHERE2 & " c.numfactu=l.numfactu and c.fecfactu=l.fecfactu and C.Codtipoa = L.Codtipoa And C.Numalbar = L.Numalbar"
    cadWHERE2 = cadWHERE2 & "  and c.codtipoa='ALO'"
  
    If Trim(txtMatr.Text) <> "" Then
        If SeparaCampoBusqueda("T", "bombamarca", txtMatr.Text, CadenaDesdeOtroForm) > 0 Then Err.Raise 513, , "Obteniendo cadena busqueda"
        cadWHERE2 = cadWHERE2 & "  and " & CadenaDesdeOtroForm
    End If
    
    lwTaxco.ListItems.Clear
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open cadWHERE2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If miRsAux.EOF Then
        Set ItmX = lwTaxco.ListItems.Add()
        ItmX.SubItems(5) = "ningun registro devuelto"
        ItmX.ListSubItems(5).Bold = True
    Else
        While Not miRsAux.EOF
            Set ItmX = lwTaxco.ListItems.Add()
            
            'cli fac ser fec art  desc cant imp matr kms obs ordncli ordse ordefec
            ItmX.Text = Format(miRsAux!codClien, "000000")
            ItmX.ToolTipText = miRsAux!NomClien
            ItmX.SubItems(1) = Format(miRsAux!Numfactu, "000000")
            ItmX.SubItems(2) = IIf(miRsAux!codtipom = "FA5", "L", "LM")
            ItmX.SubItems(3) = Format(miRsAux!FecFactu, "dd/mm/yyyy")
            ItmX.SubItems(4) = miRsAux!codArtic
            ItmX.SubItems(5) = miRsAux!NomArtic
            ItmX.SubItems(6) = Right(Space(10) & Format(miRsAux!cantidad, FormatoImporte), 10)
            ItmX.SubItems(7) = Right(Space(10) & Format(miRsAux!ImporteL, FormatoImporte), 10)
            ItmX.SubItems(8) = DBLet(miRsAux!bombamarca, "T")
            ItmX.ListSubItems(8).ToolTipText = DBLet(miRsAux!motorModelo, "T")
            If IsNull(miRsAux!numrepar) Then
                ItmX.SubItems(9) = " "
            Else
                ItmX.SubItems(9) = Right(Space(10) & Format(DBLet(miRsAux!numrepar, "N"), "#,##0"), 10)
            End If
            
            ItmX.SubItems(10) = DBLet(miRsAux!Observaciones, "T")
            'Para ordenaciones especiales
            ' Cliente
            ItmX.SubItems(11) = Format(miRsAux!codClien, "000000") & Format(miRsAux!FecFactu, "yyyymmdd") & miRsAux!codtipom & Format(miRsAux!Numfactu, "000000")
            'numfac
            ItmX.SubItems(12) = Format(miRsAux!Numfactu, "0000000") & miRsAux!codtipom
            ItmX.SubItems(13) = miRsAux!codtipom & Format(miRsAux!Numfactu, "0000000")
            'ordefec
            ItmX.SubItems(14) = Format(miRsAux!FecFactu, "yyyymmdd") & Format(miRsAux!Numfactu, "000000") & miRsAux!codtipom
            miRsAux.MoveNext
        Wend
    End If
    miRsAux.Close
    
eCargarDatosReparaciones:
    If Err.Number <> 0 Then MuestraError Err.Number, , Err.Description
    Set miRsAux = Nothing
End Sub

Private Sub txtMatr_KeyPress(KeyAscii As Integer)
    KEYpressGnral KeyAscii, 3, False
End Sub


Private Function HayFraSelccionada() As Boolean
    cadWhere = ""
    If lwTaxco.ListItems.Count = 0 Then
        cadWhere = "Ningund dato seleccionado"
    Else
        If lwTaxco.SelectedItem Is Nothing Then
            cadWhere = "Seleccione alguna de las factura"
        Else
            If lwTaxco.ListItems.Count = 1 Then
                If lwTaxco.SelectedItem.Text = "" Then cadWhere = "Seleccione alguna de las factura"
            End If
        End If
    End If
            
    
    
        
    If cadWhere <> "" Then
        MsgBox cadWhere, vbExclamation
        HayFraSelccionada = False
    Else
        HayFraSelccionada = True
    End If
End Function

Private Sub ImprimeFra()

    If Not HayFraSelccionada Then Exit Sub
    
    
    If Not PonerParamRPT2(IIf(True, 94, 12), "", 0, cadWhere, False, "", 0) Then Exit Sub
      
      
      
    'PUNTO VERDE
    '--------------------------------------------------------------------------
    'If vParamAplic.ArtReciclado <> "" Then
    '    cadParam = cadParam & "PuntoVerde= """ & vParamAplic.ArtReciclado & """|"
    '    numParam = numParam + 1
    'End If
      
    'Nombre fichero .rpt a Imprimir
    frmImprimir.NombreRPT = cadWhere
    cadWHERE2 = ""
    With lwTaxco.SelectedItem
    
        cadWhere = IIf(.SubItems(2) = "L", "FA5", "FAO")
    
        cadWhere = "{scafac.codtipom}='" & cadWhere & "'"
        If Not AnyadirAFormula(cadWHERE2, cadWhere) Then Exit Sub
        
        'Nº Factura
        cadWhere = "{scafac.numfactu}=" & Val(.SubItems(1))
        If Not AnyadirAFormula(cadWHERE2, cadWhere) Then Exit Sub
        
        'Fecha Factura
        cadWhere = "{scafac.fecfactu}= Date(" & Year(.SubItems(3)) & "," & Month(.SubItems(3)) & "," & Day(.SubItems(3)) & ")"
        If Not AnyadirAFormula(cadWHERE2, cadWhere) Then Exit Sub
        
        
    End With
   
     
        
       
     
     
     With frmImprimir
            'Nuevo. Febrero 2010
            .SeleccionaRPTCodigo = pRptvMultiInforme
            .FormulaSeleccion = cadWHERE2
            .OtrosParametros = "|pCodigoISO=""""|pCodigoRev=""""|PuntoVerde= ""SI""|"
            .NumeroParametros = 3
            .NombrePDF = .NombreRPT
            .SoloImprimir = False
            .EnvioEMail = False
            .NumeroCopias = 1
            .Opcion = 53
            .Titulo = "Factura"
            .Show vbModal
    End With




End Sub




Private Sub CambiaKilometros()
    If Not HayFraSelccionada Then Exit Sub
    
    
    
    
    
    
    cadWHERE2 = ""
    vCampos = ""
    
    With lwTaxco.SelectedItem
    
        cadWhere = IIf(.SubItems(2) = "L", "FA5", "FAO")
    
        cadWHERE2 = "scafac_eu.codtipom='" & cadWhere & "'"
        'Nº Factura
        cadWHERE2 = cadWHERE2 & " AND scafac_eu.numfactu=" & Val(.SubItems(1))
        'Fecha Factura
        cadWHERE2 = cadWHERE2 & " AND scafac_eu.fecfactu = " & DBSet(.SubItems(3), "F")
        cadWHERE2 = cadWHERE2 & " AND scafac_eu.codtipoa = 'ALO'"
        
        cadWhere = .Text & " - " & .ToolTipText & "|"
        cadWhere = cadWhere & .SubItems(2) & .SubItems(1) & "    de " & .SubItems(3) & "|"
        cadWhere = cadWhere & .SubItems(8)
        If .ListSubItems(8).ToolTipText <> "" Then cadWhere = cadWhere & "    Modelo  " & .ListSubItems(8).ToolTipText
        cadWhere = cadWhere & "|"
        cadWhere = cadWhere & Trim(.SubItems(9)) & "|"
        
        
        
    End With
    CadenaDesdeOtroForm = cadWhere
    frmVarios.Opcion = 17
    frmVarios.Show vbModal
    If CadenaDesdeOtroForm <> "" Then
        Screen.MousePointer = vbHourglass
        NumRegElim = Val(lwTaxco.SelectedItem.SubItems(1))
        cadWhere = lwTaxco.SelectedItem.SubItems(2)
        'scafac_eu numrepar
        cadWHERE2 = "UPDATE  scafac_eu  SET  numrepar = " & CadenaDesdeOtroForm & " WHERE " & cadWHERE2
        If ejecutar(cadWHERE2, False) Then
            cmdBusMatr_Click
            For davidNumalbar = 1 To lwTaxco.ListItems.Count
                If Val(lwTaxco.ListItems(davidNumalbar).SubItems(1)) = NumRegElim Then
                    
                    If cadWhere = lwTaxco.ListItems(davidNumalbar).SubItems(2) Then
                        lwTaxco.ListItems(davidNumalbar).Selected = True
                        Set lwTaxco.SelectedItem = lwTaxco.ListItems(davidNumalbar)
                        lwTaxco.ListItems(davidNumalbar).EnsureVisible
                        Exit For
                    End If
                End If
            Next
        End If
        
    End If
    cadWHERE2 = ""
    cadWhere = ""
    davidNumalbar = 0
End Sub


Private Sub ImprimeSeleccionReparaciones()
    Screen.MousePointer = vbHourglass
    conn.Execute "DELETE FROM tmpinformes WHERE codusu = " & vUsu.Codigo
    
    cadWHERE2 = ""
    cadWhere = ""
    vCampos = "INSERT INTO tmpinformes  (codusu,codigo1,campo1,fecha1,campo2,nombre1,nombre2,importe1,importe2,nombre3,importe3,obser ) VALUES "
    For NumRegElim = 1 To lwTaxco.ListItems.Count
        With lwTaxco.ListItems(NumRegElim)
            cadWhere = ", (" & vUsu.Codigo & "," & .Text & "," & .SubItems(1) & "," & DBSet(.SubItems(3), "F") & "," & IIf(.SubItems(2) = "L", 0, 1)
            cadWhere = cadWhere & "," & DBSet(.SubItems(4), "T") & "," & DBSet(.SubItems(5), "T") & "," & DBSet(.SubItems(6), "N")
            cadWhere = cadWhere & "," & DBSet(.SubItems(7), "N") & "," & DBSet(.SubItems(8), "T") & ","
            If Trim(.SubItems(9)) = "" Then
                cadWhere = cadWhere & "null"
            Else
                cadWhere = cadWhere & DBSet(.SubItems(9), "N", "N")
            End If
            cadWhere = cadWhere & "," & DBSet(.SubItems(10), "T") & ")"
    
            cadWHERE2 = cadWHERE2 & cadWhere
            
        End With
        If Len(cadWHERE2) > 10000 Then
            cadWHERE2 = Mid(cadWHERE2, 2)
            cadWHERE2 = vCampos & cadWHERE2
            conn.Execute cadWHERE2
            cadWHERE2 = ""
        End If
    Next
    If cadWHERE2 <> "" Then
        cadWHERE2 = Mid(cadWHERE2, 2)
        cadWHERE2 = vCampos & cadWHERE2
        conn.Execute cadWHERE2
    End If
    Screen.MousePointer = vbDefault
    cadWHERE2 = ""
    cadWhere = ""
    vCampos = ""
    With frmImprimir
        .FormulaSeleccion = "{tmpinformes.codusu} = " & vUsu.Codigo
        .OtrosParametros = "|pEmpresa=""" & vParam.NombreEmpresa & """|"
        .NumeroParametros = 2 'numParam

        .SoloImprimir = False
        .EnvioEMail = False
        .Opcion = 3000
        .Titulo = Me.Caption
        .NombreRPT = "taxListaReparaF.rpt"
        .ConSubInforme = False
        .MostrarTreeDesdeFuera = True
        .Show vbModal
    End With
    
    
    
End Sub




'************************************************************************************************************************
'************************************************************************************************************************
' Frame actividades
'************************************************************************************************************************
Private Sub CargaListActividades()
Dim ItmX


    cadWHERE2 = "Select * from sactiv where codactiv in (" & Mid(cadWhere, 2) & ") ORDER BY codactiv"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open cadWHERE2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set ItmX = Me.lwActividad.ListItems.Add
        'El primer campo será codtipom si llamamos desde Ventas
        ' y será codprove si llamamos desde Compras
        ItmX.Text = Format(miRsAux!codactiv, "0000")
        ItmX.SubItems(1) = miRsAux!nomactiv
        ItmX.Checked = True
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing


End Sub

Private Sub cargaFactPrevisualizacion()
Dim Sql As String
Dim ItmX
Dim Total1 As Currency

    Total1 = 0

    ListView6.SmallIcons = frmPpal.ImgListPpal
    ListView7.SmallIcons = frmPpal.ImgListPpal
    ListView8.SmallIcons = frmPpal.ImgListPpal
    ListView9.SmallIcons = frmPpal.ImgListPpal
    
    
    Me.ListView6.ListItems.Clear
    Me.ListView7.ListItems.Clear
    Me.ListView8.ListItems.Clear
    Me.ListView9.ListItems.Clear

    ' clientes
    Label11.Caption = "Cargando clientes"
    Sql = Replace(cadWhere, "'XXXX'", "scaalb.codclien, sclien.nomclien")
    If cadWHERE2 <> "" Then Sql = Sql & cadWHERE2
    Sql = Sql & " group by 1,2 order by 1"

    Set miRsAux = New ADODB.Recordset
    miRsAux.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set ItmX = Me.ListView6.ListItems.Add
        'El primer campo será codtipom si llamamos desde Ventas
        ' y será codprove si llamamos desde Compras
        
        ItmX.SmallIcon = 0
        
        ItmX.Text = Format(miRsAux.Fields(0), "000000")
        ItmX.SubItems(1) = miRsAux.Fields(1)
        ItmX.ListSubItems(1).ToolTipText = miRsAux.Fields(1)
        ItmX.SubItems(2) = Format(miRsAux.Fields(2), "###,###,##0.00")
        
        Total1 = Total1 + miRsAux.Fields(2)
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    
    ' formas de pago
    Label11.Caption = "Cargando Formas de Pago"
    Sql = Replace(cadWhere, "'XXXX'", "scaalb.codforpa, sforpa.nomforpa")
    Sql = Sql & " inner join sforpa on scaalb.codforpa = sforpa.codforpa "
    If cadWHERE2 <> "" Then Sql = Sql & cadWHERE2
    Sql = Sql & " group by 1,2 order by 1"

    Set miRsAux = New ADODB.Recordset
    miRsAux.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set ItmX = Me.ListView7.ListItems.Add
        'El primer campo será codtipom si llamamos desde Ventas
        ' y será codprove si llamamos desde Compras
        
        ItmX.SmallIcon = 0
        
        ItmX.Text = Format(miRsAux.Fields(0), "0000")
        ItmX.SubItems(1) = miRsAux.Fields(1)
        ItmX.ListSubItems(1).ToolTipText = miRsAux.Fields(1)
        ItmX.SubItems(2) = Format(miRsAux.Fields(2), "###,###,##0.00")
        
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    
    
    ' agentes
    Label11.Caption = "Cargando Agentes"
    Sql = Replace(cadWhere, "'XXXX'", "scaalb.codagent, sagent.nomagent")
    Sql = Sql & " inner join sagent on scaalb.codagent = sagent.codagent "
    If cadWHERE2 <> "" Then Sql = Sql & cadWHERE2
    Sql = Sql & " group by 1,2 order by 1"

    Set miRsAux = New ADODB.Recordset
    miRsAux.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set ItmX = Me.ListView8.ListItems.Add
        'El primer campo será codtipom si llamamos desde Ventas
        ' y será codprove si llamamos desde Compras
        
        ItmX.SmallIcon = 0
        
        ItmX.Text = Format(miRsAux.Fields(0), "0000")
        ItmX.SubItems(1) = miRsAux.Fields(1)
        ItmX.ListSubItems(1).ToolTipText = miRsAux.Fields(1)
        ItmX.SubItems(2) = Format(miRsAux.Fields(2), "###,###,##0.00")
        
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    
    ' actividad
    Label11.Caption = "Cargando Actividades"
    Sql = Replace(cadWhere, "'XXXX'", "sclien.codactiv, sactiv.nomactiv")
    Sql = Sql & " inner join sactiv on sclien.codactiv = sactiv.codactiv "
    If cadWHERE2 <> "" Then Sql = Sql & cadWHERE2
    Sql = Sql & " group by 1,2 order by 1"

    Set miRsAux = New ADODB.Recordset
    miRsAux.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set ItmX = Me.ListView9.ListItems.Add
        'El primer campo será codtipom si llamamos desde Ventas
        ' y será codprove si llamamos desde Compras
        
        ItmX.SmallIcon = 0
        
        ItmX.Text = Format(miRsAux.Fields(0), "000")
        ItmX.SubItems(1) = miRsAux.Fields(1)
        ItmX.ListSubItems(1).ToolTipText = miRsAux.Fields(1)
        ItmX.SubItems(2) = Format(miRsAux.Fields(2), "###,###,##0.00")
        
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    
    Text2.Text = Format(Total1, "###,###,##0.00")

    Label11.Caption = ""

End Sub

Private Sub CargaListArticulosProv()
Dim ItmX

    ListView5.SmallIcons = frmPpal.ImgListPpal
    
    
    Me.ListView5.ListItems.Clear

    cadWHERE2 = "Select slispr.codartic, sartic.nomartic, slispr.precioac from slispr inner join sartic on slispr.codartic = sartic.codartic and " & cadWhere
    
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open cadWHERE2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set ItmX = Me.ListView5.ListItems.Add
        'El primer campo será codtipom si llamamos desde Ventas
        ' y será codprove si llamamos desde Compras
        
        ItmX.SmallIcon = 0
        
        ItmX.Text = miRsAux!codArtic
        ItmX.SubItems(1) = miRsAux!NomArtic
        ItmX.SubItems(2) = Format(miRsAux!precioac, "###,##0.0000")
        
        ItmX.Checked = True
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing

End Sub

Private Sub cargaempresasbloquedas()
Dim IT As ListItem
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim Errores As String

    On Error GoTo Ecargaempresasbloquedas
    Set Rs = New ADODB.Recordset
    Sql = "select empresasariges.codempre,nomempre,nomresum,usuarioempresasariges.codempre bloqueada from usuarios.empresasariges left join usuarios.usuarioempresasariges on "
    Sql = Sql & " empresasariges.codempre = usuarioempresasariges.codempre And (usuarioempresasariges.codusu = " & Parametros & " Or codusu Is Null)"
    '[Monica] solo ariagro
    Sql = Sql & " WHERE ariges like 'ariges%' "
    Sql = Sql & " ORDER BY empresasariges.codempre"
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not Rs.EOF
        Errores = Format(Rs!codempre, "00000")
        Sql = "C" & Errores
        
        If IsNull(Rs!bloqueada) Then
            'Va al list de la derecha
            Set IT = ListView99(0).ListItems.Add(, Sql)
            IT.SmallIcon = 1
        Else
            Set IT = ListView99(1).ListItems.Add(, Sql)
            IT.SmallIcon = 2
        End If
        IT.Text = Errores
        IT.SubItems(1) = Rs!nomempre
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    Errores = ""
    Exit Sub
Ecargaempresasbloquedas:
    MuestraError Err.Number, Err.Description
    Me.cmdBloqEmpre(0).Enabled = False
    Errores = ""
    Set Rs = Nothing
End Sub





Private Sub CargaPuntosCaducados()


    Label2(9).Caption = "Prepara datos "
    Label2(9).Refresh

    'Empezamos a cargar datos
    Label2(9).Caption = "Leyendo movimientos puntos"
    Label2(9).Refresh
    Sql = " select smovalpuntos.codclien,sum(smovalpuntos.puntos) caduca,sclien.puntos PuntosCliente,nomclien"
    Sql = Sql & " ,max(if(concepto=3,fechaalb,'2001-01-01')) feccad  "
    Sql = Sql & " ,max(if(concepto=1,fechaalb,'2001-01-01')) feccanj  "
    Sql = Sql & " FROM smovalpuntos inner join sclien on smovalpuntos.codclien=sclien.codclien"
    Sql = Sql & " WHERE sclien.Puntos > 0 AND fechaalb <= "
    Sql = Sql & DBSet(DateAdd("d", -vParamAplic.DiasCaducidadPuntos, Now), "F")
    If cadWhere <> "" Then Sql = Sql & " AND smovalpuntos.codclien = " & cadWhere
    Sql = Sql & " group by smovalpuntos.codclien"
    Sql = Sql & " order by smovalpuntos.codclien"

    Set miRsAux = New ADODB.Recordset
    miRsAux.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = -1
    NE = 0
    While Not miRsAux.EOF
    
            If miRsAux!PuntosCliente >= miRsAux!caduca Then
                Importe = miRsAux!caduca
                OK = 1
            Else
                Importe = miRsAux!PuntosCliente
                OK = 2
            End If

    
    
            If Importe > 0 Then
                'Insert
                NumRegElim = miRsAux!codClien
                NE = NE + 1
                ListView10.ListItems.Add , "K" & Format(NE, "00000"), Format(NumRegElim, "0000")
                ListView10.ListItems(NE).SubItems(1) = miRsAux!NomClien
                ListView10.ListItems(NE).SubItems(2) = Format(miRsAux!PuntosCliente, FormatoCantidad)
                ListView10.ListItems(NE).SubItems(3) = Format(miRsAux!caduca, FormatoCantidad)
                
                ListView10.ListItems(NE).SubItems(4) = Format(Importe, FormatoCantidad)
                Sql = " "
                If Year(CDate(miRsAux!feccanj)) > 2010 Then Sql = Format(miRsAux!feccanj, "dd/mm/yyyy")
                ListView10.ListItems(NE).SubItems(5) = Sql
    
                
                Sql = " "
                If Year(CDate(miRsAux!feccad)) > 2019 Then Sql = Format(miRsAux!feccad, "dd/mm/yyyy")
                ListView10.ListItems(NE).SubItems(6) = Sql
    
    
    
            End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If NE > 0 Then
        If cadWhere <> "" Then Me.Caption = Me.Caption & "  " & ListView10.ListItems(NE).SubItems(1)
            
    End If
    
    Label2(9).Caption = ""
    If ListView10.ListItems.Count = 0 Then
        Sql = "Ningun punto a caducar "
        If cadWhere <> "" Then Sql = Sql & "  para el cliente."
        MsgBox Sql, vbExclamation
        Me.cmdPuntosCaducados(0).Enabled = False
    End If
    
    
    
End Sub



Private Sub CaducarPuntos()

    For NumRegElim = 1 To Me.ListView10.ListItems.Count
        If ListView10.ListItems(NumRegElim).Checked Then
            Label2(9).Caption = ListView10.ListItems(NumRegElim).SubItems(1)
            Label2(9).Refresh
            
            
            Sql = DevuelveDesdeBD(conAri, "max(numero)", "smovalpuntos", "codclien", ListView10.ListItems(NumRegElim).Text)
            Sql = Val(Sql) + 1
        
            ''codclien,numero,codtipom,numalbar,fechaalb,concepto,puntos,fecMov,observaciones
            Importe = -1 * ImporteFormateado(ListView10.ListItems(NumRegElim).SubItems(4))
            NE = vParamAplic.DiasCaducidadPuntos
            Errores = DateAdd("d", -NE, Now)                '3: CADUCAER puntos
            Sql = "(" & ListView10.ListItems(NumRegElim).Text & "," & Sql & ",'',0," & DBSet(Errores, "F") & ",3," & DBSet(Importe, "N", "N")
            Sql = Sql & "," & DBSet(Now, "FH") & "," & DBSet("Realizado por " & vUsu.Login, "T") & ")"
            
            Sql = "INSERT INTO smovalpuntos(codclien,numero,codtipom,numalbar,fechaalb,concepto,puntos,fecMov,observaciones) VALUES " & Sql
            If ejecutar(Sql, False) Then
                Sql = "+"
                If Importe < 0 Then
                    Sql = "-"
                    Importe = Abs(Importe)
                End If
                Sql = "UPDATE sclien set puntos = puntos " & Sql & DBSet(Importe, "N")
                Sql = Sql & " WHERE codclien =" & ListView10.ListItems(NumRegElim).Text
                conn.Execute Sql
            End If
        End If
    Next
    

    
End Sub


Private Sub hazImprimirCaducidad()
    
    conn.Execute "DELETE FROM tmpinformes WHERE codusu = " & vUsu.Codigo
    
    cadWHERE2 = ""
    cadWhere = ""
    vCampos = "INSERT INTO tmpinformes  (codusu,codigo1,nombre1,importe1,importe2,importe3,fecha1,fecha2 ) VALUES "
    For NumRegElim = 1 To ListView10.ListItems.Count
        With ListView10.ListItems(NumRegElim)
            cadWhere = ", (" & vUsu.Codigo & "," & .Text & "," & DBSet(.SubItems(1), "T") & "," & DBSet(.SubItems(2), "N")
            cadWhere = cadWhere & "," & DBSet(.SubItems(3), "N") & "," & DBSet(.SubItems(4), "N")
            cadWhere = cadWhere & "," & DBSet(Trim(.SubItems(5)), "F", "S") & "," & DBSet(Trim(.SubItems(6)), "F", "S") & ")"
    
            cadWHERE2 = cadWHERE2 & cadWhere
            
        End With
    Next
    
    cadWHERE2 = Mid(cadWHERE2, 2)
    cadWHERE2 = vCampos & cadWHERE2
    conn.Execute cadWHERE2

    
    cadWHERE2 = ""
    cadWhere = ""
    vCampos = ""
    With frmImprimir
        .FormulaSeleccion = "{tmpinformes.codusu} = " & vUsu.Codigo
        .OtrosParametros = "|pEmpresa=""" & vParam.NombreEmpresa & """|"
        .NumeroParametros = 2 'numParam

        .SoloImprimir = False
        .EnvioEMail = False
        .Opcion = 3000
        .Titulo = Me.Caption
        .NombreRPT = "rPuntosCaducados.rpt"
        .ConSubInforme = False
        .MostrarTreeDesdeFuera = False
        .Show vbModal
    End With
        
End Sub



Private Sub CargaAnticipoProveedor()

    Label2(12).Caption = cadWHERE2

    Sql = " select * FROM sproveanticipo WHERE   codprove=" & cadWhere & " AND  descontado =0"
    Sql = Sql & " order by fechaant,idanticipo"

    Set miRsAux = New ADODB.Recordset
    miRsAux.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NE = 0
    While Not miRsAux.EOF
    
        'Insert
        NE = NE + 1
        ListView11.ListItems.Add , "K" & Format(miRsAux!idAnticipo, "0000")
        ListView11.ListItems(NE).Text = Format(miRsAux!idAnticipo, "0000")
        ListView11.ListItems(NE).SubItems(1) = miRsAux!numdocum
        ListView11.ListItems(NE).SubItems(2) = Format(miRsAux!fechaant, "dd/mm/yyyy")
        
        ListView11.ListItems(NE).SubItems(3) = Format(miRsAux!Importe, FormatoCantidad)
        
                
        
    
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    

    
    
End Sub



Private Sub cargarDatosImportarPedidosEXCEL()
Dim Ruta As String


    On Error GoTo ecargarDatosImportarPedidosEXCEL
    
    
    lwPedidosFontenas(1).ListItems.Clear
    lwPedidosFontenas(0).ListItems.Clear
    
    
    
    Ruta = Me.Tag '"C:\PedidosAriadna"
    If Dir(Ruta, vbDirectory) = "" Then MkDir Ruta
    'Procesados
    If Dir(Ruta & "\Procesados", vbDirectory) = "" Then MkDir Ruta & "\Procesados"
    
    
    
    
    
    
    
    Label2(16).Caption = "Cargando datos .."
    Label2(16).Refresh
    CadenaDesdeOtroForm = "    "
    
    '1) Ficheros pendientes de PROCESAR
    '   1) CSV directo a importar
    Sql = Dir(Ruta & "\*.*", vbDirectory)
    Do While Sql <> ""   ' Inicia el bucle.
        ' Ignora el directorio actual y el que lo abarca.
        If Sql <> "." And Sql <> ".." Then
            OK = InStrRev(Sql, ".")
            If OK > 0 Then
                NE = 0
                cadWHERE2 = UCase(Mid(Sql, OK + 1))
                If cadWHERE2 = "XLS" Or cadWHERE2 = "XLSX" Then
                    'EXCEL
                    NE = 6
                Else
                    If cadWHERE2 = "CSV" Then NE = 5
                    
                End If
                
                If NE > 0 Then Me.lwPedidosFontenas(1).ListItems.Add , , Sql, , NE
              
            End If
        End If

        Sql = Dir   ' Obtiene siguiente entrada.
    Loop
    
    
    Label2(16).Caption = "Leyendo BD"
    Label2(16).Refresh
    
    Sql = "  select codigo,fechaped,count(*) lineas from slipedXLS group by 1,2  order by 1"
    '1.- Cargamos datos pendientes de procesar
    '  A)   Agrupacion
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NE = 0
    Sql = ""
    While Not miRsAux.EOF
        NE = NE + 1
        Set ItmX = lwPedidosFontenas(0).ListItems.Add(, "K" & Format(miRsAux!Codigo, "0000"), , , 9)
        ItmX.Text = Format(miRsAux!Codigo, "0000")
        ItmX.SubItems(1) = Format(miRsAux!FechaPed, "dd/mm/yyyy")
        ItmX.SubItems(2) = CStr(Format(miRsAux!Lineas, "00"))
        Sql = Sql & ", " & miRsAux!Codigo
        ItmX.Tag = 0
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If Sql <> "" Then
        Sql = "Select numpedcl from scaped where numpedcl in (" & Mid(Sql, 2) & ") ORDER BY 1"
        miRsAux.Open Sql, conn, adOpenKeyset, adLockPessimistic, adCmdText
        
        For NE = 1 To lwPedidosFontenas(0).ListItems.Count
            Sql = "numpedcl = " & lwPedidosFontenas(0).ListItems(NE).Text
            miRsAux.Find Sql, , adSearchForward, 1
            If miRsAux.EOF Then
                lwPedidosFontenas(0).ListItems(NE).Tag = 1
                lwPedidosFontenas(0).ListItems(NE).ForeColor = vbRed
                lwPedidosFontenas(0).ListItems(NE).ListSubItems(1).ForeColor = vbRed
            End If
        Next
    End If
    
    If lwPedidosFontenas(0).ListItems.Count > 0 Then
        lwPedidosFontenas(0).SelectedItem = lwPedidosFontenas(0).ListItems(1)
        PonerDatosPedidosFontenas
    End If




ecargarDatosImportarPedidosEXCEL:
    If Err.Number <> 0 Then MuestraError Err.Number, , Err.Description
    Set miRsAux = Nothing
    Label2(16).Caption = ""
    
End Sub


Private Sub PonerDatosPedidosFontenas()

    

    Me.lwPedidosFontenas(2).ListItems.Clear
    If lwPedidosFontenas(0).SelectedItem Is Nothing Then Exit Sub
    
    Label2(16).Caption = "Leyendo pedido"
    Label2(16).Refresh
        
    
    
    Sql = "  select slipedXLS.*,sartic.codartic s_codart,sartic.nomartic s_nomar from slipedXLS left join sartic on slipedXLS.codartic=sartic.codartic"
    Sql = Sql & " WHERE codigo = " & Mid(lwPedidosFontenas(0).SelectedItem.Key, 2)
    
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not miRsAux.EOF
        Set ItmX = lwPedidosFontenas(2).ListItems.Add
        ItmX.Text = miRsAux!codAlmac
        ItmX.SubItems(1) = miRsAux!codArtic
        ItmX.SubItems(2) = miRsAux!NomArtic
        ItmX.SubItems(3) = miRsAux!servidas
        ItmX.SubItems(4) = DBLet(miRsAux!numLote, "T")
        ItmX.ToolTipText = miRsAux!NomArtic
        ItmX.ListSubItems(2).ToolTipText = miRsAux!NomArtic
        ItmX.Tag = 0
        If IsNull(miRsAux!s_codart) Then
           For NE = 1 To ItmX.ListSubItems.Count
                ItmX.ListSubItems(NE).ForeColor = vbRed
            Next
            ItmX.Tag = 1
        Else
            If miRsAux!s_nomar <> miRsAux!NomArtic Then
                ItmX.ListSubItems(2).ForeColor = vbRed
                ItmX.ListSubItems(2).ToolTipText = miRsAux!s_nomar
            End If
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close

    Label2(16).Caption = ""
    
End Sub
Private Sub LanzaProcesoImportacionXLS()
Dim N As Byte


    On Error GoTo eLanzaProcesoImportacionXLS
        
        
    
    If MsgBox("¿Lanzar proceso EXCEL-CSV ?", vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    
    Sql = Me.Tag '"C:\PedidosAriadna"
    
    
    'Primero renombrar
    If LCase(lwPedidosFontenas(1).SelectedItem.Text) <> "convertir.xlsx" Then Name Me.Tag & "\" & lwPedidosFontenas(1).SelectedItem.Text As Me.Tag & "\convertir.xlsx"
    
    
    Sql = App.Path & "\convertir.bat"
    If Dir(Sql, vbArchive) = "" Then Err.Raise 513, , "No tiene el programa de conversion. "
        
           
            
        
    Screen.MousePointer = vbHourglass
    Shell Sql, vbNormalFocus
    Espera 0.5
    Screen.MousePointer = vbHourglass
    N = 0
    NumRegElim = 0
    Sql = ""
    Do
        
        Screen.MousePointer = vbHourglass
        Label2(16).Caption = NE & " seg"
        Label2(16).Refresh
        Sql = Dir(Me.Tag & "\*.convertir.xlsx")
        If Sql <> "" Then
            N = N + 1
            Espera 1
        Else
            N = 16
        End If
    Loop Until N > 15
        
    'OK procesamos el fihero
    N = 0
    NumRegElim = 0
    Sql = ""

    Do
        Sql = Dir(Me.Tag & "\*.csv")
        If Sql <> "" Then
            'Hay un csv. Lo proceso
            N = N + 1
            
            ImportarFontenasCSV False, CStr(Sql)
            
            NumRegElim = 20  'Ya he preocsado un fichero
        Else
            Espera 0.95
            Label2(16).Caption = N
        End If
        NumRegElim = NumRegElim + 1
        
    Loop Until NumRegElim > 25
    
    
    
    
    
    
    
    If N = 0 Then MsgBox "No se ha generado ningún csv", vbExclamation
    
    
eLanzaProcesoImportacionXLS:
    If Err.Number <> 0 Then MuestraError Err.Number, , Err.Description
    Label2(16).Caption = ""
    Screen.MousePointer = vbDefault
    Me.MousePointer = vbDefault
End Sub




Private Sub ImportarFontenasCSV(ConPregunta As Boolean, Fichero As String)


    On Error GoTo eImportarFontenasCSV

    If ConPregunta Then
        Sql = "Va a importar el fichero: " & lwPedidosFontenas(1).SelectedItem.Text & vbCrLf & vbCrLf & "¿Continuar?"
        If MsgBox(Sql, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    
    NE = FreeFile
    Sql = Me.Tag & "\" & Fichero
    Open Sql For Input As #NE
    
    
    Errores = ""
    
    
    cadWHERE2 = ""
    OK = 0
    NumRegElim = 0
    While Not EOF(NE)
        OK = OK + 1
        Line Input #NE, Sql
        
        'FONTENAS;2022-02-08;258;2;PSH2E14RC1O;FR 1000ML ACONDICIONADOR OZONE SWEET PSH;12;12;1;LOTE 1
        codArtic = Split(Sql, ";")
        
        If OK = 1 Then
            'primera linea.
            'Si es texto... NO la insertamos
            If UBound(codArtic) > 0 Then
                If IsNumeric(codArtic(2)) Then OK = OK + 1  'es valido
            End If
        End If
        If OK > 1 Then
            
            
            If UBound(codArtic) <> 9 Then
                Errores = Errores & "Lin " & Format(OK, "000") & "   NºCampos incorrectos"
            Else
                'Comprobaciones
                'Formato numero
                
                
                If Trim(codArtic(5)) = "" Then codArtic(5) = "0"
                
                vCampos = ""
                If Not IsNumeric(codArtic(2)) Then vCampos = " pedido"
                If Not IsNumeric(codArtic(9)) Then vCampos = " almacen"
                If Not IsNumeric(codArtic(5)) Then vCampos = " pedidas"
                If Not IsNumeric(codArtic(6)) Then vCampos = " servidas"
                If Not IsNumeric(codArtic(7)) Then vCampos = " bultos"
                'If Not EsFechaOK(codArtic(1)) Then vCampos = " fecha"
                
                If Trim(codArtic(3)) = "" Then vCampos = " articulo"
                'If Trim(codArtic(9)) = "" Then vCampos = " lote"
                
                'Si ha ido bien
                If vCampos <> "" Then
                    Errores = Errores & "Lin " & Format(OK, "000") & "   ERROR " & vCampos & vbCrLf
                Else
                    'Comprobacion BASICA
                    If Val(codArtic(2)) <> NumRegElim Then
                        vCampos = DevuelveDesdeBD(conAri, "Codigo", "slipedxls", "codigo", codArtic(2))
                        If vCampos <> "" Then Errores = Errores & "Lin " & Format(OK, "000") & "   YA existe el pedido EXCEL " & vCampos & vbCrLf
                        
                        NumRegElim = Val(codArtic(2))
                    End If
                    
                    
                    'Por tema de velocidad debieraos ir "a tramos", pero de moemento va asi
                    vCampos = DevuelveDesdeBD(conAri, "codartic", "sartic", "codartic", codArtic(3), "T")
                    If vCampos = "" Then Errores = Errores & "Lin " & Format(OK, "000") & "   No existe articulo: " & codArtic(3) & vbCrLf
                    
                    
                    
                    
                    
                    'slipedxls                          codigo ,fechaped,numlinea,codalmac
                     cadWHERE2 = cadWHERE2 & ", (" & NumRegElim & "," & DBSet(codArtic(1), "F") & "," & OK & "," & codArtic(9)
                     '                          codartic,nomartic,cantidad
                     cadWHERE2 = cadWHERE2 & "," & DBSet(codArtic(3), "T") & "," & DBSet(codArtic(4), "T") & "," & DBSet(codArtic(5), "N")
                    ',servidas,numbultos,numlote,
                     cadWHERE2 = cadWHERE2 & "," & DBSet(codArtic(6), "N") & "," & DBSet(codArtic(7), "T") & "," & DBSet(codArtic(8), "T")
                     'fechahora,usuario,fichero)
                     cadWHERE2 = cadWHERE2 & "," & DBSet(Now, "FH") & "," & DBSet(vUsu.Login, "T") & "," & DBSet(lwPedidosFontenas(1).SelectedItem.Text, "T") & ")"
                    
                End If
            End If
        
        End If
    Wend
    Close #NE
    If Errores <> "" Then
        MsgBox Errores, vbExclamation
    Else
        If cadWHERE2 = "" Then
            MsgBox "Fichero vacio", vbExclamation
        Else
            cadWHERE2 = Mid(cadWHERE2, 2) 'primera coma
            'slipedxls  codigo ,fechaped,numlinea,codalmac,codartic,nomartic,cantidad,servidas,numbultos,numlote,fechahora,usuario,fichero)
            Sql = "INSERT INTO slipedxls  (codigo ,fechaped,numlinea,codalmac,codartic,nomartic,cantidad,servidas,numbultos,numlote,fechahora,usuario,fichero) VALUES " & cadWHERE2
            conn.Execute Sql
        End If
    End If
    
    
    Sql = Me.Tag & "\" & Fichero
    cadWHERE2 = Me.Tag & "\Procesados\" & Format(Now, "yymmdd_hhnn") & Fichero
    FileCopy Sql, cadWHERE2
    Kill Sql
    Espera 1

eImportarFontenasCSV:
    If Err.Number <> 0 Then MuestraError Err.Number, , Err.Description
    Screen.MousePointer = vbDefault
End Sub




Private Sub cargarDatosAlbaranes()
    ListView12.ListItems.Clear
    ListView12.SortKey = 7
    Label2(17).Caption = "Leyendo BBDD    "
    Label2(17).Refresh
    
    Sql = "SELECT scaalb.codtipom,scaalb.numalbar,fechaalb,nomclien,numtermi,factursn,sum(importel) bases,nomforpa"
    Sql = Sql & " FROM scaalb INNER JOIN sforpa ON scaalb.codforpa=sforpa.codforpa "
    Sql = Sql & " INNER JOIN slialb ON scaalb.codtipom=slialb.codtipom and scaalb.numalbar=slialb.numalbar"
    Sql = Sql & " WHERE " & cadWhere
    Sql = Sql & " group by 1,2"
    
    
    Set miRsAux = New ADODB.Recordset
    NumRegElim = 0
    miRsAux.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set ItmX = Me.ListView12.ListItems.Add()
        
        ItmX.Text = "    " & miRsAux!codtipom
        ItmX.SubItems(1) = Format(miRsAux!Numalbar, "0000000")
        ItmX.SubItems(2) = Format(miRsAux!FechaAlb, "dd/mm/yyyy")
        ItmX.SubItems(3) = miRsAux!NomClien
        ItmX.SubItems(4) = Mid(miRsAux!nomforpa, 1, 10)
        ItmX.ListSubItems(4).ToolTipText = miRsAux!nomforpa
        ItmX.SubItems(5) = IIf(IsNull(miRsAux!NumTermi), " ", miRsAux!NumTermi)
        ItmX.SubItems(6) = Format(miRsAux!bases, FormatoImporte)
        
        'Para el orden
        ItmX.SubItems(7) = miRsAux!codtipom & Format(miRsAux!Numalbar, "0000000")
        ItmX.SubItems(8) = IIf(miRsAux!bases < 0, " ", "-") & Format(Abs(miRsAux!bases) * 100, "000000000")
        ItmX.SubItems(9) = Format(miRsAux!FechaAlb, "yyyymmdd") & miRsAux!codtipom & Format(miRsAux!Numalbar, "0000000")
        ItmX.Checked = miRsAux!factursn = 1
        ItmX.Tag = miRsAux!factursn 'para ver que actualizamos despues
    
    
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    
    Label2(17).Caption = ""
    
End Sub
