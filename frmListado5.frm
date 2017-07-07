VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmListado5 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Listado"
   ClientHeight    =   7830
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   13080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameTelefonia 
      Height          =   3015
      Left            =   4920
      TabIndex        =   186
      Top             =   1440
      Visible         =   0   'False
      Width           =   4095
      Begin VB.TextBox txtNumero 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   2640
         TabIndex        =   188
         Text            =   "Text1"
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox txtNumero 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   2640
         TabIndex        =   187
         Text            =   "Text1"
         Top             =   1080
         Width           =   855
      End
      Begin VB.CommandButton cmdContratoTelef 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   1680
         TabIndex        =   189
         Top             =   2400
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   18
         Left            =   2880
         TabIndex        =   190
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Meses duración"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   17
         Left            =   360
         TabIndex        =   193
         Top             =   1560
         Width           =   1110
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Importe terminal"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   16
         Left            =   360
         TabIndex        =   192
         Top             =   1080
         Width           =   1185
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "Datos contrato"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   18
         Left            =   120
         TabIndex        =   191
         Top             =   360
         Width           =   3855
      End
   End
   Begin VB.Frame FramePreguntaDevoluciones 
      Height          =   2175
      Left            =   3240
      TabIndex        =   180
      Top             =   0
      Visible         =   0   'False
      Width           =   5175
      Begin VB.CommandButton cmdDevolucion 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   2640
         TabIndex        =   185
         Top             =   1560
         Width           =   975
      End
      Begin VB.OptionButton optDevol 
         Caption         =   "Factura rectificativa"
         Height          =   195
         Index           =   1
         Left            =   720
         TabIndex        =   183
         Top             =   1080
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.OptionButton optDevol 
         Caption         =   "Albarán venta"
         Height          =   195
         Index           =   0
         Left            =   3240
         TabIndex        =   184
         Top             =   1080
         Width           =   1815
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   17
         Left            =   3960
         TabIndex        =   181
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "Generar devolución"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   17
         Left            =   600
         TabIndex        =   182
         Top             =   360
         Width           =   3855
      End
   End
   Begin VB.Frame FrameListadoComprasClientes 
      Height          =   5775
      Left            =   240
      TabIndex        =   173
      Top             =   720
      Width           =   10695
      Begin VB.CommandButton cmdTraerLineaCompraCliente 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   8280
         TabIndex        =   176
         Top             =   5160
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   16
         Left            =   9480
         TabIndex        =   177
         Top             =   5160
         Width           =   975
      End
      Begin MSComctlLib.ListView lw 
         Height          =   3615
         Index           =   4
         Left            =   120
         TabIndex        =   175
         Top             =   1440
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   6376
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
            Text            =   "Tipo"
            Object.Width           =   1095
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Factura"
            Object.Width           =   1589
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Fecha"
            Object.Width           =   1854
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Precio"
            Object.Width           =   1765
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Orig."
            Object.Width           =   972
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Cantidad"
            Object.Width           =   1676
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Dto1  "
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Dto2"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "Importe"
            Object.Width           =   1781
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Nombre"
            Object.Width           =   4586
         EndProperty
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Tahoma"
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
         Left            =   1320
         TabIndex        =   179
         Top             =   960
         Width           =   60
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Artículo: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   210
         Index           =   15
         Left            =   240
         TabIndex        =   178
         Top             =   960
         Width           =   840
      End
      Begin VB.Label Label9 
         Caption         =   "Compras efectuadas por el cliente cliente"
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
         Height          =   345
         Index           =   16
         Left            =   2880
         TabIndex        =   174
         Top             =   240
         Width           =   5415
      End
   End
   Begin VB.Frame FrameMarjalChipos 
      Height          =   3135
      Left            =   2040
      TabIndex        =   160
      Top             =   480
      Visible         =   0   'False
      Width           =   6015
      Begin VB.CheckBox chkCamposSocios 
         Caption         =   "Formato firma socio"
         Height          =   195
         Index           =   1
         Left            =   3120
         TabIndex        =   172
         Top             =   2160
         Width           =   2055
      End
      Begin VB.CheckBox chkCamposSocios 
         Caption         =   "Excluir campos baja"
         Height          =   195
         Index           =   0
         Left            =   480
         TabIndex        =   163
         Top             =   2160
         Width           =   2055
      End
      Begin VB.TextBox txtDescClie 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   7
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   170
         Top             =   1560
         Width           =   3495
      End
      Begin VB.TextBox txtCliente 
         Height          =   285
         Index           =   7
         Left            =   1320
         TabIndex        =   162
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox txtCliente 
         Height          =   285
         Index           =   6
         Left            =   1320
         TabIndex        =   161
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox txtDescClie 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   6
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   167
         Top             =   1200
         Width           =   3495
      End
      Begin VB.CommandButton cmdCamposSocio 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3600
         TabIndex        =   164
         Top             =   2520
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   15
         Left            =   4800
         TabIndex        =   165
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   8
         Left            =   600
         TabIndex        =   171
         Top             =   1560
         Width           =   420
      End
      Begin VB.Image imgCliente 
         Height          =   240
         Index           =   7
         Left            =   1080
         Picture         =   "frmListado5.frx":0000
         Top             =   1560
         Width           =   240
      End
      Begin VB.Image imgCliente 
         Height          =   240
         Index           =   6
         Left            =   1080
         Picture         =   "frmListado5.frx":0102
         Top             =   1200
         Width           =   240
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Socio/Cliente"
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
         Index           =   14
         Left            =   240
         TabIndex        =   169
         Top             =   840
         Width           =   1125
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   7
         Left            =   600
         TabIndex        =   168
         Top             =   1200
         Width           =   450
      End
      Begin VB.Label Label9 
         Caption         =   "Listado campos socios"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   15
         Left            =   1200
         TabIndex        =   166
         Top             =   360
         Width           =   3885
      End
   End
   Begin VB.Frame FrameContadoresAgua 
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   6015
      Begin VB.CommandButton cmdContadorAgua 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3720
         TabIndex        =   5
         Top             =   3120
         Width           =   975
      End
      Begin VB.OptionButton optVarios 
         Caption         =   "Cliente"
         Height          =   255
         Index           =   1
         Left            =   3120
         TabIndex        =   4
         Top             =   2520
         Width           =   1455
      End
      Begin VB.OptionButton optVarios 
         Caption         =   "Contador"
         Height          =   255
         Index           =   0
         Left            =   1440
         TabIndex        =   3
         Top             =   2520
         Width           =   1455
      End
      Begin VB.TextBox txtCliente 
         Height          =   285
         Index           =   1
         Left            =   1200
         TabIndex        =   2
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox txtDescClie 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   1
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1560
         Width           =   3495
      End
      Begin VB.TextBox txtDescClie 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   0
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1200
         Width           =   3495
      End
      Begin VB.TextBox txtCliente 
         Height          =   285
         Index           =   0
         Left            =   1200
         TabIndex        =   1
         Top             =   1200
         Width           =   855
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   0
         Left            =   4800
         TabIndex        =   6
         Top             =   3120
         Width           =   975
      End
      Begin VB.Image imgCliente 
         Height          =   240
         Index           =   1
         Left            =   960
         Picture         =   "frmListado5.frx":0204
         Top             =   1560
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   480
         TabIndex        =   12
         Top             =   1560
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   126
         Left            =   480
         TabIndex        =   10
         Top             =   1200
         Width           =   450
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
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
         Index           =   34
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   585
      End
      Begin VB.Image imgCliente 
         Height          =   240
         Index           =   0
         Left            =   960
         Picture         =   "frmListado5.frx":0306
         Top             =   1200
         Width           =   240
      End
      Begin VB.Label Label9 
         Caption         =   "Listado contadores agua"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   0
         Left            =   1080
         TabIndex        =   7
         Top             =   360
         Width           =   3510
      End
   End
   Begin VB.Frame FrameAlbaranesClientes 
      Height          =   6375
      Left            =   600
      TabIndex        =   155
      Top             =   360
      Visible         =   0   'False
      Width           =   7095
      Begin VB.CommandButton cmdSelecAlbaran 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4800
         TabIndex        =   159
         Top             =   5880
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   14
         Left            =   5880
         TabIndex        =   157
         Top             =   5880
         Width           =   975
      End
      Begin MSComctlLib.ListView lw 
         Height          =   5055
         Index           =   3
         Left            =   240
         TabIndex        =   156
         Top             =   720
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   8916
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Tipo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Codigo"
            Object.Width           =   1941
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Fecha"
            Object.Width           =   2011
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Importe"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Nº Factura"
            Object.Width           =   1764
         EndProperty
      End
      Begin VB.Label Label9 
         Caption         =   "Albaranes(facturas) cliente"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   14
         Left            =   360
         TabIndex        =   158
         Top             =   240
         Width           =   3885
      End
   End
   Begin VB.Frame FrameAlbaranesInternos 
      Height          =   3735
      Left            =   1200
      TabIndex        =   139
      Top             =   0
      Visible         =   0   'False
      Width           =   6015
      Begin VB.CommandButton cmdListadoAlbInt 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3720
         TabIndex        =   144
         Top             =   3120
         Width           =   975
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   9
         Left            =   4560
         TabIndex        =   143
         Top             =   2520
         Width           =   1215
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   8
         Left            =   1440
         TabIndex        =   142
         Top             =   2520
         Width           =   1215
      End
      Begin VB.TextBox txtCliente 
         Height          =   285
         Index           =   5
         Left            =   1320
         TabIndex        =   141
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox txtDescClie 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   5
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   150
         Top             =   1440
         Width           =   3495
      End
      Begin VB.TextBox txtDescClie 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   4
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   146
         Top             =   1080
         Width           =   3495
      End
      Begin VB.TextBox txtCliente 
         Height          =   285
         Index           =   4
         Left            =   1320
         TabIndex        =   140
         Top             =   1080
         Width           =   855
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   13
         Left            =   4800
         TabIndex        =   145
         Top             =   3120
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   13
         Left            =   3600
         TabIndex        =   154
         Top             =   2520
         Width           =   465
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   9
         Left            =   4200
         Picture         =   "frmListado5.frx":0408
         Top             =   2520
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Desde"
         Height          =   195
         Index           =   12
         Left            =   600
         TabIndex        =   153
         Top             =   2520
         Width           =   465
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
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
         Index           =   13
         Left            =   240
         TabIndex        =   152
         Top             =   2160
         Width           =   495
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   8
         Left            =   1080
         Picture         =   "frmListado5.frx":0493
         Top             =   2520
         Width           =   240
      End
      Begin VB.Image imgCliente 
         Height          =   240
         Index           =   5
         Left            =   1080
         Picture         =   "frmListado5.frx":051E
         Top             =   1440
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   6
         Left            =   600
         TabIndex        =   151
         Top             =   1440
         Width           =   420
      End
      Begin VB.Label Label9 
         Caption         =   "Listado albaranes internos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   13
         Left            =   1200
         TabIndex        =   149
         Top             =   240
         Width           =   3795
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   5
         Left            =   600
         TabIndex        =   148
         Top             =   1080
         Width           =   450
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
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
         Index           =   12
         Left            =   240
         TabIndex        =   147
         Top             =   720
         Width           =   585
      End
      Begin VB.Image imgCliente 
         Height          =   240
         Index           =   4
         Left            =   1080
         Picture         =   "frmListado5.frx":0620
         Top             =   1080
         Width           =   240
      End
   End
   Begin VB.Frame FrameFitoCampos 
      Height          =   4695
      Left            =   4080
      TabIndex        =   115
      Top             =   0
      Visible         =   0   'False
      Width           =   6135
      Begin VB.OptionButton optFitoCampos 
         Caption         =   "Cliente - campos"
         Height          =   255
         Index           =   1
         Left            =   3240
         TabIndex        =   124
         Top             =   3600
         Width           =   2175
      End
      Begin VB.OptionButton optFitoCampos 
         Caption         =   "Campos - Cliente"
         Height          =   255
         Index           =   0
         Left            =   960
         TabIndex        =   123
         Top             =   3600
         Value           =   -1  'True
         Width           =   2175
      End
      Begin VB.CommandButton cmdFitoCampos 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3600
         TabIndex        =   125
         Top             =   4080
         Width           =   975
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   7
         Left            =   4560
         TabIndex        =   119
         Top             =   2280
         Width           =   1215
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   6
         Left            =   1560
         TabIndex        =   118
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Frame FrameCodtipom 
         Height          =   615
         Left            =   360
         TabIndex        =   133
         Top             =   2760
         Width           =   5535
         Begin VB.CheckBox chkCodtipom 
            Caption         =   "Servicios"
            Height          =   195
            Index           =   2
            Left            =   4200
            TabIndex        =   122
            Tag             =   "FAS"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CheckBox chkCodtipom 
            Caption         =   "Internas"
            Height          =   195
            Index           =   1
            Left            =   2640
            TabIndex        =   121
            Tag             =   "FAI"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CheckBox chkCodtipom 
            Caption         =   "Ventas"
            Height          =   195
            Index           =   0
            Left            =   1080
            TabIndex        =   120
            Tag             =   "FAV"
            Top             =   240
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.Label lblDpto 
            AutoSize        =   -1  'True
            Caption         =   "Facturas"
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
            Index           =   11
            Left            =   0
            TabIndex        =   138
            Top             =   0
            Width           =   735
         End
      End
      Begin VB.TextBox txtDescClie 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   3
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   131
         Top             =   1440
         Width           =   3495
      End
      Begin VB.TextBox txtCliente 
         Height          =   285
         Index           =   3
         Left            =   1440
         TabIndex        =   117
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox txtCliente 
         Height          =   285
         Index           =   2
         Left            =   1440
         TabIndex        =   116
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox txtDescClie 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   2
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   127
         Top             =   960
         Width           =   3495
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   12
         Left            =   4800
         TabIndex        =   126
         Top             =   4080
         Width           =   975
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   7
         Left            =   4200
         Picture         =   "frmListado5.frx":0722
         Top             =   2280
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   11
         Left            =   3600
         TabIndex        =   137
         Top             =   2280
         Width           =   465
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   6
         Left            =   1200
         Picture         =   "frmListado5.frx":07AD
         Top             =   2280
         Width           =   240
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
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
         Index           =   10
         Left            =   360
         TabIndex        =   136
         Top             =   1920
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Desde"
         Height          =   195
         Index           =   10
         Left            =   720
         TabIndex        =   135
         Top             =   2280
         Width           =   465
      End
      Begin VB.Label lblIndicador 
         Caption         =   "I"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   134
         Top             =   4080
         Width           =   2895
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   720
         TabIndex        =   132
         Top             =   1440
         Width           =   420
      End
      Begin VB.Image imgCliente 
         Height          =   240
         Index           =   3
         Left            =   1200
         Picture         =   "frmListado5.frx":0838
         Top             =   1440
         Width           =   240
      End
      Begin VB.Label Label9 
         Caption         =   "Fitosanitarios x Campos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   12
         Left            =   1320
         TabIndex        =   130
         Top             =   240
         Width           =   3510
      End
      Begin VB.Image imgCliente 
         Height          =   240
         Index           =   2
         Left            =   1200
         Picture         =   "frmListado5.frx":093A
         Top             =   960
         Width           =   240
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
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
         Index           =   9
         Left            =   360
         TabIndex        =   129
         Top             =   600
         Width           =   585
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   720
         TabIndex        =   128
         Top             =   960
         Width           =   450
      End
   End
   Begin VB.Frame FrameSelProveedores 
      Height          =   6615
      Left            =   6480
      TabIndex        =   104
      Top             =   360
      Visible         =   0   'False
      Width           =   6495
      Begin VB.CommandButton cmdSelProvee 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4080
         TabIndex        =   108
         Top             =   6120
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   10
         Left            =   5160
         TabIndex        =   107
         Top             =   6120
         Width           =   975
      End
      Begin MSComctlLib.ListView lw 
         Height          =   5535
         Index           =   1
         Left            =   240
         TabIndex        =   106
         Top             =   480
         Width           =   5895
         _ExtentX        =   10398
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Codigo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripcion"
            Object.Width           =   6068
         EndProperty
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   3
         Left            =   600
         Picture         =   "frmListado5.frx":0A3C
         Top             =   6240
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   2
         Left            =   240
         Picture         =   "frmListado5.frx":0B86
         Top             =   6240
         Width           =   240
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "Seleccionar proveedores"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   10
         Left            =   1440
         TabIndex        =   105
         Top             =   120
         Width           =   3855
      End
   End
   Begin VB.Frame FrameComprasTratamientos 
      Height          =   2535
      Left            =   0
      TabIndex        =   93
      Top             =   4560
      Visible         =   0   'False
      Width           =   5655
      Begin VB.CheckBox chkVarios 
         Caption         =   "Detalla artículos"
         Height          =   255
         Index           =   2
         Left            =   1440
         TabIndex        =   97
         Top             =   1560
         Width           =   2415
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   5
         Left            =   4080
         TabIndex        =   96
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton cmdComprasTratamientos 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3120
         TabIndex        =   98
         Top             =   1920
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   9
         Left            =   4320
         TabIndex        =   99
         Top             =   1920
         Width           =   975
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   4
         Left            =   1440
         TabIndex        =   95
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label lblIndicador 
         Caption         =   "I"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   103
         Top             =   2040
         Width           =   2895
      End
      Begin VB.Label Label1 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   9
         Left            =   3120
         TabIndex        =   102
         Top             =   1080
         Width           =   465
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   5
         Left            =   3720
         Picture         =   "frmListado5.frx":0CD0
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Desde"
         Height          =   195
         Index           =   8
         Left            =   360
         TabIndex        =   101
         Top             =   1080
         Width           =   465
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
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
         Index           =   7
         Left            =   240
         TabIndex        =   100
         Top             =   720
         Width           =   495
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   4
         Left            =   960
         Picture         =   "frmListado5.frx":0D5B
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "Ajuste compras tratamientos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   9
         Left            =   720
         TabIndex        =   94
         Top             =   240
         Width           =   4275
      End
   End
   Begin VB.Frame FrameGessocial 
      Height          =   3855
      Left            =   120
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   6375
      Begin VB.Frame FrameFechaBaja 
         Caption         =   "Frame1"
         Height          =   735
         Left            =   960
         TabIndex        =   53
         Top             =   2040
         Visible         =   0   'False
         Width           =   4215
         Begin VB.TextBox txtFecha 
            Height          =   285
            Index           =   3
            Left            =   1680
            TabIndex        =   54
            Top             =   240
            Width           =   1215
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   3
            Left            =   1080
            Picture         =   "frmListado5.frx":0DE6
            Top             =   240
            Width           =   240
         End
         Begin VB.Label lblDpto 
            AutoSize        =   -1  'True
            Caption         =   "Fecha"
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
            Index           =   3
            Left            =   480
            TabIndex        =   55
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.OptionButton optVarios 
         Caption         =   "Baja"
         Height          =   255
         Index           =   4
         Left            =   2640
         TabIndex        =   52
         Top             =   3360
         Width           =   1215
      End
      Begin VB.Frame FrameGasol 
         Height          =   1215
         Left            =   120
         TabIndex        =   19
         Top             =   1200
         Visible         =   0   'False
         Width           =   6135
         Begin VB.ComboBox cboEntidades 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   600
            Width           =   5775
         End
         Begin VB.Label Label1 
            Caption         =   "Colectivo"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   21
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.OptionButton optVarios 
         Caption         =   "Actualizar"
         Height          =   255
         Index           =   3
         Left            =   1320
         TabIndex        =   18
         Top             =   3360
         Width           =   1335
      End
      Begin VB.OptionButton optVarios 
         Caption         =   "Crear"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   17
         Top             =   3360
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   1
         Left            =   5160
         TabIndex        =   15
         Top             =   3240
         Width           =   975
      End
      Begin VB.CommandButton cmdGessocial 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4080
         TabIndex        =   14
         Top             =   3240
         Width           =   975
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   1
         Left            =   480
         TabIndex        =   16
         Top             =   240
         Width           =   5145
      End
   End
   Begin VB.Frame FrameSeleccionarFamilia 
      Height          =   5415
      Left            =   240
      TabIndex        =   77
      Top             =   240
      Visible         =   0   'False
      Width           =   6615
      Begin VB.CommandButton cmdSeleccionarFamilia 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3840
         TabIndex        =   81
         Top             =   4920
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   7
         Left            =   5040
         TabIndex        =   78
         Top             =   4920
         Width           =   975
      End
      Begin MSComctlLib.ListView lw 
         Height          =   4095
         Index           =   0
         Left            =   240
         TabIndex        =   80
         Top             =   720
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   7223
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Codigo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripcion"
            Object.Width           =   6068
         EndProperty
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   0
         Left            =   240
         Picture         =   "frmListado5.frx":0E71
         Top             =   4920
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   1
         Left            =   600
         Picture         =   "frmListado5.frx":0FBB
         Top             =   4920
         Width           =   240
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "Seleccionar familias"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   7
         Left            =   960
         TabIndex        =   79
         Top             =   240
         Width           =   3855
      End
   End
   Begin VB.Frame FrameCambioProveedorPedido 
      Height          =   7215
      Left            =   0
      TabIndex        =   109
      Top             =   0
      Visible         =   0   'False
      Width           =   9615
      Begin VB.CommandButton cmdCambiarProvePedido 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   7200
         TabIndex        =   114
         Top             =   6720
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   11
         Left            =   8280
         TabIndex        =   111
         Top             =   6720
         Width           =   975
      End
      Begin MSComctlLib.ListView lw 
         Height          =   5775
         Index           =   2
         Left            =   120
         TabIndex        =   112
         Top             =   840
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   10186
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
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Codigo"
            Object.Width           =   3775
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripcion"
            Object.Width           =   7832
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Precio"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Dto1"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Dto2"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Importe"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Observa"
            Object.Width           =   4304
         EndProperty
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Líneas a cambiar"
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
         Index           =   8
         Left            =   120
         TabIndex        =   113
         Top             =   600
         Width           =   1425
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   5
         Left            =   480
         Picture         =   "frmListado5.frx":1105
         Top             =   6840
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   4
         Left            =   120
         Picture         =   "frmListado5.frx":124F
         Top             =   6840
         Width           =   240
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Cambiar proveedor en pedido"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   11
         Left            =   4560
         TabIndex        =   110
         Top             =   240
         Width           =   4815
      End
   End
   Begin VB.Frame FrameGesocialCambioSituacion 
      Height          =   5895
      Left            =   120
      TabIndex        =   58
      Top             =   0
      Visible         =   0   'False
      Width           =   5175
      Begin VB.TextBox txtModificable 
         Height          =   1575
         Index           =   0
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   59
         Text            =   "frmListado5.frx":1399
         Top             =   3480
         Width           =   4695
      End
      Begin VB.TextBox txtNoModificable 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   0
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   63
         Top             =   960
         Width           =   4695
      End
      Begin VB.TextBox txtNoModificable 
         BackColor       =   &H80000018&
         Height          =   1365
         Index           =   1
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   62
         Top             =   1680
         Width           =   4695
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Guardar"
         Height          =   375
         Index           =   5
         Left            =   3960
         TabIndex        =   60
         Top             =   5400
         Width           =   975
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Observaciones"
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
         Index           =   6
         Left            =   240
         TabIndex        =   66
         Top             =   3240
         Width           =   1245
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Situación"
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
         Index           =   5
         Left            =   240
         TabIndex        =   65
         Top             =   1440
         Width           =   780
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Asociado"
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
         Index           =   4
         Left            =   240
         TabIndex        =   64
         Top             =   720
         Width           =   765
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "Cambio situación asociado"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   5
         Left            =   600
         TabIndex        =   61
         Top             =   240
         Width           =   3855
      End
   End
   Begin VB.Frame FrameAguaMod 
      Height          =   3735
      Left            =   1680
      TabIndex        =   82
      Top             =   0
      Visible         =   0   'False
      Width           =   4815
      Begin VB.TextBox txtNumero 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   3720
         TabIndex        =   83
         Text            =   "Text1"
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox txtNoModificable 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   4
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   89
         Top             =   1560
         Width           =   2655
      End
      Begin VB.TextBox txtNoModificable 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   3
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   87
         Top             =   1080
         Width           =   2655
      End
      Begin VB.CommandButton cmdCambiarConsumo 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   2400
         TabIndex        =   84
         Top             =   3000
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   8
         Left            =   3600
         TabIndex        =   85
         Top             =   3000
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Consumo"
         Height          =   255
         Index           =   7
         Left            =   3000
         TabIndex        =   92
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Lectura modificada"
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
         Index           =   6
         Left            =   1560
         TabIndex        =   91
         Top             =   2040
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Lectura actual"
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   90
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Lectura anterior"
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   88
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "Modificar lectura facturada"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   8
         Left            =   600
         TabIndex        =   86
         Top             =   360
         Width           =   3855
      End
   End
   Begin VB.Frame FrameSubRPT 
      Height          =   3015
      Left            =   1440
      TabIndex        =   67
      Top             =   120
      Visible         =   0   'False
      Width           =   6255
      Begin VB.CommandButton cmdSubRPT 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3840
         TabIndex        =   70
         Top             =   2400
         Width           =   975
      End
      Begin VB.TextBox txtModificable 
         Height          =   285
         Index           =   2
         Left            =   1320
         MaxLength       =   30
         ScrollBars      =   2  'Vertical
         TabIndex        =   69
         Text            =   "Text1"
         Top             =   2000
         Width           =   2055
      End
      Begin VB.TextBox txtModificable 
         Height          =   285
         Index           =   1
         Left            =   1320
         MaxLength       =   100
         ScrollBars      =   2  'Vertical
         TabIndex        =   68
         Text            =   "Text1"
         Top             =   1440
         Width           =   4695
      End
      Begin VB.TextBox txtNoModificable 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   2
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   73
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   6
         Left            =   5040
         TabIndex        =   71
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Informe"
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   76
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Descripcion"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   75
         Top             =   1500
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Linea"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   74
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "Informe asociado"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   6
         Left            =   1200
         TabIndex        =   72
         Top             =   240
         Width           =   3855
      End
   End
   Begin VB.Frame FrameActualizaSdtoFm 
      Height          =   3495
      Left            =   0
      TabIndex        =   40
      Top             =   1800
      Width           =   6375
      Begin VB.TextBox txtDescFamia 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   0
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   56
         Top             =   1920
         Width           =   3495
      End
      Begin VB.TextBox txtFamia 
         Height          =   285
         Index           =   0
         Left            =   1440
         TabIndex        =   43
         Top             =   1920
         Width           =   1095
      End
      Begin VB.CommandButton cmdSdtofmInsert 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3480
         TabIndex        =   45
         Top             =   2880
         Width           =   975
      End
      Begin VB.CheckBox chkVarios 
         Caption         =   "Insertar sólo los nuevos"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   44
         Top             =   2640
         Width           =   2415
      End
      Begin VB.TextBox txtActiv 
         Height          =   285
         Index           =   1
         Left            =   1440
         TabIndex        =   42
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox txtDescActiv 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   1
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   48
         Top             =   1440
         Width           =   3375
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   2
         Left            =   1440
         TabIndex        =   41
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   4
         Left            =   4680
         TabIndex        =   46
         Top             =   2880
         Width           =   975
      End
      Begin VB.Image imgFamilia 
         Height          =   240
         Index           =   0
         Left            =   1080
         Picture         =   "frmListado5.frx":139F
         Top             =   1920
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
         Index           =   18
         Left            =   240
         TabIndex        =   57
         Top             =   1920
         Width           =   600
      End
      Begin VB.Label lblIndicador 
         Caption         =   "Label2"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   51
         Top             =   3000
         Width           =   2895
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
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
         Index           =   2
         Left            =   240
         TabIndex        =   50
         Top             =   840
         Width           =   495
      End
      Begin VB.Image imgActividad 
         Height          =   240
         Index           =   1
         Left            =   1080
         Picture         =   "frmListado5.frx":14A1
         Top             =   1440
         Width           =   240
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Actividad"
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
         Index           =   1
         Left            =   240
         TabIndex        =   49
         Top             =   1440
         Width           =   795
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   2
         Left            =   840
         Picture         =   "frmListado5.frx":15A3
         Top             =   840
         Width           =   240
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "Actualizar descuentos familia/marca"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   4
         Left            =   240
         TabIndex        =   47
         Top             =   240
         Width           =   5865
      End
   End
   Begin VB.Frame FrameDtoAsginar 
      Height          =   3135
      Left            =   120
      TabIndex        =   30
      Top             =   0
      Visible         =   0   'False
      Width           =   6135
      Begin VB.CommandButton cmdCargaDtoFamiliaActiv 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3840
         TabIndex        =   34
         Top             =   2520
         Width           =   975
      End
      Begin VB.CheckBox chkVarios 
         Caption         =   "Insertar sólo los nuevos"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   33
         Top             =   2280
         Width           =   2415
      End
      Begin VB.ComboBox cboTipoDescuento 
         Height          =   315
         Index           =   0
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   1560
         Width           =   3615
      End
      Begin VB.TextBox txtDescActiv 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   0
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   960
         Width           =   3375
      End
      Begin VB.TextBox txtActiv 
         Height          =   285
         Index           =   0
         Left            =   1680
         TabIndex        =   31
         Top             =   960
         Width           =   735
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   3
         Left            =   4920
         TabIndex        =   35
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Tipo descuento"
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
         Index           =   0
         Left            =   240
         TabIndex        =   39
         Top             =   1620
         Width           =   1290
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Actividad"
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
         Index           =   39
         Left            =   240
         TabIndex        =   38
         Top             =   960
         Width           =   795
      End
      Begin VB.Image imgActividad 
         Height          =   240
         Index           =   0
         Left            =   1320
         Picture         =   "frmListado5.frx":162E
         Top             =   960
         Width           =   240
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "Actualizar desde familias"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   3
         Left            =   480
         TabIndex        =   36
         Top             =   240
         Width           =   5145
      End
   End
   Begin VB.Frame FrameEliminarPresupuestos 
      Height          =   2775
      Left            =   120
      TabIndex        =   26
      Top             =   0
      Visible         =   0   'False
      Width           =   5175
      Begin VB.CommandButton cmdElimPresu 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   2640
         TabIndex        =   24
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   1
         Left            =   3480
         TabIndex        =   23
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   0
         Left            =   1080
         TabIndex        =   22
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   2
         Left            =   3720
         TabIndex        =   25
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   2640
         TabIndex        =   29
         Top             =   1200
         Width           =   450
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   1
         Left            =   3240
         Picture         =   "frmListado5.frx":1730
         Top             =   1200
         Width           =   240
      End
      Begin VB.Label Label9 
         Caption         =   "Eliminar presupuestos FAZ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   2
         Left            =   720
         TabIndex        =   28
         Top             =   360
         Width           =   3825
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   27
         Top             =   1200
         Width           =   450
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   840
         Picture         =   "frmListado5.frx":17BB
         Top             =   1200
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmListado5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public OpcionListado As Integer

Public OtrosDatos As String

    '==== Opciones ====
    '=============================
    '   0.-  Listados contaqdores de agua
    '   YA NO ESTA 1.-  GESSOCIAL    Crear asociado en la seccion    'NOVIEMBRE 2014. Lo quito de aqui
    '   2.-  Eliminar presupuestos "HERBELCA" Pide las fechas para la siguiente pantalla
    
    '   3.- Asignar descuentos  actividad desde sfamia
    '   4.- Insertar en sdtofm desde sactivdto
        
    '   5.- Gessocial. Cambio situacion socio(alta-baja-situacion o essocio)
    
    '   6.- Subreports de la SCRYST
    
    '   7.- Seleccion de familias
    '        Vendran las familias en otrosdatos y cadenadesdeotrofrm mostrara las que quiera
    
    '   8.- Modificar consumo en facturas de agua
    
    '   9.- Alzira. Ajuste compras tratamientos
    
    '   10.- Seleccion PROVEEDORES.  A partir de un select.....
        
    '   11.- Cambiar proveedor del pedido despues de una simulacion
      
    '   12.- Fito santiarios por campos
    '   13.- Listado albaranes internos (perosnalizable)
    
    
    '   14.- EULER. Devolverá un ALBARAN
    
    '   15.- Marjal-Chipos.  Informe campos socio
    
    '   16.- Listado compras cliente desde una fecha
    '   17.- devoluciones. Preguta si pasa a albaran o a frt
    
    '   18.-  Telefonia,  Impresion contrato.  Importet terminal y meses
    
    
Private WithEvents frmCli As frmFacClientes3
Attribute frmCli.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmMtoFamilia As frmAlmFamiliaArticulo
Attribute frmMtoFamilia.VB_VarHelpID = -1

Dim miSQL As String
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private CadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para el frmImprimir
Private cadNomRPT As String 'Nombre del informe a Imprimir
Private conSubRPT As Boolean 'Si el informe tiene subreports
Private cadPDFrpt As String 'Nombre del informe a enviar por email
Private vMostrarTree As Boolean

Private PrimVez As Boolean

Private auxiliar As String  ' Para quitar proveedores serviara para guardar cuales quito


Private Sub cboTipoDescuento_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub chkCamposSocios_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub chkCodtipom_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub chkVarios_KeyPress(Index As Integer, KeyAscii As Integer)
     KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub cmdCambiarConsumo_Click()
    If txtNumero(0).Text = "" Then Exit Sub
    
    miSQL = RecuperaValor(OtrosDatos, 3)
    If Val(miSQL) <> Val(Label1(7).Tag) Then
        
        If MsgBox("¿Seguro que desea cambiar el consumo?", vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
        Screen.MousePointer = vbHourglass
        conn.BeginTrans
        If Not HacerUpdateConsumo Then
            conn.RollbackTrans
        Else
            conn.CommitTrans
            CadenaDesdeOtroForm = "OK"
        End If
        Screen.MousePointer = vbDefault
    End If
    Unload Me
End Sub

Private Sub cmdCambiarProvePedido_Click()
    miSQL = ""
    For NumRegElim = 1 To lw(2).ListItems.Count
        If lw(2).ListItems(NumRegElim).Checked Then miSQL = miSQL & ", " & lw(2).ListItems(NumRegElim).Tag
    Next
    If miSQL = "" Then
        MsgBox "Seleccione alguna linea de articulo para cambiar de proveedor", vbExclamation
    Else
        CadenaDesdeOtroForm = miSQL
        Unload Me
    End If
End Sub

Private Sub cmdCamposSocio_Click()
    'a
    
    InicializarVbles True

    If Me.chkCamposSocios(0).Value = 1 Then
        miSQL = "sclienhuertos.fecbajas"
        cadSelect = "  (" & miSQL & ") is null"
        cadFormula = " isnull({" & miSQL & "})"
    End If


    If txtCliente(6).Text <> "" Or txtCliente(7).Text <> "" Then
        miSQL = " Cliente: "
        cadTitulo = "{sclienhuertos.codclien}"
        If Not PonerDesdeHasta(cadTitulo, "CLI", 6, 7, miSQL) Then Exit Sub
        If cadPDFrpt <> "" Then cadPDFrpt = cadPDFrpt & "              "
        cadPDFrpt = cadPDFrpt & miSQL
    End If
    If Me.chkCamposSocios(0).Value = 1 Then cadPDFrpt = Trim(cadPDFrpt & "     " & chkCamposSocios(0).Caption)
        
    CadParam = CadParam & "DesdeHasta=""" & Trim(cadPDFrpt) & """|"
    numParam = numParam + 1
    
    Screen.MousePointer = vbHourglass
    If Not HayRegParaInforme("sclienhuertos", cadSelect, True) Then
        MsgBox "No hay datos para mostrar con estos valores", vbExclamation

    Else
    
        cadTitulo = "Listado campos socios"
        
        If chkCamposSocios(1).Value = 1 Then
            cadNomRPT = "marListadoCli.rpt"
        Else
            cadNomRPT = "marListado.rpt"
        End If
        cadPDFrpt = ""
        conSubRPT = False
        
        LlamarImprimir

    End If
    Screen.MousePointer = vbDefault
        
End Sub

Private Sub cmdCancelar_Click(Index As Integer)
    If Index = 5 Then
        'GESSOCIAL
        'Cambio de situacion
        'Pone SALIR.
        CadenaDesdeOtroForm = Trim(txtModificable(0).Text)
    
    ElseIf Index = 7 Then
        CadenaDesdeOtroForm = ""
    ElseIf Index = 11 Then
        CadenaDesdeOtroForm = ""
    ElseIf Index = 17 Then
        CadenaDesdeOtroForm = ""
    ElseIf Index = 18 Then
        CadenaDesdeOtroForm = ""
    End If
    Unload Me
    
    
End Sub

Private Sub cmdCargaDtoFamiliaActiv_Click()
    miSQL = ""
    
    If Me.txtActiv(0).Text = "" Then miSQL = miSQL & "-Actividad"
    If Me.cboTipoDescuento(0).ListIndex < 0 Then miSQL = miSQL & "-Tipo descuento"
    If miSQL <> "" Then
        MsgBox "Faltan campos: " & vbCrLf & miSQL, vbExclamation
        Exit Sub
    End If
    
    If MsgBox("¿continuar con el proceso de generacion de descuentos por actividad?", vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub

    
    'SI SOLO SON NUEVOS.
    ' INSERT IGNORE INTO
    ' Si son nuevos y actualizar, replace
    
    If Me.chkVarios(0).Value = 1 Then
        miSQL = " INSERT IGNORE "
    Else
        miSQL = "REPLACE "
    End If
    miSQL = miSQL & " INTO sactivdtos SELECT " & Me.txtActiv(0).Text & " codactiv,codfamia,clasifica "
    miSQL = miSQL & " FROM sfamiadtos where clasifica=" & cboTipoDescuento(0).ItemData(cboTipoDescuento(0).ListIndex)
    
    
    If ejecutar(miSQL, False) Then
        
        CadenaDesdeOtroForm = txtActiv(0).Text
        Unload Me
    End If
End Sub

Private Sub cmdComprasTratamientos_Click()
    
    numParam = 0
    If Me.txtFecha(5).Text = "" Then numParam = 5
    If Me.txtFecha(4).Text = "" Then numParam = 4
    If numParam > 0 Then
        MsgBox "Campos fecha son obligatorios", vbExclamation
        PonerFoco txtFecha(CInt(numParam))
        Exit Sub
    End If
    
    
    Me.lblIndicador(1).Caption = "Comienzo proceso"
    Screen.MousePointer = vbHourglass
    InicializarVbles True
    
    
    If GeneraDatosComprasTratamientos Then
        HaPulsadoElBotonDeImprimir = False
        cadTitulo = "Ajuste compras tratamientos"
        CadParam = CadParam & "pdh1=""Fechas: " & txtFecha(4).Text & " - " & txtFecha(5).Text & """|"
        numParam = numParam + 1
        
        CadParam = CadParam & "Detalle=" & Abs(Me.chkVarios(2).Value) & "|"
        numParam = numParam + 1
        vMostrarTree = True
        cadNomRPT = "rAjuCompTra.rpt"   'cadPDFrpt & ".rpt"
        cadPDFrpt = ""
        conSubRPT = False
        
        LlamarImprimir
        
        If HaPulsadoElBotonDeImprimir Then
            miSQL = "Va a generar el apunte. Continuar?"
            If MsgBox(miSQL, vbQuestion + vbYesNoCancel) = vbYes Then
                GenerarApunteAjusteTratamientos
                Unload Me
            End If
            
       End If
    End If
    
    Me.lblIndicador(1).Caption = ""
    Screen.MousePointer = vbDefault
    
        
End Sub

Private Sub cmdContadorAgua_Click()

    InicializarVbles True

    If txtCliente(0).Text <> "" Or txtCliente(1).Text <> "" Then
        miSQL = " Cliente: "
        cadTitulo = "{aguacontadores.codclien}"
        If Not PonerDesdeHasta(cadTitulo, "CLI", 0, 1, miSQL) Then Exit Sub
        If cadPDFrpt <> "" Then cadPDFrpt = cadPDFrpt & "              "
        cadPDFrpt = cadPDFrpt & miSQL
        
        CadParam = CadParam & "DesdeHasta=""" & cadPDFrpt & """|"
        numParam = numParam + 1
    End If
    
    
    Screen.MousePointer = vbHourglass
    If Not HayRegParaInforme("aguacontadores", cadSelect, True) Then
        MsgBox "No hay datos para mostrar con estos valores", vbExclamation

    Else
    
        cadTitulo = "Listado contadores agua"
        If optVarios(0).Value Then
            cadPDFrpt = "rAgua1"
        Else
            cadPDFrpt = "rAgua2"
            vMostrarTree = True
        End If
        cadNomRPT = cadPDFrpt & ".rpt"
        cadPDFrpt = ""
        conSubRPT = False
        
        LlamarImprimir

    End If
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub cmdContratoTelef_Click()

    CadenaDesdeOtroForm = Trim(txtNumero(1).Text) & "|" & Trim(txtNumero(2).Text) & "|"
    Unload Me
End Sub

Private Sub cmdDevolucion_Click()
    CadenaDesdeOtroForm = IIf(Me.optDevol(0).Value, "ALV", "ART")
    Unload Me
End Sub

Private Sub cmdElimPresu_Click()
    miSQL = "1=1"
    If Me.txtFecha(0).Text <> "" Then miSQL = miSQL & " AND scafac.fecfactu >=" & DBSet(Me.txtFecha(0).Text, "F")
    If Me.txtFecha(1).Text <> "" Then miSQL = miSQL & " AND scafac.fecfactu <=" & DBSet(Me.txtFecha(1).Text, "F")
    
    cadFormula = DevuelveDesdeBD(conAri, "count(*)", "scafac", miSQL & " AND codtipom ", "FAZ", "T")
    If cadFormula = "" Then cadFormula = "0"
    If Val(cadFormula) = 0 Then
        MsgBox "Ningun dato a eliminar", vbExclamation
    Else
        CadenaDesdeOtroForm = miSQL
        frmVarios3.Opcion = 5
        frmVarios3.Show vbModal
        Unload Me
        
    End If
End Sub



Private Sub cmdFitoCampos_Click()

        
    
    
    InicializarVbles True
    Screen.MousePointer = vbHourglass
    If GenerarFitoCampos Then
        vMostrarTree = True
        cadTitulo = "Campos- Fitosantiarios"
        
        cadNomRPT = "rCamposFito"
        If optFitoCampos(1).Value Then
            cadNomRPT = cadNomRPT & "cli"
            cadTitulo = "Cliente - " & cadTitulo
        End If
        cadNomRPT = cadNomRPT & ".rpt"
        cadPDFrpt = cadNomRPT
        conSubRPT = False
        cadFormula = "{tmpinformes.codusu} = " & vUsu.codigo
        LlamarImprimir
    
    End If
    Me.lblIndicador(2).Caption = ""
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdListadoAlbInt_Click()
    InicializarVbles True
    
    cadNomRPT = DevuelveDesdeBDNew(conAri, "scryst", "documrpt", "codcryst", "82", "N")
    
    cadSelect = "(scaalb.codtipom)='ALI'"
    cadFormula = "{scaalb.codtipom}='ALI'"
    
    If txtFecha(8).Text <> "" Or txtFecha(9).Text <> "" Then
        miSQL = " Fecha: "
        If Not PonerDesdeHasta("{scaalb.fechaalb}", "F", 8, 9, miSQL) Then Exit Sub
        cadPDFrpt = cadPDFrpt & miSQL
    End If

    If txtCliente(4).Text <> "" Or txtCliente(5).Text <> "" Then
        miSQL = " Cliente: "
        If Not PonerDesdeHasta("{scaalb.codclien}", "CLI", 4, 5, miSQL) Then Exit Sub
        If cadPDFrpt <> "" Then cadPDFrpt = cadPDFrpt & "      "
        cadPDFrpt = cadPDFrpt & miSQL
        

    End If
    

    CadParam = CadParam & "dh=""" & cadPDFrpt & """|"
    numParam = numParam + 1
    
    
    Screen.MousePointer = vbHourglass
    If HayRegParaInforme("scaalb", cadSelect, False) Then
        
    
        cadTitulo = "Listado albaranes internos"
        
        cadPDFrpt = ""
        conSubRPT = False
        
        LlamarImprimir

    End If
    Screen.MousePointer = vbDefault

End Sub

Private Sub cmdSdtofmInsert_Click()
    If Me.txtFecha(2).Text = "" Then Exit Sub
    
    If Me.txtActiv(1).Text = "" And txtFamia(0).Text = "" Then Exit Sub
    
    miSQL = "Va a actualizar los descuentos familia / marca por actividad. " & vbCrLf
    If Me.txtActiv(1).Text <> "" Then miSQL = miSQL & " ACTIVIDAD: " & txtActiv(1).Text & "-" & Me.txtDescActiv(1).Text & vbCrLf
    If txtFamia(0).Text <> "" Then miSQL = miSQL & " FAMILIA: " & txtFamia(0).Text & "-" & Me.txtDescFamia(0).Text & vbCrLf
    
    miSQL = miSQL & vbCrLf & "¿Desea continuar con el proceso?"
    If MsgBox(miSQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    
    'OK, vamos p'alla
    Screen.MousePointer = vbHourglass
    lblIndicador(0).Caption = "Preparando BD"
    lblIndicador(0).Refresh
    conn.Execute "DELETE FROM tmpinformes where codusu =" & vUsu.codigo
    
    
    'Cargamos todos los posibles descuentos
    lblIndicador(0).Caption = "Leyendo familia actividad"
    lblIndicador(0).Refresh
    
    miSQL = "insert into tmpinformes(codusu,codigo1,campo1,importe1,fecha1)"
    miSQL = miSQL & " SELECT " & vUsu.codigo & ", sactivdtos.codactiv,sactivdtos.codfamia,dtoline1"
    miSQL = miSQL & "," & DBSet(txtFecha(2).Text, "F")
    miSQL = miSQL & " From sactivdtos, sfamiadtos WHERE  sfamiadtos.codfamia=sactivdtos.codfamia AND"
    miSQL = miSQL & " sfamiadtos.clasifica=sactivdtos.clasifica "
    If Me.txtActiv(1).Text <> "" Then miSQL = miSQL & " AND sactivdtos.codactiv = " & Me.txtActiv(1).Text
    If txtFamia(0).Text <> "" Then miSQL = miSQL & " AND sfamiadtos.codfamia = " & txtFamia(0).Text
    conn.Execute miSQL
    
    'Si ha puesto solo los nuevos veo cuales tengo que borrar de la temporal
    If Me.chkVarios(1).Value = 1 Then
        Set miRsAux = New ADODB.Recordset
        lblIndicador(0).Caption = "Comprobando valores existente"
        lblIndicador(0).Refresh
        miSQL = "Select codfamia from sdtofm where  "
        miSQL = miSQL & " codclien IS NULL AND codmarca is null "
        If Me.txtActiv(1).Text <> "" Then miSQL = miSQL & " AND codactiv =" & Me.txtActiv(1).Text
        If txtFamia(0).Text <> "" Then miSQL = miSQL & " AND codfamia = " & txtFamia(0).Text
        
        miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        miSQL = ""
        While Not miRsAux.EOF
            lblIndicador(0).Caption = "Fam: " & miRsAux!Codfamia
            lblIndicador(0).Refresh
            miSQL = miSQL & ", " & miRsAux!Codfamia
            If Len(miSQL) > 400 Then
                miSQL = Mid(miSQL, 2)
                miSQL = vUsu.codigo & " AND campo1 IN (" & miSQL & ")"
                miSQL = "DELETE FROM tmpinformes WHERE codusu = " & miSQL
                conn.Execute miSQL
                miSQL = ""
            End If
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        If miSQL <> "" Then
            miSQL = Mid(miSQL, 2)
            miSQL = vUsu.codigo & " AND campo1 IN (" & miSQL & ")"
            miSQL = "DELETE FROM tmpinformes WHERE codusu = " & miSQL
            conn.Execute miSQL
        End If
    Else
        'QUIERE METERLOS TODOS
        'Borro de sdtofm con codactiv  e inserto desde tmpinformes
        lblIndicador(0).Caption = "Eliminando registros anteriores"
        lblIndicador(0).Refresh
        miSQL = "DELETE from sdtofm WHERE codclien is null and codmarca is null"
        If Me.txtActiv(1).Text <> "" Then miSQL = miSQL & " AND codactiv =" & txtActiv(1).Text

        If txtFamia(0).Text <> "" Then miSQL = miSQL & " AND codfamia = " & txtFamia(0).Text
        conn.Execute miSQL
    End If
    
    'INSERTAMOS desde tmpinformes
    lblIndicador(0).Caption = "Insertando en descuentos"
    lblIndicador(0).Refresh
    miSQL = "INSERT INTO sdtofm (codclien,codfamia,codmarca,fechadto,dtoline1,dtoline2,codactiv,dtoEsp) SELECT"
    miSQL = miSQL & " null,campo1,null,fecha1,importe1,0,codigo1,0 FROM tmpinformes where codusu = " & vUsu.codigo
    conn.Execute miSQL
    
    
    
    Unload Me
    Screen.MousePointer = vbDefault
    
End Sub


Private Sub cmdSelecAlbaran_Click()
    If lw(3).ListItems.Count = 0 Then Exit Sub
    If lw(3).SelectedItem Is Nothing Then Exit Sub
    
    With lw(3).SelectedItem
        CadenaDesdeOtroForm = .Text & "|" & .SubItems(1) & "|" & .SubItems(2) & "|"
    End With
    Unload Me
End Sub

Private Sub cmdSeleccionarFamilia_Click()
    
    
    miSQL = ","
    For NumRegElim = 1 To lw(0).ListItems.Count
        If lw(0).ListItems(NumRegElim).Checked Then miSQL = miSQL & lw(0).ListItems(NumRegElim).Text & ","
    Next
    If miSQL = "," Then
        MsgBox "Seleccione alguna familia", vbExclamation
    Else
        CadenaDesdeOtroForm = miSQL
        Unload Me
    End If
End Sub

Private Sub cmdSelProvee_Click()
    
    miSQL = ""
    For NumRegElim = 1 To lw(1).ListItems.Count
        If lw(1).ListItems(NumRegElim).Checked Then miSQL = miSQL & "," & lw(1).ListItems(NumRegElim).Text
    Next
    If miSQL = "" Then
        MsgBox "Seleccione alguna proveedor", vbExclamation
    Else
        CadenaDesdeOtroForm = miSQL
        
        
        'Proceso la cadena para saber guardar cuales he quiado de los que hay
        LeerGuardarSeleccionProveedoresAQuitar False
        
        Unload Me
    End If
End Sub

Private Sub cmdSubRPT_Click()
    If Trim(txtModificable(1).Text) = "" Or Trim(txtModificable(2).Text) = "" Then
        MsgBox "Campos obligatorios", vbExclamation
        Exit Sub
    End If
    
    CadenaDesdeOtroForm = txtNoModificable(2).Text & "|" & txtModificable(1).Text & "|" & txtModificable(2).Text & "|"
    
    
    Unload Me
End Sub


Private Sub Command1_Click()
    
End Sub

Private Sub cmdTraerLineaCompraCliente_Click()
    If lw(4).ListItems.Count = 0 Then Exit Sub
    If lw(4).SelectedItem Is Nothing Then Exit Sub
    If lw(4).SelectedItem.Tag = "" Then
        MsgBox "No se puedesen seleccionar cantidades negativas", vbExclamation
        Exit Sub
    End If
    
    'codtipom numfactu fecfactu codtipoa numalbar numlinea
    With lw(4).SelectedItem
    
        If Mid(.Text, 1, 1) = "F" Then
            'Es una factura
            miSQL = "codtipom = " & DBSet(.Text, "T") & " AND numfactu = " & .SubItems(1) & " AND fecfactu =" & DBSet(.SubItems(2), "F")
            miSQL = miSQL & " AND " & .Tag
            miSQL = "Select * from slifac WHERE " & miSQL
        Else
            miSQL = "Select * from slialb WHERE " & .Tag
        
        End If
    End With
    
    'Cerramos mirsaux y lo abrimos con el sql
    miRsAux.Close
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    CadenaDesdeOtroForm = "OK"
    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimVez Then
        PrimVez = False
        If OpcionListado = 14 Then CargaAlbaranesFacturaClienteEuler
        If OpcionListado = 16 Then CargarFacturasVentaCliente
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim indice As Byte

    Me.Icon = frmPpal.Icon
    PrimVez = True
    
    FrameContadoresAgua.visible = False
    FrameGessocial.visible = False
    FrameEliminarPresupuestos.visible = False
    FrameDtoAsginar.visible = False
    FrameActualizaSdtoFm.visible = False
    FrameGesocialCambioSituacion.visible = False
    FrameSubRPT.visible = False
    FrameSeleccionarFamilia.visible = False
    FrameAguaMod.visible = False
    FrameComprasTratamientos.visible = False
    FrameSelProveedores.visible = False
    FrameCambioProveedorPedido.visible = False
    FrameFitoCampos.visible = False
    FrameAlbaranesInternos.visible = False
    FrameAlbaranesClientes.visible = False
    FrameMarjalChipos.visible = False
    FrameListadoComprasClientes.visible = False
    FramePreguntaDevoluciones.visible = False
    FrameTelefonia.visible = False

    indice = OpcionListado
    Me.Caption = "Listado"
    Select Case OpcionListado
    Case 0
        PonerFrameVisible FrameContadoresAgua
        optVarios(0).Value = True
    
    Case 1
        PonerFrameVisible FrameGessocial
        optVarios(2).Value = True
        
        FrameFechaBaja.BorderStyle = 0
        
        
    Case 2
        PonerFrameVisible FrameEliminarPresupuestos
        
    Case 3
        PonerFrameVisible FrameDtoAsginar

        CargarCombo_Tabla cboTipoDescuento(0), "sfamiatipodto", "clasifica", "nombre"
        
    Case 4
        PonerFrameVisible FrameActualizaSdtoFm
        lblIndicador(0).Caption = ""
    Case 5
        PonerFrameVisible FrameGesocialCambioSituacion
        Me.txtNoModificable(0).Text = RecuperaValor(OtrosDatos, 1)
        Me.txtNoModificable(1).Text = RecuperaValor(OtrosDatos, 2)
        Me.txtModificable(0).Text = ""
        
    Case 6
        PonerFrameVisible FrameSubRPT
        '  OtrosDatos   -> NUEVO|Codigo|Descrip|RPT|
        '   si es nuevo solo importa el codigo
        conSubRPT = RecuperaValor(Me.OtrosDatos, 1) = "0" 'MODIFICAR
        txtNoModificable(2).Text = RecuperaValor(Me.OtrosDatos, 2)
        txtModificable(1).Text = RecuperaValor(Me.OtrosDatos, 3)
        txtModificable(2).Text = RecuperaValor(Me.OtrosDatos, 4)
    Case 7
        PonerFrameVisible FrameSeleccionarFamilia
        CargaFamilias
        
    Case 8
        PonerFrameVisible Me.FrameAguaMod
        Me.txtNumero(0).Text = ""
        
        'OtrosDatos:    consumo anterior|con actuaql|m3|numfact fecfact para update|
        
        txtNoModificable(3).Text = RecuperaValor(OtrosDatos, 1)
        txtNoModificable(4).Text = RecuperaValor(OtrosDatos, 2)
        Label1(7).Tag = RecuperaValor(OtrosDatos, 3)
        Label1(7).Caption = Label1(7).Tag & " m3"
        Me.Caption = "Consumo"
        
    Case 9
        PonerFrameVisible FrameComprasTratamientos
        Me.Caption = "Tratamientos"
        lblIndicador(1).Caption = ""
        
    Case 10
         
        PonerFrameVisible FrameSelProveedores
        Me.Caption = "Proveedor"
        CargaProveedores
    Case 11
        PonerFrameVisible FrameCambioProveedorPedido
        Me.Caption = "Pedido proveedor"
        CargaLineasPedidoProveedor
        
    Case 12
        PonerFrameVisible FrameFitoCampos
        miSQL = Format(DateAdd("m", -2, Now), "/mm/yyyy")
        Me.txtFecha(6).Text = "01" & miSQL
        
        lblIndicador(2).Caption = ""
    Case 13
        PonerFrameVisible FrameAlbaranesInternos
        
    Case 14
        PonerFrameVisible FrameAlbaranesClientes
    Case 15
    
        PonerFrameVisible FrameMarjalChipos
    Case 16
        PonerFrameVisible FrameListadoComprasClientes
        
    Case 17
        PonerFrameVisible FramePreguntaDevoluciones
        
    Case 18
        PonerFrameVisible FrameTelefonia
        limpiar Me
        txtNumero(2).Text = "24"
    End Select
    
    Me.cmdCancelar(CInt(indice)).Cancel = True
End Sub




Private Sub PonerFrameVisible(ByRef F As Frame)
    F.Top = 0
    F.Left = 120
    F.visible = True
    Me.Height = F.Height + 480
    Me.Width = F.Width + 240
End Sub

Private Sub InicializarVbles(AñadireElDeEmpresa As Boolean)
    cadFormula = ""
    cadSelect = ""
    CadParam = "|"
    numParam = 0
    cadTitulo = ""
    cadNomRPT = ""
    conSubRPT = False
    cadPDFrpt = ""
    vMostrarTree = False
    If AñadireElDeEmpresa Then
        CadParam = CadParam & "pEmpresa=""" & vEmpresa.nomempre & """|"
        numParam = numParam + 1
    End If
End Sub

Private Sub LlamarImprimir()
    With frmImprimir
        .FormulaSeleccion = cadFormula
        .OtrosParametros = CadParam
        .NumeroParametros = numParam

        .SoloImprimir = False
        .EnvioEMail = False
        .Titulo = cadTitulo
        .Opcion = 3000   'VAN TODOS EN ESTE SACO
        .NombrePDF = ""
        .NombrePDF = cadPDFrpt
        .NombreRPT = cadNomRPT
        .ConSubInforme = conSubRPT
        .MostrarTreeDesdeFuera = vMostrarTree
        .Show vbModal
    End With
End Sub



Private Sub frmB_Selecionado(CadenaDevuelta As String)
    miSQL = CadenaDevuelta
End Sub

Private Sub frmC_Selec(vFecha As Date)
    miSQL = vFecha
End Sub

Private Sub frmCli_DatoSeleccionado(CadenaSeleccion As String)
    miSQL = CadenaSeleccion
End Sub

Private Sub frmMtoFamilia_DatoSeleccionado(CadenaSeleccion As String)
    miSQL = CadenaSeleccion
End Sub

Private Sub imgActividad_Click(Index As Integer)

    Screen.MousePointer = vbHourglass
    Set frmB = New frmBuscaGrid
    frmB.vTitulo = "Activiadad"
    miSQL = "Codigo|sactiv|codactiv|N||20·"
    miSQL = miSQL & "descripcion|sactiv|nomactiv|T||45·"
    frmB.vCampos = miSQL
    frmB.vCargaFrame = False
    frmB.vDevuelve = "0|1|"
    frmB.vselElem = 1
    frmB.vConexionGrid = 1  'ODBC Ariges
    frmB.vTabla = "sactiv"
    frmB.vSQL = ""
    miSQL = ""
    frmB.Show vbModal
    Set frmB = Nothing
    Screen.MousePointer = vbDefault
    If miSQL <> "" Then
        
        txtActiv(Index).Text = RecuperaValor(miSQL, 1)
        Me.txtDescActiv(Index).Text = RecuperaValor(miSQL, 2)
       
    End If

End Sub


Private Sub imgCheck_Click(Index As Integer)
Dim Cual As Byte
Dim Chec As Boolean


    If Index < 2 Then
        'Proveedores
        Cual = 0
        Chec = True
        If (Index Mod 2) = 0 Then Chec = False
    ElseIf Index < 4 Then
        Cual = 1
        Chec = True
        If (Index Mod 2) = 0 Then Chec = False
        
    ElseIf Index < 6 Then
        Cual = 2
        Chec = True
        If (Index Mod 2) = 0 Then Chec = False

'    ElseIf Index < 8 Then
'        '6 7
'        Cual = 4
'        Chec = (Index Mod 2) = 1
'    Else
'        '8 9
'        Cual = 8
'        Chec = (Index Mod 2) = 1
    End If
         
    For NumRegElim = 1 To lw(Cual).ListItems.Count
        lw(Cual).ListItems(NumRegElim).Checked = Chec
    Next

   
End Sub

Private Sub imgCliente_Click(Index As Integer)
    Screen.MousePointer = vbHourglass
    miSQL = ""
    Set frmCli = New frmFacClientes3
    frmCli.DatosADevolverBusqueda = "0|1|"
    frmCli.Show vbModal
    Set frmCli = Nothing
    If miSQL <> "" Then
        Me.txtCliente(Index).Text = RecuperaValor(miSQL, 1)
        Me.txtDescClie(Index).Text = RecuperaValor(miSQL, 2)
    End If
End Sub



Private Sub imgFamilia_Click(Index As Integer)
    miSQL = ""
    Set frmMtoFamilia = New frmAlmFamiliaArticulo
    frmMtoFamilia.DatosADevolverBusqueda = "0|1"
    frmMtoFamilia.Show vbModal
    Set frmMtoFamilia = Nothing
    If miSQL <> "" Then
        txtFamia(Index).Text = RecuperaValor(miSQL, 1)
        txtDescFamia(Index).Text = RecuperaValor(miSQL, 2)
        miSQL = ""
    End If
End Sub

Private Sub imgFecha_Click(Index As Integer)
   miSQL = ""
   Set frmC = New frmCal
   frmC.Fecha = Now
   If txtFecha(Index).Text <> "" Then
        If IsDate(txtFecha(Index).Text) Then frmC.Fecha = CDate(txtFecha(Index).Text)
   End If
   frmC.Show vbModal
   Set frmC = Nothing
   If miSQL <> "" Then txtFecha(Index).Text = Format(miSQL, "dd/mm/yyyy")
End Sub




Private Sub lw_DblClick(Index As Integer)
    If Index = 4 Then
        cmdTraerLineaCompraCliente_Click
    Else
        cmdSelecAlbaran_Click
    End If
End Sub

Private Sub lw_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then cmdTraerLineaCompraCliente_Click
End Sub

Private Sub optFitoCampos_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub optVarios_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub





Private Function PonerDesdeHasta(campo As String, Tipo As String, indD As Byte, indH As Byte, param As String) As Boolean
Dim devuelve As String
Dim cad As String
Dim Subtipo As String 'F: fecha   N: numero   T: texto  H: HORA
Dim TDes As TextBox
Dim THas As TextBox
Dim DesD As TextBox 'Descripcion DESDE
Dim DesH As TextBox '    "       HASTA

    PonerDesdeHasta = False
    
    Select Case Tipo
    Case "F"
        'Campos fecha
        Set TDes = txtFecha(indD)
        Set THas = txtFecha(indH)
        Subtipo = "F"
        'If indD = 27 Or indH = 28 Then Subtipo = "FH"
    Case "CLI"
        'Cliente
        Set TDes = txtCliente(indD)
        Set THas = txtCliente(indH)
        Set DesD = txtDescClie(indD)
        Set DesH = txtDescClie(indH)
        Subtipo = "N"

'
'    Case "PRO"
'        Set TDes = txtCodProve(indD)
'        Set THas = txtCodProve(indH)
'        Set DesD = txtDescProve(indD)
'        Set DesH = txtDescProve(indH)
'        Subtipo = "N"
'
'    Case "ART"
'
'        Set TDes = txtArticulo(indD)
'        Set THas = txtArticulo(indH)
'        Set DesD = txtDescArticulo(indD)
'        Set DesH = txtDescArticulo(indH)
'        Subtipo = "T"
'    Case "AGT"
'        Set TDes = txtAgente(indD)
'        Set THas = txtAgente(indH)
'        Set DesD = txtDescAgente(indD)
'        Set DesH = txtDescAgente(indH)
'        Subtipo = "N"
'
'    Case "ALP"
'        'Numero albaran proveedores
'
''        Set TDes = txtNumAlbar(indD)
''        Set THas = txtNumAlbar(indH)
'        Subtipo = "T"
'
'    Case "TRA"
'        'TRABAJADOR
'
'        Set TDes = txtCodTraba(indD)
'        Set THas = txtCodTraba(indH)
'        Subtipo = "N"
'
'        Set DesD = txtDescTraba(indD)
'        Set DesH = txtDescTraba(indH)
'
        
        
 
'
'    Case "FAM"
'        'FAMILIA
'
'        Set TDes = Me.txtFamia(indD)
'        Set THas = txtFamia(indH)
'        Subtipo = "N"
'        Set DesD = txtDescFamia(indD)
'        Set DesH = txtDescFamia(indH)
'
'
'    Case "MAR"
'
'        Set TDes = Me.txtmarca(indD)
'        Set THas = txtmarca(indH)
'        Subtipo = "N"
'        Set DesD = txtDescmarca(indD)
'        Set DesH = txtDescmarca(indH)
'
'    Case "ACT"
'        'ACTIVIADD
'
'        Set TDes = Me.txtcodactiv(indD)
'        Set THas = txtcodactiv(indH)
'        Subtipo = "N"
'        If indD = 5 Then
'            'llamadas
'            Set DesD = txtDescActiv(indD)
'            Set DesH = txtDescActiv(indH)
'        End If
'
    End Select
    
    devuelve = CadenaDesdeHasta(TDes.Text, THas.Text, campo, Subtipo)
    If devuelve = "Error" Then
        PonerFoco TDes
        Exit Function
    End If
    If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Function
    
    If Subtipo <> "F" And Subtipo <> "FH" Then
        'Fecha para Crystal Report
        If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Function
    Else
        'Fecha para la Base de Datos
        cad = CadenaDesdeHastaBD(TDes.Text, THas.Text, campo, Subtipo)
        If Not AnyadirAFormula(cadSelect, cad) Then Exit Function
    End If
    
    If devuelve <> "" Then
        If param <> "" Then
            'Parametro Desde/Hasta
            CadParam = CadParam & AnyadirParametroDH(param, TDes, THas, DesD, DesH) & """|"
            numParam = numParam + 1
        End If
        PonerDesdeHasta = True
    End If
End Function


Private Function AnyadirParametroDH(cad As String, ByRef TextoDESDE As TextBox, TextoHasta As TextBox, ByRef TD As TextBox, ByRef TH As TextBox) As String
On Error Resume Next
    
     If TextoDESDE.Text <> "" Then
        cad = cad & "desde " & TextoDESDE.Text
        If TD.Text <> "" Then cad = cad & " - " & TD.Text
    End If
    If TextoHasta.Text <> "" Then
        cad = cad & "  hasta " & TextoHasta.Text
        If TH <> "" Then cad = cad & " - " & TH.Text
    End If
    
    AnyadirParametroDH = cad
    If Err.Number <> 0 Then Err.Clear
End Function






Private Sub txtActiv_Gotfocus(Index As Integer)
    ConseguirFoco txtActiv(Index), 3
End Sub

Private Sub txtActiv_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    If (Shift And vbCtrlMask) > 0 Then
        'Caption = KeyCode
        If KeyCode = 65 Then imgActividad_Click Index
    End If
    
End Sub

Private Sub txtActiv_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtActiv_LostFocus(Index As Integer)
    txtActiv(Index).Text = Trim(txtActiv(Index).Text)
    cadTitulo = ""
    miSQL = ""
    If txtActiv(Index).Text <> "" Then
        If IsNumeric(txtActiv(Index).Text) Then
            cadTitulo = DevuelveDesdeBD(conAri, "nomactiv", "sactiv", "codactiv", txtActiv(Index).Text, "N")
            
            If Index <= 1 And cadTitulo = "" Then miSQL = "No existe la actividad"
            
        Else
            miSQL = "Campo numerico"
        End If
    End If
    Me.txtDescActiv(Index).Text = cadTitulo
    If miSQL <> "" Then
        MsgBox miSQL, vbExclamation
        txtActiv(Index).Text = ""
        PonerFoco txtActiv(Index)
    End If
End Sub




Private Sub txtCliente_GotFocus(Index As Integer)
    ConseguirFoco txtCliente(Index), 3
End Sub

Private Sub txtCliente_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If (Shift And vbCtrlMask) > 0 Then
        'Caption = KeyCode
        If KeyCode = 65 Then imgCliente_Click Index
    End If
End Sub

Private Sub txtCliente_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtCliente_LostFocus(Index As Integer)
Dim Descri As String
    
    Descri = ""
    txtCliente(Index).Text = Trim(txtCliente(Index).Text)
    If txtCliente(Index).Text <> "" Then
        If Not IsNumeric(txtCliente(Index).Text) Then
            MsgBox "Campo codigo cliente debe ser numérico", vbExclamation
            PonerFoco txtCliente(Index)
        Else
            Descri = DevuelveDesdeBD(conAri, "nomclien", "sclien", "codclien", txtCliente(Index).Text, "N")
            If Descri = "" Then
                MsgBox "No existe el cliente : " & txtCliente(Index).Text, vbExclamation
            End If
        End If
    End If
    Me.txtDescClie(Index).Text = Descri
    
End Sub



'******************************************************************************
'******************************************************************************
'
'   GESOCIAL
'



Private Sub txtFamia_GotFocus(Index As Integer)
    ConseguirFoco txtFamia(Index), 3
End Sub

Private Sub txtFamia_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If (Shift And vbCtrlMask) > 0 Then
        'Caption = KeyCode
        If KeyCode = 65 Then imgFamilia_Click Index
    End If
End Sub

Private Sub txtFamia_KeyPress(Index As Integer, KeyAscii As Integer)
     KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtFamia_LostFocus(Index As Integer)
Dim codigo  As String
    txtFamia(Index).Text = Trim(txtFamia(Index).Text)
    codigo = ""
    miSQL = ""
    If txtFamia(Index).Text <> "" Then
        If IsNumeric(txtFamia(Index).Text) Then
            codigo = DevuelveDesdeBD(conAri, "nomfamia", "sfamia", "codfamia", txtFamia(Index).Text, "N")
            If codigo = "" Then
                MsgBox "El codigo no pertence a ninguna familia", vbExclamation
                If Index = 0 Then
                    txtFamia(Index).Text = ""
                    txtDescFamia(Index).Text = ""
                End If
            End If
        Else
            miSQL = "Campo numerico"
        End If
    End If
     
    Me.txtDescFamia(Index).Text = codigo
    If miSQL <> "" Then
        MsgBox miSQL, vbExclamation
        txtFamia(Index).Text = ""
        PonerFoco txtFamia(Index)
    End If
End Sub


Private Sub txtFecha_GotFocus(Index As Integer)
    ConseguirFoco txtFecha(Index), 3
End Sub

Private Sub txtFecha_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtFecha_LostFocus(Index As Integer)
Dim T As String
    txtFecha(Index).Text = Trim(txtFecha(Index).Text)
    If txtFecha(Index).Text <> "" Then
        T = txtFecha(Index).Text
        If EsFechaOK(T) Then
            txtFecha(Index).Text = T
        Else
            MsgBox "Fecha con formato incorrecto: " & txtFecha(Index).Text, vbExclamation
            txtFecha(Index).Text = ""
            PonerFoco txtFecha(Index)
        End If
    End If
    
End Sub



Private Sub txtModificable_KeyPress(Index As Integer, KeyAscii As Integer)
      KEYpressGnral KeyAscii, 2, True
End Sub



Private Sub CargaFamilias()
    Set miRsAux = New ADODB.Recordset
    
    miSQL = "Select codfamia,nomfamia from sfamia where codfamia in (" & OtrosDatos & ") ORDER BY codfamia"
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    lw(0).ListItems.Clear
    While Not miRsAux.EOF
        lw(0).ListItems.Add , , miRsAux!Codfamia
        NumRegElim = NumRegElim + 1
        lw(0).ListItems(NumRegElim).SubItems(1) = miRsAux!nomfamia
        lw(0).ListItems(NumRegElim).Checked = True
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    
End Sub





Private Sub CargaProveedores()
    Set miRsAux = New ADODB.Recordset
    
    LeerGuardarSeleccionProveedoresAQuitar True
    
    
    miRsAux.Open OtrosDatos, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    miSQL = ""
    While Not miRsAux.EOF
        miSQL = miSQL & ", " & miRsAux!Codigo1
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    miSQL = Mid(miSQL, 2)
    
    miSQL = "Select codprove,nomprove from sprove where codprove in (" & miSQL & ") ORDER BY codprove"
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    lw(1).ListItems.Clear
    While Not miRsAux.EOF
        lw(1).ListItems.Add , , miRsAux!Codprove
        NumRegElim = NumRegElim + 1
        lw(1).ListItems(NumRegElim).SubItems(1) = miRsAux!nomprove
        cadTitulo = "|" & miRsAux!Codprove & "|"
        If InStr(1, auxiliar, cadTitulo) = 0 Then
            lw(1).ListItems(NumRegElim).Checked = True
        Else
            lw(1).ListItems(NumRegElim).Checked = False
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    
End Sub

Private Sub txtNumero_GotFocus(Index As Integer)
    
     ConseguirFoco txtNumero(Index), 3
    
End Sub

Private Sub txtNumero_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtNumero_LostFocus(Index As Integer)
Dim J As Long
 txtNumero(Index).Text = Trim(txtNumero(Index).Text)
    miSQL = ""
    If txtNumero(Index).Text <> "" Then
        
        Select Case Index
        Case 0
            If Not PonerFormatoEntero(txtNumero(Index)) Then
                txtNumero(Index).Text = ""
                
            Else
                'Calulamos la diferencia
                
                J = Val(Trim(Mid(Me.txtNoModificable(3).Text, 11))) 'quitamos la fecha y queda la lectura
                J = Val(txtNumero(Index).Text) - J
                If J < 0 Then
                    MsgBox "Consumo negativo", vbExclamation
                    J = 0
                Else
                    miSQL = txtNumero(Index).Text
                    
                End If
            End If
        Case 1
        
            If Not PonerFormatoDecimal(txtNumero(Index), 3) Then txtNumero(Index).Text = ""
        Case 2
            If Not PonerFormatoEntero(txtNumero(Index)) Then txtNumero(Index).Text = ""
        
        End Select
    Else
        
    End If
    If Index = 0 Then
        txtNumero(0).Text = miSQL
        If J = 0 Then J = RecuperaValor(OtrosDatos, 3)
        Label1(7).Tag = J
        Label1(7).Caption = Label1(7).Tag & " m3"
        If txtNumero(0).Text = "" Then PonerFoco txtNumero(0)
        
        
    End If
End Sub




Private Function HacerUpdateConsumo() As Boolean
Dim Rc As ADODB.Recordset
Dim L As Integer
Dim Consumo As Integer
Dim ini As Integer
Dim fin As Integer
Dim Meses As Integer
Dim Normal As Boolean   'Versus industrial
Dim Tramo As Integer
Dim SeFactura As Boolean
    HacerUpdateConsumo = False
    On Error GoTo eHacerUpdateConsumo
    
    
    cadSelect = ""
    cadFormula = RecuperaValor(OtrosDatos, 4)  'WHERE de numfavctu fecfactu...
    Set Rc = New ADODB.Recordset
    
    
    
    
    'Las modificaciones del consumo se reflejan en
    ' Linea 1-4     CONSUMO AGUA por bloques
    ' Linea 6-9     ALCANTARILLADO
    ' Linea 20 Consumo para la impresion en report (importe=0)
    ' Linea 21   Canon cuota consumo
    
    'Ademas habra que podificar la ultima lectura de contador
    
    miSQL = RecuperaValor(OtrosDatos, 4)
    miSQL = "Select  cantidad from slifac WHERE " & miSQL
    miSQL = miSQL & " AND numlinea in (5,10,12,25)"   'lleva el periodo
    Rc.Open miSQL, conn, adOpenKeyset, adLockPessimistic, adCmdText
    If Rc.EOF Then Err.Raise 513, , "Imposible obtener periodo(EOF)"
    ini = -1
    Normal = True 'Todo bien
    Do
        Meses = Rc!cantidad
        If ini < 0 Then
            ini = Rc!cantidad
        Else
            If ini <> Meses Then Normal = False
        End If
        Rc.MoveNext
    Loop Until Rc.EOF
    Rc.Close
    
    If Not Normal Then Err.Raise 513, , "Imposible obtener periodo(cant. distinta lineas 5,10,12,25)"
    
    
    'Si tiene cada uno de los bloques entonces actualizaremos
    'OtrosDatos:    consumo anterior|con actuaql|m3|numfact fecfact para update|
    
    'Consumo para report
    Consumo = Val(Label1(7).Tag)
    
    'Indica el consumo
    miSQL = "UPDATE slifac set cantidad = " & Consumo & ",numbultos=" & Consumo & " WHERE " & cadFormula
    miSQL = miSQL & " AND numlinea=20"
    conn.Execute miSQL
    'Canon Generalita sobre consumo
    miSQL = "Select  numlinea,nomartic,ampliaci,precioar from slifac WHERE " & cadFormula
    miSQL = miSQL & " AND numlinea =21"
    Rc.Open miSQL, conn, adOpenKeyset, adLockPessimistic, adCmdText
    If Not Rc.EOF Then
        miSQL = "UPDATE slifac set cantidad = " & Consumo & ",numbultos=" & Consumo
        miSQL = miSQL & ", importel=round(precioar * " & Consumo & ",2)"
        miSQL = miSQL & " WHERE " & cadFormula
        miSQL = miSQL & " AND numlinea=21"
        conn.Execute miSQL
    End If
    Rc.Close
    
    'Consumo por tramos
    miSQL = "Select  numlinea,nomartic,ampliaci,precioar from slifac WHERE " & cadFormula
    miSQL = miSQL & " AND numlinea between 1 and 4"
    Rc.Open miSQL, conn, adOpenKeyset, adLockPessimistic, adCmdText
    
    L = 0
    While Not Rc.EOF
        L = L + 1
        Rc.MoveNext
    Wend
    
    
    
    cadSelect = "Leyendo tramos factura"
    If L > 0 Then
        'Si no tiene lineas es que no se factura
        Normal = L = 3   'Si no es industrial
    
        Rc.MoveFirst
        ini = 0
        
    
        'OK vamos a ver los tramos
        
        cadSelect = Rc!Ampliaci
        'Tres tramos
        L = InStr(1, cadSelect, "m3")
        cadSelect = Trim(Mid(cadSelect, 1, L - 1))
        L = InStrRev(cadSelect, " ")
        cadSelect = Trim(Mid(cadSelect, L))
        L = Val(cadSelect)
        If L = 0 Then Err.Raise 513, , "Imposible obtener tramo(I)"
                    
        If Normal Then
            fin = Meses * L
        Else
            fin = L
        End If
        Tramo = fin
        If Consumo <= fin Then
            fin = Consumo
            Consumo = 0
        Else
            Consumo = Consumo - fin
        End If
        miSQL = "UPDATE slifac set cantidad = " & fin & ",numbultos=" & fin
        miSQL = miSQL & ", importel=round(precioar * " & fin & ",2)"
        miSQL = miSQL & " WHERE " & cadFormula
        miSQL = miSQL & " AND numlinea=1"
        conn.Execute miSQL
        


    
        'Segunda linea (para ambos Normal e industrial
        If Consumo = 0 Then
            fin = 0
            
        Else
            
            
            If Normal Then
                cadSelect = "Tramo II"
                'Solo para NORMAL vemos el segundo nivel de consumo
                Rc.MoveNext
                cadSelect = Rc!Ampliaci
                'Tres tramos
                L = InStr(1, cadSelect, "m3")
                cadSelect = Trim(Mid(cadSelect, 1, L - 1))
                L = InStrRev(cadSelect, " ")
                cadSelect = Trim(Mid(cadSelect, L))
                L = Val(cadSelect)
                If L = 0 Then Err.Raise 513, , "Imposible obtener tramo(II)"
                    
                fin = Meses * L
          
                fin = fin - Tramo
                If fin <= 0 Then Err.Raise 513, , "Obtener tramo(II). Valor 0"
                
                If Consumo >= fin Then
                    Consumo = Consumo - fin
                Else
                    fin = Consumo
                    Consumo = 0
                End If
            Else
                'Indutria. DOS tramos
                fin = Consumo   'lo que quede va al tramo 2
                Consumo = 0
            End If
                  
        End If
        
        
        
        
        
        miSQL = "UPDATE slifac set cantidad = " & fin & ",numbultos=" & fin
        miSQL = miSQL & ", importel=round(precioar * " & fin & ",2)"
        miSQL = miSQL & " WHERE " & cadFormula
        miSQL = miSQL & " AND numlinea=2"
        conn.Execute miSQL
    
        'Tercera linea SOLO domestico
        If Normal Then
            If Consumo = 0 Then
                fin = 0
            Else
                fin = Consumo
            End If
            miSQL = "UPDATE slifac set cantidad = " & fin & ",numbultos=" & fin
            miSQL = miSQL & ", importel=round(precioar * " & fin & ",2)"
            miSQL = miSQL & " WHERE " & cadFormula
            miSQL = miSQL & " AND numlinea=3"
            conn.Execute miSQL
        End If
    End If
    Rc.Close
    
    
    
    '****************************************************************************************************
    'Alcantarillado
    miSQL = "Select  numlinea,nomartic,ampliaci,precioar from slifac WHERE " & cadFormula
    miSQL = miSQL & " AND numlinea between 6 and 8"
    Rc.Open miSQL, conn, adOpenKeyset, adLockPessimistic, adCmdText
    
    L = 0
    While Not Rc.EOF
        L = L + 1
        Rc.MoveNext
    Wend
    
    If L > 0 Then
        'Se le han facturado alncatarillado
    
        Rc.MoveFirst
        ini = 0
        Consumo = Val(Label1(7).Tag)
    
        'OK vamos a ver los tramos
        cadSelect = Rc!Ampliaci
        'Tres tramos
        L = InStr(1, cadSelect, "m3")
        cadSelect = Trim(Mid(cadSelect, 1, L - 1))
        L = InStrRev(cadSelect, " ")
        cadSelect = Trim(Mid(cadSelect, L))
        L = Val(cadSelect)
        If L = 0 Then Err.Raise 513, , "Imposible obtener tramo alcant(I)"
            
        fin = Meses * L
        Tramo = fin
        If Consumo <= fin Then
            fin = Consumo
            Consumo = 0
        Else
            Consumo = Consumo - fin
        End If
        miSQL = "UPDATE slifac set cantidad = " & fin & ",numbultos=" & fin
        miSQL = miSQL & ", importel=round(precioar * " & fin & ",2)"
        miSQL = miSQL & " WHERE " & cadFormula
        miSQL = miSQL & " AND numlinea=6"
        conn.Execute miSQL
        


    
        'Segunda linea (para ambos Normal e industrial
        If Consumo = 0 Then
            fin = 0
            
        Else
            
            
           
            cadSelect = "Tramo II alcantarillado"
            
            Rc.MoveNext
            cadSelect = Rc!Ampliaci
            'Tres tramos
            L = InStr(1, cadSelect, "m3")
            cadSelect = Trim(Mid(cadSelect, 1, L - 1))
            L = InStrRev(cadSelect, " ")
            cadSelect = Trim(Mid(cadSelect, L))
            L = Val(cadSelect)
            If L = 0 Then Err.Raise 513, , "Imposible obtener tramo(II) alcantarillado"
                
            fin = Meses * L
            
            fin = fin - Tramo
            If fin <= 0 Then Err.Raise 513, , "Obtener tramo(II) Alcantarillado. Valor <=0"
            
            If Consumo >= fin Then
                Consumo = Consumo - fin
            Else
                fin = Consumo
                Consumo = 0
            End If
     
            
            
            miSQL = "UPDATE slifac set cantidad = " & fin & ",numbultos=" & fin
            miSQL = miSQL & ", importel=round(precioar * " & fin & ",2)"
            miSQL = miSQL & " WHERE " & cadFormula
            miSQL = miSQL & " AND numlinea=7"
            conn.Execute miSQL
        
            'Tercera linea
            
            If Consumo = 0 Then
                fin = 0
            Else
                fin = Consumo
            End If
            miSQL = "UPDATE slifac set cantidad = " & fin & ",numbultos=" & fin
            miSQL = miSQL & ", importel=round(precioar * " & fin & ",2)"
            miSQL = miSQL & " WHERE " & cadFormula
            miSQL = miSQL & " AND numlinea=8"
            conn.Execute miSQL

        
        
        End If
    End If
    Rc.Close
    
    'Vemos el numero de conador
    miSQL = "Select referenc,fecfactu from scafac1 WHERE " & cadFormula
    Rc.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    cadPDFrpt = Rc!referenc
    cadNomRPT = Format(Rc!FecFactu, FormatoFecha)
    Rc.Close
    
    'Updates que faltan.
    'Scafac1 donde pondra lecutar actual
    Consumo = Val(Label1(7).Tag)
    miSQL = Mid(Me.txtNoModificable(4).Text, 1, 10) & " " & Me.txtNumero(0).Text
    miSQL = "UPDATE scafac1 set observa2 = '" & miSQL & "'"
    miSQL = miSQL & " WHERE " & cadFormula
    conn.Execute miSQL
    
    
    miSQL = "UPDATE aguahcolecturas set lec_actual = " & Me.txtNumero(0).Text
    miSQL = miSQL & " WHERE contador = " & DBSet(cadPDFrpt, "T") & " AND fecha_factura ='" & cadNomRPT & "'"
    conn.Execute miSQL
    
    miSQL = "UPDATE aguacontadores  set lec_anterior = " & Me.txtNumero(0).Text
    miSQL = miSQL & " WHERE contador = " & DBSet(cadPDFrpt, "T")
    conn.Execute miSQL



    Set LOG = New cLOG
    miSQL = Mid(Me.txtNoModificable(4).Text, 11) & " / " & Me.txtNumero(0).Text
    
    miSQL = "FRA MOD[CONSUMO] " & miSQL & vbCrLf & cadFormula
    miSQL = Replace(miSQL, "codtipom", "")
    miSQL = Replace(miSQL, "numfactu", "")
    miSQL = Replace(miSQL, "fecfactu", "")
    miSQL = Replace(miSQL, " and =", "")
    miSQL = Replace(miSQL, "'", " ")
    LOG.Insertar 8, vUsu, miSQL
    Set LOG = Nothing
    
    
    
    HacerUpdateConsumo = True
    
eHacerUpdateConsumo:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description & vbCrLf & cadSelect
    Set Rc = Nothing
End Function



Private Function GeneraDatosComprasTratamientos() As Boolean
Dim RN As ADODB.Recordset
Dim total As Currency

    On Error GoTo eGeneraDatosComprasTratamientos
    
    Set miRsAux = New ADODB.Recordset
    Set RN = New ADODB.Recordset
    
    Me.lblIndicador(1).Caption = "Obteniendo ventas articulo"
    Me.lblIndicador(1).Refresh
    miSQL = "DELETE FROM tmpinformes WHERE codusu = " & vUsu.codigo
    conn.Execute miSQL
    
'''''''    miSQL = "select " & vUsu.Codigo & ", @rownum:=@rownum+1 AS rownum ,slifac.codartic,codfamia,sartic.nomartic,"
'''''''    miSQL = miSQL & " sum(cantidad*sartic.preciouc) from slifac,sartic,(SELECT @rownum:=0) r"
'''''''    miSQL = miSQL & " where slifac.codartic=sartic.codartic AND (codtipom,numfactu,fecfactu) IN"
'''''''    miSQL = miSQL & " (Select codtipom,numfactu,fecfactu from scafac1 where fecfactu between "
'''''''    miSQL = miSQL & DBSet(txtFecha(4).Text, "F") & " AND " & DBSet(txtFecha(5).Text, "F")
'''''''
'''''''    miSQL = miSQL & " and referenc like 'Parte:%') group by codartic"
'''''''    'miSQL = miSQL & " and referenc like '%') group by codartic"

'    miSQL = "select " & vUsu.Codigo & ", @rownum:=@rownum+1 AS rownum ,slifac.codartic,cliAbono,codfamia,sartic.nomartic,"
'    miSQL = miSQL & " sum(cantidad*sartic.preciouc) from scafac,scafac1,slifac,sartic,sclien  WHERE scafac1.Codtipom = "
'    miSQL = miSQL & " scafac.Codtipom And scafac1.NumFactu = scafac.NumFactu And scafac1.FecFactu = scafac.FecFactu"
'    miSQL = miSQL & " AND scafac1.Codtipom = slifac.Codtipom And scafac1.NumFactu = slifac.NumFactu And scafac1.FecFactu = slifac.FecFactu"
'    miSQL = miSQL & " and scafac1.numalbar=slifac.numalbar and scafac1.codtipoa=slifac.codtipoa and slifac.codartic=sartic.codartic"
'    miSQL = miSQL & " and scafac.fecfactu between " & DBSet(txtFecha(4).Text, "F") & " AND " & DBSet(txtFecha(5).Text, "F")
'    miSQL = miSQL & " and referenc like 'Parte:%' and scafac.codclien=sclien.codclien"
'    miSQL = miSQL & " group by slifac.codartic,cliAbono"

    
    miSQL = "select slifac.codartic,sum(cantidad) cuantos from scafac1,slifac WHERE "
    miSQL = miSQL & " scafac1.Codtipom = slifac.Codtipom And scafac1.NumFactu = slifac.NumFactu And scafac1.FecFactu = slifac.FecFactu"
    miSQL = miSQL & " and scafac1.numalbar=slifac.numalbar and scafac1.codtipoa=slifac.codtipoa "
    miSQL = miSQL & " and scafac1.fecfactu between " & DBSet(txtFecha(4).Text, "F") & " AND " & DBSet(txtFecha(5).Text, "F")
    miSQL = miSQL & " and referenc like 'Parte:%' "
    miSQL = miSQL & " group by slifac.codartic"
    miRsAux.Open miSQL, conn, adOpenKeyset, adLockOptimistic, adCmdText
    If miRsAux.EOF Then
        MsgBox "Ningun dato ", vbExclamation
        GoTo eGeneraDatosComprasTratamientos
    End If
    
    
    miSQL = ""
    While Not miRsAux.EOF
        miSQL = miSQL & ", " & DBSet(miRsAux!codArtic, "T")
        miRsAux.MoveNext
    Wend
    
    miRsAux.MoveFirst
    
    Me.lblIndicador(1).Caption = "Abriendo articulos"
    Me.lblIndicador(1).Refresh

    
    miSQL = Mid(miSQL, 2)
    miSQL = "Select codartic,preciouc,codfamia,nomartic from sartic where codartic IN (" & miSQL & ")"
    RN.Open miSQL, conn, adOpenKeyset, adLockPessimistic, adCmdText
    
    miSQL = ""
    While Not miRsAux.EOF
        NumRegElim = NumRegElim + 1
        Me.lblIndicador(1).Caption = "Art: " & miRsAux!codArtic
        Me.lblIndicador(1).Refresh
        
        
        'tmpinformes(codusu,codigo1,nombre1,campo2,campo1,nombre2,importe1) " & miSQL
        
        miSQL = miSQL & ", (" & vUsu.codigo & "," & NumRegElim & "," & DBSet(miRsAux!codArtic, "T") & ",0,"
    
        RN.Find "codartic = " & DBSet(miRsAux!codArtic, "T"), , adSearchForward, 1
        'NO PUEDE SER EOF
        If Not IsNull(miRsAux!Cuantos) Then
            total = DBLet(RN!precioUC, "N")
            total = Round(total * miRsAux!Cuantos, 2)
        Else
            total = 0
        End If
        
        
        miSQL = miSQL & RN!Codfamia & "," & DBSet(RN!NomArtic, "T") & "," & DBSet(total, "N") & ")"
        
        
        If Len(miSQL) > 10000 Then
            miSQL = Mid(miSQL, 2)
            miSQL = "INSERT INTO tmpinformes(codusu,codigo1,nombre1,campo2,campo1,nombre2,importe1) VALUES " & miSQL
            conn.Execute miSQL
            miSQL = ""
        End If
        
        miRsAux.MoveNext
    Wend
    
    If miSQL <> "" Then
        miSQL = Mid(miSQL, 2)
        miSQL = "INSERT INTO tmpinformes(codusu,codigo1,nombre1,campo2,campo1,nombre2,importe1) VALUES " & miSQL
        conn.Execute miSQL
    End If
    
    Me.lblIndicador(1).Caption = "Cerrando"
    Me.lblIndicador(1).Refresh
    miRsAux.Close
    RN.Close
    Espera 0.5
    
    'miSQL = DevuelveDesdeBD(conAri, "count(*)", "tmpinformes", "codusu", CStr(vUsu.Codigo))
    If NumRegElim > 0 Then
        GeneraDatosComprasTratamientos = True
    Else
        MsgBox "No existen datos entre las fechas", vbExclamation
    End If
    
    
    
    
    
    
eGeneraDatosComprasTratamientos:
    If Err.Number <> 0 Then
        MuestraError Err.Number, , Err.Description
        GeneraDatosComprasTratamientos = False
    End If
    Set miRsAux = Nothing
    Set RN = Nothing
End Function



Private Sub GenerarApunteAjusteTratamientos()
Dim Mc As Contadores
Dim ImporteTotal As Currency

    On Error GoTo eGenerarApunteAjusteTratamientos

    ResultadoFechaContaOK = EsFechaOKConta(Now, True)
    If ResultadoFechaContaOK > 0 Then
        If ResultadoFechaContaOK <> 4 Then MsgBox MensajeFechaOkConta, vbExclamation
        Exit Sub
    End If
    
    Set miRsAux = New ADODB.Recordset
    numParam = 0
    If vEmpresa.FechaFin <= CDate(Now) Then numParam = 1
    
    
    Set Mc = New Contadores
    cadPDFrpt = "Obteniendo contador asientos"
    If Mc.ConseguirContador("0", numParam = 0, True) = 1 Then Err.Raise 513, , cadPDFrpt
    

    

    miRsAux.Open "Select * from advparametros", conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    'Ni puede ser eof ni puede no tener valor(NULL)
    cadPDFrpt = miRsAux!DiarioAjustes
    numParam = miRsAux!CoceptoAjustes
    miRsAux.Close
    
    
    'Cabecera
    '--------------------------------------
    Me.lblIndicador(1).Caption = "Apunte"
    Me.lblIndicador(1).Refresh
    
    miSQL = "Ajuste compras tratamientos(Ariges). Periodo : " & txtFecha(4).Text & " - " & txtFecha(5).Text
    miSQL = miSQL & vbCrLf & "Realizado por " & vUsu.Nombre & " el " & Now
    miSQL = "," & DBSet(miSQL, "T", "S") & ")"
    miSQL = cadPDFrpt & ",'" & Format(Now, FormatoFecha) & "'," & Mc.Contador & miSQL
    miSQL = "INSERT INTO cabapu (numdiari, fechaent, numasien, obsdiari) VALUES (" & miSQL
    ConnConta.Execute miSQL
    
    cadTitulo = "INSERT INTO linapu(numdiari,fechaent,numasien,linliapu,codmacta,numdocum,codconce,ampconce,timporteD,timporteH,ctacontr,idcontab,punteada,traspasado) VALUES "
    
    'linapu(numdiari,fechaent,numasien,linliapu,codmacta,numdocum,codconce,ampconce,timporteD,timporteH,ctacontr,idcontab,punteada,traspasado)
    cadFormula = ", (" & cadPDFrpt & ",'" & Format(Now, FormatoFecha) & "'," & Mc.Contador & ","
    
    'apunte al HABER de las cuentas de la familia
    'y un unico apunte al DEBE por el total a la nueva cuenta generica de compras para tratamientos.

    'miSQL = "Select campo1,campo2,nomfamia, ctaventa,ctavtaser,ctavent1,ctavtaseralt,sum(importe1) as impor "
    miSQL = "Select campo1,nomfamia,ctacompr ,ctacomprser,sum(importe1) as impor "
    miSQL = miSQL & " from tmpinformes,sfamia where tmpinformes.campo1=sfamia.codfamia AND codusu = " & vUsu.codigo & " group by campo1"

    
    
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    cadSelect = ""
    ImporteTotal = 0
    
    While Not miRsAux.EOF
        Me.lblIndicador(1).Caption = "Linea: " & NumRegElim
        Me.lblIndicador(1).Refresh
    
        numParam = 0
        If DBLet(miRsAux!Impor, "N") <> 0 Then
'            'Si utiliza cuentas o cuenta alternativa  y es distinta la de servicios de la normal
'            If miRsAux!campo2 = 1 Then
'                'ALTERNATIVA
'                If miRsAux!ctavent1 <> miRsAux!ctavtaseralt Then
'                    cadNomRPT = miRsAux!ctavtaseralt
'                    cadPDFrpt = miRsAux!ctavent1
'                    numParam = 1
'                End If
'            Else
'                'Cuentas normales, vaos, las no alternativas
'                If miRsAux!ctaventa <> miRsAux!ctavtaser Then
'                    cadNomRPT = miRsAux!ctavtaser
'                    cadPDFrpt = miRsAux!ctaventa
'                    numParam = 1
'                End If
'
'            End If
            cadNomRPT = miRsAux!ctacompr
            cadPDFrpt = DBLet(miRsAux!ctacomprser, "T")
            If cadPDFrpt <> "" Then
                If cadNomRPT <> cadPDFrpt Then numParam = 1
            End If

        End If
    
        If numParam = 1 Then
    
            NumRegElim = NumRegElim + 1
            
            ImporteTotal = ImporteTotal + miRsAux!Impor
            
            'linliapu,codmacta,numdocum,codconce,
            miSQL = NumRegElim & ",'" & cadNomRPT & "','Fam:" & Format(miRsAux!campo1, "0000") & "'," & numParam
            
            'ampconce,timporteD,timporteH,
            miSQL = miSQL & "," & DBSet(miRsAux!nomfamia, "T") & ",NULL," & DBSet(miRsAux!Impor, "N") & ","
            
            'ctacontr,idcontab,punteada,traspasado)
            miSQL = miSQL & DBSet(cadPDFrpt, "T") & ",'CONTAB',0,0)"
            miSQL = cadFormula & miSQL
            cadSelect = cadSelect & miSQL
            
            'Contra apunte
            NumRegElim = NumRegElim + 1
            'linliapu,codmacta,numdocum,codconce,
            miSQL = NumRegElim & ",'" & cadPDFrpt & "','Fam:" & Format(miRsAux!campo1, "0000") & "'," & numParam
            
            'ampconce,timporteD,timporteH,
            miSQL = miSQL & "," & DBSet(miRsAux!nomfamia, "T") & "," & DBSet(miRsAux!Impor, "N") & ",NULL,"
            
            'ctacontr,idcontab,punteada,traspasado)
            miSQL = miSQL & DBSet(cadNomRPT, "T") & ",'CONTAB',0,0)"
            miSQL = cadFormula & miSQL
            cadSelect = cadSelect & miSQL
            
            
            
            If Len(cadSelect) > 20000 Then
                miSQL = Mid(cadSelect, 2)
                miSQL = cadTitulo & miSQL
                ConnConta.Execute miSQL
                cadSelect = ""
            End If
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    If cadSelect <> "" Then
            miSQL = Mid(cadSelect, 2)
            miSQL = cadTitulo & miSQL
            ConnConta.Execute miSQL
    End If
    
    
    
    
    If NumRegElim > 0 Then
        'Debe ser >0 SIEMPRE
        
        miSQL = "El asiento esta en la introduccion de apuntes." & vbCrLf & vbCrLf & "Fecha: " & Format(Now, "dd/mm/yyyy") & vbCrLf & "Número: " & Mc.Contador
        MsgBox miSQL, vbInformation
        Me.lblIndicador(1).Caption = "Proceso finalizado"
    Else
        miSQL = "DELETE FROM cabapu where numasien=" & Mc.Contador & " and fechaent =" & DBSet(Now, "F")
        ConnConta.Execute miSQL
        
    End If
        
    Set miRsAux = Nothing
    
eGenerarApunteAjusteTratamientos:
    If Err.Number <> 0 Then
        MuestraError Err.Number
        
    End If
    Set miRsAux = Nothing
End Sub



'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
'  MODIFICAR PROVEEDOR EN PEDIDO DESPUES DE UNA SIMULACION
'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
Private Sub CargaLineasPedidoProveedor()
    Set miRsAux = New ADODB.Recordset
    
    miSQL = "select tmpslipreu.*,artvario from tmpslipreu,sartic where tmpslipreu.codartic=sartic.codartic and codusu = " & vUsu.codigo & " ORDER BY numlinea"
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    lw(2).ListItems.Clear
    While Not miRsAux.EOF
        lw(2).ListItems.Add , , miRsAux!codArtic
        NumRegElim = NumRegElim + 1
        lw(2).ListItems(NumRegElim).SubItems(1) = miRsAux!NomArtic
        lw(2).ListItems(NumRegElim).Checked = True
        
        
        If miRsAux!codAlmac = 1 Then
            miSQL = "Descuentos"
        ElseIf miRsAux!codAlmac = 2 Then
            miSQL = "Precio"
        Else
            If miRsAux!artvario = 1 Then
                miSQL = "Art. varios"
            Else
                miSQL = "Igual que original"
            End If
        End If
        lw(2).ListItems(NumRegElim).SubItems(6) = miSQL
        lw(2).ListItems(NumRegElim).Tag = miRsAux!numlinea
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
End Sub



'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
'  FITO CAMPOS
'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------

Private Function GenerarFitoCampos() As Boolean
Dim RN As ADODB.Recordset

    On Error GoTo eGenerarFitoCampos
    GenerarFitoCampos = False
    
    Set miRsAux = New ADODB.Recordset
    Set RN = New ADODB.Recordset
    
    Me.lblIndicador(2).Caption = "Obteniendo ventas articulo"
    Me.lblIndicador(2).Refresh
    miSQL = "DELETE FROM tmpinformes WHERE codusu = " & vUsu.codigo
    conn.Execute miSQL
    
    'Tipo factura
    cadPDFrpt = ""
    cadTitulo = ""
    For numParam = 0 To 2
        If chkCodtipom(numParam).Value = 1 Then
            cadPDFrpt = cadPDFrpt & ", '" & Me.chkCodtipom(numParam).Tag & "'"
            cadTitulo = cadTitulo & "- " & Me.chkCodtipom(numParam).Caption
        End If
    Next numParam
    
    If cadPDFrpt <> "" Then
        cadPDFrpt = " AND slifaccampos.codtipom IN (" & Mid(cadPDFrpt, 2) & ")"
        cadTitulo = "Facturas: " & Mid(cadTitulo, 2)
    End If
    numParam = 0
    

    miSQL = " slifaccampos.codtipom = scafac1.codtipom"
    miSQL = miSQL & " and slifaccampos.numfactu = scafac1.numfactu and slifaccampos.fecfactu = scafac1.fecfactu"
    miSQL = miSQL & " and slifaccampos.codtipoa = scafac1.codtipoa and slifaccampos.numalbar = scafac1.numalbar"
    miSQL = miSQL & " and scafac.codtipom = scafac1.codtipom and scafac.numfactu = scafac1.numfactu"
    miSQL = miSQL & "  and scafac.fecfactu = scafac1.fecfactu"
    
    
    If cadPDFrpt <> "" Then miSQL = miSQL & cadPDFrpt
  

    
    If txtFecha(6).Text <> "" Or txtFecha(7).Text <> "" Then
        cadNomRPT = " Fecha: "
        If Not PonerDesdeHasta("{scafac.fecfactu}", "F", 6, 7, cadNomRPT) Then Exit Function
        If cadTitulo <> "" Then cadTitulo = cadTitulo & "              "
        cadTitulo = cadTitulo & cadNomRPT
        
        
    End If
    ' """ + chr(13) + """
    If txtCliente(2).Text <> "" Or txtCliente(3).Text <> "" Then
        cadPDFrpt = " Cliente: "
        If Not PonerDesdeHasta("{scafac.codclien}", "CLI", 2, 3, cadNomRPT) Then Exit Function
        If Len(cadTitulo) > 60 Then
            cadTitulo = cadTitulo & """ + chr(13) + """
        Else
            cadTitulo = cadTitulo & String(10, " ")
        End If
        cadTitulo = Trim(cadTitulo & cadNomRPT)
        
    End If
    CadParam = CadParam & "DesdeHasta=""" & cadTitulo & """|"
    numParam = numParam + 1
    
    
    
       
    
    cadTitulo = " FROM slifaccampos ,scafac1,scafac WHERE " & miSQL
    If cadSelect <> "" Then
        cadSelect = Replace(cadSelect, "{", "")
        cadSelect = Replace(cadSelect, "}", "")
        cadSelect = " AND " & cadSelect
    End If
    cadTitulo = cadTitulo & cadSelect
    'Me lo guarda:
    CadenaDesdeOtroForm = cadTitulo
    
    
    cadTitulo = "Select slifaccampos.*,codclien,fechaalb " & cadTitulo & " ORDER BY codclien"
     
    miRsAux.Open cadTitulo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    cadTitulo = ""
    While Not miRsAux.EOF
        NumRegElim = NumRegElim + 1
        If NumRegElim > 65535 Then NumRegElim = 1
        
        '                   codclien  campo secuenc fra     nomvari  partida  campook cant   artfito fecfac  fecalb
        'tmpinformes(codusu,codigo1,campo1,campo2,nombre1,nombre2,nombre3,porcen1,importe1,obser,fecha1 ,fecha2,)
        
        miSQL = ", (" & vUsu.codigo & "," & miRsAux!codClien & "," & miRsAux!codCampo & "," & NumRegElim & ",'"
        miSQL = miSQL & Mid(miRsAux!codtipom & "   ", 1, 3) & Format(miRsAux!Numfactu, "000000") & "',"
        'La tengo guardada en las lineas del campo
        If IsNull(miRsAux!nomvarie) Or IsNull(miRsAux!nompartida) Then
            'Lo cojo de ARIAGRO
            cadNomRPT = "NULL,NULL,1"
        Else
            cadNomRPT = DBSet(miRsAux!nomvarie, "T") & "," & DBSet(miRsAux!nompartida, "T") & ",0"
        End If
        miSQL = miSQL & cadNomRPT & ",0,NULL,"
        miSQL = miSQL & DBSet(miRsAux!FecFactu, "F") & "," & DBSet(miRsAux!FechaAlb, "F") & ")"
        
        cadTitulo = cadTitulo & miSQL
        
        If Len(cadTitulo) > 30000 Then
            cadTitulo = Mid(cadTitulo, 2)
            miSQL = "INSERT INTO tmpinformes(codusu,codigo1,campo1,campo2,nombre1,nombre2,nombre3,porcen1,importe1,obser,fecha1 ,fecha2) VALUES  " & cadTitulo
            conn.Execute miSQL
            cadTitulo = ""
        End If
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
     
    If cadTitulo <> "" Then
        cadTitulo = Mid(cadTitulo, 2)
        miSQL = "INSERT INTO tmpinformes(codusu,codigo1,campo1,campo2,nombre1,nombre2,nombre3,porcen1,importe1,obser,fecha1 ,fecha2) VALUES  " & cadTitulo
        conn.Execute miSQL
        cadTitulo = ""
    End If
     
    If NumRegElim = 0 Then
        MsgBox "Ningun dato generado", vbExclamation
        GoTo eGenerarFitoCampos
    End If
     
     
     
    'Ajustamos los campos que no tenemos NOmvarie nompartida
    miSQL = "Select distinct(campo1) from tmpinformes where codusu =" & vUsu.codigo & " AND porcen1<>0 ORDER BY campo1"
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    cadTitulo = ""
    miSQL = ""
    NumRegElim = 0
    While Not miRsAux.EOF
        miSQL = miSQL & ", " & miRsAux!campo1
        NumRegElim = NumRegElim + 1
        miRsAux.MoveNext
        If NumRegElim > 30 Then
            cadTitulo = cadTitulo & Mid(miSQL, 2) & "|"
            miSQL = ""
            NumRegElim = 0
        End If
    Wend
    miRsAux.Close

    If miSQL <> "" Then cadTitulo = cadTitulo & Mid(miSQL, 2) & "|"
    
    cadNomRPT = "select rcampos.codcampo, rpartida.nomparti, variedades.nomvarie,rcampos.codclien,rsocios.codsocio,rsocios.nomsocio,rcampos.codsitua"
    cadNomRPT = cadNomRPT & " from ((@#rcampos inner join @#rpartida on rcampos.codparti = rpartida.codparti)"
    cadNomRPT = cadNomRPT & " inner join @#variedades on rcampos.codvarie = variedades.codvarie)"
    cadNomRPT = cadNomRPT & " inner join @#rsocios on rsocios.codsocio=rcampos.codsocio"
    
    cadNomRPT = Replace(cadNomRPT, "@#", vParamAplic.Ariagro & ".") & " AND rcampos.codcampo IN ("
    
    
    
    Do
        NumRegElim = InStr(1, cadTitulo, "|")
        If NumRegElim > 0 Then
            Me.lblIndicador(2).Caption = "Campos " & Len(cadTitulo)
            Me.lblIndicador(2).Refresh
    
            miSQL = Mid(cadTitulo, 1, NumRegElim - 1)
            cadTitulo = Mid(cadTitulo, NumRegElim + 1)
            
            miSQL = cadNomRPT & miSQL & ")"
            miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
            miSQL = "UPDATE tmpinformes set nombre2=@@ , nombre3=## , porcen1=0 WHERE codusu = " & vUsu.codigo & " AND campo1 = "
            While Not miRsAux.EOF
                cadPDFrpt = Replace(miSQL, "@@", DBSet(miRsAux!nomparti, "T"))
                cadPDFrpt = Replace(cadPDFrpt, "##", DBSet(miRsAux!nomvarie, "T"))
                cadPDFrpt = cadPDFrpt & miRsAux!codCampo
                conn.Execute cadPDFrpt
                miRsAux.MoveNext
            Wend
            miRsAux.Close
        Else
            cadTitulo = ""
        End If
       
    Loop Until cadTitulo = ""
    
    
    Me.lblIndicador(2).Caption = "Lineas fitosanitarios"
    Me.lblIndicador(2).Refresh
    
    miSQL = Replace(CadenaDesdeOtroForm, "slifaccampos", "slifac")
    miSQL = miSQL & " and numlote<>''"
    miSQL = "Select slifac.*   " & miSQL
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    cadTitulo = ""
    While Not miRsAux.EOF
        cadPDFrpt = Mid(miRsAux!codtipom & "   ", 1, 3) & Format(miRsAux!Numfactu, "000000")
        If cadTitulo <> cadPDFrpt Then
            'Vemos si hay que updatear
            If cadTitulo <> "" Then
                Me.lblIndicador(2).Caption = "Fra: " & cadTitulo
                Me.lblIndicador(2).Refresh
                miSQL = "UPDATE tmpinformes set obser = " & DBSet(cadNomRPT, "T")
                miSQL = miSQL & " WHERE codusu = " & vUsu.codigo & " AND nombre1= '" & cadTitulo & "'"
                conn.Execute miSQL
            End If
            cadTitulo = cadPDFrpt
            cadNomRPT = ""
        End If
        
        vMostrarTree = True
        If cadNomRPT <> "" Then
            If Len(cadNomRPT) > 240 Then
            
                vMostrarTree = False
            Else
                cadNomRPT = cadNomRPT & vbCrLf
            End If
        
        End If
        If vMostrarTree Then cadNomRPT = cadNomRPT & miRsAux!NomArtic & "(" & miRsAux!cantidad & ")"


        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    GenerarFitoCampos = True
    
    
    
    
eGenerarFitoCampos:
    If Err.Number <> 0 Then
        MuestraError Err.Number, , Err.Description
        GenerarFitoCampos = False
    End If
    Set miRsAux = Nothing
    Set RN = Nothing
End Function




Private Sub LeerGuardarSeleccionProveedoresAQuitar(Leer As Boolean)
Dim NF As Integer

    On Error GoTo eLeerGuardarSeleccionProveedoresAQuitar
    
    miSQL = App.Path & "\provquit.dat"
    If Leer Then
        auxiliar = ""
        If Dir(miSQL, vbArchive) <> "" Then
            NF = FreeFile
            Open miSQL For Input As #NF
            Line Input #NF, auxiliar
            Close #NF
            

        End If
        If auxiliar = "" Then auxiliar = "|"
    Else
    
        '1.- En auxiliar tenemos los que QUITA
        '2.- Comprobaremos que de los que quita de normal, no lo ha vuelto a seleccionar
        '3.- Comprobaremos que de los que ha quitado estan en auxiliar
        
        
        For NumRegElim = 1 To lw(1).ListItems.Count
            If Not lw(1).ListItems(NumRegElim).Checked Then
                cadSelect = "|" & lw(1).ListItems(NumRegElim).Text & "|"
                If InStr(1, auxiliar, cadSelect) > 0 Then
                    'YA esta
                Else
                    auxiliar = auxiliar & lw(1).ListItems(NumRegElim).Text & "|"
                End If
            End If
        Next
        
        For NumRegElim = 1 To lw(1).ListItems.Count
            If lw(1).ListItems(NumRegElim).Checked Then
                'Vemos si de los que no seelccionaba lo ha vuelto a marcar
                cadSelect = "|" & lw(1).ListItems(NumRegElim).Text & "|"
                NF = InStr(1, auxiliar, cadSelect)
                If NF > 0 Then
                    
                    cadTitulo = Mid(auxiliar, 1, NF)    'Dejo el pipe con esta
                    NF = InStr(NF + 1, auxiliar, "|")
                    If NF > 0 Then
                        cadPDFrpt = Mid(auxiliar, NF + 1) 'quito el pipe
                    Else
                        cadPDFrpt = ""
                    End If
                    auxiliar = cadTitulo & cadPDFrpt
                End If
            End If
        Next
        
    
    
    
    
    
            NF = FreeFile
            Open miSQL For Output As #NF
            Print #NF, auxiliar
            Close #NF
    

    End If


    Exit Sub
eLeerGuardarSeleccionProveedoresAQuitar:
    MuestraError Err.Number, Err.Description
    auxiliar = ""
End Sub




Private Sub CargaAlbaranesFacturaClienteEuler()
    On Error GoTo eCargaAlbaranesFacturaClienteEuler
    
    lw(3).ListItems.Clear
    miSQL = " select scaalb.codtipom,scaalb.numalbar,fechaalb,sum(importel),-1 from scaalb,slialb"
    miSQL = miSQL & " where scaalb.codtipom = slialb.codtipom And scaalb.NumAlbar = slialb.NumAlbar"
    miSQL = miSQL & " and codclien=" & OtrosDatos & " group by 1,2"
    miSQL = miSQL & " Union"
    miSQL = miSQL & " select scafac1.codtipoa,numalbar,fechaalb,brutofac,scafac.NumFactu  from scafac, scafac1 where"
    miSQL = miSQL & " scafac.codtipom = scafac1.codtipom And scafac.NumFactu = scafac1.NumFactu And "
    miSQL = miSQL & " scafac.FecFactu = scafac1.FecFactu and codclien=" & OtrosDatos
    miSQL = miSQL & " order by 1,3 desc,2"
    
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    While Not miRsAux.EOF
        
        lw(3).ListItems.Add , , miRsAux.Fields(0)
        NumRegElim = NumRegElim + 1
        lw(3).ListItems(NumRegElim).SubItems(1) = Format(miRsAux.Fields(1), "0000000")
        lw(3).ListItems(NumRegElim).SubItems(2) = Format(miRsAux.Fields(2), "dd/mm/yyyy")
        lw(3).ListItems(NumRegElim).SubItems(3) = Format(miRsAux.Fields(3), FormatoImporte)
        
        'lw(3).ListItems(NumRegElim).SubItems(4) = IIf(miRsAux.Fields(4) = 1, "Si", " ")
        If miRsAux.Fields(4) = -1 Then
            lw(3).ListItems(NumRegElim).SubItems(4) = " "
        Else
            lw(3).ListItems(NumRegElim).SubItems(4) = Format(miRsAux.Fields(4), "000000")
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    CadenaDesdeOtroForm = ""
eCargaAlbaranesFacturaClienteEuler:
    If Err.Number <> 0 Then MuestraError Err.Number
    Set miRsAux = Nothing
End Sub




Private Sub CargarFacturasVentaCliente()
Dim EnAlbaranes As Boolean
   
    'El recordset lo hemos abierto en form albaranes
    'Para que si no hay resultados, NO abra el este form para no mostrar res
    'Set miRsAux = New ADODB.Recordset
    'miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    'Si es articulo de varios
    Label4(9).Caption = ""
    
    'Articulo de varios
    CadParam = "artvario"
    miSQL = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", RecuperaValor(Me.OtrosDatos, 1), "T", CadParam)
    Label4(9).Caption = miRsAux!codArtic & " - " & miSQL
    conSubRPT = CadParam = "1"
    
    
    '----------------------------------------------------------------------------
    '----------------------------------------------------------------------------
    ' Primero las facturas,y para herbelca vamos tambien a albaranes
    '----------------------------------------------------------------------------
    '----------------------------------------------------------------------------
    NumRegElim = 0
    lw(4).ListItems.Clear
    
    
    
    
    If InStr(1, miRsAux.Source, "slifac") > 0 Then
    
        
        While Not miRsAux.EOF
            lw(4).ListItems.Add , , miRsAux!codtipom
            NumRegElim = NumRegElim + 1
            lw(4).ListItems(NumRegElim).SubItems(1) = Format(miRsAux!Numfactu, "00000")
            lw(4).ListItems(NumRegElim).SubItems(2) = Format(miRsAux!FecFactu, "dd/mm/yyyy")
            
            lw(4).ListItems(NumRegElim).SubItems(3) = Format(miRsAux!precioar, FormatoPrecio)
            lw(4).ListItems(NumRegElim).SubItems(4) = DBLet(miRsAux!origpre, "T") & " "
            
            
            lw(4).ListItems(NumRegElim).SubItems(5) = Format(miRsAux!cantidad, FormatoCantidad)
            miSQL = " "
            If miRsAux!dtoline1 > 0 Then miSQL = Format(miRsAux!dtoline1, FormatoDescuento)
            lw(4).ListItems(NumRegElim).SubItems(6) = miSQL
            miSQL = " "
            If miRsAux!dtoline2 > 0 Then miSQL = Format(miRsAux!dtoline2, FormatoDescuento)
            lw(4).ListItems(NumRegElim).SubItems(7) = miSQL
        
            lw(4).ListItems(NumRegElim).SubItems(8) = Format(miRsAux!ImporteL, FormatoImporte)
            
            
            'Si es articulo de varios lo especifico
            miSQL = " "
            If conSubRPT Then miSQL = miSQL & miRsAux!NomArtic
            lw(4).ListItems(NumRegElim).SubItems(9) = miSQL
            
            'codtipom numfactu fecfactu codtipoa numalbar numlinea
            miSQL = "codtipoa = " & DBSet(miRsAux!codtipoa, "T") & " AND numalbar = " & miRsAux!NumAlbar & " AND numlinea =" & miRsAux!numlinea
            lw(4).ListItems(NumRegElim).Tag = miSQL
            
            'Si es negativo:
            If miRsAux!cantidad < 0 Then
                lw(4).ListItems(NumRegElim).ForeColor = vbRed
                For numParam = 1 To lw(4).ColumnHeaders.Count - 1
                    lw(4).ListItems(NumRegElim).ListSubItems(numParam).ForeColor = vbRed
                Next
                lw(4).ListItems(NumRegElim).Tag = ""   'Sera cantidad negativa. Para que cuando lo seleccionen, no lo pueda devovler
            End If
            miRsAux.MoveNext
        Wend
                
                
        If vParamAplic.NumeroInstalacion = 2 Then
            'Vamos a abrir albaranes
            miRsAux.Close
                    
            numParam = InStr(1, miRsAux.Source, " codclien =")
            cadFormula = Mid(miRsAux.Source, numParam + 11)
            numParam = InStr(1, LCase(cadFormula), " and ")
            cadFormula = Mid(cadFormula, 1, numParam)
            
            miSQL = " Select slialb.*,fechaalb from slialb,scaalb     where  scaalb.codtipom=slialb.codtipom and"
            miSQL = miSQL & " scaalb.NumAlbar = slialb.NumAlbar  AND codclien = " & cadFormula
            miSQL = miSQL & " AND codartic<>" & DBSet(vParamAplic.ArtReciclado, "T")
            miSQL = miSQL & " AND scaalb.codtipom <>'ALZ' "    'para quitar los que no sean albaranes
            miSQL = miSQL & " AND codartic = " & DBSet(RecuperaValor(Me.OtrosDatos, 1), "T")
            miSQL = miSQL & " ORDER BY fechaalb,numlinea"
            miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
            EnAlbaranes = True
        End If
    Else
        EnAlbaranes = True
    End If
    
    If EnAlbaranes Then
        
        
        While Not miRsAux.EOF
            lw(4).ListItems.Add , , miRsAux!codtipom
            NumRegElim = NumRegElim + 1
            lw(4).ListItems(NumRegElim).SubItems(1) = Format(miRsAux!NumAlbar, "00000")
            lw(4).ListItems(NumRegElim).SubItems(2) = Format(miRsAux!FechaAlb, "dd/mm/yyyy")
            
            lw(4).ListItems(NumRegElim).SubItems(3) = Format(miRsAux!precioar, FormatoPrecio)
            lw(4).ListItems(NumRegElim).SubItems(4) = DBLet(miRsAux!origpre, "T") & " "
            
            
            lw(4).ListItems(NumRegElim).SubItems(5) = Format(miRsAux!cantidad, FormatoCantidad)
            miSQL = " "
            If miRsAux!dtoline1 > 0 Then miSQL = Format(miRsAux!dtoline1, FormatoDescuento)
            lw(4).ListItems(NumRegElim).SubItems(6) = miSQL
            miSQL = " "
            If miRsAux!dtoline2 > 0 Then miSQL = Format(miRsAux!dtoline2, FormatoDescuento)
            lw(4).ListItems(NumRegElim).SubItems(7) = miSQL
        
            lw(4).ListItems(NumRegElim).SubItems(8) = Format(miRsAux!ImporteL, FormatoImporte)
            
            
            'Si es articulo de varios lo especifico
            miSQL = " "
            If conSubRPT Then miSQL = miSQL & miRsAux!NomArtic
            lw(4).ListItems(NumRegElim).SubItems(9) = miSQL
            
            'codtipom numfactu fecfactu codtipoa numalbar numlinea
            miSQL = "codtipom = " & DBSet(miRsAux!codtipom, "T") & " AND numalbar = " & miRsAux!NumAlbar & " AND numlinea =" & miRsAux!numlinea
            lw(4).ListItems(NumRegElim).Tag = miSQL
            
            'Si es negativo:
            If miRsAux!cantidad < 0 Then
                lw(4).ListItems(NumRegElim).ForeColor = vbRed
                For numParam = 1 To lw(4).ColumnHeaders.Count - 1
                    lw(4).ListItems(NumRegElim).ListSubItems(numParam).ForeColor = vbRed
                Next
                lw(4).ListItems(NumRegElim).Tag = ""   'Sera cantidad negativa. Para que cuando lo seleccionen, no lo pueda devovler
            End If
            miRsAux.MoveNext
        Wend
        
   
    End If
    Set lw(4).SelectedItem = lw(4).ListItems(1)
    DoEvents
    lw(4).SetFocus
    
    
    
    numParam = 0
End Sub


