VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmVariosNew 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Varios"
   ClientHeight    =   9195
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11565
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9195
   ScaleWidth      =   11565
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameHcoAlmacen 
      Height          =   8100
      Left            =   0
      TabIndex        =   147
      Top             =   0
      Visible         =   0   'False
      Width           =   11415
      Begin VB.CheckBox Check1 
         Caption         =   "�ltimo mes"
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
         Left            =   9585
         TabIndex        =   154
         Top             =   315
         Width           =   1635
      End
      Begin VB.CommandButton cmdCancelar 
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
         Index           =   6
         Left            =   10125
         TabIndex        =   149
         Top             =   7305
         Width           =   1095
      End
      Begin VB.CommandButton CmdAcepHcoAlm 
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
         Left            =   8925
         TabIndex        =   148
         Top             =   7305
         Width           =   1095
      End
      Begin MSComctlLib.ListView ListView6 
         Height          =   2970
         Left            =   135
         TabIndex        =   152
         Top             =   720
         Width           =   11130
         _ExtentX        =   19632
         _ExtentY        =   5239
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "N�mero"
            Object.Width           =   2611
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Fecha"
            Object.Width           =   1799
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Nombre"
            Object.Width           =   7955
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Plazos"
            Object.Width           =   1270
         EndProperty
      End
      Begin MSComctlLib.ListView ListView7 
         Height          =   2970
         Left            =   135
         TabIndex        =   153
         Top             =   3870
         Width           =   11130
         _ExtentX        =   19632
         _ExtentY        =   5239
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         Enabled         =   0   'False
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
            Text            =   "Telefono"
            Object.Width           =   2611
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Cliente"
            Object.Width           =   1799
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Nombre"
            Object.Width           =   7955
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Plazos"
            Object.Width           =   1270
         EndProperty
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Seleccionar Movimientos Almac�n"
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
         Left            =   135
         TabIndex        =   151
         Top             =   240
         Width           =   4875
      End
      Begin VB.Label lblInd 
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   180
         TabIndex        =   150
         Top             =   7335
         Width           =   7350
      End
   End
   Begin VB.Frame FrameVtaPlazosBorrar 
      Height          =   7575
      Left            =   1320
      TabIndex        =   129
      Top             =   0
      Visible         =   0   'False
      Width           =   8655
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cerrar"
         Height          =   375
         Index           =   16
         Left            =   7200
         TabIndex        =   131
         Top             =   6960
         Width           =   1215
      End
      Begin VB.CommandButton cmdBorrarVtaPlazosFinalizadas 
         Caption         =   "Borrar"
         Height          =   375
         Left            =   5880
         TabIndex        =   130
         Top             =   6960
         Width           =   1215
      End
      Begin MSComctlLib.ListView ListView5 
         Height          =   5895
         Left            =   240
         TabIndex        =   132
         Top             =   840
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   10398
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Telefono"
            Object.Width           =   2611
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Cliente"
            Object.Width           =   1799
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Nombre"
            Object.Width           =   7955
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Plazos"
            Object.Width           =   1270
         EndProperty
      End
      Begin VB.Label lblInd 
         Caption         =   "lblInd"
         Height          =   255
         Index           =   2
         Left            =   1320
         TabIndex        =   134
         Top             =   7080
         Width           =   2895
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   5
         Left            =   720
         Picture         =   "frmVariosNew.frx":0000
         ToolTipText     =   "Puntear al haber"
         Top             =   7080
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   4
         Left            =   360
         Picture         =   "frmVariosNew.frx":014A
         ToolTipText     =   "Quitar al haber"
         Top             =   7080
         Width           =   240
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Borrar venta plazos finalizadas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   345
         Index           =   16
         Left            =   240
         TabIndex        =   133
         Top             =   360
         Width           =   6795
      End
   End
   Begin VB.Frame FrameComprasAnyo 
      Height          =   2655
      Left            =   480
      TabIndex        =   91
      Top             =   120
      Visible         =   0   'False
      Width           =   6735
      Begin VB.CommandButton cmdComprasMeses 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3960
         TabIndex        =   94
         Top             =   2040
         Width           =   1095
      End
      Begin VB.TextBox txtNumero 
         Height          =   285
         Index           =   0
         Left            =   1440
         TabIndex        =   92
         Text            =   "Text1"
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox txtProv 
         Height          =   285
         Index           =   0
         Left            =   1440
         TabIndex        =   93
         Text            =   "Text1"
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox txtProvD 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   0
         Left            =   2400
         TabIndex        =   96
         Text            =   "Text1"
         Top             =   1560
         Width           =   3975
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   11
         Left            =   5160
         TabIndex        =   95
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "A�o"
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
         Left            =   240
         TabIndex        =   99
         Top             =   1080
         Width           =   330
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Compras por meses"
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
         Left            =   1440
         TabIndex        =   98
         Top             =   360
         Width           =   4635
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Proveedor"
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
         Left            =   240
         TabIndex        =   97
         Top             =   1560
         Width           =   885
      End
      Begin VB.Image imgProv 
         Height          =   240
         Index           =   0
         Left            =   1200
         Top             =   1560
         Width           =   240
      End
   End
   Begin VB.Frame FrameEstadisticasConsultas 
      Height          =   3855
      Left            =   120
      TabIndex        =   25
      Top             =   120
      Visible         =   0   'False
      Width           =   6855
      Begin VB.CommandButton cmdListConsultaPedido 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4200
         TabIndex        =   30
         Top             =   3120
         Width           =   1095
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   1
         Left            =   4320
         TabIndex        =   29
         Text            =   "Text1"
         Top             =   2520
         Width           =   1095
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   0
         Left            =   1800
         TabIndex        =   28
         Text            =   "Text1"
         Top             =   2520
         Width           =   1095
      End
      Begin VB.TextBox txtArt 
         Height          =   285
         Index           =   3
         Left            =   1560
         TabIndex        =   27
         Text            =   "Text1"
         Top             =   1560
         Width           =   1455
      End
      Begin VB.TextBox txtArtD 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   3
         Left            =   3000
         TabIndex        =   36
         Text            =   "Text1"
         Top             =   1560
         Width           =   3735
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Salir"
         Height          =   375
         Index           =   2
         Left            =   5400
         TabIndex        =   31
         Top             =   3120
         Width           =   1095
      End
      Begin VB.TextBox txtArt 
         Height          =   285
         Index           =   2
         Left            =   1560
         TabIndex        =   26
         Text            =   "Text1"
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox txtArtD 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   2
         Left            =   3000
         TabIndex        =   32
         Text            =   "Text1"
         Top             =   1200
         Width           =   3735
      End
      Begin VB.Image imgFecha 
         Height          =   255
         Index           =   1
         Left            =   3720
         Top             =   2520
         Width           =   255
      End
      Begin VB.Image imgFecha 
         Height          =   255
         Index           =   0
         Left            =   1200
         Top             =   2520
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   5
         Left            =   3240
         TabIndex        =   40
         Top             =   2520
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Desde"
         Height          =   255
         Index           =   4
         Left            =   600
         TabIndex        =   39
         Top             =   2520
         Width           =   615
      End
      Begin VB.Label Label4 
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
         Index           =   1
         Left            =   240
         TabIndex        =   38
         Top             =   2160
         Width           =   495
      End
      Begin VB.Image imgArticulo 
         Height          =   240
         Index           =   3
         Left            =   1200
         Top             =   1560
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   3
         Left            =   480
         TabIndex        =   37
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Estad�sticas consultas art�culo / cliente"
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
         Left            =   360
         TabIndex        =   35
         Top             =   240
         Width           =   6195
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Articulos"
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
         TabIndex        =   34
         Top             =   840
         Width           =   750
      End
      Begin VB.Image imgArticulo 
         Height          =   240
         Index           =   2
         Left            =   1200
         Top             =   1200
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Desde"
         Height          =   255
         Index           =   2
         Left            =   480
         TabIndex        =   33
         Top             =   1200
         Width           =   615
      End
   End
   Begin VB.Frame FrameImpresionFacturasDirectas 
      Height          =   2295
      Left            =   120
      TabIndex        =   8
      Top             =   0
      Width           =   5535
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Salir"
         Height          =   375
         Index           =   0
         Left            =   4200
         TabIndex        =   11
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label lblImpr 
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   5175
      End
      Begin VB.Label lblImpr 
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   10
         Top             =   1080
         Width           =   3975
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Impresi�n facturas"
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
         TabIndex        =   9
         Top             =   240
         Width           =   2955
      End
   End
   Begin VB.Frame FrameFormaEnvio 
      Height          =   3495
      Left            =   120
      TabIndex        =   41
      Top             =   0
      Width           =   6135
      Begin VB.ListBox ListEnvio 
         Height          =   1815
         Left            =   240
         TabIndex        =   44
         Top             =   960
         Width           =   5655
      End
      Begin VB.CommandButton cmdFormaEnvio 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   4560
         TabIndex        =   43
         Top             =   3000
         Width           =   1335
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Forma de envio"
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
         Left            =   1680
         TabIndex        =   42
         Top             =   360
         Width           =   2835
      End
   End
   Begin VB.Frame FrameListArticulos 
      Height          =   6855
      Left            =   120
      TabIndex        =   21
      Top             =   0
      Visible         =   0   'False
      Width           =   6855
      Begin MSComctlLib.ListView ListView1 
         Height          =   5655
         Left            =   120
         TabIndex        =   23
         Top             =   600
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   9975
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.CommandButton cmdAccionListview 
         Caption         =   "Elimin&ar"
         Height          =   375
         Left            =   4080
         TabIndex        =   6
         Top             =   6360
         Width           =   1215
      End
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "Cerrar"
         Height          =   375
         Left            =   5520
         TabIndex        =   7
         Top             =   6360
         Width           =   1095
      End
      Begin VB.Label lblElim 
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   24
         Top             =   6480
         Width           =   2415
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   0
         Left            =   240
         Picture         =   "frmVariosNew.frx":0294
         ToolTipText     =   "Quitar al haber"
         Top             =   6480
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   1
         Left            =   600
         Picture         =   "frmVariosNew.frx":03DE
         ToolTipText     =   "Puntear al haber"
         Top             =   6480
         Width           =   240
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Eliminar art�culos"
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
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   2955
      End
   End
   Begin VB.Frame FrameSelClien 
      Height          =   6975
      Left            =   240
      TabIndex        =   45
      Top             =   120
      Visible         =   0   'False
      Width           =   7095
      Begin VB.CommandButton cmdClientes 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4080
         TabIndex        =   49
         Top             =   6480
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   4
         Left            =   5520
         TabIndex        =   48
         Top             =   6480
         Width           =   1095
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   5655
         Left            =   240
         TabIndex        =   47
         Top             =   720
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   9975
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
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
            Object.Width           =   2470
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   6703
         EndProperty
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   3
         Left            =   720
         Picture         =   "frmVariosNew.frx":0528
         ToolTipText     =   "Puntear al haber"
         Top             =   6480
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   2
         Left            =   240
         Picture         =   "frmVariosNew.frx":0672
         ToolTipText     =   "Quitar al haber"
         Top             =   6480
         Width           =   240
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Seleccionar clientes"
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
         Left            =   360
         TabIndex        =   46
         Top             =   240
         Width           =   2955
      End
   End
   Begin VB.Frame FrRectifcadoStocks 
      Height          =   6375
      Left            =   120
      TabIndex        =   87
      Top             =   0
      Visible         =   0   'False
      Width           =   10575
      Begin VB.TextBox txtStock 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5055
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   89
         Text            =   "frmVariosNew.frx":07BC
         Top             =   720
         Width           =   10215
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Salir"
         Height          =   375
         Index           =   10
         Left            =   9240
         TabIndex        =   88
         Top             =   5880
         Width           =   1095
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Art�culos stock rectificado"
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
         Left            =   2160
         TabIndex        =   90
         Top             =   240
         Width           =   3765
      End
   End
   Begin VB.Frame FrameDtoProve 
      Height          =   2775
      Left            =   120
      TabIndex        =   75
      Top             =   0
      Visible         =   0   'False
      Width           =   6495
      Begin VB.CommandButton cmdActualiDtoProv 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3720
         TabIndex        =   78
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox txtDecimal 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   4680
         TabIndex        =   77
         Text            =   "Text1"
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox txtDecimal 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1680
         TabIndex        =   76
         Text            =   "Text1"
         Top             =   1560
         Width           =   855
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   9
         Left            =   5040
         TabIndex        =   79
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Descuento 2"
         Height          =   255
         Index           =   7
         Left            =   3600
         TabIndex        =   86
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Descuento 1"
         Height          =   255
         Index           =   6
         Left            =   600
         TabIndex        =   85
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Cliente"
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
         Left            =   1200
         TabIndex        =   84
         Top             =   1080
         Width           =   5175
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Pedido"
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
         TabIndex        =   83
         Top             =   1080
         Width           =   570
      End
      Begin VB.Label Label4 
         Caption         =   "Cliente"
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
         Left            =   1200
         TabIndex        =   82
         Top             =   720
         Width           =   5175
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Proveedor"
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
         TabIndex        =   81
         Top             =   720
         Width           =   885
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Actualizar descuentos pedido proveedor"
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
         Left            =   120
         TabIndex        =   80
         Top             =   240
         Width           =   6075
      End
   End
   Begin VB.Frame FrameDHArticulo 
      Height          =   3735
      Left            =   120
      TabIndex        =   13
      Top             =   0
      Width           =   6855
      Begin VB.TextBox txtProv 
         Height          =   285
         Index           =   2
         Left            =   1320
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   2520
         Width           =   975
      End
      Begin VB.TextBox txtProvD 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   2
         Left            =   2280
         TabIndex        =   114
         Text            =   "Text1"
         Top             =   2520
         Width           =   3975
      End
      Begin VB.TextBox txtProv 
         Height          =   285
         Index           =   1
         Left            =   1320
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox txtProvD 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   1
         Left            =   2280
         TabIndex        =   113
         Text            =   "Text1"
         Top             =   2160
         Width           =   3975
      End
      Begin VB.CommandButton cmdEliminarArticulos 
         Caption         =   "Buscar"
         Height          =   375
         Left            =   4080
         TabIndex        =   4
         Top             =   3120
         Width           =   1095
      End
      Begin VB.TextBox txtArtD 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   1
         Left            =   2880
         TabIndex        =   19
         Text            =   "Text1"
         Top             =   1440
         Width           =   3735
      End
      Begin VB.TextBox txtArt 
         Height          =   285
         Index           =   1
         Left            =   1320
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox txtArtD 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   0
         Left            =   2880
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   1080
         Width           =   3735
      End
      Begin VB.TextBox txtArt 
         Height          =   285
         Index           =   0
         Left            =   1320
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   1080
         Width           =   1455
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Salir"
         Height          =   375
         Index           =   1
         Left            =   5400
         TabIndex        =   5
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Proveedor"
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
         Left            =   120
         TabIndex        =   117
         Top             =   1800
         Width           =   885
      End
      Begin VB.Label Label1 
         Caption         =   "Desde"
         Height          =   255
         Index           =   9
         Left            =   360
         TabIndex        =   116
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   8
         Left            =   360
         TabIndex        =   115
         Top             =   2520
         Width           =   615
      End
      Begin VB.Image imgProv 
         Height          =   240
         Index           =   2
         Left            =   1080
         Top             =   2520
         Width           =   240
      End
      Begin VB.Image imgProv 
         Height          =   240
         Index           =   1
         Left            =   1080
         Top             =   2160
         Width           =   240
      End
      Begin VB.Label lblElim 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   3240
         Width           =   3855
      End
      Begin VB.Label Label1 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   18
         Top             =   1440
         Width           =   615
      End
      Begin VB.Image imgArticulo 
         Height          =   240
         Index           =   1
         Left            =   1080
         Top             =   1440
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Desde"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   16
         Top             =   1080
         Width           =   615
      End
      Begin VB.Image imgArticulo 
         Height          =   240
         Index           =   0
         Left            =   1080
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Articulos"
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
         Index           =   36
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   750
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Eliminar art�culos"
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
         Left            =   1440
         TabIndex        =   14
         Top             =   240
         Width           =   2955
      End
   End
   Begin VB.Frame FrameElimCambiarFecFactura 
      Height          =   2655
      Left            =   1680
      TabIndex        =   104
      Top             =   2160
      Visible         =   0   'False
      Width           =   5655
      Begin VB.CommandButton cmdCambiFecReestbFact 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   2760
         TabIndex        =   112
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Frame FrameNuevaFecFac 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   615
         Left            =   720
         TabIndex        =   109
         Top             =   1320
         Width           =   4215
         Begin VB.TextBox txtFecha 
            Height          =   285
            Index           =   3
            Left            =   1800
            TabIndex        =   110
            Text            =   "Text1"
            Top             =   240
            Width           =   1095
         End
         Begin VB.Image imgFecha 
            Height          =   255
            Index           =   3
            Left            =   1440
            Top             =   240
            Width           =   255
         End
         Begin VB.Label Label4 
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
            Index           =   11
            Left            =   600
            TabIndex        =   111
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.OptionButton optElimFact 
         Caption         =   "Eliminar - reest. albar�n"
         Height          =   195
         Index           =   1
         Left            =   3120
         TabIndex        =   108
         Top             =   960
         Width           =   2055
      End
      Begin VB.OptionButton optElimFact 
         Caption         =   "Cambiar fecha factura"
         Height          =   195
         Index           =   0
         Left            =   600
         TabIndex        =   107
         Top             =   960
         Value           =   -1  'True
         Width           =   2055
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Salir"
         Height          =   375
         Index           =   13
         Left            =   4080
         TabIndex        =   105
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Modificar fecha / Eliminar facturas"
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
         Left            =   120
         TabIndex        =   106
         Top             =   240
         Width           =   5475
      End
   End
   Begin VB.Frame FrameCambioTipoAlbaran 
      Height          =   4455
      Left            =   3000
      TabIndex        =   124
      Top             =   1080
      Visible         =   0   'False
      Width           =   5655
      Begin VB.CommandButton cmdCambiarTipoAlbaran 
         Caption         =   "&Cam&biar"
         Height          =   375
         Left            =   2760
         TabIndex        =   128
         Top             =   3720
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   3
         Left            =   4200
         TabIndex        =   125
         Top             =   3720
         Width           =   1215
      End
      Begin MSComctlLib.ListView ListView4 
         Height          =   2535
         Left            =   360
         TabIndex        =   126
         Top             =   960
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   4471
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Tipo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripcion"
            Object.Width           =   6068
         EndProperty
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Cambiar tipo de albar�n"
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
         Left            =   840
         TabIndex        =   127
         Top             =   360
         Width           =   4635
      End
   End
   Begin VB.Frame FramaTaxcoCambioKm 
      Height          =   4455
      Left            =   2640
      TabIndex        =   135
      Top             =   1200
      Visible         =   0   'False
      Width           =   6015
      Begin VB.TextBox txtNumero 
         Alignment       =   2  'Center
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
         Left            =   1800
         TabIndex        =   136
         Top             =   3240
         Width           =   1815
      End
      Begin VB.CommandButton cmdTaxcoEstabKm 
         Caption         =   "Establecer"
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
         Left            =   3000
         TabIndex        =   137
         Top             =   3840
         Width           =   1215
      End
      Begin VB.TextBox txtNoEditable 
         BackColor       =   &H00FFFFFF&
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
         Height          =   390
         Index           =   2
         Left            =   240
         TabIndex        =   145
         Top             =   2520
         Width           =   5535
      End
      Begin VB.TextBox txtNoEditable 
         BackColor       =   &H00FFFFFF&
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
         Height          =   390
         Index           =   0
         Left            =   240
         TabIndex        =   144
         Top             =   960
         Width           =   5535
      End
      Begin VB.TextBox txtNoEditable 
         BackColor       =   &H00FFFFFF&
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
         Height          =   390
         Index           =   1
         Left            =   240
         TabIndex        =   143
         Top             =   1740
         Width           =   3495
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Salir"
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
         Index           =   17
         Left            =   4440
         TabIndex        =   138
         Top             =   3840
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Kil�metros"
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
         Height          =   240
         Index           =   16
         Left            =   600
         TabIndex        =   146
         Top             =   3240
         Width           =   1050
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Veh�culo"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   15
         Left            =   240
         TabIndex        =   142
         Top             =   2280
         Width           =   825
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   14
         Left            =   240
         TabIndex        =   141
         Top             =   1500
         Width           =   765
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   13
         Left            =   240
         TabIndex        =   140
         Top             =   720
         Width           =   675
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Establecer kil�metros"
         BeginProperty Font 
            Name            =   "Verdana"
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
         Left            =   1320
         TabIndex        =   139
         Top             =   240
         Width           =   3825
      End
   End
   Begin VB.Frame FrameFamDto 
      Height          =   4255
      Left            =   135
      TabIndex        =   68
      Top             =   90
      Visible         =   0   'False
      Width           =   6825
      Begin VB.CommandButton cmdActuDtoFamMar 
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
         Left            =   4275
         TabIndex        =   74
         Top             =   3660
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancelar 
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
         Index           =   8
         Left            =   5595
         TabIndex        =   70
         Top             =   3660
         Width           =   1095
      End
      Begin MSComctlLib.ListView ListView3 
         Height          =   2455
         Left            =   120
         TabIndex        =   71
         Top             =   1080
         Width           =   6575
         _ExtentX        =   11589
         _ExtentY        =   4339
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Tipo"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Desc"
            Object.Width           =   5733
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Dto1"
            Object.Width           =   1766
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Dto2"
            Object.Width           =   1766
         EndProperty
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Familia"
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
         Height          =   240
         Index           =   4
         Left            =   120
         TabIndex        =   73
         Top             =   720
         Width           =   690
      End
      Begin VB.Label lblInd 
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   72
         Top             =   3780
         Width           =   2895
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Actualizar descuentos familia/marca"
         BeginProperty Font 
            Name            =   "Verdana"
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
         Left            =   120
         TabIndex        =   69
         Top             =   240
         Width           =   6420
      End
   End
   Begin VB.Frame FrameGenDtoCli 
      Height          =   2775
      Left            =   1680
      TabIndex        =   57
      Top             =   120
      Visible         =   0   'False
      Width           =   6495
      Begin VB.CommandButton cmdCrearDtos 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3360
         TabIndex        =   63
         Top             =   2160
         Width           =   1335
      End
      Begin VB.CheckBox chkDtoCliente 
         Caption         =   "Solo insertar nuevos"
         Height          =   255
         Left            =   960
         TabIndex        =   62
         Top             =   1800
         Width           =   1935
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Salir"
         Height          =   375
         Index           =   7
         Left            =   4920
         TabIndex        =   64
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox txtCliD 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   0
         Left            =   2280
         TabIndex        =   66
         Text            =   "Text1"
         Top             =   840
         Width           =   3975
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   2
         Left            =   1320
         TabIndex        =   61
         Text            =   "Text1"
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox txtCli 
         Height          =   285
         Index           =   0
         Left            =   1320
         TabIndex        =   60
         Text            =   "Text1"
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblIndDto 
         Caption         =   "Label2"
         Height          =   255
         Left            =   240
         TabIndex        =   67
         Top             =   2280
         Width           =   2295
      End
      Begin VB.Label Label4 
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
         Left            =   120
         TabIndex        =   65
         Top             =   1320
         Width           =   495
      End
      Begin VB.Image imgFecha 
         Height          =   255
         Index           =   2
         Left            =   960
         Top             =   1320
         Width           =   255
      End
      Begin VB.Image imgCli 
         Height          =   240
         Index           =   0
         Left            =   960
         Top             =   840
         Width           =   240
      End
      Begin VB.Label Label4 
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
         Index           =   2
         Left            =   120
         TabIndex        =   59
         Top             =   840
         Width           =   585
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Generar descuentos cliente"
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
         Left            =   960
         TabIndex        =   58
         Top             =   240
         Width           =   4635
      End
   End
   Begin VB.Frame FrameFenollar 
      Height          =   2415
      Left            =   1920
      TabIndex        =   118
      Top             =   2760
      Visible         =   0   'False
      Width           =   5055
      Begin VB.CheckBox chkFenollar 
         Caption         =   "Valorado"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   3000
         TabIndex        =   123
         Top             =   1080
         Width           =   1695
      End
      Begin VB.CheckBox chkFenollar 
         Caption         =   "Portes"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   122
         Top             =   1080
         Width           =   1935
      End
      Begin VB.CommandButton cmdImprAlbaFenoll 
         Caption         =   "&Aceptar"
         Default         =   -1  'True
         Height          =   375
         Left            =   2040
         TabIndex        =   121
         Top             =   1800
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Salir"
         Height          =   375
         Index           =   14
         Left            =   3600
         TabIndex        =   120
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Impresion albaranes"
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
         Left            =   960
         TabIndex        =   119
         Top             =   240
         Width           =   2940
      End
   End
   Begin VB.Frame FrameZona 
      Height          =   5535
      Left            =   120
      TabIndex        =   50
      Top             =   0
      Visible         =   0   'False
      Width           =   6015
      Begin VB.CommandButton cmdZonaxAlb 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3480
         TabIndex        =   56
         Top             =   4920
         Width           =   1095
      End
      Begin VB.CommandButton cmdpedxZon 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3480
         TabIndex        =   54
         Top             =   4920
         Width           =   1095
      End
      Begin MSComctlLib.TreeView Tv1 
         Height          =   3855
         Left            =   120
         TabIndex        =   53
         Top             =   840
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   6800
         _Version        =   393217
         Style           =   7
         Checkboxes      =   -1  'True
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Salir"
         Height          =   375
         Index           =   5
         Left            =   4680
         TabIndex        =   52
         Top             =   4920
         Width           =   1095
      End
      Begin VB.Label lblInd 
         Caption         =   "Label2"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   55
         Top             =   5040
         Width           =   2895
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Seleccionar zona / pedido"
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
         Left            =   720
         TabIndex        =   51
         Top             =   240
         Width           =   4875
      End
   End
   Begin VB.Frame FrameRenting 
      Height          =   6015
      Left            =   360
      TabIndex        =   100
      Top             =   120
      Visible         =   0   'False
      Width           =   6255
      Begin VB.TextBox txtRenting 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4575
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   103
         Text            =   "frmVariosNew.frx":07C2
         Top             =   720
         Width           =   5775
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cerrar"
         Height          =   375
         Index           =   12
         Left            =   4920
         TabIndex        =   102
         Top             =   5520
         Width           =   1095
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Fechas incorrectas facturacion rentings "
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
         Left            =   240
         TabIndex        =   101
         Top             =   240
         Width           =   5715
      End
   End
End
Attribute VB_Name = "frmVariosNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event DatoSeleccionado(CadenaSeleccion As String)

Public Opcion As Byte
    ' 0.-   Impresion de facturas directas (tipo 4tonda)
    ' 1.-   Eliminar articulos masiva
    ' 2.-   Estadisticas consultas (archivo-facturacion-pedidos-consulta precio/cliente

    ' 3.-   Eleccion del metodo de envio para los albaranes

    ' 4.-   Ver clientes para a�adir acciones comerciales
    
    ' 5.-   Pedidos x Zona. Selecionar las zonas
    ' 6.-   Albaranes trnasporte x codzona (dentro del albaran)
    
    ' 7.-   Generar dtos familia/marca/cliente
    
    ' 8.-   Actualizar en sdtofm para una familia y unos tipos(que se mostraran)
    
    ' 9.-   Modific los descuentos de todas las lineas de un pedido de proveedor
    
    ' 10.-  Listado de articulos rectificados por stock
        
    ' 11.-  Compras por a�o/proveedor
    
    ' 12.-  Ver rentings con fechas mal para facturar
    
    ' 13.-  Eliminar factura  pasando a albaranes  -  Cambiar fecha factura
    
    
    ' 14.  Imprimersion albaranes FENOLLAR
    
    
    ' 15.  Cambiar tipo Albaran (EULER)
    ' 16.  Quitar campos venta plazos
    ' 17.  Taxco. Establecer kilometros
    
    '100.  Ayuda al traspaso de almacenes
    '101.  Ayuda al movimientos de almacen
Public CADENA As String
    
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmA As frmBasico2
Attribute frmA.VB_VarHelpID = -1
Private WithEvents frmCli As frmBasico2
Attribute frmCli.VB_VarHelpID = -1
Private WithEvents frmPr As frmBasico2 '%=%=frmComProveedores
Attribute frmPr.VB_VarHelpID = -1


Private cad As String
Private SePuedeCerrar As Boolean   'Puede llevar DoEvents
Private PrimeraVez1 As Byte   '0.- Primera vez, 1.- cargando datos en forma_activate  2.- Fin carga




Private Sub cboTipoDtoFamia_KeyPress(KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub Check1_Click()
    Screen.MousePointer = vbHourglass
    
    CargaMovimientosAlmacen
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdAccionListview_Click()
Dim T1 As Single

    CadenaDesdeOtroForm = ""
    For NumRegElim = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(NumRegElim).Checked Then CadenaDesdeOtroForm = CadenaDesdeOtroForm & "1"
    Next
    
    If CadenaDesdeOtroForm = "" Then
        MsgBox "Seleccione algun art�culo para eliminar", vbInformation
    Else
        CadenaDesdeOtroForm = Len(CadenaDesdeOtroForm)
        CadenaDesdeOtroForm = "Va a eliminar " & CadenaDesdeOtroForm & " art�culo(s).   �Continuar?"
        If MsgBox(CadenaDesdeOtroForm, vbQuestion + vbYesNo) = vbNo Then CadenaDesdeOtroForm = ""
    End If
    If CadenaDesdeOtroForm = "" Then Exit Sub
    
    
    
    
    
    'AHora eliminamos
    'Y el log de acciones
    '------------------------------------------------------------------------------
    '  LOG de acciones
    Set LOG = New cLOG
    
    
    '-----------------------------------------------------------------------------
    
    Screen.MousePointer = vbHourglass
    lblElim(1).Caption = ""
    For NumRegElim = ListView1.ListItems.Count To 1 Step -1
        If ListView1.ListItems(NumRegElim).Checked Then
            T1 = Timer
            ListView1.ListItems(NumRegElim).EnsureVisible
            conn.BeginTrans
            If EliminarArticulo(ListView1.ListItems(NumRegElim).Text, lblElim(1)) Then
                LOG.Insertar 7, vUsu, ListView1.ListItems(NumRegElim).Text & " " & ListView1.ListItems(NumRegElim).SubItems(1)
                conn.CommitTrans
                'QUitamos del nodo
                ListView1.ListItems.Remove ListView1.ListItems(NumRegElim).Index
                T1 = 1.5 - (Timer - T1)
                If T1 > 0 Then Espera T1
                
            Else
                'NO se ha podido eliminar
                conn.RollbackTrans
                ListView1.ListItems(NumRegElim).Bold = True
                ListView1.ListItems(NumRegElim).ForeColor = vbRed
                ListView1.ListItems(NumRegElim).Checked = False
            End If
        End If
    Next
    lblElim(1).Caption = ""
    Screen.MousePointer = vbDefault
    Set LOG = Nothing
    If ListView1.ListItems.Count = 0 Then
        SePuedeCerrar = True
        Unload Me
    End If
End Sub

Private Sub CmdAcepHcoAlm_Click()
Dim NroMov As String

    If ListView6.SelectedItem <> "" Then
        RaiseEvent DatoSeleccionado(ListView6.SelectedItem)
        Unload Me
    End If
End Sub

Private Sub cmdActualiDtoProv_Click()
    If txtDecimal(0).Text = "" Or txtDecimal(1).Text = "" Then
        MsgBox "Especifique valor para ambos descuentos", vbExclamation
        Exit Sub
    End If
    
    If MsgBox("Desea actualizar los descuentos de la oferta del proveedor?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    CadenaDesdeOtroForm = ImporteFormateado(txtDecimal(0)) & "|" & ImporteFormateado(txtDecimal(1)) & "|"
    Unload Me
    
End Sub

Private Sub cmdActuDtoFamMar_Click()
    'Vamos p'alla
    cad = ""
    For NumRegElim = 1 To ListView3.ListItems.Count
        If ListView3.ListItems(NumRegElim).Checked Then cad = cad & "1"
    Next
    
    If cad = "" Then
        MsgBox "Ning�n valor seleccionado", vbExclamation
        Exit Sub
    End If
    
    cad = Len(cad)
    If Val(cad) = 1 Then
        cad = "el descuento seleccionado"
    Else
        cad = "los " & cad & " descuentos seleccionados"
    End If
    cad = "Va a actualizar la tabla de descuentos/familia-marca para " & cad
    If MsgBox(cad, vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    
    CadenaDesdeOtroForm = Label4(4).Caption  'para el log
    lblInd(1).Caption = ""

    For NumRegElim = 1 To ListView3.ListItems.Count
        lblInd(1).Caption = ListView3.ListItems(NumRegElim).SubItems(1)
        lblInd(1).Refresh
    
        cad = DBSet(ListView3.ListItems(NumRegElim).SubItems(2), "N")
    
    
        cad = "update sdtofm set dtoline1= " & cad
        '11 Octubre 2011. El dtoline 2 NO lo hacia
        cad = cad & ", dtoline2= " & DBSet(ListView3.ListItems(NumRegElim).SubItems(3), "N")
        cad = cad & " where codfamia=" & Label4(4).Tag & " and codmarca is null and codclien in ("
        cad = cad & " select codclien from sactivdtos,sclien where sclien.codactiv="
        cad = cad & " sactivdtos.codactiv and codfamia=" & Label4(4).Tag
        cad = cad & " and clasifica=" & ListView3.ListItems(NumRegElim).Text
        cad = cad & ") "
        ejecutar cad, False
        cad = ListView3.ListItems(NumRegElim).SubItems(1) & " -> " & ListView3.ListItems(NumRegElim).SubItems(2)
        CadenaDesdeOtroForm = CadenaDesdeOtroForm & vbCrLf & cad
    
    Next
    
    Set LOG = New cLOG
    LOG.Insertar 16, vUsu, CadenaDesdeOtroForm
    Set LOG = Nothing
    lblInd(1).Caption = "Proceso finalizado "
    lblInd(1).Refresh
    Espera 0.5
    Unload Me
End Sub

Private Sub cmdBorrarVtaPlazosFinalizadas_Click()
    
'    For NumRegElim = 1 To ListView5.ColumnHeaders.Count
'        Debug.Print ListView5.ColumnHeaders(NumRegElim).Width
'    Next
'
    
    cad = ""
    For NumRegElim = 1 To ListView5.ListItems.Count
        If ListView5.ListItems(NumRegElim).Checked Then cad = cad & "X"
    Next
    If Len(cad) = 0 Then
        MsgBox "No ha seleccionado ningun telefono", vbExclamation
        Exit Sub
    End If
    NumRegElim = Len(cad)
    CadenaDesdeOtroForm = PonerTrabajadorConectado(cad)
    If CadenaDesdeOtroForm = "" Then
        MsgBox "Imposible obtener trabajador conectado", vbExclamation
        Exit Sub
    End If
    
    cad = "Va e quitar los venta a plazos de " & NumRegElim & " lineas de telefono seleccionadas."
    cad = cad & vbCrLf & vbCrLf & "�Continuar?"
    If MsgBox(cad, vbQuestion + vbYesNoCancel) <> vbYes Then
        CadenaDesdeOtroForm = ""
        Exit Sub
    End If
    
    
        
    
    
    Screen.MousePointer = vbHourglass
    For NumRegElim = ListView5.ListItems.Count To 1 Step -1
        If ListView5.ListItems(NumRegElim).Checked Then
            
            lblInd(2).Caption = "Eliminando de " & Me.ListView5.ListItems(NumRegElim).Text
            lblInd(2).Refresh
    
            
            cad = DBSet(vUsu.Login, "T") & ",##@#FECH#@##,"
            cad = "INSERT INTO scrmacciones(usuario,fechora,codclien,agente,codtraba,estado,tipo,medio,observaciones) VALUES (" & cad
            cad = cad & Me.ListView5.ListItems(NumRegElim).SubItems(1) & ",1," & CadenaDesdeOtroForm & ",2,6,'Otros',"
            cad = cad & DBSet(Me.ListView5.ListItems(NumRegElim).Tag, "T") & ")"
            davidCodtipom = Replace(cad, "##@#FECH#@##", DBSet(Now, "FH"))
            If Not ejecutar(davidCodtipom, True) Then
                'Puede dar error por temas de segundos para un mismo cliente
                Espera 1.1
                davidCodtipom = Replace(cad, "##@#FECH#@##", DBSet(Now, "FH"))
                ejecutar davidCodtipom, False
            End If
    
            cad = "UPDATE sclientfno SET  ArtPlazos =null,PlazosMeses =null,ImportePlazo=null, PlazosOrigen=null "
            cad = cad & " WHERE codclien =" & Me.ListView5.ListItems(NumRegElim).SubItems(1) & " AND IdTelefono =" & DBSet(Me.ListView5.ListItems(NumRegElim).Text, "T")
            ejecutar cad, False

            ListView5.ListItems.Remove ListView5.ListItems(NumRegElim).Index
        End If
    Next
    CadenaDesdeOtroForm = ""
    davidCodtipom = ""
    Screen.MousePointer = vbDefault
    Unload Me
    
End Sub

Private Sub cmdCambiarTipoAlbaran_Click()
    cad = ""
     ListView4.Tag = ""
     CadenaDesdeOtroForm = ""
    For NumRegElim = 1 To ListView4.ListItems.Count
        If ListView4.ListItems(NumRegElim).Checked Then
            cad = cad & ListView4.ListItems(NumRegElim).Text
            CadenaDesdeOtroForm = ListView4.ListItems(NumRegElim).Text
            ListView4.Tag = ListView4.ListItems(NumRegElim).Text & " - " & ListView4.ListItems(NumRegElim).SubItems(1)
        End If
    Next
    If cad = "" Or Len(cad) > 3 Then
        CadenaDesdeOtroForm = ""
        If Len(cad) > 3 Then cad = " (y solo uno)"
        cad = "Seleccion uno " & cad & " tipo de albaran"
        MsgBox cad, vbExclamation
        Exit Sub
    End If
    
    cad = "Va a cambiar el tipo de albaran a: " & vbCrLf & vbCrLf & ListView4.Tag & vbCrLf & vbCrLf & "�Continuar?"
    If MsgBox(cad, vbQuestion + vbYesNoCancel) <> vbYes Then
        CadenaDesdeOtroForm = ""
        Exit Sub
    End If
    
    Unload Me
        
End Sub

Private Sub cmdCambiFecReestbFact_Click()
Dim cControlFra As CControlFacturaContab
Dim B1 As Boolean
    

    If Me.optElimFact(0).Value Then
    
        If txtFecha(3).Text = "" Then Exit Sub
    
        'Comprobamos la fecha NUEVA que ha puesto
        Set cControlFra = New CControlFacturaContab
        cad = ""
        'Con estos dos NO dejo pasar
        CadenaDesdeOtroForm = cControlFra.FechaCorrectaContabilizazion(ConnConta, txtFecha(3))
        If CadenaDesdeOtroForm <> "" Then cad = cad & "- " & CadenaDesdeOtroForm & vbCrLf
        CadenaDesdeOtroForm = cControlFra.FechaCorrectaIVA(ConnConta, txtFecha(3))
        If CadenaDesdeOtroForm <> "" Then cad = cad & "- " & CadenaDesdeOtroForm & vbCrLf
        CadenaDesdeOtroForm = ""
        
        If cad <> "" Then
            B1 = True
            If vParamAplic.PuedeModificarAriconta Then
                If CDate(txtFecha(3).Text) < vEmpresa.FechaIni Then
                    B1 = True 'Fecha anterior a fecha ejercicio. NO se toca
                Else
                    B1 = False
                End If
            End If
        
            If B1 Then
                MsgBox cad, vbExclamation
                Set cControlFra = Nothing
                Exit Sub
            End If
        End If
        

        If cControlFra.FechaMenorUltimaFacturaCliente(ConnConta, txtFecha(3), Me.cmdCambiFecReestbFact.Tag) Then
            If CadenaDesdeOtroForm <> "" Then cad = cad & "-Anterior a cfactura contabilizada " & vbCrLf
        End If
        Set cControlFra = Nothing
        
        CadenaDesdeOtroForm = ""
        
        If cad <> "" Then
            cad = cad & "�Continuar el proceso?"
            
            If MsgBox(cad, vbExclamation + vbYesNo) <> vbYes Then Exit Sub
        
        End If
    
        
        cad = "establecer como fecha factura: " & Me.txtFecha(3).Text
    Else
    
        If FrameNuevaFecFac.Tag = "1" Then
            MsgBox "No se puede deshacer factura de telefon�a", vbExclamation
            Exit Sub
        End If
    
        cad = "eliminar factura y reestablecer los albaranes facturados"
        
    End If
    cad = "Va a " & cad & vbCrLf & vbCrLf & vbCrLf
    cad = cad & " NO se realizaran acciones sobre Arimoney ni Ariconta " & vbCrLf & vbCrLf
    cad = cad & " **** Se grabar� el registro de acciones *** " & vbCrLf
    cad = cad & vbCrLf & vbCrLf & "Introduzca el password para continuar"
    
    
    If vParamAplic.NumeroInstalacion = vbFenollar Then
        If MsgBox("�Continuar con el proceso?", vbQuestion + vbYesNoCancel) = vbYes Then cad = "ARIADNA"
    Else
        cad = InputBox(cad, "Seguridad")
    End If
    If UCase(cad) <> "ARIADNA" Then Exit Sub
        
        
    If Me.optElimFact(0).Value Then
        CadenaDesdeOtroForm = Me.txtFecha(3).Text
    Else
        CadenaDesdeOtroForm = "OK"
    End If
    Unload Me
    
End Sub

Private Sub cmdCancelar_Click(Index As Integer)
    
    If Opcion = 0 Then
        'Esta haciendo cosas. Preguntar si cerrar
        If Not SePuedeCerrar Then
            If MsgBox("Seguro que desea finalizar el proceso?", vbQuestion + vbYesNo) = vbYes Then SePuedeCerrar = True
            Exit Sub
        End If
    ElseIf Opcion = 7 Then
        CadenaDesdeOtroForm = "" 'por si acaso han utlizado la variable
    ElseIf Opcion = 13 Then
        CadenaDesdeOtroForm = "" 'por si acaso han utlizado la variable
    ElseIf Opcion = 100 Then
        CadenaDesdeOtroForm = ""
    ElseIf Opcion = 101 Then
        CadenaDesdeOtroForm = ""
    End If
    
    SePuedeCerrar = True
    Unload Me
End Sub

Private Sub cmdCerrar_Click()
    SePuedeCerrar = True
    Unload Me
End Sub

Private Sub cmdClientes_Click()
        cad = ""
        For NumRegElim = 1 To ListView2.ListItems.Count
            If ListView2.ListItems(NumRegElim).Checked Then cad = cad & ", " & CStr(Val(ListView2.ListItems(NumRegElim).Text))
        Next NumRegElim
        If cad = "" Then
            MsgBox "Seleccione algun dato", vbExclamation
            Exit Sub
        End If
        CadenaDesdeOtroForm = Mid(cad, 2) 'le quito la primera coma
        Unload Me
End Sub

Private Sub cmdComprasMeses_Click()
Dim campo As String
    
   
    
    cad = ""
    
    'El campo A�O es obligarotorio
    txtNumero(0).Text = Trim(txtNumero(0).Text)
    If txtNumero(0).Text = "" Then
        MsgBox "Debe seleccionar una a�o para el informe.", vbInformation
        Exit Sub
    End If
    campo = "year({scafpc.fecrecep})"
    campo = campo & " = " & Me.txtNumero(0).Text
    cad = cad & "pAnyo=""" & "A�o: " & txtNumero(0).Text & """|"
    
 
    
    'Campo seleccion de un CLIENTE
    txtProv(0).Text = Trim(txtProv(0).Text)
    If txtProv(0).Text <> "" Then
        campo = campo & " AND ({scafpc.codprove} =" & txtProv(0).Text & ")"
        'Pasar el cliente solicitado como parametro
        cad = cad & "pDHCliente=""" & "Proveedor: " & txtProv(0).Text & " - " & txtProvD(0).Text & """|"
    Else
        'Mostrar en el informe el total del A�o Anterior
        campo = "(" & campo & " OR year({scafpc.fecrecep}) = " & CInt(txtNumero(0).Text) - 1 & ")"
        
        cad = cad & "pDHCliente=""" & "Proveedores: Todos" & """|"
    End If
    
    
    

    If Not HayRegParaInforme("scafpc", campo) Then Exit Sub
    
    
    'Borro los datos temporales,por si acaso se hubiera quedado
    BorrarTempInformes
    
    'Generar la temporal con los totales por a�o, mes y cliente (tmpinformes)
    If Not TempComprasMeses(campo, txtNumero(0).Text) Then
        'Borrar los registros generados por el usuario de la temporal
        BorrarTempInformes
        Exit Sub
    End If
    
    
    

        

    
    
    
    'Pasar nombre de la Empresa como parametro
    cad = "|pEmpresa=""" & vEmpresa.nomempre & """|" & cad
    

    
    With frmImprimir
        .FormulaSeleccion = "{tmpinformes.codusu} = " & vUsu.Codigo
        .OtrosParametros = cad
        .NumeroParametros = 3

        .SoloImprimir = False
        .EnvioEMail = False
        .Opcion = 2003
        .Titulo = "Compras por meses"
        .NombreRPT = "rCompasxMesGra.rpt"
        .ConSubInforme = True
        .Show vbModal
    End With
    
    'Borrar los registros generados por el usuario de la temporal
    BorrarTempInformes
End Sub



Private Sub cmdCrearDtos_Click()
Dim Actividad As String
    'Todos los datos bien
    cad = ""
    If txtCliD(0).Text = "" Xor txtCliD(0).Text = "" Then
        cad = cad & "-Error en cliente"
    Else
        If txtCli(0).Text = "" Then cad = cad & "-Falta cliente"
    End If
    
    If txtFecha(2).Text = "" Then cad = cad & "-Falta fecha "

    
    
    
    If cad <> "" Then
        cad = "Error en datos: " & vbCrLf & cad
        MsgBox cad, vbExclamation
        Exit Sub
    End If
    
'
    'Veremos si tiene dtos para la actividad
    Actividad = DevuelveDesdeBD(conAri, "codactiv", "sclien", "codclien ", txtCli(0).Text, "N")
 
    If Actividad = "" Then Actividad = "-1"
    cad = DevuelveDesdeBD(conAri, "count(*)", "sactivdtos", "codactiv", Actividad, "N")
    If cad = "" Then cad = "0"
    If Val(cad) = 0 Then
        MsgBox "No hay ningun descuentos para la actividad:" & Actividad, vbExclamation
        Exit Sub
    End If
    
    

    'OK adelante
    'Ala pues, alla vamos
    cad = "DELETE FROM tmpgendtos  WHERE codusu = " & vUsu.Codigo
    conn.Execute cad
    
  
    'Cargo con los temporales
    cad = "INSERT INTO tmpgendtos(codusu,codfamia,codmarca,clasifica,dtoline1)"
    cad = cad & " SELECT " & vUsu.Codigo & ", sfamiadtos.codfamia,codmarca,sactivdtos.clasifica,dtoline1 "
    cad = cad & " FROM sactivdtos,sfamiadtos,sfamiatipodto WHERE "
    cad = cad & " sactivdtos.codfamia=sfamiadtos.codfamia AND sfamiadtos.clasifica=sactivdtos.clasifica"
    cad = cad & " AND sfamiatipodto.clasifica=sactivdtos.clasifica AND sactivdtos.codactiv = " & Actividad

    conn.Execute cad
    
    
    

    
    'hacer insert de los que queden
    CadenaDesdeOtroForm = ""
    frmFacDtosAsigtmp.vClien = txtCli(0).Text
    frmFacDtosAsigtmp.vFec = CDate(txtFecha(2).Text)
    frmFacDtosAsigtmp.vSoloNuevos = chkDtoCliente.Value = 1
    frmFacDtosAsigtmp.Show vbModal
    
    'Grabamos esta variable para que en el mentenimiento carge el grid con estos valores
    If CadenaDesdeOtroForm <> "" Then CadenaDesdeOtroForm = txtCli(0).Text & "|" & txtFecha(2).Text & "|"
    Unload Me
End Sub




Private Sub cmdEliminarArticulos_Click()
Dim SQL As String
Dim IT As ListItem

    '
    lblElim(0).Caption = "Cargando datos"
    lblElim(0).Refresh
    
    'Eliminamos los datos de tmpnseries
    conn.Execute "DELETE FROM tmpnseries where codusu = " & vUsu.Codigo
    
    
    'Cargamos tmpnseries con los articulos del desde hasta
    SQL = ""
    If Me.txtArt(0).Text <> "" Then SQL = SQL & " AND  codartic >=" & DBSet(txtArt(0).Text, "T")
    If Me.txtArt(1).Text <> "" Then SQL = SQL & " AND codartic <=" & DBSet(txtArt(1).Text, "T")
    If Me.txtProv(1).Text <> "" Then SQL = SQL & " AND  codprove >=" & txtProv(1).Text
    If Me.txtProv(2).Text <> "" Then SQL = SQL & " AND codprove <=" & txtProv(2).Text
    
    
    
    If SQL <> "" Then SQL = Mid(SQL, 5)  'QUITAMOS EL PRIMER AND
    
    
    If SQL <> "" Then SQL = " WHERE " & SQL
    SQL = " SELECT " & vUsu.Codigo & ",codartic,0,0 FROM sartic " & SQL
    SQL = "insert into `tmpnseries` (`codusu`,`codartic`,`numlinealb`,`numlinea`) " & SQL
    conn.Execute SQL
    
    
    Set miRsAux = New ADODB.Recordset
    
    'Eliminamos de tmpnseries los articulos que seguro estan en
    ' alba, fact....
    EliminandoArticulos_Paso1
    
    
    'Ya tengo los articulos. Vere cuales puedo borrar
    lblElim(0).Caption = "Obteniendo registros"
    lblElim(0).Refresh
    
    SQL = "Select tmpnseries.codartic,nomartic from tmpnseries,sartic where codusu = " & vUsu.Codigo
    SQL = SQL & " AND tmpnseries.codartic=sartic.codartic"
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If miRsAux.EOF Then
        lblElim(0).Caption = ""
        MsgBox "No existen registros", vbExclamation
        miRsAux.Close
        Set miRsAux = Nothing
        Exit Sub
    End If
    
    'Ajustamos los tama�os para cargar el LISTVIEW
    CargaColumnas
    NumRegElim = (Screen.Width - FrameListArticulos.Width - 420) \ 2
    Me.Left = NumRegElim
    NumRegElim = (Screen.Height - FrameListArticulos.Height - 360) \ 2
    Me.Top = NumRegElim
    Me.FrameDHArticulo.visible = False
    PonerFrameVisible Me.FrameListArticulos
    Me.lblTitulo(1).Caption = "Eliminar art�culos"
    DoEvents
    
    'Vamos cargando los registros
    While Not miRsAux.EOF
        Set IT = ListView1.ListItems.Add()
        IT.Text = miRsAux!codArtic
        IT.SubItems(1) = miRsAux!NomArtic
        IT.Checked = True
        'Sig
        miRsAux.MoveNext
    Wend
End Sub

Private Sub cmdFormaEnvio_Click()
Dim i As Integer

    If ListEnvio.ListIndex < 0 Then Exit Sub
    cad = ListEnvio.List(ListEnvio.ListIndex)
    i = InStrRev(cad, "(")
    cad = Trim(Mid(cad, i + 1))
    i = InStrRev(cad, ")")
    cad = Mid(cad, 1, i - 1) 'quitamos el ultmio parentesis
    CadenaDesdeOtroForm = cad
    
    i = InStrRev(ListEnvio.List(ListEnvio.ListIndex), "(")
    cad = Mid(ListEnvio.List(ListEnvio.ListIndex), 1, i - 1)  'quito el precio kilo
    
    i = Val(Mid(cad, 1, 10))
    
    cad = Trim(Mid(cad, 11))
    
    CadenaDesdeOtroForm = i & "|" & cad & "|" & CadenaDesdeOtroForm & "|"
    
    'Desde kilo
    cad = ListEnvio.List(ListEnvio.ListIndex)
    i = InStrRev(ListEnvio.List(ListEnvio.ListIndex), "Desde :")
    cad = Mid(cad, i + 7)
    cad = Trim(Mid(cad, 1, Len(cad) - 2)) 'Le kito kg
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & cad & "|"
    SePuedeCerrar = True
    Unload Me
End Sub

Private Sub cmdImprAlbaFenoll_Click()
    CadenaDesdeOtroForm = chkFenollar(0).Value & "|" & chkFenollar(1).Value & "|"
    Unload Me
End Sub

Private Sub cmdListConsultaPedido_Click()
Dim Aux As String


    cad = ""
    Aux = CadenaDesdeHastaBD(txtArt(2).Text, txtArt(3).Text, "codartic", "T")
    If Aux <> "" Then cad = Aux
    
    'La fecha
    Aux = CadenaDesdeHastaBD(txtFecha(0).Text, txtFecha(1).Text, "DiaHora", "FH")
    If Aux <> "" Then
        If cad <> "" Then cad = cad & " AND "
        cad = cad & Aux
    End If
        
    If Not HayRegParaInforme("sconsulta", cad) Then Exit Sub
    
    
    'Para el informe
    cad = ""
    Aux = CadenaDesdeHasta(txtArt(2).Text, txtArt(3).Text, "{sconsulta.codartic}", "T")
    If Aux <> "" Then cad = Aux
    
    'La fecha
    Aux = CadenaDesdeHasta(txtFecha(0).Text, txtFecha(1).Text, "{sconsulta.DiaHora}", "FH")
    If Aux <> "" Then
        If cad <> "" Then cad = cad & " AND "
        cad = cad & Aux
    End If
    
    
    
    
    With frmImprimir
        .FormulaSeleccion = cad
        .OtrosParametros = ""
        .NumeroParametros = 0

        .SoloImprimir = False
        .EnvioEMail = False
        .Opcion = 2002
        .Titulo = "Estadistica consultas pedidos"
        .NombreRPT = "rFacConsuPrecioArt.rpt"
        .ConSubInforme = False
        .Show vbModal
    End With
    
    
    
End Sub

Private Sub cmdpedxZon_Click()
Dim N As Node
Dim i As Byte

    If Tv1.Nodes.Count = 0 Then Exit Sub
    'Nos recorreemos el tv1 por si a desmarcado alguno
    'Lo borraremos de la tabla tmpsliped
    lblInd(0).Caption = "Preparando datos"
    lblInd(0).Refresh
    
    i = 0
    Set N = Tv1.Nodes(1)
    While Not N Is Nothing
        NumRegElim = -1
        If Not N.Checked Then
            NumRegElim = N.Index
            conn.Execute "DELETE from tmpsliped where codusu = " & vUsu.Codigo & " AND codzona = " & Mid(N.Key, 2)
        Else
            i = 1
        End If
        
        Set N = N.Next
        If NumRegElim > 0 Then Tv1.Nodes.Remove NumRegElim
    Wend
    
    If i = 0 Then
        MsgBox "Nada seleccionado", vbExclamation
        lblInd(0).Caption = ""
        Exit Sub
    End If
    
    'Haremos los selectsum de las salmac
    Screen.MousePointer = vbHourglass
    Me.FrameZona.Enabled = False
    i = 0
    If PonerRestoDatos Then i = 1
        
    Me.FrameZona.Enabled = True
    Screen.MousePointer = vbDefault
    lblInd(0).Caption = ""
    If i = 1 Then
        CadenaDesdeOtroForm = DevuelveDesdeBD(conAri, "count(*)", "tmpsliped", "codusu", CStr(vUsu.Codigo))
        If CadenaDesdeOtroForm = "" Then CadenaDesdeOtroForm = "0"
        If Val(CadenaDesdeOtroForm) = 0 Then
            MsgBox "Ningun dato para mostrar", vbExclamation
            i = 0
        End If
    End If
    If i = 1 Then
        If PonerParamRPT2(48, cad, i, CadenaDesdeOtroForm, False, "", 0) Then
            With frmImprimir
                .ConSubInforme = False
                .FormulaSeleccion = "{tmpsliped.codusu} = " & vUsu.Codigo
                .NombreRPT = CadenaDesdeOtroForm
                .NombrePDF = pPdfRpt
                .Titulo = "List. pedidos por zonas"
                .OtrosParametros = cad
                .NumeroParametros = CInt(i)
                .Opcion = 2003 'Esta libre
                
                i = vParamAplic.NumCop_PedZona
                .NumeroCopias = i
                .Show vbModal
            End With
        End If
    End If
    
End Sub

Private Sub cmdTaxcoEstabKm_Click()
    CadenaDesdeOtroForm = ""
    txtNumero(1).Text = Trim(txtNumero(1).Text)
    If txtNumero(1).Text = "" Then Exit Sub
    
    If Not EsEnteroNew(txtNumero(1).Text) Then
        txtNumero(1).Text = ""
        Exit Sub
    End If
    
    cad = txtNumero(1).Text
    cad = Replace(cad, ".", "")
    NumRegElim = CLng(cad)
    
    If NumRegElim <= 0 Then
        MsgBox "kilometros debe ser mayor que cero", vbExclamation
    Else
        If NumRegElim > 1000000 Then
            If MsgBox("Kilometros suerior a un millon. �Continuar?", vbQuestion + vbYesNoCancel) <> vbYes Then
                PonerFoco txtNumero(1)
                Exit Sub
            End If
        End If
        
        
        'OK
        CadenaDesdeOtroForm = NumRegElim
        Unload Me
    End If
    
    
End Sub

Private Sub cmdZonaxAlb_Click()
Dim N
    If Tv1.Nodes.Count = 0 Then Exit Sub
    'Nos recorreemos el tv1 por si a desmarcado alguno
    'Lo borraremos de la tabla tmpsliped
    lblInd(0).Caption = "Devuelve datos"
    lblInd(0).Refresh
    

    Set N = Tv1.Nodes(1)
    NumRegElim = 0  'Los nodos NO chequeados
    cad = ""
    While Not N Is Nothing
        
        If Not N.Checked Then
            NumRegElim = NumRegElim + 1
        Else
            cad = cad & ", " & Mid(N.Key, 2)
        End If
        
        Set N = N.Next
    Wend
    
    If cad = "" Then
        MsgBox "Nada seleccionado", vbExclamation
        lblInd(0).Caption = ""
        Exit Sub
    End If
    
    'Ahora si estan todos los nodos seleccionados  no hace falta que haga
    'en el select un codzona in (1,2... etc)
    'Si son todos, son todos. No ponemos una condicion mas
    If NumRegElim = 0 Then
        CadenaDesdeOtroForm = ""
    Else
        CadenaDesdeOtroForm = Mid(cad, 2)
    End If
    Unload Me
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Activate()
    If PrimeraVez1 = 0 Then
        PrimeraVez1 = 1
        
        Select Case Opcion
        Case 0
            'Se pone a imprimir las facturas
            HacerImpresionFacturas
            
        Case 3
            ListEnvio.SetFocus
        Case 4
            CargaClientes
            
            
        Case 5
            'En cadedesdeotroform llevo si muestro solo los articulos que tenga stock
            lblTitulo(5).Tag = CadenaDesdeOtroForm
            CargaZonas
        
        Case 6
            CargaZonasAlbaranTransporte
            
        Case 8
            CargarFamiliaDtos
        Case 10
            CargaArticulosStockRectificado
        Case 12
            Me.txtRenting.Text = CadenaDesdeOtroForm
            CadenaDesdeOtroForm = ""
        Case 13
             If vParamAplic.NumeroInstalacion = vbFenollar Then optElimFact(1).Value = True
        
        Case 15
            CargaTiposALbaran
        
        Case 16
            CargaVtaPlazosFinalizados
            
        Case 100, 101
            Check1.Value = RecuperaValor(CADENA, 3)
            Screen.MousePointer = vbHourglass
            CargaMovimientosAlmacen
            Screen.MousePointer = vbDefault
        End Select
        PrimeraVez1 = 2
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub CargarIconos()
Dim i As Image


    For Each i In Me.imgArticulo
         i.Picture = frmPpal.imgListComun.ListImages(19).Picture
         i.ToolTipText = "Articulo"
    Next
    For Each i In Me.imgCli
         i.Picture = frmPpal.imgListComun.ListImages(19).Picture
         i.ToolTipText = "Cliente"
    Next
    For Each i In Me.imgFecha
         i.Picture = frmPpal.imgListComun.ListImages(23).Picture
         i.ToolTipText = "fecha"
    Next
    For Each i In Me.imgProv
         i.Picture = frmPpal.imgListComun.ListImages(19).Picture
         i.ToolTipText = "Proveedor"
    Next

End Sub

Private Sub Form_Load()
Dim IndexOpcion As Integer

    Me.Icon = frmPpal.Icon
    PrimeraVez1 = 0
    
    limpiar Me
    CargarIconos
    FrameListArticulos.visible = False
    FrameDHArticulo.visible = False
    Me.FrameImpresionFacturasDirectas.visible = False
    Me.FrameEstadisticasConsultas.visible = False
    FrameFormaEnvio.visible = False
    FrameSelClien.visible = False
    FrameZona.visible = False
    FrameGenDtoCli.visible = False
    FrameFamDto.visible = False
    FrameDtoProve.visible = False
    FrRectifcadoStocks.visible = False
    FrameComprasAnyo.visible = False
    FrameRenting.visible = False
    FrameFenollar.visible = False
    FrameVtaPlazosBorrar.visible = False
    Me.FramaTaxcoCambioKm.visible = False
    Me.FrameHcoAlmacen.visible = False
    SePuedeCerrar = True
    IndexOpcion = Opcion
    Select Case Opcion
    Case 0
        PonerFrameVisible Me.FrameImpresionFacturasDirectas
    Case 1
        PonerFrameVisible FrameDHArticulo
    Case 2
        PonerFrameVisible Me.FrameEstadisticasConsultas
    Case 3
        'Metodo de envio
        'En cadena deotro form llevo las lineas para seelccionar una de ellas
        SePuedeCerrar = False
        PonerFrameVisible FrameFormaEnvio
        CargaFormasEnvioPosibles
        IndexOpcion = -1  'No miostrara cancel
    Case 4
        PonerFrameVisible Me.FrameSelClien
        
    Case 5, 6
        cmdZonaxAlb.visible = Opcion = 6
        Me.cmdpedxZon.visible = Opcion = 5
        PonerFrameVisible FrameZona
        lblInd(0).Caption = ""
        IndexOpcion = 5  'para los dos, el cancelar es el mismo
    Case 7
        'Generar dtos para un cliente
        PonerFrameVisible FrameGenDtoCli
       
        Me.txtFecha(2).Text = Format(Now, "dd/mm/yyyy")
        lblIndDto.Caption = ""
        
    Case 8
        'Actualizar DTOS FAMILIA MARCA
        'FrameFamDto
        PonerFrameVisible FrameFamDto
        
        lblInd(1).Caption = ""
        
    Case 9
        'Actualizar dtos proveedor
        PonerFrameVisible FrameDtoProve
        
        'Cadenadesdeotroform  llevara los datos para los labels. Al momento lo pongo a ""
        Me.Label4(6).Caption = RecuperaValor(CadenaDesdeOtroForm, 1)
        Me.Label4(8).Caption = RecuperaValor(CadenaDesdeOtroForm, 2)
        CadenaDesdeOtroForm = ""
        
    Case 10
        PonerFrameVisible FrRectifcadoStocks
    Case 11
        PonerFrameVisible FrameComprasAnyo
    Case 12
        PonerFrameVisible FrameRenting
        
    Case 13
        PonerFrameVisible FrameElimCambiarFecFactura
        Me.txtFecha(3).Tag = RecuperaValor(CadenaDesdeOtroForm, 1)
        Me.cmdCambiFecReestbFact.Tag = CStr(RecuperaValor(CadenaDesdeOtroForm, 2))
        FrameNuevaFecFac.Tag = CStr(RecuperaValor(CadenaDesdeOtroForm, 3))  '0 NO es factura telefonia 1: Si
        
        CadenaDesdeOtroForm = ""
        
        
    Case 14
        PonerFrameVisible FrameFenollar
        
        If CadenaDesdeOtroForm = "ALZ" Then
            chkFenollar(0).Value = 0
            chkFenollar(1).Value = 1
        Else
            
            chkFenollar(0).Value = 1
            chkFenollar(0).visible = False
            chkFenollar(1).Value = 0
        End If
        cmdCancelar(14).visible = False
        cmdImprAlbaFenoll.Default = True
        PonerFocoBtn cmdImprAlbaFenoll
        CadenaDesdeOtroForm = ""
    Case 15
        PonerFrameVisible FrameCambioTipoAlbaran
        
    Case 16
        PonerFrameVisible FrameVtaPlazosBorrar
        
        Me.lblInd(2).Caption = ""
    Case 17
        PonerFrameVisible FramaTaxcoCambioKm
        For NumRegElim = 1 To 3
            txtNoEditable(NumRegElim - 1).Text = RecuperaValor(CadenaDesdeOtroForm, CInt(NumRegElim))
        Next
        txtNumero(1).Text = RecuperaValor(CadenaDesdeOtroForm, CInt(NumRegElim))  'es 4 numregelim
        CadenaDesdeOtroForm = ""
        
    '[Monica]22/10/2020:
    Case 100, 101
        PonerFrameVisible FrameHcoAlmacen
        lblInd(3).Caption = ""
        If Opcion = 100 Then
            lblTitulo(18).Caption = "Seleccionar Traspaso de Almac�n"
        End If
        
        
    End Select
    
    'If Opcion <> 3 Then cmdCancelar(Opcion).Cancel = True
    On Error Resume Next
    If IndexOpcion >= 0 Then cmdCancelar(IndexOpcion).Cancel = True
    
    If IndexOpcion = 14 Then PonerFocoBtn cmdImprAlbaFenoll
    
    
    If Err.Number <> 0 Then Err.Clear
End Sub



Private Sub PonerFrameVisible(Fr As Frame)
    Fr.visible = True
    Fr.Top = 0
    Fr.Left = 120
    Me.Height = Fr.Height + 480
    Me.Width = Fr.Width + 320
End Sub




Private Sub HacerImpresionFacturas()
Dim i As Integer
Dim fin As Boolean
    SePuedeCerrar = False
    
    Me.lblImpr(0).Caption = "Leyendo datos"
    lblImpr(0).Refresh
    Espera 0.25
    Set miRsAux = New ADODB.Recordset
    
    miRsAux.Open "Select count(*) from scafac WHERE " & CadenaDesdeOtroForm, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    If Not miRsAux.EOF Then NumRegElim = DBLet(miRsAux.Fields(0), "N")
    miRsAux.Close
    If NumRegElim = 0 Then Exit Sub
    
    CadenaDesdeOtroForm = "Select codtipom, numfactu, fecfactu, nomclien from scafac where " & CadenaDesdeOtroForm
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & " ORDER BY fecfactu,numfactu"
    
    miRsAux.Open CadenaDesdeOtroForm, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    i = 0
    fin = False
    While Not fin
        i = i + 1
        Me.lblImpr(1).Caption = "Fac. " & Format(miRsAux!Numfactu, "00000") & " de " & Format(miRsAux!FecFactu, "dd/mm/yyyy") & "     " & Mid(miRsAux!NomClien, 1, 20)
        lblImpr(1).Refresh
        Me.lblImpr(0).Caption = "Registro: " & i & "   de   " & NumRegElim
        lblImpr(0).Refresh
    
        'IMprimimos la factura
        ReImprimirDirectoFact " scafac.codtipom ='" & miRsAux!codtipom & "' AND scafac.numfactu = " & miRsAux!Numfactu
    
        DoEvents
        If SePuedeCerrar Then
            fin = True  'Han pulsado cancelar
        Else
            'Siguiente
            miRsAux.MoveNext
            fin = miRsAux.EOF
        End If
        If i Mod 50 = 25 Then Me.Refresh
            
        
    Wend
    If miRsAux.EOF Then
        'Significa que ha acabado toda la impresion. Con lo cual
        'pongo CadenaDesdeOtroForm="" para que el form de reimpresion lo cierre
        CadenaDesdeOtroForm = ""
    End If
    miRsAux.Close
    Set miRsAux = Nothing
    SePuedeCerrar = True
    Unload Me  'Y cierro
End Sub

Private Sub Form_Unload(Cancel As Integer)

    If Not SePuedeCerrar Then Cancel = 1
    
    
End Sub


Private Sub imgSel_Click(Index As Integer)

End Sub

Private Sub frmA_DatoSeleccionado(CadenaSeleccion As String)
    cad = CadenaSeleccion
End Sub

Private Sub frmC_Selec(vFecha As Date)
    cad = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmCli_DatoSeleccionado(CadenaSeleccion As String)
    cad = CadenaSeleccion
End Sub

Private Sub frmPr_DatoSeleccionado(CadenaSeleccion As String)
    cad = CadenaSeleccion
End Sub

Private Sub imgArticulo_Click(Index As Integer)
        cad = ""
        Set frmA = New frmBasico2
'        frmA.DesdeTPV = False
'        'frmA.DatosADevolverBusqueda3 = "@1@"
'        frmA.Show vbModal
        AyudaArticulos frmA, txtArt(Index)
        Set frmA = Nothing
        If cad <> "" Then
            Me.txtArt(Index).Text = RecuperaValor(cad, 1)
            Me.txtArtD(Index).Text = RecuperaValor(cad, 2)
        End If
End Sub

Private Sub imgCheck_Click(Index As Integer)
    If Index < 2 Then
        'LISTVIEW 1
        For NumRegElim = 1 To ListView1.ListItems.Count
            ListView1.ListItems(NumRegElim).Checked = Index = 1
        Next NumRegElim
        
    ElseIf Index < 4 Then
        For NumRegElim = 1 To ListView2.ListItems.Count
            ListView2.ListItems(NumRegElim).Checked = Index = 3
        Next NumRegElim
    ElseIf Index < 6 Then
        For NumRegElim = 1 To ListView5.ListItems.Count
            ListView5.ListItems(NumRegElim).Checked = Index = 5
        Next NumRegElim
    End If
End Sub

Private Sub imgCli_Click(Index As Integer)
    cad = ""
    Set frmCli = New frmBasico2
    AyudaClientes frmCli, txtCli(Index)
    Set frmCli = Nothing
    If cad <> "" Then
        txtCli(Index).Text = RecuperaValor(cad, 1)
        Me.txtCliD(Index).Text = RecuperaValor(cad, 2)
    End If
End Sub

Private Sub imgFecha_Click(Index As Integer)
    cad = ""
    Set frmC = New frmCal
    frmC.Fecha = Now
    If txtFecha(Index).Text <> "" Then frmC.Fecha = CDate(txtFecha(Index).Text)
    frmC.Show vbModal
    Set frmC = Nothing
    If cad <> "" Then txtFecha(Index).Text = cad
End Sub





Private Sub imgProv_Click(Index As Integer)
    cad = ""
'    Set frmPr = New frmComProveedores
'    frmPr.DatosADevolverBusqueda = "0|1|"
'    frmPr.Show vbModal
    Set frmPr = New frmBasico2
    AyudaProveedores frmPr, txtProv(Index)
    Set frmPr = Nothing
    If cad <> "" Then
        txtProv(Index).Text = RecuperaValor(cad, 1)
        Me.txtProvD(Index).Text = RecuperaValor(cad, 2)
    End If

End Sub
  
Private Sub ListView4_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    For NumRegElim = 1 To ListView4.ListItems.Count
        If ListView4.ListItems(NumRegElim).Text <> Item.Text Then ListView4.ListItems(NumRegElim).Checked = False
    Next
End Sub



Private Sub ListView6_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If Opcion = 101 Then
        CargarLineasMovimientosAlmacen Item.Text
    Else
        CargarLineasTraspasoAlmacen Item.Text
    End If
End Sub

Private Sub optElimFact_Click(Index As Integer)
    FrameNuevaFecFac.visible = Index = 0
End Sub

Private Sub Tv1_NodeCheck(ByVal Node As MSComctlLib.Node)
    If PrimeraVez1 < 2 Then Exit Sub   'Solo cuando ya este miostrado todo
    
    If Node.Parent Is Nothing Then
        'Nodo padre
        
    Else
        'Tiene padre
        
    End If
End Sub

Private Sub txtArt_GotFocus(Index As Integer)
    ConseguirFoco txtArt(Index), 3
End Sub

Private Sub txtArt_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtArt_LostFocus(Index As Integer)
Dim C As String

    txtArt(Index).Text = Trim(txtArt(Index).Text)
    If txtArt(Index).Text = "" Then
        C = ""
    Else
        C = DevuelveDesdeBDNew(conAri, "sartic", "nomartic", "codartic", txtArt(Index).Text, "T")
        If C = "" Then
            'El articulo no existe. SI fuera obligado ponerlo es aqui donde habria que poner el ocdigo
            
        End If
    End If
    txtArtD(Index).Text = C
End Sub



Private Sub EliminandoArticulos_Paso1()
Dim C As String
Dim SQL As String
Dim Aux As String
Dim nt As Integer
Dim J As Byte

    If Me.txtArt(0).Text <> "" Then SQL = SQL & " codartic >=" & DBSet(txtArt(0).Text, "T")
    If Me.txtArt(1).Text <> "" Then
        If SQL <> "" Then SQL = SQL & " AND "
        SQL = SQL & " codartic <=" & DBSet(txtArt(1).Text, "T")
    End If
    If SQL <> "" Then SQL = " WHERE " & SQL
    
     
    'El stock
    lblElim(0).Caption = "Almacenes"
    lblElim(0).Refresh
    C = "select codartic,sum(canstock) from salmac " & SQL & " group by codartic having sum(canstock) <> 0"
    miRsAux.Open C, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
         conn.Execute "DELETE FROM tmpnseries where codusu = " & vUsu.Codigo & " AND codartic = " & DBSet(miRsAux.Fields(0), "T")
        miRsAux.MoveNext
    Wend
    miRsAux.Close
     
    
    For J = 0 To 2
        DevuelveTablasBorre J, C, Aux, nt
        For NumRegElim = 1 To nt
            
            lblElim(0).Caption = RecuperaValor(Aux, CInt(NumRegElim)) & "   -"
            If J = 0 Then
                lblElim(0).Caption = lblElim(0).Caption & "Clientes"
            ElseIf J = 1 Then
                lblElim(0).Caption = lblElim(0).Caption & "Prove"
            Else
                lblElim(0).Caption = lblElim(0).Caption & "Varios"
            End If
            lblElim(0).Refresh
            
            
            miRsAux.Open "Select codartic from " & RecuperaValor(C, CInt(NumRegElim)) & SQL & " GROUP BY codartic", conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not miRsAux.EOF
                 conn.Execute "DELETE FROM tmpnseries where codusu = " & vUsu.Codigo & " AND codartic = " & DBSet(miRsAux.Fields(0), "T")
                miRsAux.MoveNext
            Wend
            miRsAux.Close
            Me.Refresh
        Next NumRegElim
    Next J
    
End Sub


Private Sub CargaColumnas()
Dim clmX As ColumnHeader

    ListView1.ListItems.Clear
    ListView1.ColumnHeaders.Clear
    
    Select Case Opcion

    Case 1
        Me.ListView1.Checkboxes = True
        Set clmX = ListView1.ColumnHeaders.Add()
        clmX.Text = "C�digo"
        clmX.Width = 2200
        Set clmX = ListView1.ColumnHeaders.Add()
        clmX.Text = "Descripci�n"
        clmX.Width = 3600
        
    End Select
    Me.FrameListArticulos.ZOrder 1  'QUe lo traiga al frente
End Sub


 



Private Sub txtCli_GotFocus(Index As Integer)
    ConseguirFoco txtCli(Index), 3
End Sub

Private Sub txtCli_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtCli_LostFocus(Index As Integer)
    cad = ""
    txtCli(Index).Text = Trim(txtCli(Index).Text)
    cad = ""
    If txtCli(Index).Text <> "" Then
        
        If PonerFormatoEntero(txtCli(Index)) Then
            cad = DevuelveDesdeBDNew(conAri, "sclien", "nomclien", "codclien", txtCli(Index).Text, "T")
            If cad = "" Then
                'El cliente no existe. SI fuera obligado ponerlo es aqui donde habria que poner el ocdigo
                If Index = 0 Then
                    MsgBox "No existe el cliente: " & txtCli(Index).Text, vbExclamation
                    txtCli(Index).Text = ""
                    PonerFoco txtCli(Index)
                End If
            End If
        Else
            txtCli(Index).Text = ""
        End If
    End If
    txtCliD(Index).Text = cad
End Sub

Private Sub txtDecimal_GotFocus(Index As Integer)
    ConseguirFoco txtDecimal(Index), 3
End Sub

Private Sub txtDecimal_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtDecimal_LostFocus(Index As Integer)
Dim b As Boolean
Dim Aux As Currency
Dim Tipo As Single
    
    If Index = 0 Or Index = 1 Then
        Tipo = 4 'Decimal
    End If
    
    txtDecimal(Index).Text = Trim(txtDecimal(Index).Text)
    If txtDecimal(Index).Text <> "" Then
        
        b = PonerFormatoDecimal(txtDecimal(Index), Tipo)
        If b Then
            If Index = 0 Or Index = 1 Then
                'hasta 99.99
                Aux = ImporteFormateado(txtDecimal(Index))
                cad = ""
                If Aux < 0 Then
                    cad = "Importe negativo"
                ElseIf Aux >= 100 Then
                    cad = "Descuentos debe ser menor que 100"
                End If
                If cad <> "" Then
                    MsgBox cad, vbExclamation
                    b = False
                End If
            End If
        End If
        If Not b Then
            txtDecimal(Index).Text = ""
            PonerFoco txtDecimal(Index)
        End If
    End If
End Sub

Private Sub txtFecha_GotFocus(Index As Integer)
    ConseguirFoco txtFecha(Index), 3
End Sub

Private Sub txtFecha_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtFecha_LostFocus(Index As Integer)
    txtFecha(Index).Text = Trim(txtFecha(Index).Text)
    If txtFecha(Index).Text <> "" Then
        cad = txtFecha(Index).Text
        If Not EsFechaOK(cad) Then
            MsgBox "Fecha incorrecta: " & txtFecha(Index).Text, vbExclamation
            txtFecha(Index).Text = ""
            PonerFoco txtFecha(Index)
        Else
            txtFecha(Index).Text = cad
        End If
    End If
End Sub


'En cadenadesdeotroform llevo las lformas posibles. Se trata de ir poniendolas en el list
Private Sub CargaFormasEnvioPosibles()
Dim i As Integer
    
    
    While CadenaDesdeOtroForm <> ""
        i = InStr(1, CadenaDesdeOtroForm, "|")
        If i = 0 Then
            CadenaDesdeOtroForm = ""
            cad = ""
        Else
            cad = Mid(CadenaDesdeOtroForm, 1, i)
            CadenaDesdeOtroForm = Mid(CadenaDesdeOtroForm, i + 1)
            
            cad = Replace(cad, "<", "|")
                        
        End If
        If cad <> "" Then
            
            i = RecuperaValor(cad, 1)
            
            cad = Format(i, "0000") & "      " & RecuperaValor(cad, 2) & "    (" & RecuperaValor(cad, 3) & ")    Desde :" & RecuperaValor(cad, 4) & " Kg"
            ListEnvio.AddItem cad
            
        End If
    Wend
    If ListEnvio.ListCount > 0 Then ListEnvio.Selected(0) = True
End Sub



Private Sub CargaClientes()
Dim IT
    On Error GoTo ECargaClientes
    Set miRsAux = New ADODB.Recordset
    
    
    
    miRsAux.Open "select sclien.codclien,nomclien from sclien " & CadenaDesdeOtroForm, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    CadenaDesdeOtroForm = ""
    While Not miRsAux.EOF
        Set IT = ListView2.ListItems.Add()
        IT.Text = Format(miRsAux!codClien, "0000")
        IT.SubItems(1) = miRsAux!NomClien
        IT.Checked = True
    
    
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Caption = ListView2.ListItems.Count
ECargaClientes:
    If Err.Number <> 0 Then MuestraError Err.Number
    Set miRsAux = Nothing
    
End Sub



Private Sub CargaZonas()
Dim N
    Tv1.Nodes.Clear
    cad = "select codzona,numpedcl,nomzonas from tmpsliped,szonas  where codzona=codzonas and tmpsliped.codusu="
    cad = cad & vUsu.Codigo & " group by 1,2 ORDER BY 1,2"
    NumRegElim = 0
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        If NumRegElim <> miRsAux!CodZona Then
            NumRegElim = miRsAux!CodZona
            cad = DBLet(miRsAux!nomzonas, "T")
            If cad = "" Then cad = "ERROR: " & NumRegElim
            Set N = Tv1.Nodes.Add(, , "C" & NumRegElim, cad)
            N.Checked = True
        End If
        
        'Probablemente insertaremos una linea por pedido
        
        
        
        'Siguiente
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
        
End Sub


Private Sub CargaZonasAlbaranTransporte()
Dim N
    Tv1.Nodes.Clear
    cad = "select  scaalb.codzonas,nomzonas from scaalb,szonas where scaalb.codzonas=szonas.codzonas and "
    cad = cad & CadenaDesdeOtroForm & " group by scaalb.codzonas ORDER BY scaalb.codzonas"
    NumRegElim = 0
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        If NumRegElim <> miRsAux!codzonas Then
            NumRegElim = miRsAux!codzonas
            cad = DBLet(miRsAux!nomzonas, "T")
            If cad = "" Then cad = "ERROR: " & NumRegElim
            Set N = Tv1.Nodes.Add(, , "C" & NumRegElim, cad)
            N.Checked = True
        End If
        
        'Probablemente insertaremos una linea por pedido
        
        
        
        'Siguiente
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    CadenaDesdeOtroForm = "NO" 'Para que si le da a cancelar NO haga nada en frmlistado2
End Sub




Private Function PonerRestoDatos() As Boolean
Dim Col As Collection  'Iremos metiendo datos para el calculo masivo de selects
Dim StoAlm As Currency
Dim StoTot As Currency
'Dim codArt As String
Dim Completos As String
Dim SQL As String

    PonerRestoDatos = False



    On Error GoTo EPonerRestoDatos
    Set miRsAux = New ADODB.Recordset
    lblInd(0).Caption = "Cargando articulos I"
    lblInd(0).Refresh
    
    cad = "select distinct(tmpsliped.codartic) from tmpsliped,sartic where tmpsliped.codartic=sartic.codartic"
    cad = cad & " and codusu = " & vUsu.Codigo & " and sartic.artvario=0  AND ctrstock = 1"
    miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    cad = ""
    Set Col = New Collection   'Cada 20 articulos hare los calculos
    While Not miRsAux.EOF
        
        NumRegElim = NumRegElim + 1
        cad = cad & ", " & DBSet(miRsAux!codArtic, "T")
        If NumRegElim > 20 Then
            cad = Mid(cad, 2)
            Col.Add cad
            cad = ""
            NumRegElim = 0
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If NumRegElim > 0 Then
        cad = Mid(cad, 2)
        Col.Add cad
    End If
    
    
    '------------------------------------------------------------
    'Datos para el cliente
    For NumRegElim = 1 To Col.Count
        lblInd(0).Caption = "Sotck y Ped " & NumRegElim & "/" & Col.Count
        lblInd(0).Refresh
    
        'STOCK
        cad = Col.Item(NumRegElim)
        cad = "(" & cad & ")"
        cad = "select codartic,sum(canstock) total from salmac where codartic IN " & cad
        'Julio19 HERBLECA NO pone almacen 4 y 20
        If vParamAplic.NumeroInstalacion = vbHerbelca Then cad = cad & " AND codalmac<>4 and codalmac<>20"
        cad = cad & " GROUP BY codartic"
        miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            cad = "UPDATE tmpsliped set stocktot= " & TransformaComasPuntos(DBLet(miRsAux!total, "N"))
            cad = cad & " WHERE codusu = " & vUsu.Codigo & " AND codartic = " & DBSet(miRsAux!codArtic, "T")
            conn.Execute cad
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        
        'Hacemos los pedidos proveedor
        cad = Col.Item(NumRegElim)
        cad = "(" & cad & ")"
        cad = "select codartic,sum(cantidad) tot,min(fecpedpr) fec from slippr,scappr where scappr.numpedpr=slippr.numpedpr and codartic in " & cad & " group by codartic"
        miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            cad = "UPDATE tmpsliped set cantpedprov = " & TransformaComasPuntos(DBLet(miRsAux!Tot, "N"))
            cad = cad & ", fecpedprov= '" & Format(miRsAux!fec, "dd/mm/yy") & "'"
            cad = cad & " WHERE codusu = " & vUsu.Codigo & " AND codartic = " & DBSet(miRsAux!codArtic, "T")
            conn.Execute cad
            miRsAux.MoveNext
        Wend
        miRsAux.Close
    Next
    Set Col = Nothing
    DoEvents
    'Stock almacen de ese almacen
    lblInd(0).Caption = "Cargando articulos II"
    lblInd(0).Refresh
    cad = "select tmpsliped.codartic,tmpsliped.codalmac from tmpsliped,sartic where tmpsliped.codartic=sartic.codartic  and codusu = " & vUsu.Codigo & " and sartic.artvario=0 and ctrstock = 1"
    cad = cad & " GROUP BY codartic,codalmac"
    miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    cad = ""
    Set Col = New Collection   'Cada 20 articulos hare los calculos
    While Not miRsAux.EOF
        NumRegElim = NumRegElim + 1
        cad = cad & ", (" & DBSet(miRsAux!codArtic, "T") & "," & miRsAux!codAlmac & ")"
        If NumRegElim > 20 Then
            cad = Mid(cad, 2)
            Col.Add cad
            cad = ""
            NumRegElim = 0
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If NumRegElim > 0 Then
        cad = Mid(cad, 2)
        Col.Add cad
    End If
    
    
    '------------------------------------------------------------
    For NumRegElim = 1 To Col.Count
        lblInd(0).Caption = "Sotck alm " & NumRegElim & "/" & Col.Count
        lblInd(0).Refresh
        'STOCK
        cad = Col.Item(NumRegElim)
        cad = "(" & cad & ")"
        cad = "select codartic,codalmac,canstock from salmac where (codartic,codalmac) IN " & cad
        miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            cad = "UPDATE tmpsliped set stockalm= " & TransformaComasPuntos(DBLet(miRsAux!CanStock, "N"))
            cad = cad & " WHERE codusu = " & vUsu.Codigo & " AND codartic = " & DBSet(miRsAux!codArtic, "T") & " AND codalmac = " & miRsAux!codAlmac
            conn.Execute cad
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        
    Next
    
    'SOLO LOS QUE TENGAN STOCK
    If Val(lblTitulo(5).Tag) = 1 Then
        lblInd(0).Caption = "Comprobando articulos con stock"
        lblInd(0).Refresh
        
       
        
        'NUEVO. Abril 2013
        ' Borraremos
        'Insertamos en una tmp los pedidos que vamos a borrar
        
        'Veremos todos los pedidos de los cuales todas las lineas de stock son cero
        cad = "select " & vUsu.Codigo & ",'PED',numpedcl,0 FROM tmpsliped inner join sartic on tmpsliped.codartic=sartic.codartic"
        cad = cad & " WHERE codusu = " & vUsu.Codigo & " and ctrstock=1  group by numpedcl having sum(coalesce(stocktot,0))=0"
        cad = "INSERT INTO tmpstockfec(codusu,codartic,codalmac,stock) " & cad
        conn.Execute cad
        
        'De todos los pedidos que vamos a eliminar quitare aquellos que ALGUNO de los articulos
        'sea de varios
        cad = "select numpedcl FROM tmpsliped inner join sartic on tmpsliped.codartic=sartic.codartic"
        cad = cad & " WHERE codusu = " & vUsu.Codigo & " and artvario=1 "
        cad = cad & " AND numpedcl IN (select codalmac FROM tmpstockfec WHERE codusu =" & vUsu.Codigo & ")"
        cad = cad & " group by  numpedcl"
        miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            cad = "Delete from tmpstockfec WHERE codusu = " & vUsu.Codigo & " AND codalmac=" & miRsAux!NumPedcl
            miRsAux.MoveNext
            conn.Execute cad
        Wend
        miRsAux.Close
        
        
        lblInd(0).Caption = "Eliminando datos"
        lblInd(0).Refresh
        cad = "DELETE FROM tmpsliped WHERE codusu = " & vUsu.Codigo
        cad = cad & " AND numpedcl IN (select codalmac FROM tmpstockfec WHERE codusu =" & vUsu.Codigo & ")"
        conn.Execute cad
        
        
        
        lblInd(0).Caption = "Comprobando servir completo"
        lblInd(0).Refresh
        cad = "select numpedcl from scaped where  servcomp=1 AND numpedcl in( select distinct(numpedcl) FROM tmpsliped inner join sartic on tmpsliped.codartic=sartic.codartic  "
        cad = cad & " WHERE codusu =  " & vUsu.Codigo & " and ctrstock=1 AND (stocktot <=0 or stocktot is NULL)  )"
        miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            cad = "Delete from tmpsliped WHERE codusu = " & vUsu.Codigo & " AND numpedcl=" & miRsAux!NumPedcl
            miRsAux.MoveNext
            conn.Execute cad
        Wend
        miRsAux.Close
        
''''''        'AHORA. Marzo 2013
''''''        '   Solo quitaremos aquellos pedidos que todas las lineas sean 0(de stcok)
''''''        Cad = "select numpedcl FROM tmpsliped inner join sartic on tmpsliped.codartic=sartic.codartic"
''''''        Cad = Cad & " WHERE codusu = " & vUsu.Codigo & " and ctrstock=1  group by 1 having sum(coalesce(stocktot,0))=0"
''''''        miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
''''''        While Not miRsAux.EOF
''''''            Cad = "Delete from tmpsliped WHERE codusu = " & vUsu.Codigo & " AND numpedcl=" & miRsAux!numpedcl
''''''            miRsAux.MoveNext
''''''            conn.Execute Cad
''''''        Wend
''''''        miRsAux.Close
''''''
''''''        'Los que no tengan marca de control de stock lo borro
''''''        Cad = "select numpedcl,tmpsliped.codartic FROM tmpsliped inner join sartic on tmpsliped.codartic=sartic.codartic"
''''''        Cad = Cad & " WHERE codusu = " & vUsu.Codigo & " and ctrstock=0  and coalesce(stocktot,0)=0"
''''''        miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
''''''        While Not miRsAux.EOF
''''''            Cad = "Delete from tmpsliped WHERE codusu = " & vUsu.Codigo & " AND numpedcl=" & miRsAux!numpedcl & " AND codartic=" & DBSet(miRsAux!codArtic, "T")
''''''            miRsAux.MoveNext
''''''            conn.Execute Cad
''''''        Wend
''''''        miRsAux.Close
''''''        Espera 0.5
''''''
''''''        lblInd(0).Caption = "Comprobando servir completo"
''''''        lblInd(0).Refresh
''''''
''''''        Cad = "select numpedcl from scaped where  servcomp=1 AND numpedcl in( select distinct(numpedcl) FROM tmpsliped inner join sartic on tmpsliped.codartic=sartic.codartic  "
''''''        Cad = Cad & " WHERE codusu =  " & vUsu.Codigo & " and ctrstock=1 AND (stocktot <=0 or stocktot is NULL)  )"
''''''        miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
''''''        While Not miRsAux.EOF
''''''            Cad = "Delete from tmpsliped WHERE codusu = " & vUsu.Codigo & " AND numpedcl=" & miRsAux!numpedcl
''''''            miRsAux.MoveNext
''''''            conn.Execute Cad
''''''        Wend
''''''        miRsAux.Close
        
         
    End If
    
    
    'marzo 2011
    'Sacaraemos las reservas de cliente. Es decir, lo que esta en pedido de cliente
    lblInd(0).Caption = "Reservas clientes"
    lblInd(0).Refresh
    cad = "select sliped.codalmac,sliped.codartic,sum(cantidad) cantii from sliped,sartic"
    cad = cad & " where sliped.codartic=sartic.codartic  and sartic.artvario=0 and ctrstock = 1"
    cad = cad & " AND (codalmac,sliped.codartic) IN (select codalmac,codartic FROM tmpsliped where codusu = " & vUsu.Codigo & ")"
    cad = cad & " GROUP BY codartic,codalmac"
    miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        lblInd(0).Caption = "Actualizando reservas"
        lblInd(0).Refresh
        While Not miRsAux.EOF
            cad = "UPDATE tmpsliped set referart= " & TransformaComasPuntos(DBLet(miRsAux!cantii, "N"))
            cad = cad & " WHERE codusu = " & vUsu.Codigo & " AND codartic = " & DBSet(miRsAux!codArtic, "T") & " AND codalmac = " & miRsAux!codAlmac
            conn.Execute cad
            miRsAux.MoveNext
        Wend
    End If
    miRsAux.Close
    
    
    
    
    PonerRestoDatos = True
        
        
    
EPonerRestoDatos:
    If Err.Number <> 0 Then MuestraError Err.Number, "PonerRestoDatos"
    Set miRsAux = Nothing
End Function


Private Sub CargaCombo(ByRef CBO As ComboBox)

    CBO.Clear
    If Opcion = 7 Then
        cad = "Select clasifica elcodigo,nombre elNombre from sfamiatipodto ORDER BY clasifica"
    End If
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        CBO.AddItem miRsAux!ElNombre
        CBO.ItemData(CBO.NewIndex) = miRsAux!ElCodigo
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
End Sub


'***************************************************************************************
'       D E S C U E N T O S
'***************************************************************************************
Private Sub RecorrerDtosUpdateando()
Dim RD As ADODB.Recordset

    Set RD = New ADODB.Recordset
'    Cad = "select * from sfamiadtos where clasifica=" & cboTipoDtoFamia.ItemData(cboTipoDtoFamia.ListIndex)
'    RD.Open Cad, conn, adOpenKeyset, adLockPessimistic, adCmdText  'KEYSET
'    'Ya he visto antes que algun dato tenia
        
    
    'NO pongo FECHA. Si me dice actualizar, actualizo
    cad = "Select * from sdtofm where codclien = " & txtCli(0).Text
    'cad = cad & " AND fechadto <= " & DBSet(txtFecha(2).Text, "F")
    cad = cad & " AND codmarca is null"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        If Not IsNull(miRsAux!Codfamia) Then
            Me.lblIndDto.Caption = DBLet(miRsAux!Codfamia, "N")
            Me.lblIndDto.Refresh
        
            RD.Find "codfamia = " & miRsAux!Codfamia, , adSearchForward, 1
            If Not RD.EOF Then
                cad = "UPDATE sdtofm set dtoline1 = " & DBSet(RD!dtoline1, "N", "N")
                cad = cad & " ,dtoline2 = " & DBSet(RD!dtoline2, "N", "N")
                cad = cad & " ,fechadto = " & DBSet(txtFecha(2).Text, "F")
                'WHERE
                cad = cad & " WHERE codclien = " & miRsAux!codClien
                cad = cad & " AND codfamia = " & miRsAux!Codfamia
                'Marca puede ser NULL
                cad = cad & " AND codmarca "
                If IsNull(miRsAux!codmarca) Then
                    cad = cad & " IS NULL"
                Else
                    cad = cad & "  = " & miRsAux!codmarca
                End If
                conn.Execute cad
                 
                
            End If
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
 
End Sub


Private Sub CargarFamiliaDtos()
Dim IT
    
    cad = DevuelveDesdeBD(conAri, "nomfamia", "sfamia", "codfamia", CadenaDesdeOtroForm)
    cad = Format(CadenaDesdeOtroForm, "0000") & " " & cad
    Label4(4).Caption = cad
    Label4(4).Tag = CadenaDesdeOtroForm  'codfamia
    
    cad = "select sfamiadtos.clasifica,nombre,dtoline1,dtoline2 from sfamiadtos,sfamiatipodto where"
    cad = cad & " sfamiadtos.clasifica=sfamiatipodto.clasifica and codfamia=" & CadenaDesdeOtroForm
    cad = cad & " Order by sfamiadtos.clasifica"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        
        Set IT = ListView3.ListItems.Add()
        IT.Text = miRsAux!clasifica
        IT.SubItems(1) = miRsAux!Nombre
        IT.SubItems(2) = Format(miRsAux!dtoline1, FormatoDescuento)
        IT.SubItems(3) = Format(miRsAux!dtoline2, FormatoDescuento)
        IT.Checked = True
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
End Sub

Private Sub CargaArticulosStockRectificado()
Dim Aux As String
    cad = "select sartic.codprove,nomprove,tmpstockfec.codartic,nomartic,codalmac,stock,preciouc,preciomp,preciost "
    cad = cad & " from tmpstockfec,sartic,sprove where tmpstockfec.codartic=sartic.codartic and sartic.codprove=sprove.codprove"
    cad = cad & " AND tmpstockfec.codusu=" & vUsu.Codigo & " ORDER BY codprove,codartic,codalmac"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = -1
    txtStock.Text = ""
    Aux = ""
    While Not miRsAux.EOF
        If NumRegElim <> miRsAux!Codprove Then
            'Pinto el proveedor y un linea
            If Aux <> "" Then
                txtStock.Text = txtStock.Text & Aux & vbCrLf
                Aux = ""
            End If
            If txtStock.Text <> "" Then txtStock.Text = txtStock.Text & vbCrLf & vbCrLf
            txtStock.Text = txtStock & miRsAux!nomprove & " (" & miRsAux!Codprove & ")" & vbCrLf
            txtStock.Text = txtStock.Text & String(40, "-")
            NumRegElim = miRsAux!Codprove
            cad = ""
        End If
        If cad <> miRsAux!codArtic Then
            If Aux <> "" Then
                txtStock.Text = txtStock.Text & Aux & vbCrLf
                Aux = ""
            End If
            txtStock.Text = txtStock.Text & vbCrLf & Space(5) & miRsAux!NomArtic & " -> " & miRsAux!codArtic
            txtStock.Text = txtStock.Text & "  (" & Format(DBLet(miRsAux!preciost, "N"), FormatoPrecio) & ":" & Format(DBLet(miRsAux!PrecioMP, "N"), FormatoPrecio) & ":" & Format(DBLet(miRsAux!precioUC, "N"), FormatoPrecio) & ")" & vbCrLf
            cad = miRsAux!codArtic
     
        End If
        CadenaDesdeOtroForm = Format(miRsAux!stock, FormatoCantidad)
        If Len(CadenaDesdeOtroForm) > 8 Then
            CadenaDesdeOtroForm = "#" & Mid(CadenaDesdeOtroForm, 1, 8)
        Else
            CadenaDesdeOtroForm = Right("       " & CadenaDesdeOtroForm, 8)
        End If
        If Aux = "" Then
            Aux = Space(10)
        Else
            If Len(Aux) > 81 Then
                If InStr(1, Aux, vbCrLf) = 0 Then Aux = Aux & vbCrLf & Space(10)
            End If
        End If
        Aux = Aux & Format("15/" & miRsAux!codAlmac & "/2010", "mmm") & ": " & CadenaDesdeOtroForm & " "
        
        
        
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    If Aux <> "" Then txtStock.Text = txtStock.Text & Aux & vbCrLf
                
End Sub


Private Sub txtNumero_GotFocus(Index As Integer)
    ConseguirFoco txtNumero(Index), 3
End Sub

Private Sub txtNumero_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtNumero_LostFocus(Index As Integer)
    cad = ""
    txtNumero(Index).Text = Trim(txtNumero(Index).Text)
    cad = ""
    If txtNumero(Index).Text <> "" Then
        
        If PonerFormatoEntero(txtNumero(Index)) Then
            
            If Index = 0 Then
                If Val(txtNumero(Index)) > 2100 Or Val(txtNumero(Index)) < 2000 Then
                    MsgBox "A�o incorrecto: " & txtNumero(Index).Text, vbExclamation
                    txtNumero(Index).Text = ""
                    PonerFoco txtNumero(Index)
                End If
            End If
        Else
            txtNumero(Index).Text = ""
        End If
    End If
    

End Sub

Private Sub txtProv_GotFocus(Index As Integer)
    ConseguirFoco txtProv(Index), 3
End Sub

Private Sub txtProv_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtProv_LostFocus(Index As Integer)
    cad = ""
    txtProv(Index).Text = Trim(txtProv(Index).Text)
    cad = ""
    If txtProv(Index).Text <> "" Then
        
        If PonerFormatoEntero(txtProv(Index)) Then
            cad = DevuelveDesdeBDNew(conAri, "sprove", "nomprove", "codprove", txtProv(Index).Text, "T")
            If cad = "" Then
                'El cliente no existe. SI fuera obligado ponerlo es aqui donde habria que poner el ocdigo
                If Index = 0 Then
                    MsgBox "No existe el proveedor: " & txtProv(Index).Text, vbExclamation
                    txtProv(Index).Text = ""
                    PonerFoco txtProv(Index)
                End If
            End If
        Else
            txtProv(Index).Text = ""
        End If
    End If
    txtProvD(Index).Text = cad
End Sub


Private Sub CargaTiposALbaran()
    Set miRsAux = New ADODB.Recordset
    cad = "'ALV','ALR','ALO','ALE'"
    cad = "Select * from stipom where codtipom in (" & cad & ") AND codtipom <>'" & CadenaDesdeOtroForm & "' ORDER BY 1"
    CadenaDesdeOtroForm = ""
    miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    While Not miRsAux.EOF
        NumRegElim = NumRegElim + 1
        ListView4.ListItems.Add , , miRsAux!codtipom
        ListView4.ListItems(NumRegElim).SubItems(1) = miRsAux!nomtipom
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    
End Sub

Private Sub CargaVtaPlazosFinalizados()
Dim IT
    On Error GoTo ECargaClientes
    Set miRsAux = New ADODB.Recordset
    
    lblInd(2).Caption = "Leyendo registros"
    lblInd(2).Refresh
    
    cad = "select IdTelefono ,nomclien,sclientfno.codclien,ImportePlazo,PlazosMeses,ArtPlazos,PlazosOrigen,NomArtic"
    cad = cad & " from sclien INNER JOIN sclientfno ON sclientfno.codclien=sclien.codclien"
    cad = cad & " left join sartic on sclientfno.ArtPlazos=sartic.codartic where ArtPlazos<>'' AND PlazosMeses =0 ORDER BY sclientfno.codclien,idtelefono"
    
    
    
    
    miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    While Not miRsAux.EOF
        Set IT = ListView5.ListItems.Add()
        IT.Text = miRsAux!idtelefono
        IT.SubItems(1) = Format(miRsAux!codClien, "0000")
        IT.SubItems(2) = miRsAux!NomClien
        IT.SubItems(3) = DBLet(miRsAux!PlazosOrigen, "N")
        
        'Texto para la elimincion
        
        cad = vbCrLf & " Articulo: " & miRsAux!artplazos & "  " & DBLet(miRsAux!NomArtic, "T")
        cad = cad & vbCrLf & " Plazos inicial: " & DBLet(miRsAux!PlazosOrigen, "N")
        cad = cad & "       Importe:  " & Format(miRsAux!ImportePlazo, FormatoImporte)
        
        IT.Tag = CStr("[FIN PLAZOS] " & cad)
        IT.ToolTipText = cad
        IT.Checked = True
    
    
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    
ECargaClientes:
    If Err.Number <> 0 Then MuestraError Err.Number
    Set miRsAux = Nothing
    lblInd(2).Caption = ""
    
    
End Sub


Private Sub CargaMovimientosAlmacen()
Dim SQL As String
Dim RSaux As ADODB.Recordset
Dim IT As ListItem
    
    On Error GoTo ECargaClientes
    
    
    
    ListView6.ColumnHeaders.Clear
    
    If Opcion = 101 Then
        ListView6.ColumnHeaders.Add , , "N�Movim.", 1150
        ListView6.ColumnHeaders.Add , , "Fecha", 1450, 0
        ListView6.ColumnHeaders.Add , , "Hora", 1100, 0
        ListView6.ColumnHeaders.Add , , "C�digo", 900, 0
        ListView6.ColumnHeaders.Add , , "Almacen", 2000, 0
        ListView6.ColumnHeaders.Add , , "Codigo", 1000, 0
        ListView6.ColumnHeaders.Add , , "Trabajador", 3200, 0
        
        ListView7.ColumnHeaders.Clear
        
        ListView7.ColumnHeaders.Add , , "Art�culo", 1200
        ListView7.ColumnHeaders.Add , , "Nombre", 3500, 0
        ListView7.ColumnHeaders.Add , , "Cantidad", 1500, 2
        ListView7.ColumnHeaders.Add , , "Situaci�n", 1200, 0
        ListView7.ColumnHeaders.Add , , "Observacion", 3400, 0
    Else
        ListView6.ColumnHeaders.Add , , "Traspaso", 1150
        ListView6.ColumnHeaders.Add , , "Fecha", 1450, 0
        ListView6.ColumnHeaders.Add , , "Hora", 1100, 0
        ListView6.ColumnHeaders.Add , , "Alm.Origen", 2000, 0
        ListView6.ColumnHeaders.Add , , "Alm.Destino", 2000, 0
        ListView6.ColumnHeaders.Add , , "Trabajador", 3100, 0
        
        ListView7.ColumnHeaders.Clear
        
        ListView7.ColumnHeaders.Add , , "Art�culo", 1200
        ListView7.ColumnHeaders.Add , , "Nombre", 4500, 0
        ListView7.ColumnHeaders.Add , , "Cantidad", 1500, 2
        ListView7.ColumnHeaders.Add , , "Observacion", 3600, 0
    
    
    End If
    
    
    Set miRsAux = New ADODB.Recordset
    
    ListView6.ListItems.Clear
    
    If Check1.Value = 0 Then
        SQL = RecuperaValor(CADENA, 1)
    Else
        SQL = RecuperaValor(CADENA, 1) & " " & RecuperaValor(CADENA, 2)
    End If
    
    If Opcion = 101 Then
        SQL = SQL & " order by fecmovim desc, 1 desc "
    Else
        SQL = SQL & " order by fechatra desc, 1 desc "
    End If
    
    
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    CadenaDesdeOtroForm = ""
    While Not miRsAux.EOF
        Set IT = ListView6.ListItems.Add()
        If Opcion = 101 Then
            IT.Text = Format(miRsAux!codMovim, "0000000")
            IT.SubItems(1) = Format(miRsAux!fecmovim, "dd/mm/yyyy")
            IT.SubItems(2) = Format(miRsAux!hormovim, "hh:mm:ss")
            IT.SubItems(3) = Format(miRsAux!codAlmac, "000")
            IT.SubItems(4) = miRsAux!nomalmac
            IT.SubItems(5) = Format(miRsAux!CodTraba, "000000")
            IT.SubItems(6) = miRsAux!NomTraba
        Else
            IT.Text = Format(miRsAux!codtrasp, "0000000")
            IT.SubItems(1) = Format(miRsAux!fechatra, "dd/mm/yyyy")
            IT.SubItems(2) = Format(miRsAux!hormovim, "hh:mm:ss")
            IT.SubItems(3) = miRsAux!nomalmori
            IT.SubItems(4) = miRsAux!nomalmdes
            IT.SubItems(5) = miRsAux!NomTraba
        
        End If
        IT.Checked = True
    
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If ListView6.ListItems.Count > 0 Then
        If Opcion = 101 Then
            CargarLineasMovimientosAlmacen ListView6.ListItems(1).Text
        Else
            CargarLineasTraspasoAlmacen ListView6.ListItems(1).Text
        End If
    End If
    ListView6.SelectedItem = ListView6.ListItems(1)
    PonerFocoLw ListView6
    
    Caption = "Movimientos de Almacen: " & ListView6.ListItems.Count
    
    
ECargaClientes:
    If Err.Number <> 0 Then MuestraError Err.Number
    Set miRsAux = Nothing
End Sub


Private Sub CargarLineasMovimientosAlmacen(nromovim As String)
Dim SQL As String
Dim RSaux As ADODB.Recordset
Dim IT As ListItem

    ListView7.ListItems.Clear


    SQL = "select slhmov.codartic, sartic.nomartic, slhmov.cantidad, if(slhmov.tipomovi = 0,'S','E') tipo, slhmov.motimovi "
    SQL = SQL & " from (slhmov inner join sartic on slhmov.codartic = sartic.codartic) "
    SQL = SQL & " where slhmov.codmovim = " & DBSet(nromovim, "N")
    
    Set miRsAux = New ADODB.Recordset
    
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    CadenaDesdeOtroForm = ""
    While Not miRsAux.EOF
        Set IT = ListView7.ListItems.Add()
        IT.Text = miRsAux!codArtic
        IT.SubItems(1) = miRsAux!NomArtic
        IT.SubItems(2) = Format(miRsAux!cantidad, "###,##0.00")
        IT.SubItems(3) = miRsAux!Tipo
        IT.SubItems(4) = DBLet(miRsAux!motimovi, "T")
        
        IT.Checked = True
    
        miRsAux.MoveNext
    Wend
    miRsAux.Close

End Sub

Private Sub CargarLineasTraspasoAlmacen(nromovim As String)
Dim SQL As String
Dim RSaux As ADODB.Recordset
Dim IT As ListItem

    ListView7.ListItems.Clear


    SQL = "select slhtra.codartic, sartic.nomartic, slhtra.cantidad, slhtra.observa2 "
    SQL = SQL & " from (slhtra inner join sartic on slhtra.codartic = sartic.codartic) "
    SQL = SQL & " where slhtra.codtrasp = " & DBSet(nromovim, "N")
    
    Set miRsAux = New ADODB.Recordset
    
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    CadenaDesdeOtroForm = ""
    While Not miRsAux.EOF
        Set IT = ListView7.ListItems.Add()
        IT.Text = miRsAux!codArtic
        IT.SubItems(1) = miRsAux!NomArtic
        IT.SubItems(2) = Format(miRsAux!cantidad, "###,##0.00")
        IT.SubItems(3) = DBLet(miRsAux!observa2, "T")
        
        IT.Checked = True
    
        miRsAux.MoveNext
    Wend
    miRsAux.Close


End Sub

