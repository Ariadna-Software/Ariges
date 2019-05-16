VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmVarios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Varios"
   ClientHeight    =   7125
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11565
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   11565
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameFenollar 
      Height          =   2415
      Left            =   5760
      TabIndex        =   118
      Top             =   3360
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
         Text            =   "frmVarios.frx":0000
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
         Caption         =   "Año"
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
         Caption         =   "Estadísticas consultas artículo / cliente"
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
         Caption         =   "Impresión facturas"
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
         Picture         =   "frmVarios.frx":0006
         ToolTipText     =   "Quitar al haber"
         Top             =   6480
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   1
         Left            =   600
         Picture         =   "frmVarios.frx":0150
         ToolTipText     =   "Puntear al haber"
         Top             =   6480
         Width           =   240
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Eliminar artículos"
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
         Picture         =   "frmVarios.frx":029A
         ToolTipText     =   "Puntear al haber"
         Top             =   6480
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   2
         Left            =   240
         Picture         =   "frmVarios.frx":03E4
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
   Begin VB.Frame FrameFamDto 
      Height          =   3255
      Left            =   840
      TabIndex        =   68
      Top             =   240
      Visible         =   0   'False
      Width           =   5775
      Begin VB.CommandButton cmdActuDtoFamMar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   2880
         TabIndex        =   74
         Top             =   2760
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Salir"
         Height          =   375
         Index           =   8
         Left            =   4200
         TabIndex        =   70
         Top             =   2760
         Width           =   1095
      End
      Begin MSComctlLib.ListView ListView3 
         Height          =   1455
         Left            =   120
         TabIndex        =   71
         Top             =   1080
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   2566
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Tipo"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Desc"
            Object.Width           =   4057
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Dto1"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Dto2"
            Object.Width           =   1411
         EndProperty
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
         Index           =   4
         Left            =   120
         TabIndex        =   73
         Top             =   720
         Width           =   5400
      End
      Begin VB.Label lblInd 
         Caption         =   "Label2"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   72
         Top             =   2880
         Width           =   2895
      End
      Begin VB.Label lblTitulo 
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
         Index           =   7
         Left            =   120
         TabIndex        =   69
         Top             =   240
         Width           =   5475
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
         Text            =   "frmVarios.frx":052E
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
         Caption         =   "Artículos stock rectificado"
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
         Caption         =   "Eliminar artículos"
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
         Caption         =   "Eliminar - reest. albarán"
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
End
Attribute VB_Name = "frmVarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Opcion As Byte
    ' 0.-   Impresion de facturas directas (tipo 4tonda)
    ' 1.-   Eliminar articulos masiva
    ' 2.-   Estadisticas consultas (archivo-facturacion-pedidos-consulta precio/cliente

    ' 3.-   Eleccion del metodo de envio para los albaranes

    ' 4.-   Ver clientes para añadir acciones comerciales
    
    ' 5.-   Pedidos x Zona. Selecionar las zonas
    ' 6.-   Albaranes trnasporte x codzona (dentro del albaran)
    
    ' 7.-   Generar dtos familia/marca/cliente
    
    ' 8.-   Actualizar en sdtofm para una familia y unos tipos(que se mostraran)
    
    ' 9.-   Modific los descuentos de todas las lineas de un pedido de proveedor
    
    ' 10.-  Listado de articulos rectificados por stock
        
    ' 11.-  Compras por año/proveedor
    
    ' 12.-  Ver rentings con fechas mal para facturar
    
    ' 13.-  Eliminar factura  pasando a albaranes  -  Cambiar fecha factura
    
    
    ' 14.  Imprimersion albaranes FENOLLAR
    
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmA As frmAlmArticu2
Attribute frmA.VB_VarHelpID = -1
Private WithEvents frmCli As frmFacClientes3
Attribute frmCli.VB_VarHelpID = -1
Private WithEvents frmPr As frmComProveedores
Attribute frmPr.VB_VarHelpID = -1


Private cad As String
Private SePuedeCerrar As Boolean   'Puede llevar DoEvents
Private PrimeraVez1 As Byte   '0.- Primera vez, 1.- cargando datos en forma_activate  2.- Fin carga




Private Sub cboTipoDtoFamia_KeyPress(KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub cmdAccionListview_Click()
Dim T1 As Single

    CadenaDesdeOtroForm = ""
    For NumRegElim = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(NumRegElim).Checked Then CadenaDesdeOtroForm = CadenaDesdeOtroForm & "1"
    Next
    
    If CadenaDesdeOtroForm = "" Then
        MsgBox "Seleccione algun artículo para eliminar", vbInformation
    Else
        CadenaDesdeOtroForm = Len(CadenaDesdeOtroForm)
        CadenaDesdeOtroForm = "Va a eliminar " & CadenaDesdeOtroForm & " artículo(s).   ¿Continuar?"
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
                ListView1.ListItems.Remove ListView1.ListItems(NumRegElim).index
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
        MsgBox "Ningún valor seleccionado", vbExclamation
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
            cad = cad & "¿Continuar el proceso?"
            
            If MsgBox(cad, vbExclamation + vbYesNo) <> vbYes Then Exit Sub
        
        End If
    
        
        cad = "establecer como fecha factura: " & Me.txtFecha(3).Text
    Else
    
        If FrameNuevaFecFac.Tag = "1" Then
            MsgBox "No se puede deshacer factura de telefonía", vbExclamation
            Exit Sub
        End If
    
        cad = "eliminar factura y reestablecer los albaranes facturados"
        
    End If
    cad = "Va a " & cad & vbCrLf & vbCrLf & vbCrLf
    cad = cad & " NO se realizaran acciones sobre Arimoney ni Ariconta " & vbCrLf & vbCrLf
    cad = cad & " **** Se grabará el registro de acciones *** " & vbCrLf
    cad = cad & vbCrLf & vbCrLf & "Introduzca el password para continuar"
    
    
    If vParamAplic.NumeroInstalacion = vbFenollar Then
        If MsgBox("¿Continuar con el proceso?", vbQuestion + vbYesNoCancel) = vbYes Then cad = "ARIADNA"
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

Private Sub cmdCancelar_Click(index As Integer)
    
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
    
    'El campo AÑO es obligarotorio
    txtNumero(0).Text = Trim(txtNumero(0).Text)
    If txtNumero(0).Text = "" Then
        MsgBox "Debe seleccionar una año para el informe.", vbInformation
        Exit Sub
    End If
    campo = "year({scafpc.fecrecep})"
    campo = campo & " = " & Me.txtNumero(0).Text
    cad = cad & "pAnyo=""" & "Año: " & txtNumero(0).Text & """|"
    
 
    
    'Campo seleccion de un CLIENTE
    txtProv(0).Text = Trim(txtProv(0).Text)
    If txtProv(0).Text <> "" Then
        campo = campo & " AND ({scafpc.codprove} =" & txtProv(0).Text & ")"
        'Pasar el cliente solicitado como parametro
        cad = cad & "pDHCliente=""" & "Proveedor: " & txtProv(0).Text & " - " & txtProvD(0).Text & """|"
    Else
        'Mostrar en el informe el total del Año Anterior
        campo = "(" & campo & " OR year({scafpc.fecrecep}) = " & CInt(txtNumero(0).Text) - 1 & ")"
        
        cad = cad & "pDHCliente=""" & "Proveedores: Todos" & """|"
    End If
    
    
    

    If Not HayRegParaInforme("scafpc", campo) Then Exit Sub
    
    
    'Borro los datos temporales,por si acaso se hubiera quedado
    BorrarTempInformes
    
    'Generar la temporal con los totales por año, mes y cliente (tmpinformes)
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
    
    'Ajustamos los tamaños para cargar el LISTVIEW
    CargaColumnas
    NumRegElim = (Screen.Width - FrameListArticulos.Width - 420) \ 2
    Me.Left = NumRegElim
    NumRegElim = (Screen.Height - FrameListArticulos.Height - 360) \ 2
    Me.Top = NumRegElim
    Me.FrameDHArticulo.visible = False
    PonerFrameVisible Me.FrameListArticulos
    Me.lblTitulo(1).Caption = "Eliminar artículos"
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
Dim I As Integer

    If ListEnvio.ListIndex < 0 Then Exit Sub
    cad = ListEnvio.List(ListEnvio.ListIndex)
    I = InStrRev(cad, "(")
    cad = Trim(Mid(cad, I + 1))
    I = InStrRev(cad, ")")
    cad = Mid(cad, 1, I - 1) 'quitamos el ultmio parentesis
    CadenaDesdeOtroForm = cad
    
    I = InStrRev(ListEnvio.List(ListEnvio.ListIndex), "(")
    cad = Mid(ListEnvio.List(ListEnvio.ListIndex), 1, I - 1)  'quito el precio kilo
    
    I = Val(Mid(cad, 1, 10))
    
    cad = Trim(Mid(cad, 11))
    
    CadenaDesdeOtroForm = I & "|" & cad & "|" & CadenaDesdeOtroForm & "|"
    
    'Desde kilo
    cad = ListEnvio.List(ListEnvio.ListIndex)
    I = InStrRev(ListEnvio.List(ListEnvio.ListIndex), "Desde :")
    cad = Mid(cad, I + 7)
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
Dim I As Byte

    If Tv1.Nodes.Count = 0 Then Exit Sub
    'Nos recorreemos el tv1 por si a desmarcado alguno
    'Lo borraremos de la tabla tmpsliped
    lblInd(0).Caption = "Preparando datos"
    lblInd(0).Refresh
    
    I = 0
    Set N = Tv1.Nodes(1)
    While Not N Is Nothing
        NumRegElim = -1
        If Not N.Checked Then
            NumRegElim = N.index
            conn.Execute "DELETE from tmpsliped where codusu = " & vUsu.Codigo & " AND codzona = " & Mid(N.Key, 2)
        Else
            I = 1
        End If
        
        Set N = N.Next
        If NumRegElim > 0 Then Tv1.Nodes.Remove NumRegElim
    Wend
    
    If I = 0 Then
        MsgBox "Nada seleccionado", vbExclamation
        lblInd(0).Caption = ""
        Exit Sub
    End If
    
    'Haremos los selectsum de las salmac
    Screen.MousePointer = vbHourglass
    Me.FrameZona.Enabled = False
    I = 0
    If PonerRestoDatos Then I = 1
        
    Me.FrameZona.Enabled = True
    Screen.MousePointer = vbDefault
    lblInd(0).Caption = ""
    If I = 1 Then
        CadenaDesdeOtroForm = DevuelveDesdeBD(conAri, "count(*)", "tmpsliped", "codusu", CStr(vUsu.Codigo))
        If CadenaDesdeOtroForm = "" Then CadenaDesdeOtroForm = "0"
        If Val(CadenaDesdeOtroForm) = 0 Then
            MsgBox "Ningun dato para mostrar", vbExclamation
            I = 0
        End If
    End If
    If I = 1 Then
        If PonerParamRPT2(48, cad, I, CadenaDesdeOtroForm, False, "", 0) Then
            With frmImprimir
                .ConSubInforme = False
                .FormulaSeleccion = "{tmpsliped.codusu} = " & vUsu.Codigo
                .NombreRPT = CadenaDesdeOtroForm
                .NombrePDF = pPdfRpt
                .Titulo = "List. pedidos por zonas"
                .OtrosParametros = cad
                .NumeroParametros = CInt(I)
                .Opcion = 2003 'Esta libre
                
                I = vParamAplic.NumCop_PedZona
                .NumeroCopias = I
                .Show vbModal
            End With
        End If
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

Private Sub Form_activate()
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
        End Select
        PrimeraVez1 = 2
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub CargarIconos()
Dim I As Image


    For Each I In Me.imgArticulo
         I.Picture = frmPpal.imgListComun.ListImages(19).Picture
         I.ToolTipText = "Articulo"
    Next
    For Each I In Me.imgCli
         I.Picture = frmPpal.imgListComun.ListImages(19).Picture
         I.ToolTipText = "Cliente"
    Next
    For Each I In Me.imgFecha
         I.Picture = frmPpal.imgListComun.ListImages(23).Picture
         I.ToolTipText = "fecha"
    Next
    For Each I In Me.imgProv
         I.Picture = frmPpal.imgListComun.ListImages(19).Picture
         I.ToolTipText = "Proveedor"
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
Dim I As Integer
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
    I = 0
    fin = False
    While Not fin
        I = I + 1
        Me.lblImpr(1).Caption = "Fac. " & Format(miRsAux!Numfactu, "00000") & " de " & Format(miRsAux!FecFactu, "dd/mm/yyyy") & "     " & Mid(miRsAux!NomClien, 1, 20)
        lblImpr(1).Refresh
        Me.lblImpr(0).Caption = "Registro: " & I & "   de   " & NumRegElim
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
        If I Mod 50 = 25 Then Me.Refresh
            
        
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


Private Sub imgSel_Click(index As Integer)

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

Private Sub imgArticulo_Click(index As Integer)
        cad = ""
        Set frmA = New frmAlmArticu2
        frmA.DesdeTPV = False
        'frmA.DatosADevolverBusqueda3 = "@1@"
        frmA.Show vbModal
        Set frmA = Nothing
        If cad <> "" Then
            Me.txtArt(index).Text = RecuperaValor(cad, 1)
            Me.txtArtD(index).Text = RecuperaValor(cad, 2)
        End If
End Sub

Private Sub imgCheck_Click(index As Integer)
    If index < 2 Then
        'LISTVIEW 1
        For NumRegElim = 1 To ListView1.ListItems.Count
            ListView1.ListItems(NumRegElim).Checked = index = 1
        Next NumRegElim
        
    Else
        For NumRegElim = 1 To ListView2.ListItems.Count
            ListView2.ListItems(NumRegElim).Checked = index = 3
        Next NumRegElim
    End If
End Sub

Private Sub imgCli_Click(index As Integer)
    cad = ""
    Set frmCli = New frmFacClientes3
    frmCli.DatosADevolverBusqueda = "0|1|"
    frmCli.Show vbModal
    Set frmCli = Nothing
    If cad <> "" Then
        txtCli(index).Text = RecuperaValor(cad, 1)
        Me.txtCliD(index).Text = RecuperaValor(cad, 2)
    End If
End Sub

Private Sub imgFecha_Click(index As Integer)
    cad = ""
    Set frmC = New frmCal
    frmC.Fecha = Now
    If txtFecha(index).Text <> "" Then frmC.Fecha = CDate(txtFecha(index).Text)
    frmC.Show vbModal
    Set frmC = Nothing
    If cad <> "" Then txtFecha(index).Text = cad
End Sub





Private Sub imgProv_Click(index As Integer)
    cad = ""
    Set frmPr = New frmComProveedores
    frmPr.DatosADevolverBusqueda = "0|1|"
    frmPr.Show vbModal
    Set frmPr = Nothing
    If cad <> "" Then
        txtProv(index).Text = RecuperaValor(cad, 1)
        Me.txtProvD(index).Text = RecuperaValor(cad, 2)
    End If

End Sub
  
Private Sub optElimFact_Click(index As Integer)
    FrameNuevaFecFac.visible = index = 0
End Sub

Private Sub Tv1_NodeCheck(ByVal Node As MSComctlLib.Node)
    If PrimeraVez1 < 2 Then Exit Sub   'Solo cuando ya este miostrado todo
    
    If Node.Parent Is Nothing Then
        'Nodo padre
        
    Else
        'Tiene padre
        
    End If
End Sub

Private Sub txtArt_GotFocus(index As Integer)
    ConseguirFoco txtArt(index), 3
End Sub

Private Sub txtArt_KeyPress(index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtArt_LostFocus(index As Integer)
Dim C As String

    txtArt(index).Text = Trim(txtArt(index).Text)
    If txtArt(index).Text = "" Then
        C = ""
    Else
        C = DevuelveDesdeBDNew(conAri, "sartic", "nomartic", "codartic", txtArt(index).Text, "T")
        If C = "" Then
            'El articulo no existe. SI fuera obligado ponerlo es aqui donde habria que poner el ocdigo
            
        End If
    End If
    txtArtD(index).Text = C
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
        clmX.Text = "Código"
        clmX.Width = 2200
        Set clmX = ListView1.ColumnHeaders.Add()
        clmX.Text = "Descripción"
        clmX.Width = 3600
        
    End Select
    Me.FrameListArticulos.ZOrder 1  'QUe lo traiga al frente
End Sub


 



Private Sub txtCli_GotFocus(index As Integer)
    ConseguirFoco txtCli(index), 3
End Sub

Private Sub txtCli_KeyPress(index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtCli_LostFocus(index As Integer)
    cad = ""
    txtCli(index).Text = Trim(txtCli(index).Text)
    cad = ""
    If txtCli(index).Text <> "" Then
        
        If PonerFormatoEntero(txtCli(index)) Then
            cad = DevuelveDesdeBDNew(conAri, "sclien", "nomclien", "codclien", txtCli(index).Text, "T")
            If cad = "" Then
                'El cliente no existe. SI fuera obligado ponerlo es aqui donde habria que poner el ocdigo
                If index = 0 Then
                    MsgBox "No existe el cliente: " & txtCli(index).Text, vbExclamation
                    txtCli(index).Text = ""
                    PonerFoco txtCli(index)
                End If
            End If
        Else
            txtCli(index).Text = ""
        End If
    End If
    txtCliD(index).Text = cad
End Sub

Private Sub txtDecimal_GotFocus(index As Integer)
    ConseguirFoco txtDecimal(index), 3
End Sub

Private Sub txtDecimal_KeyPress(index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtDecimal_LostFocus(index As Integer)
Dim b As Boolean
Dim Aux As Currency
Dim Tipo As Single
    
    If index = 0 Or index = 1 Then
        Tipo = 4 'Decimal
    End If
    
    txtDecimal(index).Text = Trim(txtDecimal(index).Text)
    If txtDecimal(index).Text <> "" Then
        
        b = PonerFormatoDecimal(txtDecimal(index), Tipo)
        If b Then
            If index = 0 Or index = 1 Then
                'hasta 99.99
                Aux = ImporteFormateado(txtDecimal(index))
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
            txtDecimal(index).Text = ""
            PonerFoco txtDecimal(index)
        End If
    End If
End Sub

Private Sub txtFecha_GotFocus(index As Integer)
    ConseguirFoco txtFecha(index), 3
End Sub

Private Sub txtFecha_KeyPress(index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtFecha_LostFocus(index As Integer)
    txtFecha(index).Text = Trim(txtFecha(index).Text)
    If txtFecha(index).Text <> "" Then
        cad = txtFecha(index).Text
        If Not EsFechaOK(cad) Then
            MsgBox "Fecha incorrecta: " & txtFecha(index).Text, vbExclamation
            txtFecha(index).Text = ""
            PonerFoco txtFecha(index)
        Else
            txtFecha(index).Text = cad
        End If
    End If
End Sub


'En cadenadesdeotroform llevo las lformas posibles. Se trata de ir poniendolas en el list
Private Sub CargaFormasEnvioPosibles()
Dim I As Integer
    
    
    While CadenaDesdeOtroForm <> ""
        I = InStr(1, CadenaDesdeOtroForm, "|")
        If I = 0 Then
            CadenaDesdeOtroForm = ""
            cad = ""
        Else
            cad = Mid(CadenaDesdeOtroForm, 1, I)
            CadenaDesdeOtroForm = Mid(CadenaDesdeOtroForm, I + 1)
            
            cad = Replace(cad, "<", "|")
                        
        End If
        If cad <> "" Then
            
            I = RecuperaValor(cad, 1)
            
            cad = Format(I, "0000") & "      " & RecuperaValor(cad, 2) & "    (" & RecuperaValor(cad, 3) & ")    Desde :" & RecuperaValor(cad, 4) & " Kg"
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
        cad = "select codartic,sum(canstock) total from salmac where codartic IN " & cad & " GROUP BY codartic"
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


Private Sub txtNumero_GotFocus(index As Integer)
    ConseguirFoco txtNumero(index), 3
End Sub

Private Sub txtNumero_KeyPress(index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtNumero_LostFocus(index As Integer)
    cad = ""
    txtNumero(index).Text = Trim(txtNumero(index).Text)
    cad = ""
    If txtNumero(index).Text <> "" Then
        
        If PonerFormatoEntero(txtNumero(index)) Then
            
            If index = 0 Then
                If Val(txtNumero(index)) > 2100 Or Val(txtNumero(index)) < 2000 Then
                    MsgBox "Año incorrecto: " & txtNumero(index).Text, vbExclamation
                    txtNumero(index).Text = ""
                    PonerFoco txtNumero(index)
                End If
            End If
        Else
            txtNumero(index).Text = ""
        End If
    End If
    

End Sub

Private Sub txtProv_GotFocus(index As Integer)
    ConseguirFoco txtProv(index), 3
End Sub

Private Sub txtProv_KeyPress(index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtProv_LostFocus(index As Integer)
    cad = ""
    txtProv(index).Text = Trim(txtProv(index).Text)
    cad = ""
    If txtProv(index).Text <> "" Then
        
        If PonerFormatoEntero(txtProv(index)) Then
            cad = DevuelveDesdeBDNew(conAri, "sprove", "nomprove", "codprove", txtProv(index).Text, "T")
            If cad = "" Then
                'El cliente no existe. SI fuera obligado ponerlo es aqui donde habria que poner el ocdigo
                If index = 0 Then
                    MsgBox "No existe el proveedor: " & txtProv(index).Text, vbExclamation
                    txtProv(index).Text = ""
                    PonerFoco txtProv(index)
                End If
            End If
        Else
            txtProv(index).Text = ""
        End If
    End If
    txtProvD(index).Text = cad
End Sub


