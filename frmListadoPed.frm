VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmListadoPed 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   7725
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   12645
   Icon            =   "frmListadoPed.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7725
   ScaleWidth      =   12645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FramepedidoAnulado 
      Height          =   7335
      Left            =   5400
      TabIndex        =   338
      Top             =   0
      Visible         =   0   'False
      Width           =   7455
      Begin VB.CheckBox chkVarios 
         Caption         =   "Detallar"
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
         Left            =   600
         TabIndex        =   368
         Top             =   6360
         Value           =   1  'Checked
         Width           =   2085
      End
      Begin VB.TextBox txtNombre 
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
         Index           =   77
         Left            =   2550
         Locked          =   -1  'True
         TabIndex        =   366
         Text            =   "Text5"
         Top             =   5760
         Width           =   4515
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   77
         Left            =   1440
         MaxLength       =   6
         TabIndex        =   347
         Top             =   5760
         Width           =   1095
      End
      Begin VB.TextBox txtNombre 
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
         Index           =   76
         Left            =   2550
         Locked          =   -1  'True
         TabIndex        =   364
         Text            =   "Text5"
         Top             =   5280
         Width           =   4515
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   76
         Left            =   1440
         MaxLength       =   6
         TabIndex        =   346
         Top             =   5280
         Width           =   1095
      End
      Begin VB.TextBox txtNombre 
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
         Index           =   75
         Left            =   2550
         Locked          =   -1  'True
         TabIndex        =   361
         Text            =   "Text5"
         Top             =   4440
         Width           =   4515
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   75
         Left            =   1440
         MaxLength       =   6
         TabIndex        =   345
         Top             =   4440
         Width           =   1095
      End
      Begin VB.TextBox txtNombre 
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
         Index           =   74
         Left            =   2550
         Locked          =   -1  'True
         TabIndex        =   359
         Text            =   "Text5"
         Top             =   3960
         Width           =   4515
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   74
         Left            =   1440
         MaxLength       =   6
         TabIndex        =   344
         Top             =   3960
         Width           =   1095
      End
      Begin VB.TextBox txtNombre 
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
         Index           =   73
         Left            =   2790
         Locked          =   -1  'True
         TabIndex        =   356
         Text            =   "Text5"
         Top             =   3000
         Width           =   4275
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   73
         Left            =   1440
         MaxLength       =   6
         TabIndex        =   343
         Top             =   3000
         Width           =   1335
      End
      Begin VB.TextBox txtNombre 
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
         Index           =   72
         Left            =   2790
         Locked          =   -1  'True
         TabIndex        =   354
         Text            =   "Text5"
         Top             =   2520
         Width           =   4275
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   72
         Left            =   1440
         MaxLength       =   6
         TabIndex        =   342
         Top             =   2520
         Width           =   1335
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   71
         Left            =   4515
         MaxLength       =   10
         TabIndex        =   341
         Top             =   1440
         Width           =   1260
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   70
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   340
         Top             =   1440
         Width           =   1260
      End
      Begin VB.CommandButton cmdPedidosAnulados 
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
         Left            =   4560
         TabIndex        =   348
         Top             =   6600
         Width           =   1065
      End
      Begin VB.CommandButton cmdCancel 
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
         Index           =   80
         Left            =   5760
         TabIndex        =   349
         Top             =   6600
         Width           =   1065
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   1
         Left            =   1140
         Picture         =   "frmListadoPed.frx":000C
         Top             =   1440
         Width           =   240
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   52
         Left            =   1140
         Top             =   5760
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
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
         Index           =   88
         Left            =   480
         TabIndex        =   367
         Top             =   5760
         Width           =   645
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   51
         Left            =   1140
         Top             =   5280
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
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
         Index           =   87
         Left            =   480
         TabIndex        =   365
         Top             =   5280
         Width           =   645
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Incidencia"
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
         Height          =   240
         Index           =   86
         Left            =   240
         TabIndex        =   363
         Top             =   4920
         Width           =   1095
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   50
         Left            =   1140
         Top             =   4440
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
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
         Index           =   85
         Left            =   480
         TabIndex        =   362
         Top             =   4440
         Width           =   570
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   49
         Left            =   1140
         Top             =   3960
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
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
         Index           =   84
         Left            =   480
         TabIndex        =   360
         Top             =   3960
         Width           =   645
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Familia"
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
         Height          =   240
         Index           =   83
         Left            =   240
         TabIndex        =   358
         Top             =   3600
         Width           =   780
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
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
         Index           =   82
         Left            =   480
         TabIndex        =   357
         Top             =   3000
         Width           =   660
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   48
         Left            =   1140
         Top             =   3000
         Width           =   240
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   47
         Left            =   1140
         Top             =   2520
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
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
         Index           =   81
         Left            =   480
         TabIndex        =   355
         Top             =   2520
         Width           =   645
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Proveedor"
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
         Height          =   240
         Index           =   80
         Left            =   240
         TabIndex        =   353
         Top             =   2160
         Width           =   1125
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   2
         Left            =   4215
         Picture         =   "frmListadoPed.frx":0097
         Top             =   1440
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
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
         Index           =   79
         Left            =   3600
         TabIndex        =   352
         Top             =   1440
         Width           =   570
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
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
         Index           =   78
         Left            =   480
         TabIndex        =   351
         Top             =   1440
         Width           =   690
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha pedido"
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
         Height          =   240
         Index           =   77
         Left            =   240
         TabIndex        =   350
         Top             =   1080
         Width           =   1425
      End
      Begin VB.Label Label9 
         Caption         =   "Pedidos anulados"
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
         Left            =   2280
         TabIndex        =   339
         Top             =   360
         Width           =   4335
      End
   End
   Begin VB.Frame FramePreFacturar 
      Height          =   7095
      Left            =   120
      TabIndex        =   68
      Top             =   120
      Width           =   7035
      Begin VB.TextBox txtCodigo 
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
         Index           =   61
         Left            =   840
         MaxLength       =   6
         TabIndex        =   85
         Top             =   6120
         Width           =   780
      End
      Begin VB.Frame Frame16 
         Height          =   975
         Left            =   240
         TabIndex        =   296
         Top             =   3480
         Visible         =   0   'False
         Width           =   6255
         Begin VB.CheckBox chkTpPago2 
            Caption         =   "Efectivo"
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
            Left            =   1320
            TabIndex        =   303
            Top             =   240
            Value           =   1  'Checked
            Width           =   1140
         End
         Begin VB.CheckBox chkTpPago2 
            Caption         =   "Transferencia"
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
            Left            =   2520
            TabIndex        =   302
            Top             =   240
            Value           =   1  'Checked
            Width           =   1665
         End
         Begin VB.CheckBox chkTpPago2 
            Caption         =   "Talón"
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
            Left            =   4260
            TabIndex        =   301
            Top             =   240
            Value           =   1  'Checked
            Width           =   885
         End
         Begin VB.CheckBox chkTpPago2 
            Caption         =   "Pagaré"
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
            Left            =   5160
            TabIndex        =   300
            Top             =   240
            Value           =   1  'Checked
            Width           =   1020
         End
         Begin VB.CheckBox chkTpPago2 
            Caption         =   "Recibo bancario"
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
            Left            =   870
            TabIndex        =   299
            Top             =   600
            Value           =   1  'Checked
            Width           =   1890
         End
         Begin VB.CheckBox chkTpPago2 
            Caption         =   "Confirming"
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
            Left            =   2865
            TabIndex        =   298
            Top             =   600
            Value           =   1  'Checked
            Width           =   1335
         End
         Begin VB.CheckBox chkTpPago2 
            Caption         =   "Tarjeta crédito"
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
            Index           =   16
            Left            =   4245
            TabIndex        =   297
            Top             =   600
            Value           =   1  'Checked
            Width           =   1875
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Tipo pago"
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
            Height          =   240
            Index           =   12
            Left            =   120
            TabIndex        =   304
            Top             =   240
            Width           =   1050
         End
      End
      Begin VB.ComboBox cmbTipAlbaran 
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
         ItemData        =   "frmListadoPed.frx":0122
         Left            =   2535
         List            =   "frmListadoPed.frx":012F
         Style           =   2  'Dropdown List
         TabIndex        =   86
         Top             =   6120
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CheckBox chkResumenForpa 
         Caption         =   "Resumen forma de pago"
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
         Left            =   3840
         TabIndex        =   84
         Top             =   5490
         Width           =   2790
      End
      Begin VB.CheckBox chkSoloFacturar 
         Caption         =   "Solo Albaranes para facturar"
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
         Left            =   465
         TabIndex        =   83
         Top             =   5400
         Value           =   1  'Checked
         Width           =   3165
      End
      Begin VB.Frame Frame7 
         Caption         =   "Tipo Informe"
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
         Height          =   690
         Left            =   270
         TabIndex        =   188
         Top             =   4560
         Width           =   6255
         Begin VB.OptionButton OptDetalle 
            Caption         =   "Con IVA"
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
            Left            =   4815
            TabIndex        =   82
            Top             =   270
            Width           =   1230
         End
         Begin VB.OptionButton OptDetalle 
            Caption         =   "Facturacion"
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
            Left            =   2985
            TabIndex        =   81
            Top             =   270
            Width           =   1530
         End
         Begin VB.OptionButton OptDetalle 
            Caption         =   "Resumen"
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
            Left            =   1575
            TabIndex        =   80
            Top             =   270
            Width           =   1260
         End
         Begin VB.OptionButton OptDetalle 
            Caption         =   "Detalle"
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
            Left            =   360
            TabIndex        =   79
            Top             =   270
            Width           =   1005
         End
      End
      Begin VB.Frame Frame6 
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   360
         TabIndex        =   182
         Top             =   2520
         Width           =   6135
         Begin VB.TextBox txtCodigo 
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
            Index           =   33
            Left            =   1350
            MaxLength       =   6
            TabIndex        =   78
            Top             =   600
            Width           =   780
         End
         Begin VB.TextBox txtNombre 
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
            Index           =   33
            Left            =   2130
            Locked          =   -1  'True
            TabIndex        =   184
            Text            =   "Text5"
            Top             =   600
            Width           =   3885
         End
         Begin VB.TextBox txtCodigo 
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
            Index           =   32
            Left            =   1350
            MaxLength       =   6
            TabIndex        =   77
            Top             =   240
            Width           =   780
         End
         Begin VB.TextBox txtNombre 
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
            Index           =   32
            Left            =   2130
            Locked          =   -1  'True
            TabIndex        =   183
            Text            =   "Text5"
            Top             =   240
            Width           =   3885
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
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
            Index           =   17
            Left            =   360
            TabIndex        =   187
            Top             =   600
            Width           =   660
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Desde"
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
            Index           =   0
            Left            =   360
            TabIndex        =   186
            Top             =   240
            Width           =   690
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   19
            Left            =   1050
            Top             =   615
            Width           =   240
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Agente"
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
            Height          =   240
            Index           =   38
            Left            =   0
            TabIndex        =   185
            Top             =   0
            Width           =   765
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   18
            Left            =   1050
            Top             =   255
            Width           =   240
         End
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   26
         Left            =   1740
         MaxLength       =   10
         TabIndex        =   71
         Top             =   960
         Width           =   1260
      End
      Begin VB.CommandButton cmdAceptarPreFac 
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
         Left            =   4425
         TabIndex        =   87
         Top             =   6480
         Width           =   1065
      End
      Begin VB.CommandButton cmdCancel 
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
         Index           =   5
         Left            =   5595
         TabIndex        =   88
         Top             =   6480
         Width           =   1065
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   27
         Left            =   4035
         MaxLength       =   10
         TabIndex        =   72
         Top             =   960
         Width           =   1260
      End
      Begin VB.TextBox txtNombre 
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
         Index           =   30
         Left            =   2370
         Locked          =   -1  'True
         TabIndex        =   90
         Text            =   "Text5"
         Top             =   2760
         Width           =   4005
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   30
         Left            =   1710
         MaxLength       =   3
         TabIndex        =   75
         Top             =   2760
         Width           =   660
      End
      Begin VB.TextBox txtNombre 
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
         Index           =   31
         Left            =   2370
         Locked          =   -1  'True
         TabIndex        =   89
         Text            =   "Text5"
         Top             =   3120
         Width           =   4005
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   31
         Left            =   1710
         MaxLength       =   3
         TabIndex        =   76
         Top             =   3120
         Width           =   660
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   29
         Left            =   1740
         MaxLength       =   6
         TabIndex        =   74
         Top             =   2040
         Width           =   870
      End
      Begin VB.TextBox txtNombre 
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
         Index           =   29
         Left            =   2610
         Locked          =   -1  'True
         TabIndex        =   70
         Text            =   "Text5"
         Top             =   2040
         Width           =   3795
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   28
         Left            =   1740
         MaxLength       =   6
         TabIndex        =   73
         Top             =   1680
         Width           =   870
      End
      Begin VB.TextBox txtNombre 
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
         Index           =   28
         Left            =   2610
         Locked          =   -1  'True
         TabIndex        =   69
         Text            =   "Text5"
         Top             =   1680
         Width           =   3795
      End
      Begin VB.Image imgayuda 
         Height          =   240
         Index           =   1
         Left            =   4080
         ToolTipText     =   "Previsión facturacion"
         Top             =   6600
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Período facturación"
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
         Height          =   240
         Index           =   24
         Left            =   360
         TabIndex        =   305
         Top             =   5880
         Width           =   2085
      End
      Begin VB.Label lblTipAlbaran 
         Caption         =   "Tipo albarán"
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
         Index           =   0
         Left            =   2535
         TabIndex        =   228
         Top             =   5880
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
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
         Index           =   44
         Left            =   3120
         TabIndex        =   100
         Top             =   960
         Width           =   570
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   10
         Left            =   1440
         Picture         =   "frmListadoPed.frx":0162
         Top             =   960
         Width           =   240
      End
      Begin VB.Label Label9 
         Caption         =   "Prefacturación de Albaranes"
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
         Left            =   360
         TabIndex        =   99
         Top             =   240
         Width           =   6375
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Albaran"
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
         Height          =   240
         Index           =   43
         Left            =   390
         TabIndex        =   98
         Top             =   720
         Width           =   1515
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
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
         Index           =   42
         Left            =   780
         TabIndex        =   97
         Top             =   960
         Width           =   690
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   11
         Left            =   3735
         Picture         =   "frmListadoPed.frx":01ED
         Top             =   960
         Width           =   240
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   16
         Left            =   1410
         Top             =   2775
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Formas de Pago"
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
         Height          =   240
         Index           =   41
         Left            =   390
         TabIndex        =   96
         Top             =   2520
         Width           =   1740
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
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
         Index           =   40
         Left            =   705
         TabIndex        =   95
         Top             =   2760
         Width           =   690
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   17
         Left            =   1410
         Top             =   3135
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
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
         Index           =   39
         Left            =   705
         TabIndex        =   94
         Top             =   3120
         Width           =   660
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
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
         Index           =   35
         Left            =   780
         TabIndex        =   93
         Top             =   2040
         Width           =   615
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   15
         Left            =   1440
         Top             =   2040
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
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
         Index           =   34
         Left            =   780
         TabIndex        =   92
         Top             =   1680
         Width           =   645
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
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
         Height          =   240
         Index           =   33
         Left            =   390
         TabIndex        =   91
         Top             =   1440
         Width           =   765
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   14
         Left            =   1440
         Top             =   1680
         Width           =   240
      End
   End
   Begin VB.Frame FramePreFacMante 
      Height          =   7575
      Left            =   0
      TabIndex        =   103
      Top             =   0
      Width           =   8580
      Begin VB.Frame Frame2 
         Height          =   615
         Index           =   1
         Left            =   360
         TabIndex        =   237
         Top             =   6240
         Width           =   8070
         Begin VB.TextBox txtCodigo 
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
            Index           =   54
            Left            =   1995
            MaxLength       =   4
            TabIndex        =   122
            Top             =   195
            Width           =   900
         End
         Begin VB.TextBox txtNombre 
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
            Index           =   54
            Left            =   2940
            Locked          =   -1  'True
            TabIndex        =   239
            Text            =   "Text5"
            Top             =   195
            Width           =   4995
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   54
            Left            =   1650
            Top             =   240
            Width           =   240
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Centro Coste"
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
            Height          =   240
            Index           =   0
            Left            =   120
            TabIndex        =   238
            Top             =   240
            Width           =   1425
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1395
         Index           =   0
         Left            =   360
         TabIndex        =   134
         Top             =   840
         Width           =   8070
         Begin VB.TextBox txtNombre 
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
            Index           =   52
            Left            =   3285
            Locked          =   -1  'True
            TabIndex        =   150
            Text            =   "Text5"
            Top             =   615
            Width           =   4710
         End
         Begin VB.TextBox txtCodigo 
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
            Index           =   52
            Left            =   2445
            MaxLength       =   6
            TabIndex        =   112
            Top             =   615
            Width           =   780
         End
         Begin VB.TextBox txtCodigo 
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
            Index           =   47
            Left            =   2445
            MaxLength       =   4
            TabIndex        =   113
            Top             =   990
            Width           =   780
         End
         Begin VB.TextBox txtNombre 
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
            Index           =   47
            Left            =   3285
            Locked          =   -1  'True
            TabIndex        =   135
            Text            =   "Text5"
            Top             =   990
            Width           =   4710
         End
         Begin VB.TextBox txtCodigo 
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
            Index           =   44
            Left            =   2445
            MaxLength       =   10
            TabIndex        =   111
            Top             =   240
            Width           =   1260
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Cta. Prev. Cobro"
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
            Height          =   240
            Index           =   8
            Left            =   120
            TabIndex        =   152
            Top             =   600
            Width           =   1785
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   31
            Left            =   2130
            Top             =   615
            Width           =   240
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Operador"
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
            Height          =   240
            Index           =   5
            Left            =   120
            TabIndex        =   137
            Top             =   945
            Width           =   1020
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   26
            Left            =   2130
            Top             =   990
            Width           =   240
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   15
            Left            =   2130
            Picture         =   "frmListadoPed.frx":0278
            Top             =   240
            Width           =   240
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Facturación"
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
            Height          =   240
            Index           =   2
            Left            =   120
            TabIndex        =   136
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.Frame Frame1 
         Height          =   3960
         Left            =   360
         TabIndex        =   105
         Top             =   2280
         Width           =   8070
         Begin VB.Frame FrameTapa 
            Height          =   2895
            Left            =   30
            TabIndex        =   273
            Top             =   960
            Visible         =   0   'False
            Width           =   7980
            Begin VB.ComboBox cboSituMan 
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
               ItemData        =   "frmListadoPed.frx":0303
               Left            =   1455
               List            =   "frmListadoPed.frx":0310
               Style           =   2  'Dropdown List
               TabIndex        =   274
               Top             =   720
               Width           =   2895
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "Situación"
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
               Height          =   240
               Index           =   15
               Left            =   120
               TabIndex        =   275
               Top             =   720
               Width           =   975
            End
         End
         Begin VB.CheckBox chkSituFacMant 
            Caption         =   "Situación facturación"
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
            Left            =   4980
            TabIndex        =   272
            Top             =   600
            Width           =   2970
         End
         Begin VB.TextBox txtCodigo 
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
            Index           =   56
            Left            =   1380
            MaxLength       =   6
            TabIndex        =   121
            Top             =   3480
            Width           =   915
         End
         Begin VB.TextBox txtNombre 
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
            Index           =   56
            Left            =   2340
            Locked          =   -1  'True
            TabIndex        =   270
            Text            =   "Text5"
            Top             =   3480
            Width           =   5580
         End
         Begin VB.TextBox txtNombre 
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
            Index           =   55
            Left            =   2340
            Locked          =   -1  'True
            TabIndex        =   267
            Text            =   "Text5"
            Top             =   3120
            Width           =   5580
         End
         Begin VB.TextBox txtCodigo 
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
            Index           =   55
            Left            =   1380
            MaxLength       =   6
            TabIndex        =   120
            Top             =   3120
            Width           =   915
         End
         Begin VB.TextBox txtNombre 
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
            Index           =   50
            Left            =   2340
            Locked          =   -1  'True
            TabIndex        =   130
            Text            =   "Text5"
            Top             =   2160
            Width           =   5580
         End
         Begin VB.TextBox txtCodigo 
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
            Index           =   50
            Left            =   1380
            MaxLength       =   3
            TabIndex        =   118
            Top             =   2160
            Width           =   915
         End
         Begin VB.TextBox txtNombre 
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
            Index           =   51
            Left            =   2340
            Locked          =   -1  'True
            TabIndex        =   129
            Text            =   "Text5"
            Top             =   2520
            Width           =   5580
         End
         Begin VB.TextBox txtCodigo 
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
            Index           =   51
            Left            =   1380
            MaxLength       =   3
            TabIndex        =   119
            Top             =   2520
            Width           =   915
         End
         Begin VB.TextBox txtCodigo 
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
            Index           =   49
            Left            =   1380
            MaxLength       =   6
            TabIndex        =   117
            Top             =   1560
            Width           =   915
         End
         Begin VB.TextBox txtNombre 
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
            Index           =   49
            Left            =   2340
            Locked          =   -1  'True
            TabIndex        =   125
            Text            =   "Text5"
            Top             =   1560
            Width           =   5580
         End
         Begin VB.TextBox txtCodigo 
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
            Index           =   48
            Left            =   1380
            MaxLength       =   6
            TabIndex        =   116
            Top             =   1200
            Width           =   915
         End
         Begin VB.TextBox txtNombre 
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
            Index           =   48
            Left            =   2340
            Locked          =   -1  'True
            TabIndex        =   110
            Text            =   "Text5"
            Top             =   1200
            Width           =   5580
         End
         Begin VB.TextBox txtCodigo 
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
            Index           =   46
            Left            =   2085
            MaxLength       =   2
            TabIndex        =   115
            Top             =   600
            Width           =   825
         End
         Begin VB.TextBox txtNombre 
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
            Index           =   46
            Left            =   2925
            Locked          =   -1  'True
            TabIndex        =   108
            Text            =   "Text5"
            Top             =   600
            Width           =   1920
         End
         Begin VB.TextBox txtCodigo 
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
            Index           =   45
            Left            =   2085
            MaxLength       =   3
            TabIndex        =   114
            Top             =   240
            Width           =   825
         End
         Begin VB.TextBox txtNombre 
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
            Index           =   45
            Left            =   2925
            Locked          =   -1  'True
            TabIndex        =   106
            Text            =   "Text5"
            Top             =   240
            Width           =   5040
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
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
            Left            =   420
            TabIndex        =   271
            Top             =   3480
            Width           =   570
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   39
            Left            =   1080
            Top             =   3480
            Width           =   240
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Departamento"
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
            Height          =   240
            Index           =   13
            Left            =   120
            TabIndex        =   269
            Top             =   2880
            Width           =   1590
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   38
            Left            =   1080
            Top             =   3120
            Width           =   240
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Desde"
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
            Left            =   420
            TabIndex        =   268
            Top             =   3120
            Width           =   600
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   29
            Left            =   1080
            Top             =   2160
            Width           =   240
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Formas de Pago"
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
            Height          =   240
            Index           =   7
            Left            =   120
            TabIndex        =   133
            Top             =   1920
            Width           =   1740
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Desde"
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
            Left            =   420
            TabIndex        =   132
            Top             =   2160
            Width           =   600
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   30
            Left            =   1080
            Top             =   2520
            Width           =   240
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
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
            Left            =   420
            TabIndex        =   131
            Top             =   2520
            Width           =   570
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
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
            Index           =   9
            Left            =   420
            TabIndex        =   128
            Top             =   1560
            Width           =   570
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   28
            Left            =   1080
            Top             =   1560
            Width           =   240
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Desde"
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
            Index           =   8
            Left            =   420
            TabIndex        =   127
            Top             =   1200
            Width           =   600
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Cliente"
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
            Height          =   240
            Index           =   6
            Left            =   120
            TabIndex        =   126
            Top             =   960
            Width           =   765
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   27
            Left            =   1080
            Top             =   1200
            Width           =   240
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Mes a facturar"
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
            Height          =   240
            Index           =   4
            Left            =   120
            TabIndex        =   109
            Top             =   600
            Width           =   1545
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Contrato"
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
            Height          =   240
            Index           =   3
            Left            =   120
            TabIndex        =   107
            Top             =   240
            Width           =   1470
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   25
            Left            =   1785
            Top             =   240
            Width           =   240
         End
      End
      Begin VB.CommandButton cmdCancel 
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
         Index           =   7
         Left            =   7380
         TabIndex        =   124
         Top             =   6960
         Width           =   1065
      End
      Begin VB.CommandButton cmdAceptarPreFacMan 
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
         Left            =   6210
         TabIndex        =   123
         Top             =   6960
         Width           =   1065
      End
      Begin VB.Label lblFactMant 
         Caption         =   "Label5"
         Height          =   375
         Left            =   360
         TabIndex        =   231
         Top             =   6960
         Width           =   5625
      End
      Begin VB.Label Label7 
         Caption         =   "Prefacturación Mantenimientos"
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
         Left            =   360
         TabIndex        =   104
         Top             =   360
         Width           =   6375
      End
   End
   Begin VB.Frame FrameFacturar 
      Height          =   7575
      Left            =   90
      TabIndex        =   101
      Top             =   0
      Width           =   8565
      Begin VB.Frame FramTaxcoTrabajador 
         Height          =   615
         Left            =   180
         TabIndex        =   335
         Top             =   1800
         Visible         =   0   'False
         Width           =   8205
         Begin VB.TextBox txtNombre 
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
            Index           =   69
            Left            =   3510
            Locked          =   -1  'True
            TabIndex        =   337
            Text            =   "Text5"
            Top             =   195
            Width           =   4590
         End
         Begin VB.TextBox txtCodigo 
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
            Index           =   69
            Left            =   2655
            MaxLength       =   6
            TabIndex        =   140
            Top             =   195
            Width           =   825
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   46
            Left            =   2325
            Top             =   240
            Width           =   240
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Trabajador"
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
            Height          =   240
            Index           =   9
            Left            =   240
            TabIndex        =   336
            Top             =   240
            Width           =   1185
         End
      End
      Begin VB.Frame Frame15 
         Height          =   975
         Left            =   180
         TabIndex        =   236
         Top             =   5205
         Width           =   8205
         Begin VB.CheckBox chkTpPago2 
            Caption         =   "Tarjeta crédito"
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
            Left            =   5580
            TabIndex        =   295
            Top             =   600
            Value           =   1  'Checked
            Width           =   1800
         End
         Begin VB.CheckBox chkTpPago2 
            Caption         =   "Confirming"
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
            Left            =   3780
            TabIndex        =   294
            Top             =   600
            Value           =   1  'Checked
            Width           =   1650
         End
         Begin VB.CheckBox chkTpPago2 
            Caption         =   "Recibo bancario"
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
            Left            =   1575
            TabIndex        =   293
            Top             =   600
            Value           =   1  'Checked
            Width           =   1890
         End
         Begin VB.CheckBox chkTpPago2 
            Caption         =   "Pagaré"
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
            Left            =   6180
            TabIndex        =   292
            Top             =   240
            Value           =   1  'Checked
            Width           =   1110
         End
         Begin VB.CheckBox chkTpPago2 
            Caption         =   "Talón"
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
            TabIndex        =   291
            Top             =   240
            Value           =   1  'Checked
            Width           =   1290
         End
         Begin VB.CheckBox chkTpPago2 
            Caption         =   "Transferencia"
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
            Left            =   2985
            TabIndex        =   290
            Top             =   240
            Value           =   1  'Checked
            Width           =   1890
         End
         Begin VB.CheckBox chkTpPago2 
            Caption         =   "Efectivo"
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
            Left            =   1575
            TabIndex        =   289
            Top             =   240
            Value           =   1  'Checked
            Width           =   1125
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Tipo pago"
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
            Height          =   240
            Index           =   11
            Left            =   240
            TabIndex        =   288
            Top             =   240
            Width           =   1050
         End
      End
      Begin VB.Frame FrameProgress 
         Height          =   1050
         Left            =   180
         TabIndex        =   214
         Top             =   6240
         Visible         =   0   'False
         Width           =   5820
         Begin MSComctlLib.ProgressBar PB_Fact 
            Height          =   345
            Left            =   120
            TabIndex        =   215
            Top             =   600
            Width           =   5505
            _ExtentX        =   9710
            _ExtentY        =   609
            _Version        =   393216
            Appearance      =   1
         End
         Begin VB.Label lblProgess 
            Caption         =   "Iniciando el proceso ..."
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   217
            Top             =   345
            Width           =   5460
         End
         Begin VB.Label lblProgess 
            Caption         =   "Facturando:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   216
            Top             =   135
            Width           =   5340
         End
      End
      Begin VB.Frame Frame4 
         Height          =   3345
         Left            =   180
         TabIndex        =   158
         Top             =   1800
         Width           =   8205
         Begin VB.TextBox txtNombre 
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
            Index           =   42
            Left            =   3450
            Locked          =   -1  'True
            TabIndex        =   172
            Text            =   "Text5"
            Top             =   2520
            Width           =   4635
         End
         Begin VB.TextBox txtCodigo 
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
            Index           =   42
            Left            =   2670
            MaxLength       =   6
            TabIndex        =   148
            Top             =   2520
            Width           =   780
         End
         Begin VB.TextBox txtNombre 
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
            Index           =   43
            Left            =   3450
            Locked          =   -1  'True
            TabIndex        =   171
            Text            =   "Text5"
            Top             =   2880
            Width           =   4635
         End
         Begin VB.TextBox txtCodigo 
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
            Index           =   43
            Left            =   2670
            MaxLength       =   6
            TabIndex        =   149
            Top             =   2880
            Width           =   780
         End
         Begin VB.TextBox txtCodigo 
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
            Index           =   41
            Left            =   2670
            MaxLength       =   6
            TabIndex        =   147
            Top             =   2040
            Width           =   915
         End
         Begin VB.TextBox txtNombre 
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
            Index           =   41
            Left            =   3585
            Locked          =   -1  'True
            TabIndex        =   167
            Text            =   "Text5"
            Top             =   2040
            Width           =   4500
         End
         Begin VB.TextBox txtCodigo 
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
            Index           =   40
            Left            =   2670
            MaxLength       =   6
            TabIndex        =   146
            Top             =   1680
            Width           =   915
         End
         Begin VB.TextBox txtNombre 
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
            Index           =   40
            Left            =   3585
            Locked          =   -1  'True
            TabIndex        =   166
            Text            =   "Text5"
            Top             =   1680
            Width           =   4500
         End
         Begin VB.TextBox txtCodigo 
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
            Index           =   38
            Left            =   3045
            MaxLength       =   10
            TabIndex        =   144
            Top             =   1200
            Width           =   1350
         End
         Begin VB.TextBox txtCodigo 
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
            Index           =   39
            Left            =   5490
            MaxLength       =   10
            TabIndex        =   145
            Top             =   1200
            Width           =   1350
         End
         Begin VB.TextBox txtCodigo 
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
            Index           =   36
            Left            =   3135
            MaxLength       =   10
            TabIndex        =   142
            Top             =   720
            Width           =   1260
         End
         Begin VB.TextBox txtCodigo 
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
            Index           =   37
            Left            =   5580
            MaxLength       =   10
            TabIndex        =   143
            Top             =   720
            Width           =   1260
         End
         Begin VB.TextBox txtCodigo 
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
            Index           =   35
            Left            =   3615
            MaxLength       =   10
            TabIndex        =   141
            Top             =   240
            Width           =   780
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   22
            Left            =   2325
            Top             =   2520
            Width           =   240
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Forma pago"
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
            Height          =   240
            Index           =   3
            Left            =   240
            TabIndex        =   175
            Top             =   2520
            Width           =   1305
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Desde"
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
            Index           =   48
            Left            =   1620
            TabIndex        =   174
            Top             =   2520
            Width           =   645
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   23
            Left            =   2325
            Top             =   2880
            Width           =   240
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
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
            Index           =   49
            Left            =   1620
            TabIndex        =   173
            Top             =   2880
            Width           =   615
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   20
            Left            =   2325
            Top             =   1680
            Width           =   240
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   21
            Left            =   2325
            Top             =   2040
            Width           =   240
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
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
            Index           =   50
            Left            =   1620
            TabIndex        =   170
            Top             =   2040
            Width           =   615
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Desde"
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
            Index           =   51
            Left            =   1620
            TabIndex        =   169
            Top             =   1680
            Width           =   645
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Cliente"
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
            Height          =   240
            Index           =   2
            Left            =   240
            TabIndex        =   168
            Top             =   1680
            Width           =   765
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
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
            Index           =   37
            Left            =   4605
            TabIndex        =   165
            Top             =   1200
            Width           =   615
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   12
            Left            =   2775
            Picture         =   "frmListadoPed.frx":0341
            Top             =   1215
            Width           =   240
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Albaran"
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
            Height          =   240
            Index           =   1
            Left            =   240
            TabIndex        =   164
            Top             =   1200
            Width           =   1560
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Desde"
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
            Index           =   46
            Left            =   1935
            TabIndex        =   163
            Top             =   1200
            Width           =   780
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   13
            Left            =   5220
            Picture         =   "frmListadoPed.frx":03CC
            Top             =   1215
            Width           =   240
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
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
            Index           =   36
            Left            =   4650
            TabIndex        =   162
            Top             =   720
            Width           =   570
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Nº Albaran"
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
            Height          =   240
            Index           =   4
            Left            =   240
            TabIndex        =   161
            Top             =   720
            Width           =   1170
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Desde"
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
            Index           =   45
            Left            =   1935
            TabIndex        =   160
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Periodicidad Facturación"
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
            Height          =   240
            Index           =   6
            Left            =   240
            TabIndex        =   159
            Top             =   240
            Width           =   2685
         End
      End
      Begin VB.Frame Frame3 
         Height          =   1095
         Left            =   180
         TabIndex        =   154
         Top             =   720
         Width           =   8205
         Begin VB.TextBox txtCodigo 
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
            Index           =   34
            Left            =   2625
            MaxLength       =   10
            TabIndex        =   138
            Top             =   240
            Width           =   1350
         End
         Begin VB.TextBox txtNombre 
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
            Index           =   0
            Left            =   3450
            Locked          =   -1  'True
            TabIndex        =   155
            Text            =   "Text5"
            Top             =   600
            Width           =   4635
         End
         Begin VB.TextBox txtCodigo 
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
            Index           =   0
            Left            =   2625
            MaxLength       =   6
            TabIndex        =   139
            Top             =   600
            Width           =   810
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Facturación"
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
            Height          =   240
            Index           =   5
            Left            =   240
            TabIndex        =   157
            Top             =   240
            Width           =   2010
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   14
            Left            =   2325
            Picture         =   "frmListadoPed.frx":0457
            Top             =   240
            Width           =   240
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Cta.Prevista Cobro"
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
            Height          =   240
            Index           =   7
            Left            =   240
            TabIndex        =   156
            Top             =   600
            Width           =   2055
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   24
            Left            =   2325
            Top             =   645
            Width           =   240
         End
      End
      Begin VB.CommandButton cmdAceptarFac 
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
         Left            =   6150
         TabIndex        =   151
         Top             =   6840
         Width           =   1065
      End
      Begin VB.CommandButton cmdCancel 
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
         Index           =   6
         Left            =   7275
         TabIndex        =   153
         Top             =   6840
         Width           =   1065
      End
      Begin VB.Label Label10 
         Caption         =   "Facturación de Albaranes"
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
         Left            =   180
         TabIndex        =   102
         Top             =   240
         Width           =   8100
      End
      Begin VB.Label Label10 
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
         Left            =   360
         TabIndex        =   235
         Top             =   3360
         Width           =   6615
      End
   End
   Begin VB.Frame FrameGenAlbaran 
      Height          =   7215
      Left            =   480
      TabIndex        =   58
      Top             =   0
      Width           =   5835
      Begin VB.Frame FrameCanjePuntos 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   360
         TabIndex        =   322
         Top             =   4320
         Visible         =   0   'False
         Width           =   5055
         Begin VB.TextBox txtCodigo 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   68
            Left            =   3840
            TabIndex        =   50
            Text            =   "0"
            Top             =   360
            Width           =   855
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Enabled         =   0   'False
            Height          =   285
            Index           =   63
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   325
            Top             =   360
            Width           =   855
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Enabled         =   0   'False
            Height          =   285
            Index           =   4
            Left            =   480
            Locked          =   -1  'True
            TabIndex        =   323
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Puntos canjear"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   75
            Left            =   3720
            TabIndex        =   327
            Top             =   120
            Width           =   1080
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Puntos albarán"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   74
            Left            =   2040
            TabIndex        =   326
            Top             =   120
            Width           =   1080
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Puntos cliente"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   73
            Left            =   480
            TabIndex        =   324
            Top             =   120
            Width           =   1005
         End
         Begin VB.Shape Shape1 
            Height          =   735
            Left            =   0
            Top             =   0
            Width           =   5055
         End
      End
      Begin VB.Frame FrameBultosHerbelca 
         Height          =   855
         Left            =   360
         TabIndex        =   308
         Top             =   5160
         Visible         =   0   'False
         Width           =   5415
         Begin VB.ComboBox cboDestinoB 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmListadoPed.frx":04E2
            Left            =   3600
            List            =   "frmListadoPed.frx":04E9
            Style           =   2  'Dropdown List
            TabIndex        =   328
            Top             =   300
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   62
            Left            =   1680
            TabIndex        =   49
            Top             =   315
            Width           =   975
         End
         Begin VB.Label lblDestinoB 
            AutoSize        =   -1  'True
            Caption         =   "Destino"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2880
            TabIndex        =   329
            Top             =   360
            Visible         =   0   'False
            Width           =   540
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Nº bultos albaran"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   26
            Left            =   240
            TabIndex        =   309
            Top             =   360
            Width           =   1245
         End
      End
      Begin VB.Frame FrameGenAlbEuler 
         Caption         =   "             Generar albaran euler"
         Height          =   735
         Left            =   360
         TabIndex        =   306
         Top             =   3480
         Width           =   4935
         Begin VB.ComboBox cboTipoAlbaranEuler 
            Height          =   315
            ItemData        =   "frmListadoPed.frx":04F6
            Left            =   720
            List            =   "frmListadoPed.frx":0506
            Style           =   2  'Dropdown List
            TabIndex        =   38
            Top             =   360
            Width           =   2055
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Albaran"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   25
            Left            =   480
            TabIndex        =   307
            Top             =   120
            Width           =   900
         End
      End
      Begin VB.Frame FramePartes 
         Height          =   1455
         Left            =   360
         TabIndex        =   282
         Top             =   5160
         Visible         =   0   'False
         Width           =   5295
         Begin VB.ComboBox cboFacturacion 
            Height          =   315
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   47
            Top             =   960
            Width           =   2055
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   60
            Left            =   2160
            TabIndex        =   45
            Top             =   480
            Width           =   975
         End
         Begin VB.CheckBox chkFrInterna 
            Caption         =   "Interna"
            Height          =   255
            Left            =   3840
            TabIndex        =   46
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   59
            Left            =   360
            TabIndex        =   44
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Tipo facturación"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   65
            Left            =   360
            TabIndex        =   285
            Top             =   1080
            Width           =   1155
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Cantidad otros"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   64
            Left            =   2160
            TabIndex        =   284
            Top             =   180
            Width           =   1065
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Cantidad real dosis"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   63
            Left            =   120
            TabIndex        =   283
            Top             =   180
            Width           =   1365
         End
      End
      Begin VB.CheckBox chkFraMostrador 
         Caption         =   "A factura mostrador"
         Height          =   255
         Left            =   3720
         TabIndex        =   40
         Top             =   4440
         Width           =   1815
      End
      Begin VB.Frame FramRectARM 
         Caption         =   "Fra a la que rectifica"
         Height          =   735
         Left            =   720
         TabIndex        =   256
         Top             =   5160
         Width           =   4815
         Begin VB.TextBox txtFraRectifica 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   240
            TabIndex        =   258
            Top             =   240
            Width           =   3615
         End
         Begin VB.CommandButton cmdSelFraRect 
            Caption         =   "+"
            Height          =   375
            Left            =   4080
            TabIndex        =   257
            Top             =   240
            Width           =   375
         End
      End
      Begin VB.Frame FrameZona 
         BorderStyle     =   0  'None
         Caption         =   "Frame16"
         Height          =   615
         Left            =   720
         TabIndex        =   252
         Top             =   3720
         Width           =   4935
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   22
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   253
            Text            =   "Text5"
            Top             =   240
            Width           =   3735
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   22
            Left            =   420
            MaxLength       =   4
            TabIndex        =   37
            Top             =   240
            Width           =   615
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   35
            Left            =   120
            Picture         =   "frmListadoPed.frx":0541
            Top             =   240
            Width           =   240
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Zona"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   53
            Left            =   120
            TabIndex        =   254
            Top             =   0
            Width           =   360
         End
      End
      Begin VB.CheckBox chkImpHojaExped 
         Caption         =   "Imprimir Hoja Expedición"
         Height          =   255
         Left            =   1080
         TabIndex        =   43
         Top             =   5280
         Width           =   2295
      End
      Begin VB.CheckBox chkImpEtiq 
         Caption         =   "Imprimir Etiquetas"
         Height          =   255
         Left            =   3720
         TabIndex        =   42
         Top             =   4800
         Width           =   1575
      End
      Begin VB.Frame FramepedidoFactura 
         BorderStyle     =   0  'None
         Caption         =   "Frame15"
         Height          =   615
         Left            =   720
         TabIndex        =   232
         Top             =   6000
         Width           =   4815
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   5
            Left            =   1155
            Locked          =   -1  'True
            TabIndex        =   233
            Text            =   "Text5"
            Top             =   240
            Width           =   3495
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   5
            Left            =   315
            MaxLength       =   6
            TabIndex        =   48
            Top             =   240
            Width           =   735
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   0
            Left            =   0
            Picture         =   "frmListadoPed.frx":0643
            Top             =   240
            Width           =   240
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Cta prevista cobro"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   23
            Left            =   0
            TabIndex        =   234
            Top             =   0
            Width           =   1335
         End
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   25
         Left            =   2400
         MaxLength       =   10
         TabIndex        =   39
         Top             =   4440
         Width           =   1215
      End
      Begin VB.CheckBox chkImpAlbaran 
         Caption         =   "Imprimir Albaran"
         Height          =   255
         Left            =   1080
         TabIndex        =   41
         Top             =   4800
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   19
         Left            =   1140
         MaxLength       =   4
         TabIndex        =   36
         Top             =   3240
         Width           =   615
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   19
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   65
         Text            =   "Text5"
         Top             =   3240
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   18
         Left            =   1140
         MaxLength       =   4
         TabIndex        =   35
         Top             =   2520
         Width           =   615
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   18
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   63
         Text            =   "Text5"
         Top             =   2520
         Width           =   3735
      End
      Begin VB.CommandButton cmdAceptarGenAlb 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3360
         TabIndex        =   51
         Top             =   6720
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   3
         Left            =   4560
         TabIndex        =   52
         Top             =   6720
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   17
         Left            =   1140
         MaxLength       =   4
         TabIndex        =   34
         Top             =   1800
         Width           =   615
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   17
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   59
         Text            =   "Text5"
         Top             =   1800
         Width           =   3735
      End
      Begin VB.Label Label5 
         Caption         =   "Zona:"
         Height          =   255
         Left            =   240
         TabIndex        =   255
         Top             =   6720
         Width           =   2535
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Albaran"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   32
         Left            =   840
         TabIndex        =   67
         Top             =   4440
         Width           =   1155
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   2040
         Picture         =   "frmListadoPed.frx":0745
         Top             =   4440
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Forma de Envío"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   11
         Left            =   840
         TabIndex        =   66
         Top             =   3000
         Width           =   1110
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   10
         Left            =   840
         Picture         =   "frmListadoPed.frx":07D0
         Top             =   3240
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Material Preparado por"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   10
         Left            =   840
         TabIndex        =   64
         Top             =   2280
         Width           =   1650
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   9
         Left            =   840
         Picture         =   "frmListadoPed.frx":08D2
         Top             =   2520
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Pasar Pedido a "
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
         Left            =   600
         TabIndex        =   62
         Top             =   480
         Width           =   4695
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Introduzca los siguiente campos: "
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
         Left            =   600
         TabIndex        =   61
         Top             =   1200
         Width           =   2835
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Trabajador de Albaran"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   15
         Left            =   840
         TabIndex        =   60
         Top             =   1560
         Width           =   1620
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   8
         Left            =   840
         Picture         =   "frmListadoPed.frx":09D4
         Top             =   1800
         Width           =   240
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   7800
      Top             =   6480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameEstVentas 
      Height          =   3975
      Left            =   480
      TabIndex        =   218
      Top             =   720
      Width           =   7035
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   53
         Left            =   1500
         MaxLength       =   4
         TabIndex        =   222
         Top             =   1440
         Width           =   840
      End
      Begin VB.CommandButton cmdAceptarEstVentas 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4200
         TabIndex        =   224
         Top             =   3120
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   8
         Left            =   5280
         TabIndex        =   225
         Top             =   3120
         Width           =   975
      End
      Begin VB.Frame Frame14 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   480
         TabIndex        =   219
         Top             =   1800
         Width           =   6495
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   8
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   220
            Text            =   "Text5"
            Top             =   120
            Width           =   3855
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   8
            Left            =   1020
            MaxLength       =   6
            TabIndex        =   223
            Top             =   120
            Width           =   840
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   32
            Left            =   705
            Picture         =   "frmListadoPed.frx":0AD6
            Top             =   120
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
            Index           =   30
            Left            =   0
            TabIndex        =   221
            Top             =   120
            Width           =   585
         End
      End
      Begin VB.Label Label2 
         Caption         =   "Ventas por meses"
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
         Left            =   480
         TabIndex        =   227
         Top             =   480
         Width           =   4815
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
         Index           =   57
         Left            =   480
         TabIndex        =   226
         Top             =   1440
         Width           =   330
      End
   End
   Begin VB.Frame FramePedxArtic 
      Height          =   7455
      Left            =   240
      TabIndex        =   53
      Top             =   120
      Width           =   11775
      Begin VB.CheckBox chkPedxClixSemEntrega 
         Caption         =   "Listado agrupado por articulo"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   330
         Top             =   3720
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.CheckBox chkPedxClixSemEntrega 
         Caption         =   "IVA incluido"
         Height          =   255
         Index           =   2
         Left            =   3720
         TabIndex        =   29
         Top             =   6720
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.Frame FrameAsociacion 
         Height          =   975
         Left            =   480
         TabIndex        =   316
         Top             =   1080
         Visible         =   0   'False
         Width           =   6615
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   66
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   318
            Text            =   "Text5"
            Top             =   240
            Width           =   3735
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   66
            Left            =   1440
            TabIndex        =   10
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   67
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   317
            Text            =   "Text5"
            Top             =   600
            Width           =   3735
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   67
            Left            =   1440
            TabIndex        =   11
            Top             =   600
            Width           =   615
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Ruta"
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
            Index           =   72
            Left            =   0
            TabIndex        =   321
            Top             =   0
            Width           =   405
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   44
            Left            =   960
            Picture         =   "frmListadoPed.frx":0BD8
            Top             =   240
            Width           =   240
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
            Height          =   195
            Index           =   71
            Left            =   480
            TabIndex        =   320
            Top             =   240
            Width           =   450
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   45
            Left            =   960
            Picture         =   "frmListadoPed.frx":0CDA
            Top             =   600
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
            Height          =   195
            Index           =   70
            Left            =   480
            TabIndex        =   319
            Top             =   600
            Width           =   420
         End
      End
      Begin VB.Frame FrameZonaCli 
         Height          =   975
         Left            =   480
         TabIndex        =   310
         Top             =   2160
         Visible         =   0   'False
         Width           =   6615
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   65
            Left            =   1440
            TabIndex        =   9
            Top             =   600
            Width           =   615
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   65
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   314
            Text            =   "Text5"
            Top             =   600
            Width           =   3735
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   64
            Left            =   1440
            TabIndex        =   8
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   64
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   312
            Text            =   "Text5"
            Top             =   240
            Width           =   3735
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
            Height          =   195
            Index           =   69
            Left            =   480
            TabIndex        =   315
            Top             =   600
            Width           =   420
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   43
            Left            =   960
            Picture         =   "frmListadoPed.frx":0DDC
            Top             =   600
            Width           =   240
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
            Height          =   195
            Index           =   68
            Left            =   480
            TabIndex        =   313
            Top             =   240
            Width           =   450
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   42
            Left            =   960
            Picture         =   "frmListadoPed.frx":0EDE
            Top             =   240
            Width           =   240
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Zona"
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
            Index           =   67
            Left            =   0
            TabIndex        =   311
            Top             =   0
            Width           =   420
         End
      End
      Begin VB.Frame FrameTiposFactura 
         Height          =   1695
         Left            =   240
         TabIndex        =   286
         Top             =   4680
         Visible         =   0   'False
         Width           =   5895
         Begin VB.ComboBox cboAnyos 
            Height          =   315
            ItemData        =   "frmListadoPed.frx":0FE0
            Left            =   4680
            List            =   "frmListadoPed.frx":0FF0
            Style           =   2  'Dropdown List
            TabIndex        =   333
            Top             =   1320
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.ComboBox cboTipoCredito 
            Height          =   315
            Left            =   2400
            Style           =   2  'Dropdown List
            TabIndex        =   332
            Top             =   1320
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.CheckBox chkPedxClixSemEntrega 
            Caption         =   "Comparativo fechas"
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
            Index           =   4
            Left            =   120
            TabIndex        =   331
            Top             =   1380
            Width           =   2175
         End
         Begin VB.ListBox ListTipoFact 
            Height          =   960
            Left            =   1320
            Style           =   1  'Checkbox
            TabIndex        =   31
            Top             =   240
            Width           =   3975
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Años"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   76
            Left            =   4320
            TabIndex        =   334
            Top             =   1350
            Visible         =   0   'False
            Width           =   360
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Tipo factura"
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
            Index           =   66
            Left            =   120
            TabIndex        =   287
            Top             =   240
            Width           =   1020
         End
      End
      Begin VB.Frame Frame10 
         Height          =   3015
         Left            =   360
         TabIndex        =   201
         Top             =   3480
         Width           =   6375
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   58
            Left            =   2280
            Locked          =   -1  'True
            TabIndex        =   280
            Text            =   "Text5"
            Top             =   1920
            Width           =   3735
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   58
            Left            =   1560
            TabIndex        =   20
            Top             =   1920
            Width           =   615
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   57
            Left            =   2280
            Locked          =   -1  'True
            TabIndex        =   277
            Text            =   "Text5"
            Top             =   1560
            Width           =   3735
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   57
            Left            =   1560
            TabIndex        =   19
            Top             =   1560
            Width           =   615
         End
         Begin VB.CheckBox Check1chkAgrupaAg 
            Caption         =   "Agrupa agente"
            Height          =   435
            Left            =   4680
            TabIndex        =   24
            Top             =   2400
            Width           =   1455
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   24
            Left            =   1560
            TabIndex        =   18
            Top             =   960
            Width           =   615
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   24
            Left            =   2280
            Locked          =   -1  'True
            TabIndex        =   265
            Text            =   "Text5"
            Top             =   960
            Width           =   3735
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   23
            Left            =   1560
            TabIndex        =   17
            Top             =   600
            Width           =   615
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   23
            Left            =   2280
            Locked          =   -1  'True
            TabIndex        =   262
            Text            =   "Text5"
            Top             =   600
            Width           =   3735
         End
         Begin VB.Frame Frame11 
            Caption         =   " Ordenar por "
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
            Height          =   615
            Left            =   120
            TabIndex        =   204
            Top             =   2280
            Width           =   4455
            Begin VB.OptionButton OptOrdenVentas 
               Caption         =   "Vol. ventas"
               Height          =   255
               Left            =   3000
               TabIndex        =   23
               Top             =   240
               Value           =   -1  'True
               Width           =   1215
            End
            Begin VB.OptionButton OptOrdenNomclien 
               Caption         =   "Nombre cliente"
               Height          =   375
               Left            =   1560
               TabIndex        =   22
               Top             =   180
               Width           =   1455
            End
            Begin VB.OptionButton OptOrdenCodclien 
               Caption         =   "Cod. cliente"
               Height          =   255
               Left            =   120
               TabIndex        =   21
               Top             =   240
               Width           =   1215
            End
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   1
            Left            =   3840
            MaxLength       =   15
            TabIndex        =   16
            Top             =   120
            Width           =   1695
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   41
            Left            =   1200
            Picture         =   "frmListadoPed.frx":1000
            Top             =   1920
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
            Height          =   195
            Index           =   62
            Left            =   720
            TabIndex        =   281
            Top             =   1920
            Width           =   420
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   40
            Left            =   1200
            Picture         =   "frmListadoPed.frx":1102
            Top             =   1560
            Width           =   240
         End
         Begin VB.Label Label4 
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
            Index           =   61
            Left            =   120
            TabIndex        =   279
            Top             =   1320
            Width           =   795
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
            Height          =   195
            Index           =   60
            Left            =   720
            TabIndex        =   278
            Top             =   1560
            Width           =   450
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
            Height          =   195
            Index           =   58
            Left            =   720
            TabIndex        =   266
            Top             =   960
            Width           =   420
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   37
            Left            =   1200
            Picture         =   "frmListadoPed.frx":1204
            Top             =   960
            Width           =   240
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
            Height          =   195
            Index           =   56
            Left            =   720
            TabIndex        =   264
            Top             =   600
            Width           =   450
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Agente"
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
            Index           =   55
            Left            =   120
            TabIndex        =   263
            Top             =   360
            Width           =   615
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   36
            Left            =   1200
            Picture         =   "frmListadoPed.frx":1306
            Top             =   600
            Width           =   240
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "€"
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
            Index           =   19
            Left            =   5640
            TabIndex        =   203
            Top             =   120
            Width           =   105
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Mostrar Clientes con ventas superiores a"
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
            Left            =   120
            TabIndex        =   202
            Top             =   120
            Width           =   3465
         End
      End
      Begin VB.CheckBox chkDispo 
         Caption         =   "Datos de dpto/obra"
         Height          =   255
         Index           =   1
         Left            =   1920
         TabIndex        =   260
         Top             =   4320
         Width           =   1935
      End
      Begin VB.CheckBox chkDispo 
         Caption         =   "Detalle"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   259
         Top             =   4320
         Width           =   1575
      End
      Begin VB.Frame FramepedxClien 
         BorderStyle     =   0  'None
         Height          =   2295
         Left            =   5760
         TabIndex        =   241
         Top             =   4320
         Visible         =   0   'False
         Width           =   6375
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   10
            Left            =   1500
            MaxLength       =   6
            TabIndex        =   15
            Top             =   1680
            Width           =   735
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   10
            Left            =   2280
            Locked          =   -1  'True
            TabIndex        =   249
            Text            =   "Text5"
            Top             =   1680
            Width           =   3855
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   9
            Left            =   1500
            MaxLength       =   6
            TabIndex        =   14
            Top             =   1320
            Width           =   735
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   9
            Left            =   2280
            Locked          =   -1  'True
            TabIndex        =   247
            Text            =   "Text5"
            Top             =   1320
            Width           =   3855
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   7
            Left            =   1500
            MaxLength       =   6
            TabIndex        =   13
            Top             =   600
            Width           =   735
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   7
            Left            =   2280
            Locked          =   -1  'True
            TabIndex        =   245
            Text            =   "Text5"
            Top             =   600
            Width           =   3855
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   6
            Left            =   1500
            MaxLength       =   6
            TabIndex        =   12
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   6
            Left            =   2280
            Locked          =   -1  'True
            TabIndex        =   242
            Text            =   "Text5"
            Top             =   240
            Width           =   3855
         End
         Begin VB.Label Label4 
            Caption         =   "Zona"
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
            Index           =   52
            Left            =   0
            TabIndex        =   251
            Top             =   1080
            Width           =   615
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
            Height          =   195
            Index           =   47
            Left            =   600
            TabIndex        =   250
            Top             =   1680
            Width           =   420
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   34
            Left            =   1200
            Picture         =   "frmListadoPed.frx":1408
            Top             =   1680
            Width           =   240
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
            Height          =   195
            Index           =   31
            Left            =   600
            TabIndex        =   248
            Top             =   1320
            Width           =   450
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   33
            Left            =   1200
            Picture         =   "frmListadoPed.frx":150A
            Top             =   1320
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
            Height          =   195
            Index           =   29
            Left            =   600
            TabIndex        =   246
            Top             =   600
            Width           =   420
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   13
            Left            =   1200
            Picture         =   "frmListadoPed.frx":160C
            Top             =   600
            Width           =   240
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Agente"
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
            Index           =   28
            Left            =   0
            TabIndex        =   244
            Top             =   0
            Width           =   615
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
            Height          =   195
            Index           =   27
            Left            =   600
            TabIndex        =   243
            Top             =   240
            Width           =   450
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   3
            Left            =   1200
            Picture         =   "frmListadoPed.frx":170E
            Top             =   240
            Width           =   240
         End
      End
      Begin VB.ComboBox cmbTipAlbaran 
         Height          =   315
         Index           =   1
         ItemData        =   "frmListadoPed.frx":1810
         Left            =   2640
         List            =   "frmListadoPed.frx":181D
         Style           =   2  'Dropdown List
         TabIndex        =   229
         Top             =   4920
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Frame Frame12 
         BorderStyle     =   0  'None
         Height          =   2055
         Left            =   7080
         TabIndex        =   205
         Top             =   1440
         Width           =   6495
         Begin VB.Frame Frame13 
            Height          =   615
            Left            =   240
            TabIndex        =   211
            Top             =   1320
            Width           =   2655
            Begin VB.OptionButton OptResumen 
               Caption         =   "Resumen"
               Height          =   255
               Left            =   1320
               TabIndex        =   213
               Top             =   240
               Width           =   1095
            End
            Begin VB.OptionButton OptDetalle 
               Caption         =   "Detalle"
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   212
               Top             =   240
               Value           =   -1  'True
               Width           =   975
            End
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   2
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   207
            Text            =   "Text5"
            Top             =   360
            Width           =   3855
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   2
            Left            =   1380
            MaxLength       =   6
            TabIndex        =   25
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   3
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   206
            Text            =   "Text5"
            Top             =   720
            Width           =   3855
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   3
            Left            =   1380
            MaxLength       =   6
            TabIndex        =   26
            Top             =   720
            Width           =   735
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   1
            Left            =   1080
            Picture         =   "frmListadoPed.frx":1850
            Top             =   360
            Width           =   240
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Trabajador"
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
            Index           =   22
            Left            =   120
            TabIndex        =   210
            Top             =   120
            Width           =   945
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
            Height          =   195
            Index           =   21
            Left            =   480
            TabIndex        =   209
            Top             =   360
            Width           =   450
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   2
            Left            =   1080
            Picture         =   "frmListadoPed.frx":1952
            Top             =   720
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
            Height          =   195
            Index           =   20
            Left            =   480
            TabIndex        =   208
            Top             =   720
            Width           =   420
         End
      End
      Begin VB.Frame Frame5 
         BorderStyle     =   0  'None
         Height          =   1095
         Left            =   360
         TabIndex        =   176
         Top             =   2640
         Width           =   6495
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   21
            Left            =   1380
            MaxLength       =   6
            TabIndex        =   7
            Top             =   720
            Width           =   735
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   21
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   178
            Text            =   "Text5"
            Top             =   720
            Width           =   3855
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   20
            Left            =   1380
            MaxLength       =   6
            TabIndex        =   6
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   20
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   177
            Text            =   "Text5"
            Top             =   360
            Width           =   3855
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
            Height          =   195
            Index           =   12
            Left            =   480
            TabIndex        =   181
            Top             =   720
            Width           =   420
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   12
            Left            =   1080
            Picture         =   "frmListadoPed.frx":1A54
            Top             =   720
            Width           =   240
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
            Height          =   195
            Index           =   13
            Left            =   480
            TabIndex        =   180
            Top             =   360
            Width           =   450
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
            Index           =   16
            Left            =   0
            TabIndex        =   179
            Top             =   120
            Width           =   585
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   11
            Left            =   1080
            Picture         =   "frmListadoPed.frx":1B56
            Top             =   360
            Width           =   240
         End
      End
      Begin VB.Frame Frame9 
         BorderStyle     =   0  'None
         Height          =   1095
         Left            =   360
         TabIndex        =   195
         Top             =   1920
         Width           =   6375
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   13
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   197
            Text            =   "Text5"
            Top             =   360
            Width           =   3735
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   13
            Left            =   1260
            MaxLength       =   3
            TabIndex        =   2
            Top             =   360
            Width           =   495
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   14
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   196
            Text            =   "Text5"
            Top             =   720
            Width           =   3735
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   14
            Left            =   1260
            MaxLength       =   3
            TabIndex        =   3
            Top             =   720
            Width           =   495
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   4
            Left            =   960
            Picture         =   "frmListadoPed.frx":1C58
            Top             =   360
            Width           =   240
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Almacen"
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
            Left            =   120
            TabIndex        =   200
            Top             =   120
            Width           =   735
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
            Height          =   195
            Index           =   7
            Left            =   480
            TabIndex        =   199
            Top             =   360
            Width           =   450
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   5
            Left            =   960
            Picture         =   "frmListadoPed.frx":1D5A
            Top             =   720
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
            Height          =   195
            Index           =   8
            Left            =   480
            TabIndex        =   198
            Top             =   720
            Width           =   420
         End
      End
      Begin VB.Frame Frame8 
         BorderStyle     =   0  'None
         Height          =   1215
         Left            =   480
         TabIndex        =   189
         Top             =   3240
         Width           =   6975
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   15
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   191
            Text            =   "Text5"
            Top             =   480
            Width           =   4215
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   15
            Left            =   1260
            MaxLength       =   16
            TabIndex        =   4
            Top             =   480
            Width           =   1215
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   16
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   190
            Text            =   "Text5"
            Top             =   840
            Width           =   4215
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   16
            Left            =   1260
            MaxLength       =   16
            TabIndex        =   5
            Top             =   840
            Width           =   1215
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   6
            Left            =   960
            Picture         =   "frmListadoPed.frx":1E5C
            Top             =   480
            Width           =   240
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Artículo"
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
            Left            =   120
            TabIndex        =   194
            Top             =   240
            Width           =   660
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
            Height          =   195
            Index           =   3
            Left            =   480
            TabIndex        =   193
            Top             =   480
            Width           =   450
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   7
            Left            =   960
            Picture         =   "frmListadoPed.frx":1F5E
            Top             =   840
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
            Height          =   195
            Index           =   9
            Left            =   480
            TabIndex        =   192
            Top             =   840
            Width           =   420
         End
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   12
         Left            =   4140
         MaxLength       =   10
         TabIndex        =   1
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   2
         Left            =   5760
         TabIndex        =   33
         Top             =   7080
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptarPedxArtic 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4680
         TabIndex        =   32
         Top             =   7080
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   11
         Left            =   1740
         MaxLength       =   10
         TabIndex        =   0
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Frame FrameOrden1 
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   240
         TabIndex        =   240
         Top             =   6600
         Width           =   3975
         Begin VB.CheckBox chkPedxClixSemEntrega 
            Caption         =   "Obs. pedido"
            Height          =   255
            Index           =   1
            Left            =   2040
            TabIndex        =   28
            Top             =   120
            Value           =   1  'Checked
            Width           =   1815
         End
         Begin VB.CheckBox chkPedxClixSemEntrega 
            Caption         =   "Semana entrega"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   27
            Top             =   120
            Value           =   1  'Checked
            Width           =   1575
         End
         Begin VB.ComboBox cboTipocliente 
            Height          =   315
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   30
            Top             =   480
            Width           =   1815
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Tipo cliente"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   59
            Left            =   240
            TabIndex        =   276
            Top             =   510
            Width           =   810
         End
         Begin VB.Image imgayuda 
            Height          =   255
            Index           =   0
            Left            =   3240
            ToolTipText     =   "Pedidos por cliente"
            Top             =   480
            Width           =   255
         End
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
         Height          =   195
         Index           =   54
         Left            =   480
         TabIndex        =   261
         Top             =   4680
         Width           =   2850
      End
      Begin VB.Label lblTipAlbaran 
         Caption         =   "Tipo de albaranes:"
         Height          =   255
         Index           =   1
         Left            =   600
         TabIndex        =   230
         Top             =   4920
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   7
         Left            =   3840
         Picture         =   "frmListadoPed.frx":2060
         Top             =   1440
         Width           =   240
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
         Height          =   195
         Index           =   5
         Left            =   960
         TabIndex        =   57
         Top             =   1440
         Width           =   450
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Pedido"
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
         Left            =   480
         TabIndex        =   56
         Top             =   1200
         Width           =   1110
      End
      Begin VB.Label Label1 
         Caption         =   "Pedidos por Artículo"
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
         Left            =   360
         TabIndex        =   55
         Top             =   480
         Width           =   4815
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   6
         Left            =   1440
         Picture         =   "frmListadoPed.frx":20EB
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
         Height          =   195
         Index           =   2
         Left            =   3360
         TabIndex        =   54
         Top             =   1440
         Width           =   420
      End
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   0
      Left            =   12000
      Picture         =   "frmListadoPed.frx":2176
      Tag             =   "-1"
      ToolTipText     =   "Buscar almacén"
      Top             =   0
      Width           =   240
   End
End
Attribute VB_Name = "frmListadoPed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event DatoSeleccionado(CadenaSeleccion As String)

Public OpcionListado As Integer
'(ver opciones en frmListado)
      
      
      
'Alguna opcion mas
'                   1000.-  Es cuando paso pedido a albaran y este a factura en el mismo proceso
'                   1001.-  Facturar un unico albaran
'                   1010  algo de pasar a albaran
'                   1043   pasar parte trabajoa albaran
      
Public NumCod As String 'Para indicar cod. Traspaso,Movimiento, etc. que llama
                        'Para indicar nº oferta a imprimir


Public codClien As String 'Para seleccionar inicialmente las ofertas del Proveedor
                        'para paso ped a labaran llevare datos x defecto: llevo: tienecoddiren & "|" & zonacliente & "|"
                

Public FacturaSuperaImporteTicket As Boolean 'Dira si el importe es mayor que el maximo permitido pora la ley para facturas simpificadas (todo esta en parametros)


'Private HaDevueltoDatos As Boolean
Private NomTabla As String
Private NomTablaLin As String

'Private WithEvents frmMtoCartasOfe As frmFacCartasOferta
Private WithEvents frmMtoClient As frmBasico2
Attribute frmMtoClient.VB_VarHelpID = -1
Private WithEvents frmMtoAlmacen As frmAlmAlPropios
Attribute frmMtoAlmacen.VB_VarHelpID = -1
Private WithEvents frmMtoArticulo As frmBasico2
Attribute frmMtoArticulo.VB_VarHelpID = -1
Private WithEvents frmMtoTraba As frmBasico2 'frmAdmTrabajadores
Attribute frmMtoTraba.VB_VarHelpID = -1
Private WithEvents frmMtoFEnvio As frmFacFormasEnvio
Attribute frmMtoFEnvio.VB_VarHelpID = -1
Private WithEvents frmMtoFPago As frmBasico2 'frmFacFormasPago
Attribute frmMtoFPago.VB_VarHelpID = -1
Private WithEvents frmMtoAgente As frmBasico2 '%=%=frmFacAgentesCom
Attribute frmMtoAgente.VB_VarHelpID = -1
Private WithEvents frmMtoBancosPro As frmBasico2 'frmFacBancosPropios
Attribute frmMtoBancosPro.VB_VarHelpID = -1
Private WithEvents frmMtoTipCo As frmManTiposContrato
Attribute frmMtoTipCo.VB_VarHelpID = -1
Private WithEvents frmZ As frmFacZonas
Attribute frmZ.VB_VarHelpID = -1

Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1

'Private WithEvents frmB As frmBuscaGrid  'Busquedas
Private WithEvents frmF As frmCal 'Calendario de Fechas
Attribute frmF.VB_VarHelpID = -1

Private WithEvents frmLd As frmListadoOfer   'Para desde pedido a al fra de RMA seelccione los datos de deovlucion
Attribute frmLd.VB_VarHelpID = -1
Private WithEvents frmDptoEnvio As frmFacCliEnvDpto
Attribute frmDptoEnvio.VB_VarHelpID = -1

Private WithEvents frmMens As frmMensajes
Attribute frmMens.VB_VarHelpID = -1


'----- Variables para el INforme ----
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private cadParam As String
Private numParam As Byte
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private Titulo As String 'Titulo informe que se pasa a frmImprimir
Private nomRPT As String 'nombre del fichero .rpt a imprimir
Private conSubRPT As Boolean 'si tiene subinformes para enlazarlos a las tablas correctas
'-------------------------------------


Dim indCodigo As Integer 'indice para txtCodigo

Dim PrimeraVez As Boolean

Dim vContinua As Boolean


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 2, cerrar
    If cerrar Then Unload Me
End Sub




Private Sub cboFacturacion_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cboTipoAlbaranEuler_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cboTipocliente_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub



Private Sub Check1chkAgrupaAg_Click()
    'FrameTiposFactura.visible = Me.Check1chkAgrupaAg.Value = 0
End Sub

Private Sub Check1chkAgrupaAg_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkFraMostrador_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkFrInterna_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkImpAlbaran_KeyDown(KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
    
End Sub

Private Sub chkImpAlbaran_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
    
    
End Sub




Private Sub chkImpEtiq_KeyDown(KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub chkImpEtiq_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub chkImpHojaExped_KeyDown(KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub chkImpHojaExped_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub










Private Sub chkPedxClixSemEntrega_Click(Index As Integer)
      If Index = 4 Then
        If chkPedxClixSemEntrega(4).Value = 1 Then
            If Me.cboTipoCredito.ListCount = 0 Then
                If vParamAplic.OperacionesAseguradas Then
                    CargarCombo_Tabla cboTipoCredito, "stipocredito", "codTipoCredito", "nomTipoCredito", , True
                Else
                    cboTipoCredito.AddItem "Todos"
                    
                End If
            End If
            If cboAnyos.ListIndex < 0 Then cboAnyos.ListIndex = 0
            cboTipoCredito.visible = True
            cboAnyos.visible = True
            Label4(76).visible = True
        Else
            cboTipoCredito.visible = False
            cboAnyos.visible = False
            Label4(76).visible = False
        End If
    End If
End Sub

Private Sub chkPedxClixSemEntrega_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
  
End Sub

Private Sub chkSituFacMant_Click()
    FrameTapa.visible = chkSituFacMant.Value = 1
End Sub

Private Sub chkSituFacMant_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkSoloFacturar_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptarEstVentas_Click()
'Estadistica Ventas por meses
Dim Campo As String
    
    InicializarVbles
    
    'Pasar nombre de la Empresa como parametro
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1

    
    'El campo AÑO es obligarotorio
    txtCodigo(53).Text = Trim(txtCodigo(53).Text)
    If txtCodigo(53).Text = "" Then
        MsgBox "Debe seleccionar una año para el informe.", vbInformation
        Exit Sub
    Else
        Campo = "year({scafac.fecfactu})"
        cadFormula = Campo & " = " & txtCodigo(53).Text
'        campo = campo & " = " & CInt(txtCodigo(53).Text) - 1
'        cadFormula = "(" & cadFormula & " OR " & campo & ")"
        
        'Parametro del año solicitado para el informe
        'Pasar el año solicitado como parametro
        cadParam = cadParam & "pAnyo=""" & "Año: " & txtCodigo(53).Text & """|"
        numParam = numParam + 1
    End If
    
    'Campo seleccion de un CLIENTE
    txtCodigo(8).Text = Trim(txtCodigo(8).Text)
    If txtCodigo(8).Text <> "" Then
        Campo = "{scafac.codclien}"
        cadFormula = cadFormula & " AND (" & Campo & " =" & txtCodigo(8).Text & ")"
        'Pasar el cliente solicitado como parametro
        cadParam = cadParam & "pDHCliente=""" & "Cliente: " & txtCodigo(8).Text & " - " & txtNombre(8).Text & """|"
    Else
        'Mostrar en el informe el total del Año Anterior
        Campo = Campo & " = " & CInt(txtCodigo(53).Text) - 1
        cadFormula = "(" & cadFormula & " OR " & Campo & ")"
        
        cadParam = cadParam & "pDHCliente=""" & "Cliente: Todos" & """|"
    End If
    numParam = numParam + 1
    
    
    'Comprobar si hay registros para mostrar en el informe
    cadSelect = cadFormula
    If Not HayRegParaInforme("scafac", cadSelect) Then Exit Sub
    
    
    'Borro los datos temporales,por si acaso se hubiera quedado
    BorrarTempInformes
    
    'Generar la temporal con los totales por año, mes y cliente (tmpinformes)
    If Not TempVentasMeses(cadSelect, txtCodigo(53).Text) Then
        'Borrar los registros generados por el usuario de la temporal
        BorrarTempInformes
        Exit Sub
    End If
    
    cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo
    Titulo = "Ventas por meses"
'    If Me.OptTipoInf(0).Value = True Then
        nomRPT = "rFacVentasxMesGra.rpt"
'    Else
'        Exit Sub
'        nomRPT = "rFacVentasxMesTex.rpt"
'    End If
    conSubRPT = False
    
    LlamarImprimir
    
    'Borrar los registros generados por el usuario de la temporal
    BorrarTempInformes
End Sub



Private Sub cmdAceptarFac_Click()
'Facturacion de Albaranes
Dim Campo As String, Cad As String
Dim cadFrom As String
Dim cadSQL As String 'Para seleccionar los Albaranes del rango seleccion
                      'que no se van a facturar
Dim CambiamosConta As Boolean
Dim Seguir As Boolean
Dim RT As ADODB.Recordset
Dim UnoSolo As Boolean
Dim vCli As CCliente
Dim Pregunta As Boolean



    InicializarVbles
    
    CambiamosConta = False
    '--- Comprobar q los campos tienen valor
    If Trim(txtCodigo(34).Text) = "" Then 'Fecha factura
        MsgBox "El campo Fecha Factura debe tener valor.", vbExclamation
        Exit Sub
    End If
    
    If Trim(txtCodigo(0).Text) = "" Or txtNombre(0).Text = "" Then 'Banco propio
        MsgBox "El campo cta. prevista de cobro debe tener valor.", vbExclamation
        Exit Sub
    End If
    
    
    'Junio 2014. Para el proceso de facturacion hemos añadido TIPOS de pago
    'Alguno debe estar marcado
    If OpcionListado = 52 Then
        Cad = ""
        For indCodigo = 0 To 6   'LOS SEIS PRIMEROS
            If Me.chkTpPago2(indCodigo).Value = 1 Then Cad = "1"
        Next indCodigo
        If Cad = "" Then
            MsgBox "Seleccione algun tipo de pago", vbExclamation
            Exit Sub
        End If
    End If
    


    'FechaOK
    ResultadoFechaContaOK = EsFechaOKConta(CDate(Me.txtCodigo(34).Text), True)
    If ResultadoFechaContaOK <> 0 Then
        If ResultadoFechaContaOK <> 4 Then MsgBox MensajeFechaOkConta, vbExclamation
        Exit Sub
    End If
    
    'Mayo 2013
    'Fechas contabilizacion de facturas
    Cad = "concat(anofactu,'|',perfactu,'|')"
    
    cadFrom = DevuelveDesdeBD(conConta, "periodos", "parametros", "1", "1", "N", Cad)
    If cadFrom <> "" Then
        indCodigo = 1
        If cadFrom = "0" Then indCodigo = 3
        Campo = RecuperaValor(Cad, 2)
        'MEs
        If Campo <> "" Then
            cadFrom = CStr(CByte(Campo) * indCodigo)
            Cad = RecuperaValor(Cad, 1)
            indCodigo = DiasMes(CByte(cadFrom), CInt(Cad))
            Cad = indCodigo & "/" & Format(cadFrom, "00") & "/" & Cad
            If CDate(Cad) > CDate(txtCodigo(34).Text) Then
                Cad = "El periodo de facturacion del IVA ya esta cerrado.  ¿Desea continuar?"
                If MsgBox(Cad, vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
            End If
        End If
    End If
    
    
    If FacturaSuperaImporteTicket Then
        Cad = "NO"
        If codClien = "ALM" Then
            Cad = ""
        Else
            If codClien = "ALV" And OpcionListado = 222 And vParamAplic.NumeroInstalacion = vbHerbelca Then Cad = ""
        End If
        If Cad <> "" Then
            MsgBox "Supera importe tickets.     Sólo puden hacer facturas directas las facturas de mostrador", vbExclamation
            MsgBox "El programa continuara generando la factura.   AVISE soporte tecnico", vbExclamation
            FacturaSuperaImporteTicket = False
        End If
    End If
    
    If Not ComprobarSecuencialFactura Then Exit Sub
    
    
    
    
    
    
    
    indCodigo = 0
    Cad = ""
    Campo = ""
    cadFrom = ""
    UnoSolo = False
    '--- Seleccinar los Albaranes que cumplen los criterios introducidos
    If OpcionListado <> 222 Then 'Facturas Ventas (FACTURACION)
                                 '222: Facturas de Mostrador/Rectificativa
        'Desde/Hasta Nº ALBARAN
        '-------------------------
        If txtCodigo(36).Text <> "" Or txtCodigo(37).Text <> "" Then
            Campo = NomTabla & ".numalbar"
            Cad = ""
            If Not PonerDesdeHasta(Campo, "N", 36, 37, Cad) Then Exit Sub
        End If
    
        'Desde/Hasta FECHA del ALBARAN
        '--------------------------------------------
        If txtCodigo(38).Text <> "" Or txtCodigo(39).Text <> "" Then
            'Para MySQL
            Campo = "scaalb.fechaalb"
            Cad = CadenaDesdeHastaBD(txtCodigo(38).Text, txtCodigo(39).Text, Campo, "F")
            If Not AnyadirAFormula(cadSelect, Cad) Then Exit Sub
            'Para Crystal Report
            Campo = "{scaalb.fechaalb}"
            Cad = "pDHFecha=""Fecha: "
            If Not PonerDesdeHasta(Campo, "F", 38, 39, Cad) Then Exit Sub
        End If
    
        'Cadena para seleccion D/H CLIENTE
        '----------------------------------------
        If txtCodigo(40).Text <> "" Or txtCodigo(41).Text <> "" Then
            Campo = "scaalb.codclien"
            Cad = ""
            If Not PonerDesdeHasta(Campo, "N", 40, 41, Cad) Then Exit Sub
        End If
    
        'Cadena para seleccion FORMA PAGO
        '--------------------------------------------
        If txtCodigo(42).Text <> "" Or txtCodigo(43).Text <> "" Then
            Campo = "scaalb.codforpa"
            Cad = " "
            If Not PonerDesdeHasta(Campo, "N", 42, 43, Cad) Then Exit Sub
        End If

    
    
        'JUNIO 2014
        ' Tipos de pago. Si estan todos NO haremos select
        If OpcionListado = 52 Then
            
            Cad = ""
            Titulo = ""
            For numParam = 0 To 6
                If Me.chkTpPago2(numParam).Value = 1 Then
                    Cad = Cad & "1"
                    Titulo = Titulo & ", " & numParam
                End If
            Next numParam
            
            If Len(Cad) = 7 Then
                'LOS HA COGIDO TODOS. NO lo incluyo en el desde hasta
            Else
                Set RT = New ADODB.Recordset
                Titulo = Mid(Titulo, 2)
                Cad = "Select codforpa from sforpa where tipforpa in (" & Titulo & ")"
                RT.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                Cad = ""
                While Not RT.EOF
                    Cad = Cad & ", " & RT!codforpa
                    RT.MoveNext
                Wend
                RT.Close
                Set RT = Nothing
                
                If Cad = "" Then
                    'MAL. NInguna forpa de pago con ese tipo de pago. Fuerzo un -1
                    Cad = "-1"
                Else
                    Cad = Mid(Cad, 2)
                End If
                Titulo = ""
                Cad = " scaalb.codforpa IN (" & Cad & ")"
                If Not AnyadirAFormula(cadSelect, Cad) Then Exit Sub
            End If
                
            'junio 20
            'TAXCO. Facturacion final mes OTs
            If vParamAplic.NumeroInstalacion = vbTaxco Then
                If codClien = "ALO" Then
                    Cad = DevuelveDesdeBD(conAri, "trabajadorTaxcoOTMes", "eulerparam", "1", "1")
                    If Cad = "" Then
                        MsgBox "Falta configurar en parametros 'Trabajador facturacion OTs fin mes'", vbExclamation
                        Exit Sub
                    End If
                    
                    vParamAplic.TrabajadorFinMesFacturacion = CInt(Cad)
                    Cad = DevuelveDesdeBD(conAri, "nomtraba", "straba", "codtraba", Cad)
                    If Cad = "" Then
                        MsgBox "No existe trabajador asignado en parametros: " & vParamAplic.TrabajadorFinMesFacturacion, vbExclamation
                        Exit Sub
                    End If
                    
                    
               End If
            End If
                    
        End If
    
        'Otros criterios de Seleccion
        '---------------------------------------------
        'Seleccionar de la Tabla de albaranes scaalb, solo los Albaranes que sean
        'del tipo:Ventas o Reparacion o Mantenimiento
    '    cad = " scaalb.codtipom='ALV' "
        Cad = " scaalb.codtipom='" & codClien & "' " 'filtrar por tipo de albaran segun llamado de Alb.Ventas o Alb. Reparacion
        'Solo lo añadimos a CadSelect porque vamos a Facturar y no a sacar un listado
        If Not AnyadirAFormula(cadSelect, Cad) Then Exit Sub
    
    
        'Seleccionar los Albanares de la Periodicidad indicada
        If txtCodigo(35).Text <> "" Then
            Cad = " sclien.periodof=" & txtCodigo(35).Text
            If Not AnyadirAFormula(cadSelect, Cad) Then Exit Sub
            cadFrom = " scaalb INNER JOIN sclien ON scaalb.codclien=sclien.codclien "
        End If
        
    Else
        'Facturar UNA solo
        Cad = ""
        If codClien = "ALM" And vParamAplic.EntradaRapidaFacturasMostrador Then Cad = "NO"
        
        
        
        
        
        If FramTaxcoTrabajador.visible Then
           
                
            If txtCodigo(69).Text = "" Or (txtCodigo(69).Text = "" Xor txtNombre(69).Text = "") Then
                MsgBox "Falta trabajador conectado", vbExclamation
                PonerFoco txtCodigo(69)
                Exit Sub
            End If
            
            
            If txtCodigo(0).Text = 2 Then
                
                'HA puesto CREDITO. No deberia poner este valor si el cliente NO lo tiene
                Campo = "codtipom= " & DBSet(codClien, "T") & " AND numalbar"
                Campo = DevuelveDesdeBD(conAri, "codclien", "scaalb", Campo, NumCod)
                If Campo = "" Then
                    MsgBox "Error leyendo albaran. "
                    Exit Sub
                End If
                
                Set vCli = New CCliente
                If vCli.LeerDatos(Campo) Then
                    If vCli.ForPago <> 2 Then
                        MsgBox "No se puede hacer crédito al cliente", vbExclamation
                        Campo = ""
                    Else
                        If vCli.ClienteBloqueado(2, False) Then
                            Campo = ""
                        Else
                            If vCli.Observaciones <> "" Then
                                MsgBox vCli.Observaciones, vbInformation
                            Else
                                If MsgBox("Vas a hacer un CREDITO, ¿Continuar?", vbQuestion + vbYesNoCancel) <> vbYes Then Campo = ""
                            End If
                        End If
                        
                        
                        
                        
                    End If
                Else
                    Campo = ""
                End If
                Set vCli = Nothing
                If Campo = "" Then Exit Sub
                
            Else
                'Si no es tarjeta
            End If
            
            
            
            
            'Facturanod UNA en taxco, el codtraba, y la forma de oago, la actualizamos
            Campo = ""
            If txtCodigo(0).Text = 1 Then Campo = "1"
            If txtCodigo(0).Text = 2 Then Campo = "2"
            If txtCodigo(0).Text = 3 Then Campo = "15"
            If TaxcoFacurarUnAlbaranALVIC_Facturado <> "" Then Campo = ""
            If Campo <> "" Then
                Campo = "UPDATE scaalb set codforpa=" & Campo
                Campo = Campo & ", codtraba =" & txtCodigo(69).Text
                Campo = Campo & " WHERE codtipom='" & codClien & "' AND numalbar=" & NumCod
                ejecutar Campo, False
                Campo = ""
            End If
            
          
        End If
        
        
        
        If Cad = "" Then
            If MsgBox("Generar la factura?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        End If
        
        If codClien <> "ALY" Then
            'Lo que habia antes Arbirl 2021
            'en la llamada reutilizamos las vbles codclien y NumCod para guardar tipomov y numalbar.
            cadFormula = "{scaalb.codtipom}='" & codClien & "' AND scaalb.numalbar=" & NumCod
            cadSelect = cadFormula
    
        Else
            cadFormula = " (scaalb.codtipom,scaalb.numalbar) IN (Select codtipoa,numalbar from "
            cadFormula = cadFormula & " sproyectolin WHERE codtipom='" & codClien & "' AND numproyec=" & NumCod & ")"
            cadSelect = cadFormula
        End If
        UnoSolo = True
    End If
    
    
    
    cadSQL = cadSelect
                                                                
    'Pequeña comprobacion de los centros de coste
    If vEmpresa.TieneAnalitica Then
        Cad = "select count(*) from slialb where codccost is null and (codtipom,numalbar) in ("
        Cad = Cad & "select codtipom,numalbar from scaalb where "
        Cad = Cad & cadSelect
        Cad = Cad & " AND  scaalb.factursn=1 )"
        Cad = Replace(Cad, "{", "(")
        Cad = Replace(Cad, "}", ")")
        NumRegElim = CInt(NumRegistros(Cad))
        If NumRegElim > 0 Then
             Cad = "Existen lineas de albaran(" & NumRegElim & ") sin asignar centro de coste"
             Cad = Cad & vbCrLf & vbCrLf & Space(30) & "¿Continuar?"
             If MsgBox(Cad, vbQuestion + vbYesNo) = vbNo Then Exit Sub
        End If
        
    End If
    
                                                        
                                                                
                                                                
                                                                'Septiembre 2009
    'Seleccionar los Albaranes que tiene scaalb.factursn=1     y TENGAN lineas
    Cad = " {scaalb.factursn=1} "
  
        
    'cad = cad & " and (scaalb.codtipom,scaalb.numalbar) in (select codtipom,numalbar from slialb group by codtipom,numalbar)"
    If Not UnoSolo Then Cad = Cad & " and (scaalb.codtipom,scaalb.numalbar) in (select distinct codtipom,numalbar from slialb)"
    
    
    
    
    
    
    If Not AnyadirAFormula(cadSelect, Cad) Then Exit Sub
        
    
    AnyadirAFormula cadFormula, Cad
    
    
    '--- Comprobar q se han Seleccionados registros de Albaran con esos criterios
    Cad = "Select count(*) " ' & NomTabla & " INNER JOIN " & nomTablaLin
    If cadFrom = "" Then cadFrom = " scaalb "
    Cad = Cad & " FROM " & cadFrom

    If Not HayRegParaInforme(cadFrom, cadSelect) Then Exit Sub
    
    If Not HayRegParaInforme(cadFrom, Campo, True) Then
        MsgBox "Albaranes para facturar sin lineas", vbExclamation
        Exit Sub
    End If
    Campo = ""
    'Verificar si con los criterios seleccionados (PARA VENTAS)
    'seleccionar si quedan en el rango Albaranes que no se van a Facturar
    'y mostrar mensaje
    Pregunta = False
    If OpcionListado <> 222 Then
        'Seleccionar los Albaranes que tiene scaalb.factursn=0
        Campo = " scaalb.factursn=0 "
        If Not AnyadirAFormula(cadSQL, Campo) Then Exit Sub
        cadSQL = Cad & " WHERE " & cadSQL
        If RegistrosAListar(cadSQL) > 0 Then
            'Mostrar los Albaranes que no se van a Facturar
            cadSQL = Replace(cadSQL, "count(*)", "scaalb.codtipom,scaalb.numalbar,scaalb.fechaalb,scaalb.codclien,scaalb.nomclien")
            frmMensajes.OpcionMensaje = 12
            frmMensajes.cadWhere = cadSQL
            frmMensajes.Show vbModal
            If frmMensajes.vCampos = "0" Then Exit Sub
            Pregunta = True
        End If
    End If
    
    Cad = Cad & " WHERE " & cadSelect
    

    
    
    'Pasar Albaranes a Facturas
    If InStr(Cad, "sclien") <> 0 Then 'hay JOIN con sclien
        Cad = Replace(Cad, "count(*)", "scaalb.*, sclien.periodof")
    Else
        Cad = Replace(Cad, "count(*)", "*")
    End If







    'Albarananes EN B
    If codClien = "ALZ" Then
        If Not AbrirConexionConta(True) Then
            Cad = "Error MUY grave." & vbCrLf & "Error conectando con BD: " & vParamAplic.ContabilidadB
            MsgBox Cad, vbCritical
            End
            Exit Sub
        End If
        CambiamosConta = True
    End If


    'TAXCO. Facturacion gasolinera.
    '   Hay clientes que la suma de facturas es cero.
    '   Ejmplo. Albaran de + 20    Albaran de -20
    '       Con lo cual la factura es cero. Y TAXCO no quiere que entren facturas a cero, con lo cual
    '       desmarcaremos facturarsn y volveremos a lanzar el proceso
    If codClien = "ALD" Then
        If vParamAplic.NumeroInstalacion = vbTaxco Then
            Screen.MousePointer = vbHourglass
            Dim MIaUX As String
            MIaUX = "select codclien,sum(importel) from slialb,scaalb where scaalb.codtipom=slialb.codtipom and scaalb.numalbar=slialb.numalbar AND "
            MIaUX = MIaUX & cadSelect
            MIaUX = MIaUX & " group by codclien having sum(importel)=0"
            Set miRsAux = Nothing
            Set miRsAux = New ADODB.Recordset
            miRsAux.Open MIaUX, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            MIaUX = ""
            While Not miRsAux.EOF
                If MIaUX = "" Then
                    conn.Execute "DELETE FROM tmpinformes where codusu =" & vUsu.Codigo
                    Espera 0.5
                End If
                MIaUX = "insert INTO tmpinformes(CodUsu , Codigo1, campo1, campo2, nombre1, nombre2, nombre3, obser, Importe1, fecha1)"
                MIaUX = MIaUX & " select " & vUsu.Codigo & " ,scaalb.numalbar,codclien,numlinea,nomclien,codartic,nomartic,ampliaci,importel,fechaalb  "
                MIaUX = MIaUX & " from slialb,scaalb where scaalb.codtipom=slialb.codtipom and scaalb.numalbar=slialb.numalbar AND "
                MIaUX = MIaUX & cadSelect & " AND codclien = " & miRsAux!codClien
                               
                conn.Execute MIaUX

                
                miRsAux.MoveNext
            
            Wend
            miRsAux.Close
            Set miRsAux = Nothing
            
            If MIaUX <> "" Then
                'Abrimos informes
                cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo
                cadParam = "pEmpresa =""" & vEmpresa.nomresum & """|"
                numParam = 1
                Titulo = "Facturacion a cero"
                nomRPT = "rFacFacturasCero.rpt"
                LlamarImprimir
                
                If MsgBox("Desea continuar facturando?", vbQuestion + vbYesNoCancel) <> vbYes Then
                    Screen.MousePointer = vbDefault
                    Exit Sub
                Else
                    Pregunta = True
                End If
                
            End If
            
            
        End If
    End If

    '--- Mostrar Barra de PRogreso
    If OpcionListado <> 222 Then '222: Facturas Mostrador/Rectificativa
                                 '52: Facturas de Venta
                                 'Facturas Reparacion
        
        
        
        'Fitosnatiarios. Comprobar que todos los albaranes llevan puesto carnet de manipulador
        'y  lotes correctamente
        If vParamAplic.ManipuladorFitosanitarios2 Then
            Screen.MousePointer = vbHourglass
            Dim AuxCadena As String
          
            AuxCadena = ""
            If Not ComprobarFitosAlbaranesFacturasCliente(AuxCadena, cadSelect) Then AuxCadena = "NO"
            Screen.MousePointer = vbDefault
            
            If AuxCadena <> "" Then
                AuxCadena = App.Path & "\errfacFito.txt"
                
                AuxCadena = "Hay incidencias en fitosanitarios. Vea el fichero " & AuxCadena
                AuxCadena = AuxCadena & vbCrLf & vbCrLf & "¿Continuar de igual modo? "
                If MsgBox(AuxCadena, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
                Pregunta = True
            End If
        End If
        
        If vParamAplic.NumeroInstalacion = vbHerbelca Then
            Screen.MousePointer = vbHourglass
            AuxCadena = ""
            If Not ComprobarPrecioMinimoFacturacion(AuxCadena, cadSelect) Then AuxCadena = "NO"
            Screen.MousePointer = vbDefault
            If AuxCadena <> "" Then
                AuxCadena = App.Path & "\errfacFito.txt"
                
                AuxCadena = "Hay precios inferiores al precio míminmo. Ver fichero:  " & AuxCadena
                
                If vUsu.Nivel = 0 Then AuxCadena = AuxCadena & vbCrLf & vbCrLf & "¿Continuar de igual modo? "
                If MsgBox(AuxCadena, IIf(vUsu.Nivel = 0, vbQuestion + vbYesNoCancel, vbExclamation)) <> vbYes Then Exit Sub
                Pregunta = True
            End If
        End If
        
        If Not Pregunta Then
            'NO SE HA HECHO pregunta. Preguntamos si quiere seguir
            AuxCadena = "¿Continuar con el proceso de facturacion con los valores seleccionados?"
            If MsgBox(AuxCadena, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
        End If
        
        Me.Height = Me.Height + 300
        Me.FrameFacturar.Height = Me.FrameFacturar.Height + 300
        Me.FrameProgress.visible = True
        Me.FrameProgress.Top = 6250
        Me.PB_Fact.Left = 200
        Me.PB_Fact.Value = 0
        Me.lblProgess(1).Caption = "Inicializando el proceso..."
        
        
        
        'Si vamos a facturar albaranes "B" tenemos que cerrar la conexion CONTA y abrirla contra la
        'Segunda BD que nos indica
        
    End If
    

    '--- Traspasa Albaranes a Facturas
   'proceso normal
    'Fecha de la factura, Cta Prevista de Cobro
    Screen.MousePointer = vbHourglass
    Seguir = True
    If vParamAplic.ArtPortesN <> "" Then
        If vParamAplic.TipoPortes = vbHerbelca Then
            'tipo portes HERBELCA.
            Campo = lblProgess(0).Caption
            lblProgess(0).Caption = "Portes:"
            If Not ProcesoPortesTipoHerbelca(Cad, cadSelect, lblProgess(1)) Then Seguir = False
            lblProgess(0).Caption = Campo
            Campo = ""
        End If
    End If
     
    If Seguir Then
     
        '------------------------------------------------------------------------------
        '  LOG de acciones.
        Set LOG = New cLOG
        Campo = "Albaran: " & codClien & " : " & NumCod
        LOG.Insertar 2, vUsu, Campo
        Set LOG = Nothing
        '-----------------------------------------------------------------------------

        'campo = txtCSB(0).Text & "|" & txtCSB(1).Text & "|" & txtCSB(2).Text & "|"
        
        'Esto es Marzo21
        If TaxcoFacurarUnAlbaranALVIC_Facturado <> "" Then
            'Es para TAXCO. Para cuando alguna factura se ha queado en ALbaran por un error de bloqueo o elimminando, o lo que sea
            'No pasa nunca, pero mejor tener el proceso y asi lanzarlo desde aqui
            TaxcoFacurarUnAlbaranALVIC_Facturado = Mid(TaxcoFacurarUnAlbaranALVIC_Facturado, 2)
            TraspasoFacturasGasol Cad, cadSelect, txtCodigo(34).Text, txtCodigo(0).Text, Me.PB_Fact, Me.lblProgess(1), False, codClien, TaxcoFacurarUnAlbaranALVIC_Facturado, CByte(vParamAplic.NumCopiasFacturacion), True
            TaxcoFacurarUnAlbaranALVIC_Facturado = "" 'reestablezco
        Else
            Campo = "|||"
            TraspasoAlbaranesFacturas Cad, cadSelect, txtCodigo(34).Text, txtCodigo(0).Text, Me.PB_Fact, Me.lblProgess(1), True, codClien, Campo, CByte(vParamAplic.NumCopiasFacturacion), False, False, UnoSolo, FacturaSuperaImporteTicket
        End If
    End If

    Screen.MousePointer = vbDefault
    
    If CambiamosConta Then
       'Reestablecer la conexion con la antigua conta
       AbrirConexionConta False
    End If
    '--- Ocultar Barra de Progreso
    If OpcionListado <> 222 Then '222: Facturas Mostrador
        Me.Height = Me.Height - 300
        Me.FrameFacturar.Height = Me.FrameFacturar.Height - 300
        Me.FrameProgress.visible = False
    Else
        If OpcionListado = 222 And vParamAplic.NumeroInstalacion = vbTaxco And codClien = "ALO" Then davidNumalbar = 0
    
        'Cierro y salgo
        Unload Me
    End If
End Sub

'Cuando genera UNA factura, comprobaremos que la fecha introducida es secuencial. No es anterior a una factura de esa serie
Private Function ComprobarSecuencialFactura() As Boolean

    ComprobarSecuencialFactura = False
    
    If vParamAplic.NumeroInstalacion = vbAlzira And codClien = "ART" Then
        'En alzira, las rectificativas , para usuario standard, solo puede ser de HOY
         If CDate(txtCodigo(34).Text) <> CDate(Format(Now, "dd/mm/yyyy")) Then
            cadParam = "La fecha deberia ser ser la de hoy" & vbCrLf '& " ¿Continuar?"
            
            MsgBox cadParam, vbExclamation
            txtCodigo(34).Text = Format(Now, "dd/mm/yyyy")
            Exit Function
        End If
    End If
    
        
    cadParam = DevuelveTipoFacturaDesdeAlbaran(codClien, FacturaSuperaImporteTicket)
    If cadParam <> "" Then
        cadFormula = Year(CDate(txtCodigo(34).Text))
        cadFormula = "fecfactu > " & DBSet(txtCodigo(34).Text, "F") & " AND fecfactu <= '" & cadFormula & "-12-31' AND codtipom"
        
        cadFormula = DevuelveDesdeBD(conAri, "count(*)", "scafac", cadFormula, cadParam, "T")
        If Val(cadFormula) > 0 Then
            cadFormula = "Hay facturas (" & cadFormula & ") de las serie  " & cadParam & " con fecha posterior a  " & txtCodigo(34).Text
            cadFormula = cadFormula & vbCrLf & "¿Continuar?"
            If MsgBox(cadFormula, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Function
        End If
    End If
    cadFormula = ""
    cadParam = ""
    ComprobarSecuencialFactura = True
End Function




Private Sub cmdAceptarGenAlb_Click()
'Solicitar datos para Generar Albaran a partir de un Pedido
Dim Cad As String

    'DAVID
    'Comprobar que me han puesto algun dato
    '-------------------------------------------------------------------
    Cad = ""
    If txtCodigo(17).Text = "" Or txtCodigo(18).Text = "" Or txtCodigo(19).Text = "" Or txtCodigo(25).Text = "" Then Cad = "M"
    If OpcionListado = 1000 Or OpcionListado = 1010 Then
        If txtCodigo(5).Text = "" Then Cad = "M"
        If txtNombre(5).Text = "" Then Cad = "M"
    End If
    If txtNombre(17).Text = "" Or txtNombre(18).Text = "" Or txtNombre(19).Text = "" Then Cad = "M"
    
    If Cad <> "" Then
        MsgBox "Campos obligatorios ", vbExclamation
        Exit Sub
    End If
    
    If vParamAplic.PtosAsignar > 0 Then
        If Me.FrameCanjePuntos.visible Then
            If txtCodigo(68).Text <> "" Then
                If ImporteFormateado(txtCodigo(68).Text) > CCur(txtCodigo(68).Tag) Then
                    MsgBox "No puede canjear mas de " & txtCodigo(68).Tag & " puntos", vbExclamation
                    Exit Sub
                End If
            End If
        End If
    End If
    
    
    
    If Me.FramRectARM.visible Then
        'Fra rectificativa desde ARM(teinsa)
        If Me.txtFraRectifica.Text = "" Then
            MsgBox "Debe indicar la factura a la que rectifica", vbExclamation
            Exit Sub
        End If
    End If
    
    
    If Me.FrameZona.visible Then
        Cad = ""
        If txtCodigo(22).Text = "" Xor txtNombre(22).Text = "" Then
            Cad = "Zona incorrecta"
            
        Else
            If txtCodigo(22).Text = "" Then Cad = "Zona incorrecta. Indique una"
        End If
        If Cad <> "" Then
            MsgBox Cad, vbExclamation
            Exit Sub
        End If
    End If
        
    If Me.chkFraMostrador.visible Then
    
        
        'SI NO VA A FACTURA DE MOSTRADOR AVISAMOS
        If chkFraMostrador.Value = 0 Then
        
            
            Cad = String(60, "*") & vbCrLf
            Cad = vbCrLf & Cad & Cad & Cad & vbCrLf
            
            If vParamAplic.NumeroInstalacion <> 1 Then
            
                'EN HERBERLCA FACTURAS SOLO MOSTRADOR
                If vParamAplic.NumeroInstalacion = vbHerbelca And OpcionListado = 1000 Then
                    'Mayo 2021
                    'SOLO mostrador
                    Cad = Cad & "   SOLO puede generar facturas de mostrador"
                    MsgBox Cad, vbExclamation
                    
                    Exit Sub
                Else
                    'Lo que habia
                    Cad = Cad & "    Va a generar una factura de albarán,   NO de mostrador" & "" & vbCrLf & vbCrLf & "             ¿Continuar?" & vbCrLf & Cad
                    If MsgBox(Cad, vbQuestion + vbYesNoCancel + vbDefaultButton3) <> vbYes Then Exit Sub
                End If
            End If
        End If
    End If
       
       
       
    If FramePartes.visible Then
        If Me.txtCodigo(59).Text = "" Then
            MsgBox "Introduzca litros reales", vbExclamation
            PonerFoco txtCodigo(59)
            Exit Sub
        End If
        
        'MSGBOX
        Cad = "Desea cerrar el parte de produccion con:" & vbCrLf
        Cad = Cad & "       Litros real: " & txtCodigo(59).Text & vbCrLf
        Cad = Cad & "       Cantidad: " & txtCodigo(60).Text & vbCrLf
        If MsgBox(Cad, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
        
        CadenaDesdeOtroForm = txtCodigo(59).Text & "|" & Me.chkFrInterna.Value & "|" & txtCodigo(60).Text & "|"
        ' 17 Abril 2012
        'Tipo facturacion
        CadenaDesdeOtroForm = CadenaDesdeOtroForm & Me.cboFacturacion.ListIndex & "|"
        
        
    End If
    
    
    
    
        
    'Enero 2011
    'la factura a la k rectifica la llevo a una temporal, para no crear mas variables
    If Me.FramRectARM.visible Then
        Cad = "Insert into tmpcrmmsg (`codusu`,`codigo`,`asun_obs`) values (" & vUsu.Codigo & ",'1'," & DBSet(Me.txtFraRectifica.Tag, "T") & ")"
        If Not ejecutar(Cad, False) Then Exit Sub
    End If
    
    Cad = txtCodigo(17).Text & "|"
    Cad = Cad & txtCodigo(18).Text & "|"
    Cad = Cad & txtCodigo(19).Text & "|"
    Cad = Cad & txtCodigo(25).Text & "|"
    Cad = Cad & Me.chkImpAlbaran.Value & "|"
    Cad = Cad & Me.chkImpEtiq.Value & "|"
    Cad = Cad & Me.chkImpHojaExped.Value & "|"
    'mando el banco propio
    'Octubre 2010
    '--------------------------------------------------------------
    'Ahora siempre mando el banco. Cuando es albaran NO le hago caso
    Cad = Cad & txtCodigo(5).Text & "|"
    'Y mando tb, si esta visible, la zona d eenvio
    If Me.FrameZona.visible Then Cad = Cad & txtCodigo(22).Text
    Cad = Cad & "|"
    
    '10 Enero 2011
    'Si tiene vParamAplic.FrasMostradorSerieDistinta podria pasar el pedido a
    'fra mostrador. Se le habra puesto visible el check
    If Me.chkFraMostrador.visible Then
        If vParamAplic.FrasMostradorSerieDistinta Then
            If Me.chkFraMostrador.Value = 1 Then Cad = Cad & "1" 'PASAMOS A FRAS/alb MOSTRADOR(primero a alb mostrador)
        End If
    End If
    Cad = Cad & "|"
    
    If InstalacionEsEulerTaxco Then
        Select Case Me.cboTipoAlbaranEuler.ListIndex
        Case 1
            Cad = Cad & "ALE"
        Case 2
            Cad = Cad & "ALO"
        Case 3
            Cad = Cad & "ALR"
        Case Else
            Cad = Cad & "ALV"
        End Select
        
        
    Else
         If vParamAplic.NumeroInstalacion = vbFenollar Then Cad = Cad & IIf(cboDestinoB.ListIndex = 1, "ALZ", "ALV")
         
    End If
    Cad = Cad & "|"
    
    
    
    'Numero de bultos
    If Not Me.FrameBultosHerbelca.visible Then
        Cad = Cad & "0"
    Else
        Cad = Cad & CInt(Val(Me.txtCodigo(62).Text))
    End If
    Cad = Cad & "|"
    
    'Mayo 2018
    '--------------------
    If vParamAplic.PtosAsignar > 0 Then
        If Me.FrameCanjePuntos.visible Then
            If txtCodigo(68).Text <> "" Then Cad = Cad & ImporteFormateado(txtCodigo(68).Text)
        End If
    End If
    Cad = Cad & "|"
    
    RaiseEvent DatoSeleccionado(Cad)
    
    
    
    Unload Me
End Sub




Private Sub cmdAceptarPedxArtic_Click()
'41: Informe de Pedidos por Articulo
'44: Informe de Pedidos por Cliente
'49: Informe de Albaranes por Artículo
Dim Campo As String
Dim Cad As String
Dim Sql As String
Dim cadFormula2 As String
Dim Cadselect2 As String
Dim cadSelect3 As String
Dim Indice As Integer

Dim Tocho As String
Dim fec As Date
    InicializarVbles
    cadFormula2 = ""
    Cadselect2 = ""
    cadSelect3 = ""
    
    'Pasar nombre de la Empresa como parametro
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1


    '===================================================
    '================= FORMULA =========================
    
    'Cadena para seleccion Fechas de Pedido/Albaran/Factura
    '--------------------------------------------
    'Desde/Hasta FECHA
    'para el informe 227 fecha requerida
    If OpcionListado = 227 Then
        If txtCodigo(11).Text = "" Or txtCodigo(12).Text = "" Then
            MsgBox "Los campos D/H fecha factura deben tener valor.", vbInformation
            Exit Sub
        End If
        
        If DateDiff("d", txtCodigo(11).Text, txtCodigo(12).Text) > 365 Then
            MsgBox "El intervalo de fechas no puede ser superior a un año.", vbInformation
            Exit Sub
        End If
    End If
    
    If txtCodigo(11).Text <> "" Or txtCodigo(12).Text <> "" Then
        If OpcionListado = 227 Or OpcionListado = 228 Then
            Campo = "{" & NomTabla & ".fecfactu}"
        ElseIf OpcionListado = 49 Then
            Campo = "{" & NomTabla & ".fechaalb}"
        Else
            Campo = "{" & NomTabla & ".fecpedcl}"
        End If
        Cad = "pDHFecha=""Fecha: "
        If Not PonerDesdeHasta(Campo, "F", 11, 12, Cad) Then Exit Sub
        cadSelect = CadenaDesdeHastaBD(txtCodigo(11).Text, txtCodigo(12).Text, Campo, "F")
        
        'Guardamos el periodo para calcular las ventas
        If OpcionListado = 227 Then
            cadFormula2 = cadFormula
            Cadselect2 = cadSelect
            'obtenemos el periodo anterior de ventas
            Cad = "": Sql = ""
            If txtCodigo(11).Text <> "" Then Cad = Day(txtCodigo(11).Text) & "/" & Month(txtCodigo(11).Text) & "/" & Year(txtCodigo(11).Text) - 1
            If txtCodigo(12).Text <> "" Then Sql = Day(txtCodigo(12).Text) & "/" & Month(txtCodigo(12).Text) & "/" & Year(txtCodigo(12).Text) - 1
            cadSelect3 = CadenaDesdeHastaBD(Cad, Sql, Campo, "F")
        
        ElseIf OpcionListado = 41 Or OpcionListado = 42 Then '42:Disponibilidad Stock
        'pasar D/H fecha como parametro para enlazar con la cabecera de pedidos proveedor
        'que esta como subinforme y que seleccione el mismo rango de fecha que
        'para la cabecera de pedidos de cliente
            If txtCodigo(11).Text <> "" Then
                Cad = "pFechaD=" & "Date(" & Year(txtCodigo(11).Text) & ", " & Month(txtCodigo(11).Text) & ", " & Day(txtCodigo(11).Text) & ")"
            Else
                Cad = "pFechaD=" & "Date(1900,01,01)"
            End If
            cadParam = cadParam & Cad & "|"
            numParam = numParam + 1
            If txtCodigo(12).Text <> "" Then
                Cad = "pFechaH=" & "Date(" & Year(txtCodigo(12).Text) & ", " & Month(txtCodigo(12).Text) & ", " & Day(txtCodigo(12).Text) & ")"
            Else
                Cad = "pFechaH=" & "Date(9999,12,31)"
            End If
            cadParam = cadParam & Cad & "|"
            numParam = numParam + 1
        End If
    End If
    
    'Cadena para seleccion ALMACEN
    '--------------------------------------------
    If Me.Frame9.visible Then
        If txtCodigo(13).Text <> "" Or txtCodigo(14).Text <> "" Then
            Campo = "{" & NomTablaLin & ".codalmac}"
            'Parametro Desde/Hasta Almacen
            Cad = "pDHAlmacen=""Almacen: "
            If Not PonerDesdeHasta(Campo, "N", 13, 14, Cad) Then Exit Sub
        End If
    End If
    
    
    'Cadena para seleccion ARTICULO
    '--------------------------------------------
    If Me.Frame8.visible Then
        If txtCodigo(15).Text <> "" Or txtCodigo(16).Text <> "" Then
            Campo = "{" & NomTablaLin & ".codartic}"
            'Parametro Desde/Hasta Articulo
            Cad = "pDHArticulo=""Artículo: "
             If Not PonerDesdeHasta(Campo, "T", 15, 16, Cad) Then Exit Sub
        End If
    End If
    
    'Cadena para seleccion CLIENTE
    '--------------------------------------------
    If Me.Frame5.visible Then
        If OpcionListado = 44 Then
            CadenaDesdeOtroForm = ""
            If txtCodigo(20).Text <> "" Or txtCodigo(21).Text <> "" Then
                Campo = "{" & NomTabla & ".codclien}"
                'Parametro Desde/Hasta Cliente
                Cad = "Cliente: "
                If Not PonerDesdeHasta(Campo, "N", 20, 21, Cad) Then Exit Sub
            End If
            If Cad <> "" Then CadenaDesdeOtroForm = Cad
            If chkPedxClixSemEntrega(0).Value = 0 Then
                If Me.cboTipocliente.ItemData(cboTipocliente.ListIndex) >= 0 Then
                    'Tiene seleccionado UN tipo de cliente
                    CadenaDesdeOtroForm = CadenaDesdeOtroForm & "    Tipo: " & Me.cboTipocliente.List(cboTipocliente.ListIndex)
                End If
            End If
            If CadenaDesdeOtroForm <> "" Then
                Cad = "pDHCliente=""" & CadenaDesdeOtroForm & """|"
                cadParam = cadParam & Cad
                numParam = numParam + 1
                CadenaDesdeOtroForm = ""
            End If
                
                        
            Cad = "observas=" & chkPedxClixSemEntrega(1).Value & "|"
            cadParam = cadParam & Cad
            numParam = numParam + 1
            
            
                
                
        Else
            If txtCodigo(20).Text <> "" Or txtCodigo(21).Text <> "" Then
                Campo = "{" & NomTabla & ".codclien}"
                'Parametro Desde/Hasta Cliente
                Cad = "pDHCliente=""Cliente: "
                If Not PonerDesdeHasta(Campo, "N", 20, 21, Cad) Then Exit Sub
            End If
        End If
    End If
    
    
    'Cadena para seleccion TRABAJADOR
    '--------------------------------------------
    If Me.Frame12.visible Then
        If txtCodigo(2).Text <> "" Or txtCodigo(3).Text <> "" Then
            Campo = "{scafac1.codtraba}"
            'Parametro Desde/Hasta Trabajador
            Cad = "pDHTrabajador=""Trabajador: "
            If Not PonerDesdeHasta(Campo, "N", 2, 3, Cad) Then Exit Sub
        End If
    End If
    
    
    
    '227: Listado Ventas por cliente
    'Importe ventas superior a ....
    If Me.Frame10.visible Then
        If txtCodigo(1).Text <> "" Then
            Cad = DBSet(txtCodigo(1).Text, "N")
        Else
            Cad = ""
        End If
        
        cadParam = cadParam & "pImporte=" & Cad & "|"
        numParam = numParam + 1
            
        'En este sub meteremos la actividad tb
        If txtCodigo(1).Text <> "" Or txtCodigo(57).Text <> "" Or txtCodigo(58).Text <> "" Or txtCodigo(23).Text <> "" Or txtCodigo(24).Text <> "" Then
            'seleccionar solo los clientes que el total de la BaseImp supere esa cantidad
            'y esten en el desde hasta que marcamos aqui
            If cadSelect <> "" Then Sql = Cadselect2 & " AND "
            Cad = ObtenerClientesNuevo(cadSelect, Cad)
            cadSelect = Sql & Cad
'            If cadSelect3 <> "" Then cadSelect3 = cadSelect3 & " AND "
'            cadSelect3 = cadSelect3 & cad
            If cadFormula2 <> "" Then cadFormula2 = cadFormula2 & " AND "
            cadFormula = cadFormula2 & Cad
        End If
        
        
    End If
    
    
    If OpcionListado = 49 Then
        Campo = ".numalbar"
'        cad = "{" & NomTabla & ".codtipom}='ALV'"
'        If Not AnyadirAFormula(cadFormula, cad) Then Exit Sub
'        If Not AnyadirAFormula(cadSelect, cad) Then Exit Sub
        '-- Ahora en este informe hay mas posibilidades de selección [SERVICIOS]
        If vParamAplic.Servicios Then
            Indice = cmbTipAlbaran(1).ListIndex
            If Indice < 0 Then
                MsgBox "Debe seleccionar el tipo o los tipos de alabarán a procesar", vbExclamation
                Exit Sub
            Else
                Select Case Indice
                    Case 0 ' solo ventas
                        Cad = "{" & NomTabla & ".codtipom}='ALV'"
                        If Not AnyadirAFormula(cadFormula, Cad) Then Exit Sub
                        If Not AnyadirAFormula(cadSelect, Cad) Then Exit Sub
                        Titulo = "Albaranes por artículo (Ventas)"
                    Case 1 ' solo servicios
                        Cad = "{" & NomTabla & ".codtipom}='ALS'"
                        If Not AnyadirAFormula(cadFormula, Cad) Then Exit Sub
                        If Not AnyadirAFormula(cadSelect, Cad) Then Exit Sub
                        Titulo = "Albaranes por artículo (Servicios)"
                    Case 2 ' ventas y servicios
                        Cad = " ({" & NomTabla & ".codtipom}='ALV'" & _
                                " OR {" & NomTabla & ".codtipom}='ALS')"
                        If Not AnyadirAFormula(cadFormula, Cad) Then Exit Sub
                        If Not AnyadirAFormula(cadSelect, Cad) Then Exit Sub
                        Titulo = "Albaranes por artículo (Ventas y servicios)"
                End Select
            End If
        Else
            Cad = "{" & NomTabla & ".codtipom}='ALV'"
            If Not AnyadirAFormula(cadFormula, Cad) Then Exit Sub
            If Not AnyadirAFormula(cadSelect, Cad) Then Exit Sub
            Titulo = "Albaranes por artículo (Ventas)"
        End If
        'Pasar nombre el título del informe
        cadParam = cadParam & "|pTitulo=""" & Titulo & """|"
        numParam = numParam + 1
    Else
        Campo = ".numpedcl"
        Titulo = "Listado pedidos"
    End If


    'Febrero.
    'Si lis ped x cliente
    ' y no va por sememan de entrega entonces
    If OpcionListado = 44 Then
    
        
    
        'Ha selecionado un tipo de cliente
        If Me.chkPedxClixSemEntrega(0).Value = 0 Then
            If Me.cboTipocliente.ItemData(cboTipocliente.ListIndex) >= 0 Then
                'Voy montar un select que devuelva de los clientes que esan en pedidos
                'y sean del tipo ese
                Cad = DevuelveClientesPedidosPorTipo
                If Cad = "" Then Cad = "-1"
                CadenaDesdeOtroForm = "{" & NomTabla & ".codclien} IN [" & Cad & "]"
                If Not AnyadirAFormula(cadFormula, CadenaDesdeOtroForm) Then Exit Sub
                CadenaDesdeOtroForm = "{" & NomTabla & ".codclien} IN (" & Cad & ")"
                If Not AnyadirAFormula(cadSelect, CadenaDesdeOtroForm) Then Exit Sub
                CadenaDesdeOtroForm = ""
            End If
        End If
    End If
    
    
''''''''''''''''''''    If vParamAplic.NumeroInstalacion = vbFenollar Then
''''''''''''''''''''
''''''''''''''''''''        Cad = "{" & NomTabla & ".cerrado} = 0"
''''''''''''''''''''        If Not AnyadirAFormula(cadFormula, Cad) Then Exit Sub
''''''''''''''''''''        If Not AnyadirAFormula(cadSelect, Cad) Then Exit Sub
''''''''''''''''''''
''''''''''''''''''''        If OpcionListado = 41 And Me.chkPedxClixSemEntrega(3).Value = 1 Then
''''''''''''''''''''            'Listado agrupado por articulo unicamente por articulo  almacen. Solo contro stock
''''''''''''''''''''              Cad = "{sartic.ctrstock} = 1"
''''''''''''''''''''            If Not AnyadirAFormula(cadFormula, Cad) Then Exit Sub
''''''''''''''''''''           '  cad = "{sartic.ctrstock} = 1"
''''''''''''''''''''           ' If Not AnyadirAFormula(cadSelect, cad) Then Exit Sub
''''''''''''''''''''        End If
''''''''''''''''''''    End If
    
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    If OpcionListado = 227 Then
    
        
    
        Cad = NomTabla
        
        If chkPedxClixSemEntrega(4).Value = 1 Then
            If Check1chkAgrupaAg.Value = 1 Then
                MsgBox "No se agrupa por agente", vbExclamation
                Check1chkAgrupaAg.Value = 0
            End If
            'Noviembre 2019
            'Anual
            Titulo = "Ventas por cliente anual"
            nomRPT = "rFacVentasxClienPeriodoComparativo.rpt"
        Else
            'Lo que habia
            Titulo = "Ventas por Cliente"
            If Me.Check1chkAgrupaAg.Value = 1 Then
                'Agrupa cliente
                Titulo = Titulo & " (cliente)"
                nomRPT = "rFacVentasxClienAg.rpt"
            Else
                nomRPT = "rFacVentasxClien.rpt"
            End If
            
        End If
        
        
        conSubRPT = False
    ElseIf OpcionListado = 228 Then
        Cad = NomTabla & " INNER JOIN scafac1 ON scafac.codtipom=scafac1.codtipom and scafac.fecfactu=scafac1.fecfactu and scafac.numfactu=scafac1.numfactu"
        Titulo = "Ventas por Trabajador"
        If Me.OptDetalle(2).Value = True Then 'Inf. Detalle
            
            nomRPT = DevuelveDesdeBD(conAri, "documrpt", "scryst", "codcryst", "92")
            
            If nomRPT = "" Then nomRPT = "rFacVentasxTrabaDet.rpt"
            conSubRPT = True
        ElseIf Me.OptResumen.Value = True Then 'Inf. Resum
            nomRPT = "rFacVentasxTrabaRes.rpt"
            conSubRPT = False
        End If
    Else
       Cad = NomTabla & " INNER JOIN " & NomTablaLin
       Cad = Cad & " ON " & NomTabla & Campo & "=" & NomTablaLin & Campo
       If OpcionListado = 49 Then Cad = Cad & " AND " & NomTabla & ".codtipom=" & NomTablaLin & ".codtipom "
    End If
    
    If Not HayRegParaInforme(Cad, cadSelect) Then Exit Sub
    
    
    If OpcionListado = 227 Then
        BorrarTempInformes
        
        'Pasar los datos a la tabla temporal tmpInformes y luego mostrar
        'el informe de esta tabla
        Cadselect2 = Replace(Cadselect2, "{", "")
        Cadselect2 = Replace(Cadselect2, "}", "")
        
        cadSelect3 = Replace(cadSelect3, "{", "")
        cadSelect3 = Replace(cadSelect3, "}", "")
        
        
        'JULIO 2013
        'QUe tipo de facturas entran
        Cad = ""
        Sql = ""
        'LOS SELECCIONADOS

        For Indice = 0 To Me.ListTipoFact.ListCount - 1
            'Siempre 3 carcateres
            If Me.ListTipoFact.Selected(Indice) Then
                Cad = Cad & Mid(Me.ListTipoFact.List(Indice), 1, 3) & "|"
                Sql = Sql & "X"
            End If
        Next
    
        
        If chkPedxClixSemEntrega(4).Value = 1 Then
            'Nuevo informe ANUAL
            'AÑADIREMOS DESDE HASTA tipocredito
            
        Else
            'lo que habia
              If Len(Sql) > 5 Or Len(Sql) < 1 Then
                  MsgBox "Maximo 5 tipos de facturas", vbExclamation
                   Exit Sub
              Else
                  
                  If Len(Sql) = 5 Then
                      'Si son 5 tipos de factura y NO esta agrupado por cliente es otro rpt
                      If Me.Check1chkAgrupaAg.Value = 0 Then nomRPT = "rFacVentasxClien5.rpt"
                  End If
                  Sql = ""
              End If
            
        End If
        
        Screen.MousePointer = vbHourglass
        
        
        Sql = ""
        If txtCodigo(64).Text <> "" Then Sql = Sql & " AND (sclien.codzonas)>= " & txtCodigo(64).Text
        If txtCodigo(65).Text <> "" Then Sql = Sql & " AND (sclien.codzonas)<= " & txtCodigo(65).Text
        cadSelect = cadSelect & Sql
        Cadselect2 = Cadselect2 & Sql
        cadSelect3 = cadSelect3 & Sql
        
        Sql = ""
        If txtCodigo(66).Text <> "" Then Sql = Sql & " AND (sclien.codrutas)>= " & txtCodigo(66).Text
        If txtCodigo(67).Text <> "" Then Sql = Sql & " AND (sclien.codrutas)<= " & txtCodigo(67).Text
        cadSelect = cadSelect & Sql
        Cadselect2 = Cadselect2 & Sql
        cadSelect3 = cadSelect3 & Sql
        
        
        Sql = ""
        If Me.txtCodigo(23).Text <> "" Then Sql = Sql & " AND scafac.codagent >= " & txtCodigo(23).Text
        If Me.txtCodigo(24).Text <> "" Then Sql = Sql & " AND scafac.codagent <= " & txtCodigo(24).Text
        cadSelect = cadSelect & Sql
        Cadselect2 = Cadselect2 & Sql
        cadSelect3 = cadSelect3 & Sql
        
         If chkPedxClixSemEntrega(4).Value = 1 And vParamAplic.OperacionesAseguradas Then
            
            Sql = ""
            If Me.cboTipoCredito.ListIndex >= 1 Then Sql = Sql & " AND sclien.credipriv = " & cboTipoCredito.ItemData(cboTipoCredito.ListIndex)
            cadSelect = cadSelect & Sql
            'Cadselect2 = Cadselect2 & SQL
            'cadSelect3 = cadSelect3 & SQL
           
        End If
        
        
        If chkPedxClixSemEntrega(4).Value = 1 Then
            'Noviembre 2019
            If cboAnyos.ListIndex = -1 Then cboAnyos.ListIndex = 0
            If Not TempVentasClientesPeriodoComparativo(Me.Check1chkAgrupaAg.Value = 1, cadSelect, Cadselect2, cadSelect3, Label4(54), Cad, CDate(txtCodigo(11).Text), CDate(txtCodigo(12).Text), CInt(cboAnyos.Text)) Then Exit Sub
            cadSelect3 = "0"
        Else
            'lo que hacia antes
            If Not TempVentasClientes(Me.Check1chkAgrupaAg.Value = 1, cadSelect, Cadselect2, cadSelect3, Label4(54), Cad) Then Exit Sub
        End If
        
        Cad = DevuelveDesdeBD(conAri, "count(*)", "tmpinformes", "codusu", CStr(vUsu.Codigo))
        If Cad = "" Then Cad = "0"
        If Val(Cad) = 0 Then
            MsgBox "No existen datos ", vbExclamation
            Exit Sub
        End If
        
        
        
        'Añadir como parametros el total del periodo que devuelve en cadSelect2
        'y añadir el parametro del total periodo anterior q devuelve en cadSelect3
        cadParam = cadParam & "pTotal=" & Cadselect2 & "|"
        numParam = numParam + 1
        cadParam = cadParam & "pTotalAnt=" & cadSelect3 & "|"
        numParam = numParam + 1
        
        'Añadir el parametro para el orden del informe
        'Orden del Informe
        If Me.OptOrdenCodclien.Value Then
            Cad = "{tmpinformes.codigo1}"
            Sql = "Orden: Cod. cliente"
        ElseIf Me.OptOrdenNomclien.Value Then
            Cad = "{tmpinformes.nombre1}"
            Sql = "Orden: Nombre cliente"
        ElseIf Me.OptOrdenVentas.Value Then
           
            Cad = "{@paraOrdenarVolumen}" '"{tmpinformes.importe5}"
            Sql = "Orden: Volumen ventas"
        End If
        
        
        
        If chkPedxClixSemEntrega(4).Value = 1 And vParamAplic.OperacionesAseguradas Then
            If Me.cboTipoCredito.ListIndex > 0 Then Sql = Sql & "   " & UCase(cboTipoCredito.Text)
        End If
        
        'Si lleva d/H agente
        Sql = Sql & "               "
        Campo = ""
        If txtCodigo(23).Text <> "" Then Campo = Campo & "    desde " & txtCodigo(23).Text & "  " & Me.txtNombre(23).Text
        If txtCodigo(24).Text <> "" Then Campo = Campo & "    hasta " & txtCodigo(24).Text & "  " & Me.txtNombre(24).Text
        If Campo <> "" Then Campo = "Agente " & Trim(Campo)
        Sql = Sql & Campo
        
        'Si lleva actividad
        Sql = Sql & "     "
        Campo = ""
        If txtCodigo(57).Text <> "" Then
            Campo = Campo & "    desde " & txtCodigo(57).Text
            If Len(Sql) < 115 Then Campo = Campo & "  " & Me.txtNombre(57).Text
        End If
        If txtCodigo(58).Text <> "" Then
            Campo = Campo & "    hasta " & txtCodigo(58).Text
            If Len(Sql) < 125 Then Campo = Campo & "  " & Me.txtNombre(58).Text
        End If
        If Campo <> "" Then Campo = "Act: " & Trim(Campo)
        Sql = Sql & Campo
        
        
        'Si lleva ZONA
        Sql = Sql & "     "
        Campo = ""
        If txtCodigo(64).Text <> "" Then
            Campo = Campo & "    desde " & txtCodigo(64).Text
            If Len(Sql) < 115 Then Campo = Campo & "  " & Me.txtNombre(64).Text
        End If
        If txtCodigo(65).Text <> "" Then
            Campo = Campo & "    hasta " & txtCodigo(65).Text
            If Len(Sql) < 125 Then Campo = Campo & "  " & Me.txtNombre(65).Text
        End If
        If Campo <> "" Then Campo = "Zona: " & Trim(Campo)
        Sql = Sql & Campo
        
        'Si lleva RUTA
        If txtCodigo(66).Text <> "" Or txtCodigo(67).Text <> "" Then
            
            Sql = Sql & "     "
            Campo = ""
            If txtCodigo(66).Text <> "" Or txtCodigo(67).Text <> "" Then
                Campo = txtCodigo(66).Text & " " & txtNombre(66).Text
            Else
                If txtCodigo(66).Text <> "" Then
                    Campo = Campo & "    desde " & txtCodigo(66).Text
                    If Len(Sql) < 115 Then Campo = Campo & "  " & Me.txtNombre(66).Text
                End If
                If txtCodigo(67).Text <> "" Then
                    Campo = Campo & "    hasta " & txtCodigo(67).Text
                    If Len(Sql) < 125 Then Campo = Campo & "  " & Me.txtNombre(67).Text
                End If
            End If
            If Campo <> "" Then Campo = IIf(vParamAplic.NumeroInstalacion = 2, "Asociacion:", "Ruta:") & Trim(Campo)
            Sql = Sql & Campo
        End If
        
        'JULIO 2013
        'Tipos de moviemientos que van incluidos
        Sql = Trim(Sql) & "   Fact:"
        cadSelect3 = "0"
        For Indice = 0 To Me.ListTipoFact.ListCount - 1
            'Siempre 3 carcateres
            If Me.ListTipoFact.Selected(Indice) Then
                cadSelect3 = Val(cadSelect3) + 1
                Sql = Sql & "  " & Mid(Me.ListTipoFact.List(Indice), 1, 3)
                Cadselect2 = Mid(Me.ListTipoFact.List(Indice), 1, 3)
                If Cadselect2 = "FAE" Then
                    Cadselect2 = "Exterior"
                ElseIf Cadselect2 = "FAO" Then
                    Cadselect2 = "Orden Tr."
                Else
                    'LO que habia
                    Cadselect2 = Trim(Mid(Me.ListTipoFact.List(Indice), 6))
                End If
                If InStr(1, UCase(Cadselect2), "FACTURA") > 0 Then
                    Cadselect2 = Trim(Mid(Cadselect2, InStr(1, UCase(Cadselect2), "FACTURA") + 8))
                End If
                If Len(Cadselect2) > 5 Then Cadselect2 = Mid(Cadselect2, 1, 5) & "."
                
                If chkPedxClixSemEntrega(4).Value = 1 Then
                    'No hacemos nada
                
                Else
                    cadParam = cadParam & "C" & cadSelect3 & "=""" & Cadselect2 & """|"
                    numParam = numParam + 1
                End If
                
            End If
        Next
        If chkPedxClixSemEntrega(4).Value = 1 Then
            For Indice = 2 To 5
                
                If Indice > Val(Me.cboAnyos.Text) Then
                    Cadselect2 = ""
                Else
                    fec = CDate(txtCodigo(11).Text)
                    fec = DateAdd("yyyy", -(Indice - 1), fec)
                    Cadselect2 = Format(fec, "dd/mm/yy")
                    fec = CDate(txtCodigo(12).Text)
                    fec = DateAdd("yyyy", -(Indice - 1), fec)
                    Cadselect2 = Cadselect2 & "-" & Format(fec, "dd/mm/yy")
                    
                    
                End If
                cadParam = cadParam & "C" & Indice & "=""" & Cadselect2 & """|"
                numParam = numParam + 1
            Next
        End If
        
        
        cadParam = cadParam & "pOrden=" & Cad & "|"
        numParam = numParam + 1
        cadParam = cadParam & "pCadOrden=""" & Sql & """|"
        numParam = numParam + 1
        
        
        'no le pasamos formula de seleccion porque los datos ya estan en la temporal
        'solo el usuario que genero la informacion en la temporal
        cadFormula = "{tmpinformes.codusu}=" & vUsu.Codigo
        
    ElseIf OpcionListado = 44 Then
        'If Me.optOrdePed(0).Value Then
        '    Cad = "{sliped.codartic}"
        'Else
            Cad = "{scaped.numpedcl}"
        'End If
        cadParam = cadParam & "rOrden=" & Cad & "|"
        numParam = numParam + 1
        
        
        
        'MArzo 2010
        Sql = ""
        If txtCodigo(6).Text <> "" Or txtCodigo(7).Text <> "" Then
            Campo = "{scaped.codagent}"
            'Parametro Desde/Hasta agente
            Cad = "@=""Agente: "
            If Not PonerDesdeHasta(Campo, "N", 6, 7, Cad) Then Exit Sub
            Sql = Mid(Cad, 4)
        End If
        If txtCodigo(9).Text <> "" Or txtCodigo(10).Text <> "" Then
            Campo = "{sclien.codzonas}"
            'Parametro Desde/Hasta zona
            Cad = "@=""Zonas: "
            If Not PonerDesdeHasta(Campo, "N", 9, 10, Cad) Then Exit Sub
            Cad = Mid(Cad, 4)
            Sql = Trim(Sql & "    " & Cad)
        End If
        If Sql <> "" Then
            Sql = """" & Sql & """"
            cadParam = cadParam & "pdHAgenZona= " & Sql & "|"
            numParam = numParam + 1
            Sql = ""
        End If
        
        
        'trampa
        'Si ha marcado no agrupar por semana sale otro report.
        
        If Me.chkPedxClixSemEntrega(2).Value = 1 Then
            OpcionListado = 2051
            Sql = "{scaped.numpedcl}"
            If Me.chkPedxClixSemEntrega(0).Value = 1 Then Sql = "{scaped.sementre}"
            'If cadFormula <> "" Then cadFormula = cadFormula & " AND "
            'cadFormula = cadFormula & " {tmpscapla.codusu} =" & vUsu.codigo
            cadParam = cadParam & "Grupo= " & Sql & "|"
            numParam = numParam + 1
        Else
        
            If Me.chkPedxClixSemEntrega(0).Value = 0 Then OpcionListado = 46
        End If
        
        
    ElseIf OpcionListado = 42 Then
        'DISPONIBILIDAD DE STOCKS
        '20 Enero 2010
        Screen.MousePointer = vbHourglass
        Label4(54).Caption = "Prepara datos"
        Label4(54).Refresh           'La cuolumna importel llevara el stock del almacen ppal(el 1)
        Sql = "DELETE FROM tmpsliped where codusu = " & vUsu.Codigo
        conn.Execute Sql
        
        
        'Meto en la tmp.  En codclien pondre el codprove
        Label4(54).Caption = "Proc 1"
        Label4(54).Refresh
        
                
        
        
        
        Sql = ""
        Sql = Sql & " SELECT " & vUsu.Codigo & ",0,0,codalmac,sliped.codartic,sliped.nomartic,sum(cantidad),sartic.codprove,rotacion,artvario "
        Sql = Sql & " FROM scaped,sliped,sartic WHERE scaped.numpedcl =sliped.numpedcl AND "
        Sql = Sql & " sartic.codartic=sliped.codartic AND sartic.ctrstock=1  And artvario = 0 and cerrado=0"
        
        'Si no quiere con departamentos(obra
        If Me.chkDispo(1).Value = 0 Then Sql = Sql & " AND scaped.coddirec is null"
        'El select
        If cadSelect <> "" Then Sql = Sql & " AND " & cadSelect
        Sql = Sql & " GROUP BY codalmac,sliped.codartic"
        Tocho = ""
        Set miRsAux = Nothing
        Set miRsAux = New ADODB.Recordset
        miRsAux.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            Sql = ""
            For Indice = 0 To miRsAux.Fields.Count - 1
                If Indice = 4 Or Indice = 5 Then
                    Sql = Sql & ", " & DBSet(miRsAux.Fields(Indice), "T")
                Else
                    Sql = Sql & ", " & DBSet(miRsAux.Fields(Indice), "N")
                End If
            Next Indice
            Sql = ", (" & Mid(Sql, 2) & ")"
            Tocho = Tocho & Sql
            miRsAux.MoveNext
            
            If miRsAux.EOF Then
                Indice = 0
            Else
                If Len(Tocho) > 12000 Then Indice = 0
            End If
            
            If Indice = 0 Then
                
                Tocho = Mid(Tocho, 2)
                Sql = "INSERT INTO tmpsliped(codusu,numpedcl,numlinea,codalmac,codartic,nomartic,cantidad,codclien,numbultos,codzona) VALUES " & Tocho
                conn.Execute Sql
                Tocho = ""
            End If
        Wend
        miRsAux.Close
        
        'Febrero 2013
        'Los de varios entraran . En el proceso final actualizaremos su stock a CERO
        Sql = " SELECT " & vUsu.Codigo & ",0,0,codalmac,sliped.codartic,sliped.nomartic,sum(cantidad),sartic.codprove,rotacion,artvario "
        Sql = Sql & " FROM scaped,sliped,sartic WHERE scaped.numpedcl =sliped.numpedcl AND "
        Sql = Sql & " sartic.codartic=sliped.codartic AND cerrado=0 and artvario = 1"
        'Si no quiere con departamentos(obra
        If Me.chkDispo(1).Value = 0 Then Sql = Sql & " AND scaped.coddirec is null"
        'El select
        If cadSelect <> "" Then Sql = Sql & " AND " & cadSelect
        Sql = Sql & " GROUP BY codalmac,sliped.codartic"
        
        Tocho = ""
        miRsAux.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            Sql = ""
            For Indice = 0 To miRsAux.Fields.Count - 1
                If Indice = 4 Or Indice = 5 Then
                    Sql = Sql & ", " & DBSet(miRsAux.Fields(Indice), "T")
                Else
                    Sql = Sql & ", " & DBSet(miRsAux.Fields(Indice), "N")
                End If
            Next Indice
            Sql = ", (" & Mid(Sql, 2) & ")"
            Tocho = Tocho & Sql
            miRsAux.MoveNext
            
            If miRsAux.EOF Then
                Indice = 0
            Else
                If Len(Tocho) > 12000 Then Indice = 0
            End If
            
            If Indice = 0 Then
                
                Tocho = Mid(Tocho, 2)
                Sql = "INSERT IGNORE INTO tmpsliped(codusu,numpedcl,numlinea,codalmac,codartic,nomartic,cantidad,codclien,numbultos,codzona) VALUES " & Tocho
                conn.Execute Sql
                Tocho = ""
            End If
        Wend
        miRsAux.Close
        
        
        
        
        
        
        'Acciones. Proceso final
        ObtenerValores
        Label4(54).Caption = ""
        
        
        'Pongo los nuevos valores para la cadformula
        Cad = "{tmpsliped.codusu}=" & vUsu.Codigo
        cadFormula = ""
        If Not AnyadirAFormula(cadFormula, Cad) Then Exit Sub
        'Añado si detalla o no
        cadParam = cadParam & "Detalle= " & Me.chkDispo(0).Value & "|"
        numParam = numParam + 1
        
        
        nomRPT = DevuelveDesdeBD(conAri, "documrpt", "scryst", "codcryst", "68")
        
        
        Screen.MousePointer = vbDefault
    ElseIf OpcionListado = 41 Then
        
        If chkPedxClixSemEntrega(3).Value = 1 Then OpcionListado = 2052
    End If
    
    
    LlamarImprimir
    
    
    If OpcionListado = 46 Then OpcionListado = 44 'Lo cambio por que frmImprimir ha llamado a otro report. Pongo donde estaba
    If OpcionListado = 2051 Then OpcionListado = 44 'Lo cambio por que frmImprimir ha llamado a otro report. Pongo donde estaba
    If OpcionListado = 2052 Then OpcionListado = 41 'Lo cambio por que frmImprimir ha llamado a otro report. Pongo donde estaba
End Sub


Private Sub ObtenerValores()
Dim Sql As String


    'En importe1 tendre el del almacen PPAL

    Set miRsAux = New ADODB.Recordset
    Label4(54).Caption = "Stock "
    Label4(54).Refresh
    
    
    Sql = "Select salmac.* from tmpsliped,salmac where codusu = " & vUsu.Codigo & " AND tmpsliped.codartic=salmac.codartic"
    Sql = Sql & " AND tmpsliped.codalmac=salmac.codalmac "
    miRsAux.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        If miRsAux!CanStock <> 0 Then
            Sql = "UPDATE tmpsliped set stockalm = " & CStr(Val(CCur(miRsAux!CanStock)))
            Sql = Sql & " WHERE codusu = " & vUsu.Codigo & " AND codartic=" & DBSet(miRsAux!codArtic, "T")
            Sql = Sql & " AND codalmac =" & DBSet(miRsAux!codAlmac, "N")
            conn.Execute Sql
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    DoEvents
    
    'Metemos los del almacenppal si hay arituclos de almacen distinto del ppal(herbelca)
    Label4(54).Caption = "Almacen PPAL"
    Label4(54).Refresh
    Sql = "Select distinct salmac.codartic,canstock from  tmpsliped,salmac where tmpsliped.codArtic = salmac.codArtic AND "
    '               stk del ppal          articuo en almacen<>1    NO varios
    Sql = Sql & " salmac.codalmac=1 and   tmpsliped.codalmac>1 and codzona=0 And CanStock > 0"
    miRsAux.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
      
        Sql = "UPDATE tmpsliped set importel = " & CStr(Val(CCur(miRsAux!CanStock)))
        Sql = Sql & " WHERE codusu = " & vUsu.Codigo & " AND codartic=" & DBSet(miRsAux!codArtic, "T")
        Sql = Sql & " and codalmac>1"
        conn.Execute Sql

        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    'ped prov
    Label4(54).Caption = "Ped. prov"
    Label4(54).Refresh
    Sql = "select codartic,codalmac,sum(cantidad) cuant from slippr,scappr WHERE scappr.numpedpr = slippr.numpedpr "
    If Me.chkDispo(1).Value = 0 Then Sql = Sql & " AND scappr.obra=0"
    Sql = Sql & " AND codartic IN ( select distinct(codartic) from tmpsliped WHERE codusu = " & vUsu.Codigo & ") GROUP BY 1,2"
    miRsAux.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        If miRsAux!cuant <> 0 Then
            Sql = "UPDATE tmpsliped set cantpedprov = " & CStr(Val(CCur(miRsAux!cuant)))
            Sql = Sql & " WHERE codusu = " & vUsu.Codigo & " AND codartic=" & DBSet(miRsAux!codArtic, "T")
            Sql = Sql & " AND codalmac =" & DBSet(miRsAux!codAlmac, "N")
            conn.Execute Sql
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    Label4(54).Caption = "Articulos varios"
    Label4(54).Refresh
    Sql = "UPDATE tmpsliped SET stockalm =0,stocktot=0 WHERE codusu = " & vUsu.Codigo & " AND codzona=1" 'Los de varios
    conn.Execute Sql
    
    
    Label4(54).Caption = "Disponibilidad"
    Label4(54).Refresh
    
    
    Sql = "DELETE  FROM tmpsliped WHERE codusu = " & vUsu.Codigo & " AND  stockalm +cantpedprov-round(cantidad,0)>=0"
    conn.Execute Sql
   
    
    'Abril 2013.
    'Herbelca
    If vParamAplic.RecMercanciaSoloPpal Then
        Label4(54).Caption = "Ajuste ped proveed tipo H"
        Label4(54).Refresh
        
        Sql = "select codartic,sum(cantidad) cuant from slippr,scappr WHERE scappr.numpedpr = slippr.numpedpr  AND codalmac=1 "
        If Me.chkDispo(1).Value = 0 Then Sql = Sql & " AND scappr.obra=0"
        Sql = Sql & " AND codartic IN ( select distinct(codartic) from tmpsliped WHERE codusu = " & vUsu.Codigo & " and codalmac>1) GROUP BY 1"
        miRsAux.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            
            Sql = "UPDATE tmpsliped set cantpedprov = cantpedprov + " & CStr(Val(CCur(miRsAux!cuant)))
            Sql = Sql & " WHERE codusu = " & vUsu.Codigo & " AND codartic=" & DBSet(miRsAux!codArtic, "T")
            Sql = Sql & " AND codalmac > 1 "
            conn.Execute Sql
        
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        
        
    End If
    
    
    If vParamAplic.RecMercanciaSoloPpal Then  'herbleca quitamos aqui los del almacen 1
        Sql = "DELETE  FROM tmpsliped WHERE codusu = " & vUsu.Codigo & " AND  stockalm +cantpedprov-round(cantidad,0)>=0 and codalmac=1"
        conn.Execute Sql
    End If
    
End Sub


Private Sub cmdAceptarPreFac_Click()
'Prevision de Facturacion de Albaranes
Dim Campo As String, Cad As String
Dim B As Boolean
Dim Indice As Integer
   

    If OpcionListado = 50 Then
        Cad = ""
        For indCodigo = 10 To 16   'LOS SEIS PRIMEROS
            If Me.chkTpPago2(indCodigo).Value = 1 Then Cad = "1"
        Next indCodigo
        If Cad = "" Then
            MsgBox "Seleccione algun tipo de pago", vbExclamation
            Exit Sub
        End If
    End If
    


    InicializarVbles
    B = (OpcionListado = 50)
    
    'If (Not B) Or (B And codClien = "ALV") Then
        'Pasar nombre de la Empresa como parametro
        cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
        numParam = numParam + 1
    'End If
    
    
    'Desde/Hasta FECHA del ALBARAN
    '--------------------------------------------
    If Trim(txtCodigo(26).Text) <> "" Or Trim(txtCodigo(27).Text) <> "" Then

            Campo = "scaalb.fechaalb"
            cadSelect = CadenaDesdeHastaBD(txtCodigo(26).Text, txtCodigo(27).Text, Campo, "F")
            'Para Crystal Report
            Campo = "{scaalb.fechaalb}"
            Cad = "pDHFecha=""Fecha: "
            If Not PonerDesdeHasta(Campo, "F", 26, 27, Cad) Then Exit Sub

    End If

    'Cadena para seleccion CLIENTE
    '--------------------------------------------
    If txtCodigo(28).Text <> "" Or txtCodigo(29).Text <> "" Then
      
            Campo = "{scaalb.codclien}"
            Cad = "pDHCliente=""Cliente: "
   
        If Not PonerDesdeHasta(Campo, "N", 28, 29, Cad) Then Exit Sub
    End If
  
    If B Then 'opcionlistado=50
        'Cadena para seleccion FORMA PAGO
        '--------------------------------------------
        NumCod = ""   'reutilizo
        If txtCodigo(30).Text <> "" Or txtCodigo(31).Text <> "" Then
          
            Campo = "{scaalb.codforpa}"
            Cad = "Forma Pago: "
            If Not PonerDesdeHasta(Campo, "N", 30, 31, Cad) Then Exit Sub
            NumCod = Cad
        End If
        
        'Sep 2015
        If Me.txtCodigo(61).Text <> "" Then
            Campo = "{sclien.periodof}"
            Cad = ""
            If Not PonerDesdeHasta(Campo, "N", 61, 61, Cad) Then Exit Sub
            Cad = "    Periodo: " & Me.txtCodigo(61).Text & "."
            NumCod = Trim(NumCod & Cad)
        End If
        
        
        'JUNIO 2014
        If OpcionListado = 50 Then
            Cad = ""
            Campo = ""
            For NumRegElim = 10 To 16
                If Me.chkTpPago2(NumRegElim).Value = 1 Then
                    Cad = Cad & "1"
                    Campo = Campo & ", " & NumRegElim - 10
                End If
            Next
            
            If Len(Cad) = 7 Then
                'LOS HA COGIDO TODOS. NO lo incluyo en el desde hasta
            Else
                Set miRsAux = New ADODB.Recordset
                Campo = Mid(Campo, 2)
                Cad = "Select codforpa from sforpa where tipforpa in (" & Campo & ")"
                miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                Cad = ""
                While Not miRsAux.EOF
                    Cad = Cad & ", " & miRsAux!codforpa
                    miRsAux.MoveNext
                Wend
                miRsAux.Close
                Set miRsAux = Nothing
                
                If Cad = "" Then
                    'MAL. NInguna forpa de pago con ese tipo de pago. Fuerzo un -1
                    Cad = "-1"
                Else
                    Cad = Mid(Cad, 2)
                End If
                
                
                If Not AnyadirAFormula(cadSelect, "scaalb.codforpa IN (" & Cad & ")") Then Exit Sub
                If Not AnyadirAFormula(cadFormula, "{scaalb.codforpa} IN [" & Cad & "]") Then Exit Sub
                
                'Para el rpt saber que tipo de pago llevo
                Cad = ""
                For NumRegElim = 10 To 16
                    If Me.chkTpPago2(NumRegElim).Value = 1 Then Cad = Cad & "- " & Mid(chkTpPago2(NumRegElim).Caption, 1, 5) & "."
                Next
                Cad = "        Tipo pago: " & Mid(Cad, 2)
                NumCod = Trim(NumCod & Cad)
            End If
                
                
            
            If NumCod <> "" Then
                                
                cadParam = cadParam & "pDHForpa=""" & NumCod & "|"
                numParam = numParam + 1
            End If
        End If
        
        
        
        'seleccionar los Albaranes de Venta/Repar/Mantenim.
        'seleccionamos tipo de movimiento segun albaran de venta o Reparacion (ALV,ALR)
        '-- Aqui es donde se seleccionaban los albaranes a mostrar en el informe, ahora
        '   como se pueden seleccionar diferentes combinaciones se modifica la carga de la
        '   selección (se queda en rem la antigua línea) [SERVICIOS]
        
        If vParamAplic.Servicios And codClien <> "ALR" Then
            Indice = cmbTipAlbaran(0).ListIndex
            If Indice < 0 Then
                MsgBox "Debe seleccionar el tipo o los tipos de alabarán a procesar", vbExclamation
                Exit Sub
            Else
                Select Case Indice
                    Case 0 ' solo ventas
                        Cad = " {scaalb.codtipom}='ALV' "
                        If Not AnyadirAFormula(cadFormula, Cad) Then Exit Sub
                        If Not AnyadirAFormula(cadSelect, Cad) Then Exit Sub
                    Case 1 ' solo servicios
                        Cad = " {scaalb.codtipom}='ALS' "
                        If Not AnyadirAFormula(cadFormula, Cad) Then Exit Sub
                        If Not AnyadirAFormula(cadSelect, Cad) Then Exit Sub
                    Case 2 ' ventas y servicios
                        Cad = " ({scaalb.codtipom}='ALV'" & _
                                " OR {scaalb.codtipom}='ALS') "
                        If Not AnyadirAFormula(cadFormula, Cad) Then Exit Sub
                        If Not AnyadirAFormula(cadSelect, Cad) Then Exit Sub
                End Select
            End If
        Else
        
            
            If vParamAplic.NumeroInstalacion = vbEuler Then
                'Añadimos todos los tipos de albaran en la prefacturacion
                Cad = " {scaalb.codtipom} IN ['ALV','ALR','ALE','ALO','ALM','ALD','ALB','ALT'] "
            Else
                Cad = " {scaalb.codtipom}='" & codClien & "' "
            End If
            
           
            
            If Not AnyadirAFormula(cadFormula, Cad) Then Exit Sub
            If Not AnyadirAFormula(cadSelect, Cad) Then Exit Sub
        End If
        'Seleccionar los que esten marcados para facturar
        'Seleccionar solo aquellos que el campo scaalb.factursn=1
        If Me.chkSoloFacturar.Value = 1 Then
            Cad = " {scaalb.factursn}=1 "
            If Not AnyadirAFormula(cadFormula, Cad) Then Exit Sub
            If Not AnyadirAFormula(cadSelect, Cad) Then Exit Sub
        End If
    Else
        'Cadena para seleccion AGENTE
        '--------------------------------------------
        If txtCodigo(32).Text <> "" Or txtCodigo(33).Text <> "" Then
            Campo = "{scaalb.codagent}"
            Cad = "pDHAgente="""
            If Not PonerDesdeHasta(Campo, "N", 32, 33, Cad) Then Exit Sub
        End If
        
        'Seleccionar solo aquellos que tienen Nº de Pedido para comparar los Plazos de Entrega
        Campo = " NOT ISNULL({scaalb.numpedcl}) "
        If Not AnyadirAFormula(cadFormula, Campo) Then Exit Sub
        If Not AnyadirAFormula(cadSelect, Campo) Then Exit Sub
    End If
    
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    cadSelect = "scaalb.codclien = sclien.codclien AND " & cadSelect
    If Not HayRegParaInforme("scaalb,sclien", cadSelect) Then Exit Sub
    
    If OpcionListado = 51 Then
        Titulo = "Incumplimiento Plazos de Entrega"
        nomRPT = "rFacIncumPlazos.rpt"
        
    'ENERO 2009
    ElseIf OpcionListado = 50 Then
    'ElseIf OpcionListado = 50 And codClien = "ALV" Then
        If chkResumenForpa.Value = 1 Then
            'VAMOS A MOSTRAR LA HOJA RESUMEN DE FORMAS DE PAGO
            conn.Execute "DELETE FROM tmpinformes where codusu =" & vUsu.Codigo
        
            If Me.OptDetalle(0).Value Then
                Titulo = "SELECT scaalb.codforpa, sum(slialb.importel)," & vUsu.Codigo
                Titulo = Titulo & " FROM   ((scaalb scaalb INNER JOIN sclien sclien ON scaalb.codclien=sclien.codclien) INNER JOIN"
                Titulo = Titulo & " slialb slialb ON (scaalb.codtipom=slialb.codtipom) AND (scaalb.numalbar=slialb.numalbar))"
                Titulo = Titulo & " INNER JOIN starif starif ON sclien.codtarif=starif.codlista"
            
            ElseIf Me.OptDetalle(4).Value Then
                'IVA incluido
                '
                Titulo = "SELECT  scaalb.codforpa ,sum(slialb.importel * (100 + coalesce(porceiva,0)+coalesce(porcerec,0))/100 * (100  -(scaalb.dtognral  + scaalb.dtoppago)) /100),"
                Titulo = Titulo & vUsu.Codigo
                Titulo = Titulo & " FROM   slialb slialb INNER JOIN scaalb scaalb ON "
                Titulo = Titulo & " (slialb.codtipom=scaalb.codtipom) AND (slialb.numalbar=scaalb.numalbar)"
                Titulo = Titulo & " INNER JOIN sclien sclien ON scaalb.codclien=sclien.codclien"
                Titulo = Titulo & " INNER JOIN sartic sartic ON sartic.codartic=slialb.codartic"
                Titulo = Titulo & " LEFT JOIN " & IIf(vParamAplic.ContabilidadNueva, "ariconta", "conta") & vParamAplic.NumeroConta
                Titulo = Titulo & ".tiposiva tiposiva on sartic.codigiva =tiposiva.codigiva"
                
            Else
                Titulo = "SELECT  scaalb.codforpa ,sum(slialb.importel)," & vUsu.Codigo
                Titulo = Titulo & " FROM   slialb slialb INNER JOIN scaalb scaalb ON "
                Titulo = Titulo & " (slialb.codtipom=scaalb.codtipom) AND (slialb.numalbar=scaalb.numalbar)"
                Titulo = Titulo & " INNER JOIN sclien sclien ON scaalb.codclien=sclien.codclien"
            End If
    
            If cadSelect <> "" Then Titulo = Titulo & " WHERE " & cadSelect
                
            
            Titulo = Titulo & " GROUP BY codforpa"
            Titulo = "INSERT INTO tmpinformes (codigo1,importe1,codusu) " & Titulo
            conn.Execute Titulo
        End If
    
    
        Titulo = "Previsión Facturación Ventas"
        If codClien = "ALR" Then Titulo = Titulo & "(REP)"
        If codClien = "ALO" And vParamAplic.NumeroInstalacion = vbTaxco Then Titulo = "Prevision taller"
        '-- Si estan activos los servicios hay diferentes posibilidades y el título
        '   las refleja, la variabele 'indice' lleva la información del combo seleccionado y
        '   ha sido cargada un poco más arriba [SERVICIOS]
        
        If vParamAplic.Servicios And codClien <> "ALR" Then
            Select Case Indice
                Case 0
                    Titulo = "Previsión Facturación Ventas"
                Case 1
                    Titulo = "Previsión Facturación Servicios"
                Case 2
                    Titulo = "Previsión Facturación Ventas y Servicios"
            End Select
        End If
        If Me.OptDetalle(3).Value Then Titulo = Titulo & "(Fact.)"
        If Me.OptDetalle(4).Value Then Titulo = Titulo & "(con IVA)"
    
        conSubRPT = True
        If Me.OptDetalle(0).Value = True Then
        
            If InstalacionEsEulerTaxco Then
                'EULER llamara a la suya
                nomRPT = "eulFacPrevFactDetalle.rpt"
            Else
                nomRPT = "rFacPrevFactDetalle.rpt"
            End If
        ElseIf Me.OptDetalle(1).Value = True Then
            nomRPT = "rFacPrevFactResum.rpt"
        
        ElseIf Me.OptDetalle(4).Value = True Then
            nomRPT = "rFacPrevFactIVA.rpt"
        Else
            'Nuevo Marzo 2009
            'Como se facturara, es decir, el primer nivel de agrupacion es el tipofact de scaalb
            nomRPT = "rFacPrevFactDetalleCole.rpt"
        End If
        
        Cad = "pCodUsu=" & vUsu.Codigo & "|"
        cadParam = cadParam & Cad
        numParam = numParam + 1
        
        '-- Ahora el título depende de los tipos de albaranes seleccionados [SERVICIOS]
        Cad = "pTitulo=""" & Titulo & """|"
        cadParam = cadParam & Cad
        numParam = numParam + 1
        
        
        '--  Mostrara , o no, el subreport con el resumen por forma pago
        Cad = "pVerForpa=" & Abs(chkResumenForpa.Value) & "|"
        cadParam = cadParam & Cad
        numParam = numParam + 1
        
        '--- departamentos
        'cad = "TieneDpto=" & Abs(vParamAplic.Departamento) & "|"
        Cad = "TieneDpto=" & Abs(vParamAplic.HayDeparNuevo > 0) & "|"
        cadParam = cadParam & Cad
        numParam = numParam + 1
        
        
        
        
        On Error GoTo EPreFact
        Cad = "delete from tmpstockfec where codusu=" & vUsu.Codigo
        conn.Execute Cad
        
        
        
        
        'Insertar total bonificaciones por cliente,articulo en una TEMPORAL
        Cad = "SELECT " & vUsu.Codigo & " as codusu,  slialb.codartic,scaalb.codclien,sum(slialb.cantidad) as stock "
        Cad = Cad & "FROM (((scaalb INNER JOIN slialb ON scaalb.codtipom=slialb.codtipom and scaalb.numalbar=slialb.numalbar) "
        Cad = Cad & " INNER JOIN sbonif ON slialb.codartic=sbonif.codartic ) "
        Cad = Cad & " INNER JOIN sclien ON scaalb.codclien=sclien.codclien) "
        Cad = Cad & " INNER JOIN starif ON sclien.codtarif=starif.codlista "
        Cad = Cad & "WHERE " & cadSelect
        Cad = Cad & " AND starif.bonifica=1 "
        Cad = Cad & " GROUP BY scaalb.codclien,slialb.codartic"
        
        Cad = "INSERT INTO tmpstockfec (codusu,codartic,codalmac,stock) " & Cad
        conn.Execute Cad
        

        '18/05/2021: previsualizacion de resultados de la prefacturacion en listview
        CargarPrevisualizacion cadSelect
        If Not vContinua Then Exit Sub

        
        
        B = False 'PARA QUE NO ENTRE EN LO DE ABAJO y vaya a imprimir
    End If
    

    LlamarImprimir
    
    vContinua = False
    
    If OpcionListado = 50 Then
        Cad = "delete from tmpstockfec where codusu=" & vUsu.Codigo
        conn.Execute Cad
    End If
EPreFact:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Informe Prefacturación", Err.Description
    End If
End Sub

Private Sub CargarPrevisualizacion(vSelect As String)
Dim Sql As String
            
    If Me.OptDetalle(4).Value Then
        'IVA incluido
        Sql = "SELECT 'XXXX' ,sum(slialb.importel * (100 + coalesce(porceiva,0)+coalesce(porcerec,0))/100 * (100  -(scaalb.dtognral  + scaalb.dtoppago)) /100)"
        Sql = Sql & " FROM   slialb slialb INNER JOIN scaalb scaalb ON "
        Sql = Sql & " (slialb.codtipom=scaalb.codtipom) AND (slialb.numalbar=scaalb.numalbar)"
        Sql = Sql & " INNER JOIN sclien sclien ON scaalb.codclien=sclien.codclien"
        Sql = Sql & " INNER JOIN sartic sartic ON sartic.codartic=slialb.codartic"
        Sql = Sql & " LEFT JOIN " & IIf(vParamAplic.ContabilidadNueva, "ariconta", "conta") & vParamAplic.NumeroConta
        Sql = Sql & ".tiposiva tiposiva on sartic.codigiva =tiposiva.codigiva"
    Else
        Sql = "SELECT 'XXXX', sum(slialb.importel) "
        Sql = Sql & " FROM   ((scaalb scaalb INNER JOIN sclien sclien ON scaalb.codclien=sclien.codclien) INNER JOIN"
        Sql = Sql & " slialb slialb ON (scaalb.codtipom=slialb.codtipom) AND (scaalb.numalbar=slialb.numalbar))"
    End If
    
    
    Set frmMens = New frmMensajes
    frmMens.cadWhere = Sql
    If vSelect <> "" Then
        frmMens.cadWHERE2 = " where " & vSelect
    Else
        frmMens.cadWHERE2 = ""
    End If
    frmMens.OpcionMensaje = 29
    frmMens.Show vbModal
    
    Set frmMens = Nothing

End Sub



Private Sub cmdAceptarPreFacMan_Click()
'74: PreFacturar Mantenimientos
'75: Facturar Mantenimientos
Dim Campo As String, Cad As String
Dim B As Boolean
Dim PreguntaHecha As Boolean

     InicializarVbles
     B = (OpcionListado = 74) 'Prefacturar (mostrar listado)
     If Not B Then chkSituFacMant.Value = 0 'por si acaso
     
     'Introducir el mes que se va a facturar
     If txtCodigo(46).Text = "" Then
        MsgBox "Debe introducir el mes a Facturar.", vbInformation
        Exit Sub
    End If
     
     
    'Febrero 2011
    'Ha puesto departamento.
    If Me.chkSituFacMant.Value = 0 Then
        If Me.txtCodigo(55).Text <> "" Or txtCodigo(56).Text <> "" Then
            'SI no este el mismo cliente... NO seguimos
            If txtCodigo(48).Text <> txtCodigo(49).Text Or txtCodigo(49).Text = "" Then
                MsgBox "Seleccione un cliente para poder indicar departamento", vbExclamation
                PonerFoco txtCodigo(48)
                Exit Sub
            End If
        End If
    End If
     
    If Not B Then 'Vamos a facturar
        'si vamos a facturar comprobar que la fecha de factura tiene valor
        Cad = ""
        If txtCodigo(44).Text = "" Then Cad = " - El campo Fecha Factura" & vbCrLf
            
        
        'si vamos a facturar debe haber una cta prev. de cobro
        If txtCodigo(52).Text = "" Then Cad = Cad & " - El campo Cta. Prev. de cobro " & vbCrLf
            
        
        'si vamos a facturar comprobar que el cod. de operador tiene valor
        If txtCodigo(47).Text = "" Then Cad = Cad & " - El campo operador " & vbCrLf
            
        
        
        'Si tienen analitcia y es por proyecto, la esta pidiendo en el fram2
        'Con lo cual es un campo obligatorio
        If Me.Frame2(1).visible Then
            If txtCodigo(54).Text = "" Then Cad = Cad & " - Centro de coste"
        End If
        
        If Cad <> "" Then
            MsgBox "Debe comprobar: " & vbCrLf & vbCrLf & Cad, vbExclamation
            Exit Sub
        End If
    End If
     
     
    'Pasar nombre de la Empresa como parametro
    cadParam = "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1

     
    'Cadena para seleccion TIPO CONTRATO
    '--------------------------------------------
    If txtCodigo(45).Text <> "" Then
        Campo = "{scaman.codtipco}"
'        If Not PonerDesdeHasta(campo, "N", 48, 49, cad) Then Exit Sub
        Cad = Campo & "= '" & txtCodigo(45).Text & "'"
        If Not AnyadirAFormula(cadFormula, Cad) Then Exit Sub
        
        'Parametro
        Cad = "pTipCo=""Tipo Contrato: "
        cadParam = cadParam & Cad & txtCodigo(45).Text & " - " & txtNombre(45).Text & """|"
        numParam = numParam + 1
    End If
     
     
    'Cadena para seleccion CLIENTE
    '--------------------------------------------
    CadenaDesdeOtroForm = ""
    If Me.chkSituFacMant.Value = 0 Then
        If txtCodigo(48).Text <> "" Or txtCodigo(49).Text <> "" Then
            Campo = "{scaman.codclien}"
            Cad = "Cliente: "
            If Not PonerDesdeHasta(Campo, "N", 48, 49, Cad) Then Exit Sub
            CadenaDesdeOtroForm = CadenaDesdeOtroForm & Cad
        End If
        If txtCodigo(55).Text <> "" Or txtCodigo(56).Text <> "" Then
            Campo = "{scaman.coddirec}"
            Cad = "Dpto: "
            If Not PonerDesdeHasta(Campo, "N", 55, 56, Cad) Then Exit Sub
            CadenaDesdeOtroForm = Trim(CadenaDesdeOtroForm & Space(10) & Cad)
        End If
    End If
    
    CadenaDesdeOtroForm = "pDHCliente=""" & CadenaDesdeOtroForm & """|"
    cadParam = cadParam & CadenaDesdeOtroForm
    numParam = numParam + 1
    CadenaDesdeOtroForm = ""
    
    
    'Cadena para seleccion FORMA PAGO
    '--------------------------------------------
    If Me.chkSituFacMant.Value = 0 Then
        If txtCodigo(50).Text <> "" Or txtCodigo(51).Text <> "" Then
            Campo = "{scaman.codforpa}"
            Cad = "pDHForpa=""Forma Pago: "
            If Not PonerDesdeHasta(Campo, "N", 50, 51, Cad) Then Exit Sub
        End If
    End If
            
            
    'MES A FACTURAR
    'Seleccionar solo aquellos que el campo del mes seleccionado sea no nulo
    '------------------------------------------------------------------------
    Cad = Format(txtCodigo(46).Text, "00")
    Campo = "mes" & Cad & "act"
    Cad = "(NOT ISNULL({scaman." & Campo & "})) and ({scaman." & Campo & "}<>0)"
    If Not AnyadirAFormula(cadFormula, Cad) Then Exit Sub
    'Parametro
    Cad = "pCampoMes={scaman." & Campo & "}" & "|"
    cadParam = cadParam & Cad
    numParam = numParam + 1
    Cad = "pMes=""Mes a Facturar: " & UCase(txtNombre(46).Text) & """|"
    cadParam = cadParam & Cad
    numParam = numParam + 1
    
    
    
    If B Then
        'Prevision
        Cad = "0"
        
        If txtCodigo(48).Text = txtCodigo(49).Text And txtCodigo(49).Text <> "" Then
            'Mismo cliente
            If Me.txtCodigo(55).Text <> "" Or txtCodigo(56).Text <> "" Then Cad = "1"
        End If
        Cad = "DetallaDir=" & Cad & "|"
        cadParam = cadParam & Cad
        numParam = numParam + 1
    End If
    
    
    'Comprobamos si ha seleccionado
    If Me.chkSituFacMant.Value = 1 Then
        Cad = "NumeroMes=" & txtCodigo(46).Text & "|"
        cadParam = cadParam & Cad
        numParam = numParam + 1
        
        Campo = "Facturado y pendiente"
        
        If Me.cboSituMan.ListIndex > 0 Then
            'Si pide facturado o pendiente
            
            If Me.cboSituMan.ListIndex = 1 Then
                'FACTURADOS.  Luego ult mesfac tiene que ser mayor o igual
                Cad = " >= "
                Campo = "Facturado"
            Else
                Cad = "<"
                Campo = "Pendiente"
            End If
            Cad = "({scaman.ulmesfac}" & Cad & txtCodigo(46).Text & ")"
            If Not AnyadirAFormula(cadFormula, Cad) Then Exit Sub
            
        End If
        'Reutilizare:
        Cad = "pDHForpa=""Situacion: " & Campo & """|"
        cadParam = cadParam & Cad
        numParam = numParam + 1
        
    End If
    
    
    
        
    cadSelect = cadFormula
    If Not HayRegParaInforme("scaman", cadSelect) Then Exit Sub
    
    
    'Aqui deberiamos comporbar si el periodo indicado YA esta facturado o no
    PreguntaHecha = False
    If Not B Then
        'FACTURACION
        
        ResultadoFechaContaOK = EsFechaOKConta(CDate(txtCodigo(44).Text), True)
        If ResultadoFechaContaOK > 0 Then
            If MensajeFechaOkConta <> 4 Then
                Cad = MensajeFechaOkConta & ". ¿Continuar?"
                If MsgBox(Cad, vbQuestion + vbYesNo) = vbNo Then Exit Sub
            End If
            PreguntaHecha = True
        End If
    End If
    
    
    
    'Comprobamos si hay matenimientos ya facturados
    If Me.chkSituFacMant.Value = 0 Then
        'facturacion o listado normal
    
        Cad = "Select * from " & cadFormula
        If miRsAux Is Nothing Then Set miRsAux = New ADODB.Recordset
        Cad = "SELECT scaman.codclien,nomclien,coddirec "
        Cad = Cad & " FROM scaman INNER JOIN sclien ON scaman.codclien=sclien.codclien "
        Cad = Cad & " WHERE " & cadSelect
        'Que el ultimo mes de facturado sea mayor o igual  al que voy a facturar
        Cad = Cad & " AND ulmesfac >= " & txtCodigo(46).Text
        miRsAux.Open Cad, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        Cad = ""
        While Not miRsAux.EOF
            Titulo = ""
            If Not IsNull(miRsAux!CodDirec) Then
                Titulo = DevuelveDesdeBD(conAri, "nomdirec", "sdirec", "codclien=" & miRsAux!codClien & " AND coddirec ", CStr(miRsAux!CodDirec))
                Titulo = "( " & miRsAux!CodDirec & " " & Titulo & ")"
                
            End If
            Cad = Cad & "    .- " & miRsAux!NomClien & Titulo & vbCrLf
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        Set miRsAux = Nothing
        If Cad <> "" Then
            Cad = "Los siguientes mantenimientos ya estan facturados: " & vbCrLf & Cad & vbCrLf & vbCrLf
            Cad = Cad & "¿Continuar?"
            If MsgBox(Cad, vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
            PreguntaHecha = True
        End If
    End If
    
    If B Then 'OpcionListado = 74 'NO Imprime, mostrar resultado en pantalla
        
        If Me.chkSituFacMant.Value = 0 Then
            Titulo = "Prefacturación Mantenimientos"
            nomRPT = "rManPrefacturar.rpt"
        Else
            Titulo = "Situación facturacion mantenimientos"
            nomRPT = "rManPrefacturarSitu.rpt"
        End If
        LlamarImprimir
    Else
    
        
        If Not PreguntaHecha Then
            Cad = "¿Seguro que desea seguir con el proceso?"
            If MsgBox(Cad, vbQuestion + vbYesNo) = vbNo Then Exit Sub
        End If
            
        
        '------------------------------------------------------------------------------
        '  LOG de acciones
        Set LOG = New cLOG
        Cad = "MANTENIMIENTOS: " & vbCrLf & cadSelect
        LOG.Insertar 2, vUsu, Cad
        Set LOG = Nothing
        '-----------------------------------------------------------------------------
        
        
        
        'Generar facturas de los mantenimientos seleccionados para facturar
        'cada mantenimiento genera una factura
        Cad = "SELECT scaman.codclien,scaman.coddirec,sdirec.nomdirec,nummante,fechaini,codtipco,codforpa,tipopago," & Campo & " as importe "
        'David
        'Necesito el campo concefaccl y el tipopago(mensual...)
        Cad = Cad & ", concefac"
        Cad = Cad & " FROM scaman LEFT OUTER JOIN sdirec ON scaman.codclien=sdirec.codclien AND scaman.coddirec=sdirec.coddirec "
        Cad = Cad & " WHERE " & cadSelect
        
        lblFactMant.Caption = "Obteniendo datos"
        lblFactMant.Refresh
        'Pasamos la SQL que selecciona los mantenimientos a facturar y
        'le pasamos la fecha y operador de la factura.
        If TraspasoMtosAFacturas(Cad, cadSelect, txtCodigo(44).Text, txtCodigo(47).Text, txtCodigo(52).Text, txtCodigo(46).Text, lblFactMant, txtCodigo(54).Text) Then 'Fecha de la factura, Operador
            Unload Me
        End If
        lblFactMant.Caption = ""
    End If
End Sub



Private Sub cmdCancel_Click(Index As Integer)


     If OpcionListado = 222 And vParamAplic.NumeroInstalacion = vbTaxco And (codClien = "ALO" Or codClien = "ALE" Or codClien = "ALV") Then
        'SI es taxco, ha cerrado un OT de CONTADO.
        'O factura o no sale de ahi
        If davidNumalbar > 0 Then
        
            If davidNumalbar > 2 Then
                cadFormula = "Si continua se le quitará  la marca de cerrado. ¿Continuar?"
                If MsgBox(cadFormula, vbQuestion + vbYesNoCancel) = vbYes Then
                    cadFormula = "UPDATE scaalb set factursn=0"
                    cadFormula = cadFormula & " WHERE codtipom='" & codClien & "' AND numalbar=" & NumCod
                    If ejecutar(cadFormula, False) Then davidNumalbar = 0
                    
                End If
                cadFormula = ""
            Else
                MsgBox "Ha cerrado  " & IIf(codClien = "ALO", "una orden de taller", "un albaran") & " NO de crédito. Debes facturar", vbCritical
            
                davidNumalbar = davidNumalbar + 3
            End If
            
        End If
    End If


    Unload Me
     
End Sub


Private Sub cmdPedidosAnulados_Click()
Dim Campo As String
Dim Cad As String

    'Marzo 2022
    InicializarVbles
    
    'Pasar nombre de la Empresa como parametro
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    cadSelect = " TRUE "
    '-------------------------
    If txtCodigo(70).Text <> "" Or txtCodigo(71).Text <> "" Then
        Campo = "{schped.fecpedcl}"
        Cad = "F. Pedido: "
        If Not PonerDesdeHasta(Campo, "F", 70, 71, Cad) Then Exit Sub
        
        'No se porque , las fechas NO las añade al cadselect  en este FORM !!!!!
        If txtCodigo(70).Text <> "" Then cadSelect = cadSelect & " AND " & Campo & " >= " & DBSet(txtCodigo(70).Text, "F")
        If txtCodigo(71).Text <> "" Then cadSelect = cadSelect & " AND " & Campo & " <= " & DBSet(txtCodigo(71).Text, "F")
        Cad = Cad
        
    End If
    NumCod = Cad
    Cad = ""
    'Inicidencia
    If txtCodigo(76).Text <> "" Or txtCodigo(77).Text <> "" Then
        Campo = "{schped.codincid}"
        Cad = "Incid:"
        If Not PonerDesdeHasta(Campo, "T", 76, 77, Cad) Then Exit Sub
    End If
    NumCod = Trim(NumCod & "     " & Cad)
    cadParam = cadParam & "pDH1=""" & NumCod & """|"
    numParam = numParam + 1
    
    
    
    If txtCodigo(72).Text <> "" Or txtCodigo(73).Text <> "" Then
        Campo = "{sartic.codprove}"
        Cad = "pDHProv=""Proveedor: "
        If Not PonerDesdeHasta(Campo, "N", 72, 73, Cad) Then Exit Sub
    End If
    
    If txtCodigo(74).Text <> "" Or txtCodigo(75).Text <> "" Then
        Campo = "{sartic.codfamia}"
        Cad = "pDHFam=""Familia: "
        If Not PonerDesdeHasta(Campo, "N", 74, 75, Cad) Then Exit Sub
    End If
    
    Campo = " schped INNER JOIN slhped on schped.numpedcl=slhped.numpedcl AND schped.fecpedcl=slhped.fecpedcl"
    Campo = Campo & " INNER JOIN sartic ON slhped.codartic=sartic.codartic "
    cadSelect = Replace(cadSelect, "{", "")
    cadSelect = Replace(cadSelect, "}", "")
        
    If Not HayRegParaInforme(Campo, cadSelect) Then Exit Sub
    
    
    'Ok . Al rpt
    cadParam = cadParam & "detalla=" & chkVarios(0).Value & "|"
    numParam = numParam + 1
    
    
    Titulo = "Pedidos anulados"
    nomRPT = "rPedidosAnulados.rpt"
    conSubRPT = False
    
    LlamarImprimir
        
     

End Sub

Private Sub cmdSelFraRect_Click()
        
        Set frmLd = New frmListadoOfer
        frmLd.OpcionListado = 225
        frmLd.codClien = "N"
        frmLd.Show vbModal
        Set frmLd = Nothing
        
        
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Activate()
Dim Banco As Integer

    If PrimeraVez Then
        PrimeraVez = False
        Select Case OpcionListado
            Case 41, 42, 44, 49, 227, 228 '41: Informe de Pedidos por Articulo
                        '42: Informe de Disponibilidad de Stocks
                        '44: Informe de Pedidos por Cliente
                        '49: Informe de Albaranes por Articulo
                        '227: Inf. estadistica Ventas por cliente
                PonerFoco txtCodigo(11)
            Case 43, 1000, 1010
                    '43: Generar Albaran desde Pedido (NO IMPRIME LISTADO)
                    '1000: Pedido a factura:  Piede ademas de los datos del albaran, la cta prevista
                    
                If OpcionListado = 43 And vParamAplic.NumeroInstalacion = vbFenollar Then
                     
                    cadFormula = PonerTrabajadorConectado(cadParam)
                    If cadFormula <> "" Then
'                        me.Label1
                        Me.txtCodigo(17).Text = cadFormula
                        Me.txtCodigo(18).Text = cadFormula
                        Me.txtNombre(17).Text = cadParam
                        Me.txtNombre(18).Text = cadParam
                        cadParam = "sclien.codenvio=senvio.codenvio AND codclien"
                        cadParam = DevuelveDesdeBD(conAri, "concat(sclien.codenvio,'|',nomenvio,'|')", "sclien,senvio", cadParam, CStr(davidNumalbar))
                        If cadParam <> "" Then
                            Me.txtCodigo(19).Text = RecuperaValor(cadParam, 1)
                            Me.txtNombre(19).Text = RecuperaValor(cadParam, 2)
                            
                        End If
                        
                    End If
                    cadParam = ""
                    cadFormula = ""
                End If
                    
                If FramePartes.visible = True Then
                    'Es facturar parte
                            

                    'CadenaDesdeOtroForm = 'trab,fecha,interna
                    txtCodigo_LostFocus 17
                    txtNombre(18).Text = txtNombre(17).Text
                    txtCodigo_LostFocus 19
                            
                    PonerFoco txtCodigo(25)
                Else
                    PonerFoco txtCodigo(17)
                End If
                
                
                
                
                
            Case 50, 51 '50: Prevision de Facturacion Albaranes (NO IMPRIME LISTADO)
                        '51: Inf. Incumplimiento Plazos de Entrega
                PonerFoco txtCodigo(26)
            Case 52, 222
                '52: Facturacion de Albaranes
                '222: Factura de Mostrador
                indCodigo = 0
                Banco = 0
                
                
                If vParamAplic.NumeroInstalacion = vbTaxco Then
                    'Esto lo monto porque si ha llegado a importar el fichero de ALVIC
                    'y da un fallo facturando, entonces me mete en un lio. Tengo que borrar facturas, borrar albarnaes, volver a procesar...
                    If TaxcoFacurarUnAlbaranALVIC_Facturado <> "" Then
                        'Fecha
                        txtCodigo(34).Text = RecuperaValor(TaxcoFacurarUnAlbaranALVIC_Facturado, 1)
                        Banco = 1
                        TaxcoFacurarUnAlbaranALVIC_Facturado = RecuperaValor(TaxcoFacurarUnAlbaranALVIC_Facturado, 2)
                        
                        
                    End If
                    
                End If
                    
                
                

                If vParamAplic.CodigoUnicoBancoPropio > 0 Then
                    Banco = vParamAplic.CodigoUnicoBancoPropio
                Else
                    'Fras mostrador en herbelca
                    If vParamAplic.NumeroInstalacion = 2 And codClien = "ALM" Then Banco = 7
                End If
                If Banco > 0 Then
                    txtCodigo(0).Text = Banco
                    txtNombre(0).Text = PonerNombreDeCod(txtCodigo(0), conAri, "sbanpr", "nombanpr", "codbanpr", "Bancos Propios")
                    If vParamAplic.EntradaRapidaFacturasMostrador And codClien = "ALM" Then indCodigo = 1
                End If
                If indCodigo = 0 Then
                    PonerFoco txtCodigo(34)
                Else
                    PonerFocoBtn cmdAceptarFac
                End If
                    
                    
            Case 74 '74: Previsión facturación Mantenimientos
                PonerFoco txtCodigo(45)
            Case 75 '75: Facturacion de Mantenimientos
                PonerFoco txtCodigo(44)
            Case 229 '229: Inf. estadistica ventas por meses
                PonerFoco txtCodigo(53)
        End Select
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub CargaIconosAyuda()
Dim Ima As Image
    On Error Resume Next 'mejor que no diera errores, pero bien, tampoco vamos a enfadarnos
    For Each Ima In Me.imgayuda
        Ima.Picture = frmPpal.imgListComun.ListImages(46).Picture
    Next
    Err.Clear
End Sub

Private Sub CargaIconosAyuda2()
Dim i As Integer

    On Error Resume Next
    
    For i = 12 To 31
        imgBuscarOfer(i).Picture = imgBuscar(0).Picture
    Next i
    imgBuscarOfer(38).Picture = imgBuscar(0).Picture
    imgBuscarOfer(39).Picture = imgBuscar(0).Picture
    imgBuscarOfer(46).Picture = imgBuscar(0).Picture
    imgBuscarOfer(54).Picture = imgBuscar(0).Picture
    
    
    If OpcionListado = 80 Then
        For i = 47 To 52
            imgBuscarOfer(i).Picture = imgBuscar(0).Picture
        Next i
    End If
    
    Err.Clear
End Sub


Private Sub Form_Load()
Dim H As Integer, W As Integer
Dim indFrame As Single
Dim Aux As String


    'Icono del formulario
    Me.Icon = frmPpal.Icon

    PrimeraVez = True
    limpiar Me
    CargaIconosAyuda
    CargaIconosAyuda2
    
    'Ocultar todos los Frames de Formulario
    Me.FramePedxArtic.visible = False
    Me.FrameGenAlbaran.visible = False
    Me.FramePreFacturar.visible = False
    Me.FrameFacturar.visible = False
    Me.FramePreFacMante.visible = False
    Me.FrameEstVentas.visible = False
    FramepedidoAnulado.visible = False
    CommitConexion
    
    NomTabla = "scaped"
    NomTablaLin = "sliped"
        
    Select Case OpcionListado
        'LISTADOS DE FACTURACION
        '-----------------------
            
        Case 41, 42, 44, 49, 227, 228 '41: Informe de Pedidos por Articulo
                    '42: Informe de Disponibilidad de Stocks
                    '44: Informe de Pedidos por Cliente
                    '49: Informe de Albaranes por Articulo
                    '227: Inf. estadistica Ventas por cliente
                    '228: Inf. estadistica Ventas por trabjador
            PonerFramePedxArticVisible True, H, W
            indFrame = 2 'solo para el boton cancelar
            '-- Si está activada la opción de servicios, muestra los controles que permiten
            '   el tipo o tipos de albaranes a mostrar en el informe, en caso contrario los
            '   oculta para no liar [SERVICIOS]
            lblTipAlbaran(1).visible = False
            cmbTipAlbaran(1).visible = False
            Label4(54).Caption = ""
            
            
            chkPedxClixSemEntrega(2).visible = OpcionListado = 44
            chkDispo(0).visible = OpcionListado = 42 'solo disponibilidad
            chkDispo(1).visible = OpcionListado = 42 'solo disponibilidad
            chkPedxClixSemEntrega(3).visible = OpcionListado = 41  'Agrupado tipo fenollar
            chkPedxClixSemEntrega(3).Top = 6000
            If vParamAplic.Servicios And OpcionListado = 49 Then
                lblTipAlbaran(1).visible = True
                cmbTipAlbaran(1).visible = True
            End If
            
            If OpcionListado = 49 Then 'Albaranes de Venta
                NomTabla = "scaalb"
                NomTablaLin = "slialb"
            ElseIf OpcionListado = 227 Or OpcionListado = 228 Then
                NomTabla = "scafac"
                NomTablaLin = "slifac"
                
                'poner por defecto las fechas del ejercicio contable
                
                If OpcionListado = 228 Then
                    Me.txtCodigo(11).Text = Format(Now, "dd/mm/yyyy")
                    Me.txtCodigo(12).Text = Format(Now, "dd/mm/yyyy")
                Else
                  Me.txtCodigo(11).Text = vEmpresa.FechaIni
                 Me.txtCodigo(12).Text = vEmpresa.FechaFin
                End If
                
                Label4(54).Top = 5880
                If OpcionListado = 227 Then
                    FrameTiposFactura.BorderStyle = 0
                    FrameTiposFactura.Left = Frame10.Left + 30
                    FrameTiposFactura.Top = Frame10.Top + Frame10.Height - 90
                    FrameTiposFactura.visible = True
                    
                    Label4(54).Top = cmdAceptarPedxArtic.Top + 120
                    
                    CargaListTipoFacturas
                End If
                
            End If
            
            If OpcionListado = 44 Then
                CargaComboTipoCliente
                'Para TEINSA no marco las obsrvaciones
                chkPedxClixSemEntrega(1).Value = 1
                If vParamAplic.NumeroInstalacion = 3 Then chkPedxClixSemEntrega(1).Value = 0
                
                'Para HERBELCA desmarco la semana de entrega
                chkPedxClixSemEntrega(0).Value = 1
                If vParamAplic.NumeroInstalacion = 2 Then chkPedxClixSemEntrega(0).Value = 0
                
                'Ppara fontenas marcamos con IVA
                chkPedxClixSemEntrega(2).Value = IIf(vParamAplic.NumeroInstalacion = 5, 1, 0)
            End If
            
        Case 43, 1000, 1043
                '43: Generar Albaran desde Pedido (NO IMPRIME LISTADO)
                '1000:  Pedido a factura: pide la cta prevista de cobro
                
                'En codclien llevo: tienecoddiren & "|" & zonacliente & "|" 'si es visible framezona
                If codClien <> "" Then
                    W = RecuperaValor(codClien, 3)
                Else
                    W = 0
                End If
            
                'Si tiene coddiren tendre que pedir la zona de reparto
                FramePartes.visible = False
                FrameZona.visible = False
                FrameGenAlbEuler.visible = False
                FrameBultosHerbelca.visible = False
                FrameCanjePuntos.visible = False
                
                lblDestinoB.visible = vParamAplic.NumeroInstalacion = vbFenollar
                cboDestinoB.visible = vParamAplic.NumeroInstalacion = vbFenollar
                If vParamAplic.NumeroInstalacion = vbFenollar Then cboDestinoB.ListIndex = 0
                H = FrameGenAlbaran.Height
                If W = 1 Then
                    txtCodigo(22).Text = RecuperaValor(codClien, 1)
                    txtNombre(22).Text = RecuperaValor(codClien, 2)
                    FrameZona.visible = True
                End If
                
                If OpcionListado = 43 Or OpcionListado = 1000 Then
                    If vParamAplic.PtosAsignar > 0 Then
                        
                        Aux = RecuperaValor(codClien, 4)
                        If Aux <> "" Then
                        
                            txtCodigo(4).Tag = Aux
                            Aux = RecuperaValor(codClien, 5)
                            
                            txtCodigo(63).Tag = Aux
                            If CCur(txtCodigo(4).Tag) > CCur(txtCodigo(63).Tag) Then
                                txtCodigo(68).Tag = txtCodigo(63).Tag
                            Else
                                txtCodigo(68).Tag = txtCodigo(4).Tag
                            End If
                            
                            
                            'txtCodigo(68).Text = Format(txtCodigo(68).Tag, FormatoImporte)
                            txtCodigo(68).Text = Format(0, FormatoImporte)
                            txtCodigo(63).Text = Format(txtCodigo(63).Tag, FormatoImporte)
                            txtCodigo(4).Text = Format(txtCodigo(4).Tag, FormatoImporte)
                            
                            H = H + 600 'FrameCanjePuntos.Height
                            FrameCanjePuntos.Left = 720
                            FrameCanjePuntos.Top = H - FrameCanjePuntos.Height - 240
                            FrameCanjePuntos.visible = True
                            
                            
                        End If
                        
                    End If
                End If
                                
                Label5.Caption = "" 'Zona d
                'la zona del cliente
                ' RecuperaValor(codClien, 1)
                
                W = 6515
                H = H + 340
                
                cmdAceptarGenAlb.Top = H - 465
                cmdCancel(3).Top = cmdAceptarGenAlb.Top
                Label5.Top = cmdAceptarGenAlb.Top
                
                PonerFrameVisible Me.FrameGenAlbaran, True, H, W
                txtCodigo(25).Text = Format(Now, "dd/mm/yyyy")
                indFrame = 3
                chkImpAlbaran.Caption = "Impimir "
                If OpcionListado = 1000 Then
                    Label4(32).Caption = "Fec. FACTURA"
                    Label3.Caption = "FACTURAR pedido"
                    chkImpAlbaran.Caption = chkImpAlbaran.Caption & "FACTURA"
                    If vParamAplic.NumeroInstalacion = vbHerbelca Then Me.chkFraMostrador.Value = 1
                Else
                    
                    Label4(32).Caption = "Fecha albarán"
                    chkImpAlbaran.Caption = chkImpAlbaran.Caption & "albaran"
                    If NumCod = "REP" Then
                        Label3.Caption = "Pasar Reparación a Albaran"
                    Else
                        Label3.Caption = "Pasar Pedido a Albaran"
                        
                        If InstalacionEsEulerTaxco Then
                            FrameGenAlbEuler.BorderStyle = 0
                            FrameGenAlbEuler.visible = True
                            Me.cboTipoAlbaranEuler.ListIndex = 0
                        Else
                            'Menos para partes de trabajo
                            If Me.OpcionListado <> 1043 Then Me.FrameBultosHerbelca.visible = True
                        End If
                        
                        If Me.OpcionListado = 1043 Then
                            Label3.Caption = "Pasar parte trabajo a Albaran"
                            
                            CargarComboFacturacion
                            Me.cboFacturacion.ListIndex = 1
                            
                            Me.chkImpEtiq.visible = False
                            Me.chkImpHojaExped.visible = False
                            FramePartes.visible = True
                            
                            
      
                            'CadenaDesdeOtroForm = 'trab,fecha,interna
                            txtCodigo(17).Text = RecuperaValor(CadenaDesdeOtroForm, 1)
                            txtCodigo(18).Text = RecuperaValor(CadenaDesdeOtroForm, 1)
                            txtCodigo(19).Text = vParamAplic.PorDefecto_Envio
                            txtCodigo(25).Text = RecuperaValor(CadenaDesdeOtroForm, 2)
                            txtCodigo(59).Text = "0" 'cantidad litros
                            txtCodigo(60).Text = "1,00" 'cantidad otros
                            
                            Titulo = RecuperaValor(CadenaDesdeOtroForm, 3)
                            If Titulo = "1" Then
                                Me.chkFrInterna.Value = 1
                            Else
                                Me.chkFrInterna.Value = 0
                            End If
                            
                            'Mayo 2013. Tipo facturacion
                            Titulo = RecuperaValor(CadenaDesdeOtroForm, 4)
                            Me.cboFacturacion.ListIndex = Val(Titulo)
                            
                            CadenaDesdeOtroForm = ""
                            Titulo = ""
                            OpcionListado = 43
                        End If
                        
                            
                    End If
                End If
                FramepedidoFactura.visible = (OpcionListado = 1000)
                
                
                
                
                '- Ver si hay articulo portes para imprimir hoja Expedicion
                If vParamAplic.TipoPortes = 1 Then
                    Me.chkImpHojaExped.Value = 1
                Else
                    Me.chkImpHojaExped.Value = 0
                End If
            
                'Enero 2011
                'Si es PEW o devolucion de pedidos, ni mostramos el chek de hoja de expedidcion ni imprfra ni ostias ne vinagre
                
                FramRectARM.visible = False
                txtFraRectifica.Text = ""
                chkFraMostrador.visible = False
                If NumCod = "PEW" Then
                    If OpcionListado = 1000 Then
                        FramRectARM.visible = True
                        'OK no mostraremos todas esas cosas
                        chkImpHojaExped.visible = False
                        'chkImpAlbaran.visible = False
                        chkImpEtiq.visible = False
                    End If
                Else
                    'Si opcion=1000(pasando a factura)
                    '  y tiene numero serie distinto para las facturas de mostrador, entonces
                    '  debe seleccionar si el pedido lo pasa a Mostraor(contado)
                    'vParamAplic.FrasMostradorSerieDistinta    . Los pedidos de B van a B
                    If vParamAplic.FrasMostradorSerieDistinta And NumCod <> "PEZ" Then
                        chkFraMostrador.visible = True
                        If Me.OpcionListado = 1000 Then
                            chkFraMostrador.Caption = "Factura"
                        Else
                            chkFraMostrador.Caption = "Albaran"
                        End If
                        chkFraMostrador.Caption = chkFraMostrador.Caption & " de mostrador"
                    End If
                End If
            
        Case 50, 51 '50: Prevision Facturacion de Albaranes (NO IMPRIME LISTADO)
                    '51: Inf. Incumplimiento Plazos de Entrega
            PonerFramePreFacVisible True, H, W
            indFrame = 5 'solo para el boton cancelar
            '-- Si está activada la opción de servicios, muestra los controles que permiten
            '   el tipo o tipos de albaranes a mostrar en el informe, en caso contrario los
            '   oculta para no liar [SERVICIOS]
            lblTipAlbaran(0).visible = False
            cmbTipAlbaran(0).visible = False
            If vParamAplic.Servicios Then
                lblTipAlbaran(0).visible = codClien <> "ALR"
                cmbTipAlbaran(0).visible = codClien <> "ALR"
                
                lblTipAlbaran(0).Top = IIf(OpcionListado = 50, 5880, 3880)
                cmbTipAlbaran(0).Top = IIf(OpcionListado = 50, 6120, 4120)
                lblTipAlbaran(0).Left = IIf(OpcionListado = 50, 2400, 360)
                cmbTipAlbaran(0).Left = IIf(OpcionListado = 50, 2400, 360)

                
            End If
            chkResumenForpa.visible = OpcionListado = 50
            
        Case 52, 222
                    '52: Facturacion de Albaranes
                    '222: Factura de Mostrador
                    
            PonerFrameFacVisible True, H, W
            txtCodigo(34).Text = Format(Now, "dd/mm/yyyy")
            txtCodigo(39).Text = Format(CDate(txtCodigo(34).Text) - 1, "dd/mm/yyyy")
            indFrame = 6
            
            NomTabla = "scaalb"
            NomTablaLin = "slialb"
            
            
            
            'Si es facturacion directa: 222 oculto el frame y muestro el albaran que voy a facturar
            FramTaxcoTrabajador.visible = False
            Frame4.visible = (OpcionListado = 52)
            If OpcionListado = 52 Then
                cadFormula = codClien
                If codClien = "ALD" Then cadFormula = "GASOLINERA"
                If codClien = "ALB" Then cadFormula = "TIENDA"
                
                Label10(0).Caption = "Facturación de Albaranes " & cadFormula
                Me.Frame15.Top = 5180 '--5040
                Frame15.visible = True
                cadFormula = ""
                
                
            Else
                Label10(10).Caption = "Albarán:     " & codClien & "   " & NumCod
                Me.Frame15.Top = 1800
                Frame15.visible = False
                If vParamAplic.NumeroInstalacion = vbTaxco Then
                    FramTaxcoTrabajador.visible = True
                    cadFormula = PonerTrabajadorConectado(cadParam)
                    If cadFormula <> "" Then
                        txtCodigo(69).Text = cadFormula
                        txtNombre(69).Text = cadParam
                    End If
                    
                    'DE momento BLOQUEADO PARA TODOS
                    FramTaxcoTrabajador.Enabled = False
                End If
            End If
           
            
            
            
        Case 74, 75 '74: Prefacturación Mantenimientos
                    '75: Facturacion de Mantenimientos
            lblFactMant.Caption = ""
            PonerFramePreFacManteVisible True, H, W
            indFrame = 7 'solo para el boton cancelar
            
            chkSituFacMant.visible = OpcionListado = 74
            Me.FrameTapa.visible = False
            cboSituMan.ListIndex = 0
            
        Case 229 '229: Inf. estadistica ventas por mes
            H = 4000
            W = 7035
            PonerFrameVisible Me.FrameEstVentas, True, H, W
            indFrame = 8
            
            
        Case 80
            indFrame = OpcionListado
            H = FramepedidoAnulado.Height
            W = FramepedidoAnulado.Width
            PonerFrameVisible FramepedidoAnulado, True, H, W
    End Select
    
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel(indFrame).Cancel = True
    Me.Width = W + 70
    Me.Height = H + 350
        
End Sub



Private Sub Form_Unload(Cancel As Integer)
    If OpcionListado = 227 And InstalacionEsEulerTaxco Then LeeGuardaListFacturas False
    
    If OpcionListado = 222 And vParamAplic.NumeroInstalacion = vbTaxco And (codClien = "ALO" Or codClien = "ALE" Or codClien = "ALV") Then
        'SI es taxco, ha cerrado un OT de CONTADO. O factura o no sale de ahi
        If davidNumalbar > 0 Then Cancel = 1
    End If
    FacturaSuperaImporteTicket = False 'Lo dejo aqui por si acaso para que nunca se qede precargado
End Sub


Private Sub frmB_Selecionado(CadenaDevuelta As String)
    cadFormula = CadenaDevuelta
End Sub

Private Sub frmDptoEnvio_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmF_Selec(vFecha As Date)
'Calendario de Fecha
    txtCodigo(indCodigo).Text = Format(vFecha, "dd/mm/yyyy")
End Sub


Private Sub frmLd_DatoSeleccionado(CadenaSeleccion As String)
    txtFraRectifica.Text = RecuperaValor(CadenaSeleccion, 1) & RecuperaValor(CadenaSeleccion, 2) & " de " & RecuperaValor(CadenaSeleccion, 3)
    txtFraRectifica.Tag = CadenaSeleccion
End Sub

Private Sub frmMens_DatoSeleccionado(CadenaSeleccion As String)
    vContinua = (CadenaSeleccion = "OK")
End Sub

Private Sub frmMtoAgente_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Agente
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoAlmacen_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Almacen
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoArticulo_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Articulos
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoBancosPro_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Bancos Propios
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub frmMtoClient_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Clientes
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoFEnvio_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Formas de Envio
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoFPago_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Formas de Pabo
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoTipCo_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Tipos de Contrato del Mantenimiento
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoTraba_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Trabajadores
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmZ_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub imgAyuda_Click(Index As Integer)
Dim Ayuda As String

    'Sera las ayuda. Tampoco queiero la biblia, pero,
    'si un "pelin" de ayuda no me vendria mal a mi, imaginemos a el cliente final
    Select Case Index
    Case 0
        Ayuda = "Si marca la opción 'Semana de entrega' saldrá el informe valorado y agrupado por semana de entrega" & vbCrLf
        Ayuda = Ayuda & "En caso contrario, saldrá el informe sin valorar, con stocks y pedidos a proveedor de los artículos"
        Ayuda = Ayuda & vbCrLf & vbCrLf & "Tipo de cliente válido sólo para la opción 'Sin valorar'(semana entrega desmarcado)"
        Ayuda = Ayuda & vbCrLf & vbCrLf & "- Informe CON IVA. Aplicará IVA y descuentos de cabecera a cada pedido" & vbCrLf
    
    Case 1
        Ayuda = "- Si no indica periodo de facturacion serán todos. " & vbCrLf
        Ayuda = Ayuda & "- IVA incluido. Aplicará IVA y descuentos de cabecera a cada albarán" & vbCrLf
    End Select
    Ayuda = imgayuda(Index).ToolTipText & vbCrLf & String(45, "=") & vbCrLf & Ayuda
    MsgBox Ayuda, vbInformation
End Sub





Private Sub imgBuscarOfer_Click(Index As Integer)
    Select Case Index
            
        Case 11, 12, 14, 15, 20, 21, 27, 28, 32 'Cod. CLIENTE
            Select Case Index
                Case 11, 12: indCodigo = Index + 9
                Case 14, 15: indCodigo = Index + 14
                Case 20, 21: indCodigo = Index + 20
                Case 27, 28: indCodigo = Index + 21
                Case 32: indCodigo = 8
            End Select
            Set frmMtoClient = New frmBasico2
            If Not IsNumeric(txtCodigo(indCodigo).Text) Then txtCodigo(indCodigo).Text = ""
            AyudaClientes frmMtoClient, txtCodigo(indCodigo)
            Set frmMtoClient = Nothing
            
        Case 4, 5 'Cod. ALMACEN
            If Index = 4 Then indCodigo = 13
            If Index = 5 Then indCodigo = 14
            Set frmMtoAlmacen = New frmAlmAlPropios
            frmMtoAlmacen.DatosADevolverBusqueda = "0|1|"
            If Not IsNumeric(txtCodigo(indCodigo).Text) Then txtCodigo(indCodigo).Text = ""
            frmMtoAlmacen.Show vbModal
            Set frmMtoAlmacen = Nothing
            
        Case 6, 7 'Cod. ARTICULO
            If Index = 6 Then
                indCodigo = 15
            Else
                indCodigo = 16
            End If
            Set frmMtoArticulo = New frmBasico2
            'frmMtoArticulo.DatosADevolverBusqueda3 = "@1@"
'            frmMtoArticulo.DesdeTPV = False
'            frmMtoArticulo.Show vbModal
            AyudaArticulos frmMtoArticulo, txtCodigo(indCodigo)
            Set frmMtoArticulo = Nothing
        
        Case 1, 2, 8, 9, 26 'cod. TRABAJADOR
            Select Case Index
                Case 1, 2: indCodigo = Index + 1
                Case 8, 9: indCodigo = Index + 9
                Case 46: indCodigo = 69
                Case 26:  indCodigo = 47
            End Select
            
'            Set frmMtoTraba = New frmAdmTrabajadores
'            frmMtoTraba.DatosADevolverBusqueda = "0|1|"
'            frmMtoTraba.Show vbModal
            Set frmMtoTraba = New frmBasico2
            AyudaTrabajadores frmMtoTraba, txtCodigo(indCodigo)
            Set frmMtoTraba = Nothing
            
        Case 10 'Cod. Forma de Envio
            indCodigo = 19
            Set frmMtoFEnvio = New frmFacFormasEnvio
            frmMtoFEnvio.DatosADevolverBusqueda = "0|1|"
            frmMtoFEnvio.Show vbModal
            Set frmMtoFEnvio = Nothing
            
        Case 16, 17, 22, 23, 29, 30 'Forma de PAGO
            Select Case Index
                Case 16, 17: indCodigo = Index + 14
                Case 22, 23: indCodigo = Index + 20
                Case 29, 30: indCodigo = Index + 21
            End Select
'            Set frmMtoFPago = New frmFacFormasPago
'            frmMtoFPago.DatosADevolverBusqueda = "0|1|"
            If Not IsNumeric(txtCodigo(indCodigo).Text) Then txtCodigo(indCodigo).Text = ""
'            frmMtoFPago.Show vbModal
            Set frmMtoFPago = New frmBasico2
            AyudaFormasPago frmMtoFPago, txtCodigo(indCodigo)
            Set frmMtoFPago = Nothing

        Case 3, 13, 18, 19, 36, 37 'cod AGENTE
            If Index <= 13 Then
                'D/H agente para pedido x cliente
                'MARZO 2010
                indCodigo = 7
                If Index = 3 Then indCodigo = 6
            Else
                If Index < 36 Then
                    indCodigo = Index + 14
                Else
                    indCodigo = Index - 13
                End If
            End If
'            Set frmMtoAgente = New frmFacAgentesCom
'            frmMtoAgente.DatosADevolverBusqueda = "0|1|"
            If Not IsNumeric(txtCodigo(indCodigo).Text) Then txtCodigo(indCodigo).Text = ""
'            frmMtoAgente.Show vbModal
            Set frmMtoAgente = New frmBasico2
            AyudaAgentesComerciales frmMtoAgente, txtCodigo(indCodigo), , True
            Set frmMtoAgente = Nothing
            
        Case 0, 24, 31 'Bancos Propios
            indCodigo = 0
            If Index = 31 Then
                indCodigo = 52
            ElseIf Index = 0 Then
                indCodigo = 5
            Else
                If Index = 24 Then
                    If vParamAplic.NumeroInstalacion = vbTaxco And OpcionListado = 222 Then Exit Sub
                End If
            End If
'            Set frmMtoBancosPro = New frmFacBancosPropios
'            frmMtoBancosPro.DatosADevolverBusqueda = "0|1|"
'            frmMtoBancosPro.Show vbModal
'            Set frmMtoBancosPro = Nothing
            Set frmMtoBancosPro = New frmBasico2
            AyudaBancosPropios frmMtoBancosPro, txtCodigo(indCodigo)
            Set frmMtoBancosPro = Nothing
        
        Case 25 'Tipo CONTRATO
            indCodigo = 45
            Set frmMtoTipCo = New frmManTiposContrato
            frmMtoTipCo.DatosADevolverBusqueda = "0|1|"
            frmMtoTipCo.Show vbModal
            Set frmMtoTipCo = Nothing
        Case 33, 34, 35, 42, 43
            'ZONAS
            If Index = 35 Then
                indCodigo = 22
            ElseIf Index < 35 Then
                indCodigo = (Index - 33) + 9  'txtcodigo(9)
            Else
                'Ventas por cliente
                indCodigo = (Index + 22) 'txtcodigo(64) y 65
            End If
            Set frmZ = New frmFacZonas
            frmZ.DatosADevolverBusqueda = "0|1|"
            frmZ.Show vbModal
            Set frmZ = Nothing
            
        Case 38, 39
                
                If Index = 38 Or Index = 39 Then
                    If txtCodigo(48).Text <> txtCodigo(49).Text Or txtCodigo(49).Text = "" Then
                        MsgBox "Seleccione un cliente para podre indicar departamento", vbExclamation
                        Exit Sub
                    End If
                    
                    'Tienen que existir el cliente
                    If txtNombre(48).Text = "" Then
                        MsgBox "No existe el cliente", vbExclamation
                        Exit Sub
                    End If
                    indCodigo = Index + 17
                End If
                Set frmDptoEnvio = New frmFacCliEnvDpto
                frmDptoEnvio.DireccionesEnvio = False
                If txtCodigo(Index).Text <> "" Then
                    frmDptoEnvio.VerDatoDpto = CInt(txtCodigo(Index).Text)
                Else
                    frmDptoEnvio.VerDatoDpto = -1
                End If
                frmDptoEnvio.codClien = CLng(txtCodigo(48).Text)
                frmDptoEnvio.NomClien = txtNombre(48).Text
                frmDptoEnvio.Show vbModal
                Set frmDptoEnvio = Nothing
                
        Case 40, 41
            indCodigo = Index + 17   '57 y 58
            AbrirBuscaGrid indCodigo
        
        Case 44, 45
            indCodigo = Index + 22
            AbrirBuscaGrid indCodigo
        
        'familia - proveedor inicdencia
        Case 47, 48, 49, 50, 51, 52
            indCodigo = Index + 25
            AbrirBuscaGrid indCodigo
        
        
        Case 54
            indCodigo = 54
            AbrirBuscaGrid indCodigo
    End Select
    PonerFoco txtCodigo(indCodigo)
End Sub


Private Sub imgFecha_Click(Index As Integer)
   Screen.MousePointer = vbHourglass
   Set frmF = New frmCal
   frmF.Fecha = Now
   
   Select Case Index
        Case 0 'Frame Pasar Pedido -> Albaran
            indCodigo = 25
        'Case 1 'framePedidos
        '    indCodigo = 3 'Desde
        'Case 2 'framePedidos
        '    indCodigo = 4 'Hasta
        'marzo 22
        Case 1, 2
            indCodigo = Index + 69
            
        Case 6 'FramePedxArtic
            indCodigo = 11 'Fecha Desde
        Case 7 'FramePedxArtic
            indCodigo = 12 'Fecha Hasta
        Case 9 'FramePedCompras
            indCodigo = 24 'Fecha Hasta
        Case 10 'FramePreFacturar
            indCodigo = 26
        Case 11 'FramePreFacturar
            indCodigo = 27
        Case 12 'Frame Factura
            indCodigo = 38
        Case 13 'Frame Factura
            indCodigo = 39
        Case 14 'FrameFactura
            indCodigo = 34
        Case 15 'FrameFactura
            indCodigo = 44
   End Select
   
   PonerFormatoFecha txtCodigo(indCodigo)
   If txtCodigo(indCodigo).Text <> "" Then frmF.Fecha = CDate(txtCodigo(indCodigo).Text)

   Screen.MousePointer = vbDefault
   frmF.Show vbModal
   Set frmF = Nothing
   PonerFoco txtCodigo(indCodigo)
End Sub










Private Sub lblDestinoB_Click()
     
        If vUsu.Nivel <= 1 Then
            If cboDestinoB.ListCount = 1 Then cboDestinoB.AddItem "Presupuesto"
        End If
    
End Sub

Private Sub ListTipoFact_KeyPress(KeyAscii As Integer)
     KEYpress KeyAscii
End Sub

Private Sub OptDetalle_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub OptOrdenCodclien_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub



Private Sub OptOrdenVentas_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtCodigo_GotFocus(Index As Integer)
    ConseguirFoco txtCodigo(Index), 3
End Sub

Private Sub txtCodigo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtCodigo_KeyPress(Index As Integer, KeyAscii As Integer)
'    KEYpress KeyAscii
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 28: KEYBusqueda KeyAscii, 14 'cliente
            Case 29: KEYBusqueda KeyAscii, 15 'cliente
            Case 30: KEYBusqueda KeyAscii, 16 'forma de pago
            Case 31: KEYBusqueda KeyAscii, 17 'forma de pago
            Case 32: KEYBusqueda KeyAscii, 18 'agentes
            Case 33: KEYBusqueda KeyAscii, 19 'agentes
            
            Case 26: KEYFecha KeyAscii, 10 'fecha
            Case 27: KEYFecha KeyAscii, 11 'fecha
        
            Case 40: KEYBusqueda KeyAscii, 20 'cliente
            Case 41: KEYBusqueda KeyAscii, 21 'cliente
            Case 42: KEYBusqueda KeyAscii, 22 'forma de pago
            Case 43: KEYBusqueda KeyAscii, 23 'forma de pago
            Case 69: KEYBusqueda KeyAscii, 46 'trabajador
            Case 0: KEYBusqueda KeyAscii, 24 'cta provista de cobro
            
            Case 34: KEYFecha KeyAscii, 14 'fecha
            Case 38: KEYFecha KeyAscii, 12 'fecha
            Case 39: KEYFecha KeyAscii, 13 'fecha
            Case 44: KEYFecha KeyAscii, 15 'fecha
        
            Case 45: KEYBusqueda KeyAscii, 25 'tipo de contrato
            Case 48: KEYBusqueda KeyAscii, 27 'cliente
            Case 49: KEYBusqueda KeyAscii, 28 'cliente
            Case 50: KEYBusqueda KeyAscii, 29 'forma de pago
            Case 51: KEYBusqueda KeyAscii, 30 'forma de pago
            Case 55: KEYBusqueda KeyAscii, 38 'departamento
            Case 56: KEYBusqueda KeyAscii, 39 'departamento
            Case 54: KEYBusqueda KeyAscii, 54 'centro de coste
            Case 52: KEYBusqueda KeyAscii, 31 'cta prev cobro
            Case 47: KEYBusqueda KeyAscii, 26 'operador
            
            
            
            Case 72 To 77
                     KEYBusqueda KeyAscii, Index - 25 'operador
        End Select
    Else
        KEYpress KeyAscii
    End If

End Sub


Private Sub KEYBusqueda(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgBuscarOfer_Click (Indice)
End Sub

Private Sub KEYFecha(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgFecha_Click (Indice)
End Sub



Private Sub txtCodigo_LostFocus(Index As Integer)
Dim devuelve As String
Dim codCampo As String, NomCampo As String
Dim tabla As String
      
    Select Case Index
        Case 1, 60, 68 'Importe (Decimal(12,2))
            If Not PonerFormatoDecimal(txtCodigo(Index), 1) Then
                If Index = 68 Then txtCodigo(Index) = ""
                If Index = 1 Then txtCodigo(Index) = ""
            End If
            'Puntos
            If Index = 68 Then
                If txtCodigo(Index) = "" Then txtCodigo(Index) = "0"
                If ImporteFormateado(txtCodigo(Index).Text) > CCur(txtCodigo(Index).Tag) Then
                    MsgBox "Maximo puntos a canjear: " & txtCodigo(Index).Tag, vbExclamation
                    txtCodigo(68).Text = Format(txtCodigo(68).Tag, FormatoImporte)
                End If
            End If
        Case 0, 5, 52 'Bancos Propios
            If PonerFormatoEntero(txtCodigo(Index)) Then
            
                If OpcionListado = 222 Then
                    If vParamAplic.NumeroInstalacion = vbTaxco And codClien <> "ALE" Then
                        tabla = ""
                        If Val(txtCodigo(Index)) < 1 Then tabla = "N"
                        If Val(txtCodigo(Index)) > 3 Then tabla = "N"
            
                        If tabla <> "" Then
                            MsgBox "Valores permitidos: 1-EFECTIVO   2-CREDITO     3-TARJETA", vbExclamation
                            txtCodigo(Index).Text = ""
                            txtNombre(Index).Text = ""
                            PonerFoco txtCodigo(Index)
                            Exit Sub
                        End If
                    End If
                End If
                txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, "sbanpr", "nombanpr", "codbanpr", "Bancos Propios")
                If txtCodigo(Index).Text <> "" And txtNombre(Index).Text <> "" Then
                    txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "0000")
                Else
                    PonerFoco txtCodigo(Index)
                End If
            Else
                txtNombre(Index).Text = ""
            End If
        
        'FECHA Desde Hasta
        Case 11, 12, 25, 26, 27, 34, 38, 39, 44, 70, 71
            If txtCodigo(Index).Text <> "" Then
                PonerFormatoFecha txtCodigo(Index)
                If Index = 34 And txtCodigo(34).Text <> "" Then _
                    txtCodigo(39).Text = Format(CDate(txtCodigo(34).Text) - IIf(vParamAplic.NumeroInstalacion = vbTaxco, 0, 1), "dd/mm/yyyy")
            End If
           
'            'Fecha entrega para Pedido. Poner la semana
'            If Index = 26 Then txtNombre(4).Text = CalculaSemana(CDate(txtCodigo(26).Text))
        
        Case 53, 59 'AÑO - litros reales partes trabajo
             If Not PonerFormatoEntero(txtCodigo(Index)) Then
                If Index = 59 Then txtCodigo(Index).Text = ""
             End If
        
        Case 36, 37, 62 'Nº de Pedido / Albaran
            If PonerFormatoEntero(txtCodigo(Index)) Then
                If Index <> 62 Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "0000000")
            Else
                txtCodigo(Index).Text = ""
            End If
            

        Case 35  'Num coppias y Periodicidad Facturacion
            PonerFormatoEntero txtCodigo(Index)

        Case 8, 20, 21, 28, 29, 40, 41, 48, 49 'Cod. CLIENTE
            If PonerFormatoEntero(txtCodigo(Index)) Then
                NomCampo = "nomclien"
                tabla = "sclien"
                codCampo = "codclien"
                txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, tabla, NomCampo, codCampo, "Cliente")
                If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000000")
            Else
                txtNombre(Index).Text = ""
            End If
            
        Case 13, 14 'ALMACEN
            If PonerFormatoEntero(txtCodigo(Index)) Then
                NomCampo = "nomalmac"
                tabla = "salmpr"
                codCampo = "codalmac"
                txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, tabla, NomCampo, codCampo, "Almacen")
                If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000")
            Else
                txtNombre(Index).Text = ""
            End If
     
        Case 2, 3, 17, 18, 47, 69 'Cod. Trabajador
            If PonerFormatoEntero(txtCodigo(Index)) Then
                NomCampo = "nomtraba"
                tabla = "straba"
                codCampo = "codtraba"
                txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, tabla, NomCampo, codCampo, "Trabajador")
                If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "0000")
            Else
                txtNombre(Index).Text = ""
            End If
            
        Case 19 'Cod. Envio
            NomCampo = "nomenvio"
            tabla = "senvio"
            codCampo = "codenvio"
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, tabla, NomCampo, codCampo, "Forma de Envío")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000")
            
            
        Case 30, 31, 42, 43, 50, 51 'Cod. Formas de PAGO
            If PonerFormatoEntero(txtCodigo(Index)) Then
                txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, "sforpa", "nomforpa", "codforpa", "Formas de Pago")
                If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000")
            Else
                txtNombre(Index).Text = ""
                txtCodigo(Index).Text = ""
            End If
        
        Case 6, 7, 32, 33, 23, 24 'AGENTE
            If PonerFormatoEntero(txtCodigo(Index)) Then
                txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, "sagent", "nomagent", "codagent", "Agente")
                If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "0000")
            Else
                txtNombre(Index).Text = ""
                txtCodigo(Index).Text = ""
            End If
        Case 9, 10, 22, 64, 65 'ZONA
            If PonerFormatoEntero(txtCodigo(Index)) Then
                txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, "szonas", "nomzonas", "codzonas", "Zonas")
                If Index = 22 Then
                    'Si pone ZONA tiene que existir
                    If txtNombre(Index).Text = "" Then txtCodigo(Index).Text = ""
                End If
                If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "0000")
            Else
                txtNombre(Index).Text = ""
                txtCodigo(Index).Text = ""
            End If
            
        Case 45 'TIPO CONTRATO
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, "stipco", "nomtipco", "codtipco", "Tipo Contrato", "T")
            
        Case 46 'MES a facturar
            If PonerFormatoEntero(txtCodigo(Index)) Then
                'Comprobar que el mes es correcto, valores entre 1-12
                devuelve = txtCodigo(Index).Text
                If (CByte(devuelve) >= 1) And (CByte(devuelve) <= 12) Then
                    txtNombre(Index).Text = UCase(MonthName(CLng(devuelve)))
                Else
                    MsgBox "El valor introducido no es un MES válido.(1-12).", vbInformation
                    txtCodigo(Index).Text = ""
                    PonerFoco txtCodigo(Index)
                End If
            Else
                txtNombre(Index).Text = ""
            End If
            
            
        Case 54
            'Centro de coste
            txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)
            codCampo = ""
            If txtCodigo(Index).Text <> "" Then
                
                codCampo = "nomccost"
                tabla = DevuelveDesdeBD(conConta, "codccost", IIf(vParamAplic.ContabilidadNueva, "cabccost", "ccoste"), "codccost", txtCodigo(Index).Text, "T", codCampo)
            
                If tabla = "" Then
                    MsgBox "No existe el centro de coste: " & txtCodigo(Index).Text, vbExclamation
                    
                End If
                If codCampo = "nomccost" Then codCampo = ""
                txtCodigo(Index).Text = tabla
            End If
            txtNombre(Index).Text = codCampo
            
            
            
        Case 55, 56
            'DEPARTAMENTO
            codCampo = ""
            devuelve = ""
            If txtCodigo(Index).Text = "" Then
                txtNombre(Index).Text = codCampo
                Exit Sub
            End If
            
            If Index = 55 Or Index = 56 Then
                'Departemento. Tienen que poner un UNICO cliente
                If txtCodigo(48).Text <> txtCodigo(49).Text Or txtCodigo(49).Text = "" Then
                    MsgBox "Seleccione un cliente para poder indicar departamento", vbExclamation
                    txtCodigo(Index).Text = ""
                    PonerFoco txtCodigo(48)
                    devuelve = "NO"
                End If
            End If
            If devuelve = "" Then
                If PonerFormatoEntero(txtCodigo(Index)) Then
                    codCampo = PonerNombreDeCod(txtCodigo(Index), conAri, "sdirec", "nomdirec", "codclien=" & txtCodigo(48).Text & " AND coddirec ", "Departamentos")
                    If txtCodigo(Index).Text <> "" And codCampo <> "" Then
                        txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "0000")
                    Else
                        txtCodigo(Index).Text = ""
                        PonerFoco txtCodigo(Index)
                    End If
                Else
                    txtCodigo(Index).Text = ""
                End If
            End If
            txtNombre(Index).Text = codCampo
            
        Case 57, 58
            If PonerFormatoEntero(txtCodigo(Index)) Then
                NomCampo = "nomactiv"
                tabla = "sactiv"
                codCampo = "codactiv"
                txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, tabla, NomCampo, codCampo, "Actividad")
                If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000")
            Else
                txtCodigo(Index).Text = ""
                txtNombre(Index).Text = ""
            End If
        
        Case 66, 67
             If PonerFormatoEntero(txtCodigo(Index)) Then
                NomCampo = "nomrutas"
                tabla = "srutas"
                codCampo = "codrutas"
                txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, tabla, NomCampo, codCampo, IIf(vParamAplic.NumeroInstalacion = 2, "Asociacion", "Ruta"))
                If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000")
            Else
                txtCodigo(Index).Text = ""
                txtNombre(Index).Text = ""
            End If
                
                
                
        Case 72, 73
            If PonerFormatoEntero(txtCodigo(Index)) Then
                NomCampo = "nomprove"
                tabla = "sprove"
                codCampo = "codprove"
                txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, tabla, NomCampo, codCampo, "Proveedor")
                If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000")
            Else
                txtNombre(Index).Text = ""
            End If
        
        
        Case 74, 75
            If PonerFormatoEntero(txtCodigo(Index)) Then
                NomCampo = "nomfamia"
                tabla = "sfamia"
                codCampo = "codfamia"
                txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, tabla, NomCampo, codCampo, "Familia")
                If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000")
            Else
                txtNombre(Index).Text = ""
            End If
        
        Case 76, 77
            If Trim(txtCodigo(Index)) <> "" Then
                NomCampo = "nomincid"
                tabla = "sincid"
                codCampo = "codincid"
                txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, tabla, NomCampo, codCampo, "Inicidencia", "T")
                'If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000")
            Else
                txtNombre(Index).Text = ""
            End If
        '##### Recuperar facturas ALZIRA
        Case 4 'nº factura
            PonerFocoBtn Me.cmdAceptarFac
        '#####
    End Select
End Sub



Private Sub PonerFramePedxArticVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Pone el Frame de Informe Pedidos por Articulo Visible y Ajustado al Formulario, y visualiza los controles
'necesario para cada Informe de Ofertas

    'MARZO 2010
    'Los botones de aceptar cancelar
    H = 4800
    If OpcionListado = 41 Then
        H = 6080
    Else
        If OpcionListado = 44 Then
            H = 8080
        Else
            If OpcionListado = 227 Then H = 9500 '8080
        End If
    End If
    cmdAceptarPedxArtic.Top = H
    Me.cmdCancel(2).Top = H

    'lbl
    H = 4680
    If OpcionListado = 41 Then
        H = 4880
    Else
        If OpcionListado = 44 Then
            H = 6880
        Else
            If OpcionListado = 227 Then H = 5880
        End If
    End If
    Label4(54).Top = H

    'El form
    H = 5415
    If OpcionListado = 41 Then
        H = 6575
    Else
        If OpcionListado = 44 Then
            H = 8575
        Else
            If OpcionListado = 227 Then H = 10000 '8500
        End If
    End If
    W = 7515
    

    
        
    'Ajustar Tamaño del Frame para ajustar tamaño de Formulario al del Frame
    PonerFrameVisible Me.FramePedxArtic, visible, H, W
    
    If visible = True Then
        Me.Frame5.visible = OpcionListado = 41 Or (OpcionListado = 44) Or (OpcionListado = 227) 'D/H cliente
        'D/H Artículo
        'Me.Frame8.visible = (OpcionListado <> 44) And (OpcionListado <> 227) And (OpcionListado <> 228)
        Me.Frame8.visible = (OpcionListado <> 227) And (OpcionListado <> 228)
        
        Me.Frame9.visible = (OpcionListado <> 227 And OpcionListado <> 228) 'D/H Almacen
        Me.Frame10.visible = (OpcionListado = 227)
        FrameAsociacion.visible = (OpcionListado = 227)
        FrameAsociacion.BorderStyle = 0
        Me.Frame12.visible = (OpcionListado = 228)
        FrameOrden1.visible = (OpcionListado = 44)
        FramepedxClien.visible = (OpcionListado = 44)
        FrameZonaCli.visible = (OpcionListado = 227)
        
        
        
        
        'Para que salga
        
        If OpcionListado = 44 Then 'Informe Pedido por cliente
            Me.Frame5.Top = 4320
            Me.Frame5.Left = 400
            Me.label1.Caption = "Pedidos por Cliente"
            '
            FramepedxClien.Top = 5440
            FramepedxClien.Left = 500
            Frame8.Left = 320
            
            FrameOrden1.Top = 7600
            chkPedxClixSemEntrega(2).Top = 7710
        ElseIf OpcionListado = 227 Then 'Inf. Estadistica ventas x cliente
            Me.Frame5.Top = 1800
            Me.Frame5.Left = 500
            Me.FrameZonaCli.Top = 2800
            FrameAsociacion.Top = 3800   '8000
            Me.Frame10.Top = 4800        '3800
            
            FrameZonaCli.Left = 500
            Me.label1.Caption = "Ventas por Cliente"
            Label4(4).Caption = "Fecha Factura"
            
            
            Label4(72).Caption = IIf(vParamAplic.NumeroInstalacion = vbHerbelca, "Asociacion", "Rutas")
            Frame10.BorderStyle = 0
            FrameZonaCli.BorderStyle = 0
            
            
            
            'Me.cmdAceptarPedxArtic.Top = 4850
            'Me.cmdCancel(2).Top = Me.cmdAceptarPedxArtic.Top
        ElseIf OpcionListado = 228 Then 'Inf. Estadistica ventas x trabajador
            Me.Frame12.Top = 1900
            Me.Frame12.Left = 500
            Me.label1.Caption = "Ventas por Trabajador"
            Label4(4).Caption = "Fecha Factura"
            Me.cmdAceptarPedxArtic.Top = 4150
            Me.cmdCancel(2).Top = Me.cmdAceptarPedxArtic.Top
        Else
            Me.Frame8.Top = 3120
            Me.Frame8.Left = 500
            If OpcionListado = 41 Then
                Me.label1.Caption = "Pedidos por Artículo"
                Frame5.Top = 4300
                Frame5.Left = 300
                Me.Frame8.Top = 3020
                Me.Frame8.Left = 300
            ElseIf OpcionListado = 42 Then
                Me.label1.Caption = "Disponibilidad de Stocks"
            ElseIf OpcionListado = 49 Then
                Me.label1.Caption = "Albaranes por Artículo"
                Label4(4).Caption = "Fecha Albaran"
            End If
        End If
    End If
End Sub


Private Sub PonerFramePreFacVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Pone el Frame del Prevision Facturacion Albaran Visible y Ajustado al Formulario, y visualiza los controles
Dim B As Boolean
Dim Cad As String

    H = 7095  '6355
    If OpcionListado = 51 Then 'Inf. Incum. plazos entrega
        H = 5400
        Me.cmdAceptarPreFac.Top = 4600
        Me.cmdCancel(5).Top = Me.cmdAceptarPreFac.Top
    Else
        Frame16.visible = True 'tipos de pago
    End If
    W = 6800
    'Ajustar Tamaño del Frame para ajustar tamaño de Formulario al del Frame
    PonerFrameVisible Me.FramePreFacturar, visible, H, W
    If visible = True Then
        B = (OpcionListado = 50)
        Label4(41).visible = B
        Me.imgBuscarOfer(16).visible = B
        Me.imgBuscarOfer(17).visible = B
        Me.txtCodigo(30).visible = B
        Me.txtCodigo(31).visible = B
        Me.txtNombre(30).visible = B
        Me.txtNombre(31).visible = B
        Me.Frame6.visible = Not B
        Me.Frame6.Top = 2750
        Me.Frame6.Left = 460
        
        'solo albaranes a facturar
        Me.chkSoloFacturar.visible = B
        Me.chkSoloFacturar.Value = 1
        
        'Detalle o resumen
        Me.Frame7.visible = B And codClien = "ALV"
        Me.Frame7.visible = B 'And CodClien = "ALV"
        If vParamAplic.NumeroInstalacion = 5 Then
            'En fontenas no quitamos solo facturar
            Me.OptDetalle(4).Value = True
            If B Then Me.chkSoloFacturar.Value = 0
        Else
            Me.OptDetalle(0).Value = True
        End If
        'Sept 2015. Periodo facturacion
        Me.Label4(24).visible = B
        Me.txtCodigo(61).visible = B
        Me.imgayuda(1).visible = B
        
        
        If Not B Then
            Me.Label9(0).Caption = "Incum. plazos entrega"
        Else 'Prevision Facturacion
            Select Case codClien 'aqui guardamos el tipo de movimiento
                Case "ALV": Cad = "" ' antes "(Ventas)" [SERVICIOS]
                Case "ALR": Cad = "(Reparaciones)"
                Case "ALM": Cad = "(Mantenimientos)"
            End Select
            Me.Label9(0).Caption = "Previsión de facturación " & Cad
        End If
    End If
End Sub


Private Sub PonerFrameFacVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Pone el Frame de Facturacion de Albaran Visible y Ajustado al Formulario, y visualiza los controles
Dim Cad As String

    H = 7100 + 280 '--180
    W = 8565 '--7480
    
    If visible = True Then
         Select Case codClien 'aqui guardamos el tipo de movimiento
            Case "ALV": Cad = "(Ventas)"
            Case "ALR": Cad = "(Reparaciones)"
            Case "ALM", "ART":
                If codClien = "ALM" Then
                    Cad = "(Mostrador)"
                Else
                    Cad = "(Rectificativa)"
                End If
                'Me.Frame3.Top = 1200
                Me.Frame4.visible = False
                H = 4000
                Me.cmdAceptarFac.Top = 3200
                Me.cmdCancel(6).Top = Me.cmdAceptarFac.Top
            Case "ALS": Cad = "(Servicios)"
                
        End Select
        
        
        'marzo 2016
        Me.Label10(0).Caption = "Facturación de Albaranes " & Cad
        Me.Caption = "Facturación"
    End If
    
    PonerFrameVisible Me.FrameFacturar, visible, H, W
End Sub


Private Sub PonerFramePreFacManteVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Pone el Frame del Prevision Facturacion Albaran Mantenimientos Visible y Ajustado al Formulario, y visualiza los controles
Dim B As Boolean
Dim Cad As String

    
    If visible = True Then
        B = (OpcionListado = 74) 'prefacturar
        W = 8620 '7120
        If B Then 'prefacturar
            'H = 5600
            H = 7600
        Else 'facturar
            'H = 6855
            H = 7855
            
        End If
        'H = 6855
        Me.FramePreFacMante.Height = H
        'Ajustar Tamaño del Frame para ajustar tamaño de Formulario al del Frame
        PonerFrameVisible Me.FramePreFacMante, visible, H, W
        
        If B Then 'prefacturar
            Me.Frame2(0).visible = False
            Me.Frame2(1).visible = False
            Me.Frame1.Top = Me.Frame1.Top - 800
            Me.Frame1.BorderStyle = 0
            Me.Label7(1).Caption = "Prefacturación Mantenimientos"
            Me.cmdAceptarPreFacMan.Top = H - Me.cmdAceptarPreFacMan.Height - 120
            Me.cmdCancel(7).Top = cmdAceptarPreFacMan.Top
        Else 'facturar
            Me.Label7(1).Caption = "Facturación Mantenimientos"
            Me.txtCodigo(44).Text = Format(Now, "dd/mm/yyyy")
            Me.txtCodigo(47).Text = PonerTrabajadorConectado(Cad)
            Me.txtNombre(47).Text = Cad
            B = False                            'Si es por proyecto pedira el CC, si no cojera el tel trab o la familia
            If vEmpresa.TieneAnalitica Then B = vParamAplic.ModoAnalitica = 2
            Me.Frame2(1).visible = B
        End If
        Me.lblFactMant.Top = H - Me.lblFactMant.Height - 120
    End If
    txtCodigo(54).Text = ""
End Sub



Private Function AnyadirParametroDH(Cad As String, indD As Byte, indH As Byte) As String
On Error Resume Next

    If txtCodigo(indD).Text <> "" And txtCodigo(indH).Text <> "" Then
        If txtCodigo(indD).Text = txtCodigo(indH).Text Then
            Cad = Cad & txtCodigo(indD).Text
            If txtNombre(indD).Text <> "" Then Cad = Cad & " - " & txtNombre(indD).Text
            AnyadirParametroDH = Cad
            Exit Function
        End If
    End If
    
    If txtCodigo(indD).Text <> "" Then
        Cad = Cad & "desde " & txtCodigo(indD).Text
        If txtNombre(indD).Text <> "" Then Cad = Cad & " - " & txtNombre(indD).Text
    End If
    If txtCodigo(indH).Text <> "" Then
        Cad = Cad & "  hasta " & txtCodigo(indH).Text
        If txtNombre(indH).Text <> "" Then Cad = Cad & " - " & txtNombre(indH).Text
    End If
    AnyadirParametroDH = Cad
End Function


Private Function PonerDesdeHasta(Campo As String, Tipo As String, indD As Byte, indH As Byte, param As String) As Boolean
Dim devuelve As String

    PonerDesdeHasta = False
    devuelve = CadenaDesdeHasta(txtCodigo(indD).Text, txtCodigo(indH).Text, Campo, Tipo)
    If devuelve = "Error" Then Exit Function
    If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Function
    If Tipo <> "F" Then
        If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Function
    End If
    If devuelve <> "" Then
        If param <> "" Then
            'Parametro Desde/Hasta
            cadParam = cadParam & AnyadirParametroDH(param, indD, indH) & """|"
            numParam = numParam + 1
        End If
        PonerDesdeHasta = True
    End If
End Function


Private Sub InicializarVbles()
    cadFormula = ""
    cadSelect = ""
    cadParam = ""
    numParam = 0
End Sub


Private Sub LlamarImprimir()
    With frmImprimir
        .FormulaSeleccion = cadFormula
        .OtrosParametros = cadParam
        .NumeroParametros = numParam
        
        .SoloImprimir = False
        
        .EnvioEMail = False
        .Opcion = OpcionListado
        .Titulo = Titulo
        .ConSubInforme = conSubRPT
        .NombreRPT = nomRPT  'nombre del informe
        .Show vbModal
    End With
End Sub

Private Sub txtCodigo_Validate(Index As Integer, Cancel As Boolean)
    Select Case Index
           Case 15, 16 'ARTICULO
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, "sartic", "nomartic", "codartic", "Articulo", "T")
            'If txtNombre(Index).Text = "" And txtCodigo(Index) <> "" Then Cancel = True
    End Select
End Sub




Private Function ObtenerClientesNuevo(cadW As String, Importe As String) As String
Dim Sql As String
Dim RS As ADODB.Recordset

    On Error GoTo EClientes
    
    cadW = Replace(cadW, "{", "")
    cadW = Replace(cadW, "}", "")
    
    Sql = "select scafac.codclien,scafac.nomclien,sum(baseimp1),sum(baseimp2),sum(baseimp3),sum(baseimp1)+ sum(if(isnull(baseimp2),0,baseimp2))+ sum(if(isnull(baseimp3),0,baseimp3)) as BaseImp"
    Sql = Sql & " From scafac,sclien WHERE scafac.codclien=sclien.codclien "
    If txtCodigo(57).Text <> "" Then Sql = Sql & " AND sclien.codactiv >= " & txtCodigo(57).Text
    If txtCodigo(58).Text <> "" Then Sql = Sql & " AND sclien.codactiv <= " & txtCodigo(58).Text
    'El agente
    If Me.txtCodigo(23).Text <> "" Then Sql = Sql & " AND scafac.codagent >= " & txtCodigo(23).Text
    If Me.txtCodigo(24).Text <> "" Then Sql = Sql & " AND scafac.codagent <= " & txtCodigo(24).Text
    
    
    If cadW <> "" Then Sql = Sql & " AND " & cadW
    Sql = Sql & " group by codclien "
    If Importe <> "" Then Sql = Sql & " having baseimp>" & Importe
    
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Sql = ""
    While Not RS.EOF
'        If RS!BaseImp >= CCur(Importe) Then
            Sql = Sql & RS!codClien & ","
'        End If
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    
    'Si no tiene DATOS es que ninguno entra dentro de estos registros
    If Sql = "" Then Sql = "-1-"
    
    If Sql <> "" Then
        Sql = Mid(Sql, 1, Len(Sql) - 1)
        Sql = "( {scafac.codclien} IN [" & Sql & "] )"
    End If
    ObtenerClientesNuevo = Sql
    
EClientes:
   If Err.Number <> 0 Then MuestraError Err.Number, , Err.Description
End Function






Private Sub AbrirBuscaGrid(OP As Integer)
    
    Set frmB = New frmBuscaGrid
    cadFormula = "" 'Aqui metera el valor
    Select Case OP
    Case 54
        'CEntro de coste
        If vParamAplic.ContabilidadNueva Then
            frmB.vCampos = "Codigo|ccoste|codccost|T||20·Descripción|ccoste|nomccost|T||70·"
            frmB.vTabla = "ccoste"
        Else
            frmB.vCampos = "Codigo|cabccost|codccost|T||20·Descripción|cabccost|nomccost|T||70·"
            frmB.vTabla = "cabccost"
        End If

        frmB.vSQL = ""
       
        '###A mano
        frmB.vDevuelve = "0|1|"
        frmB.vTitulo = "Centros de coste"
        frmB.vselElem = 0
        frmB.vConexionGrid = conConta
    Case 66, 67
        
        frmB.vCampos = "Codigo|srutas|codrutas|T||20·Descripción|srutas|nomrutas|T||70·"
        frmB.vTabla = "srutas"
        frmB.vSQL = ""
       
        '###A mano
        frmB.vDevuelve = "0|1|"
        frmB.vTitulo = IIf(vParamAplic.NumeroInstalacion = vbHerbelca, "Asociacion", "Rutas")
        frmB.vselElem = 0
        frmB.vConexionGrid = conAri
    
    
    
    Case 57, 58
        frmB.vCampos = "Codigo|sactiv|codactiv|T||20·Descripción|sactiv|nomactiv|T||70·"
        frmB.vTabla = "sactiv"
        frmB.vSQL = ""
       
        '###A mano
        frmB.vDevuelve = "0|1|"
        frmB.vTitulo = "Actividades"
        frmB.vselElem = 0
        frmB.vConexionGrid = conAri
    
    
    Case 72, 73
        frmB.vCampos = "Codigo|sprove|codprove|T|0000|20·Nombre|sprove|nomprove|T||70·"
        frmB.vTabla = "sprove"
        frmB.vSQL = ""
       
        '###A mano
        frmB.vDevuelve = "0|1|"
        frmB.vTitulo = "Proveedores"
        frmB.vselElem = 0
        frmB.vConexionGrid = conAri
    
    Case 74, 75
        frmB.vCampos = "Codigo|sfamia|codfamia|N|0000|20·Descripción|sfamia|nomfamia|T||70·"
        frmB.vTabla = "sfamia"
        frmB.vSQL = ""
       
        '###A mano
        frmB.vDevuelve = "0|1|"
        frmB.vTitulo = "Familias"
        frmB.vselElem = 0
        frmB.vConexionGrid = conAri
    Case 76, 77
        frmB.vCampos = "Codigo|sincid|codincid|T||20·Descripción|sincid|nomincid|T||70·"
        frmB.vTabla = "sincid"
        frmB.vSQL = ""
       
        '###A mano
        frmB.vDevuelve = "0|1|"
        frmB.vTitulo = "Incidencias"
        frmB.vselElem = 0
        frmB.vConexionGrid = conAri

    
    
    End Select
    frmB.Show vbModal
    Set frmB = Nothing
    
    
    If cadFormula <> "" Then
        'Ha devuelto algun dato
        'If Op = 54 Then
            txtCodigo(OP).Text = RecuperaValor(cadFormula, 1)
            txtNombre(OP).Text = RecuperaValor(cadFormula, 2)
        'End If
    End If
End Sub





'-----------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------
'
'   Generacion de portes TIPO HERBELCA
'
'-----------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------
Private Function ProcesoPortesTipoHerbelca(cadSQL As String, cadWhere As String, ByRef LblBar As Label) As Boolean
    
    ProcesoPortesTipoHerbelca = False
    'Primera pasada para comprobar

    If Not CargaPorteTipoHerbelca(True, cadSQL, cadWhere, LblBar) Then Exit Function
 
    'Segunda pasada para insertar los portes
    If CargaPorteTipoHerbelca(False, cadSQL, cadWhere, LblBar) Then ProcesoPortesTipoHerbelca = True
        

    
End Function


Private Function CargaPorteTipoHerbelca(Comprobar As Boolean, cadSQL As String, cadWhere As String, ByRef LblBar As Label) As Boolean
Dim RSalb As ADODB.Recordset
Dim Sql As String
Dim Codclien1 As Long
Dim ClienConPortes As Boolean
Dim cadW As String
Dim FecEnvio As Date
Dim T1 As Single
Dim DatosPortes As String  'nomartic|preciov|preciouc|

    On Error GoTo ETraspasoAlbFac

    CargaPorteTipoHerbelca = False

    'Meteremos en una de las temporales los registros que comprobando den error
    Set RSalb = New ADODB.Recordset
    
    If Comprobar Then
        Sql = "Delete from tmpsliped where codusu = " & vUsu.Codigo
        conn.Execute Sql
    Else
        Sql = "Select nomartic,preciove,preciouc from sartic where codartic = " & DBSet(vParamAplic.ArtPortesN, "T")
        RSalb.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        'No pueser eof
        DatosPortes = RSalb!NomArtic & "|"
        Sql = CStr(RSalb!PrecioVe)
        DatosPortes = DatosPortes & TransformaComasPuntos(Sql) & "|"
        Sql = CStr(RSalb!precioUC)
        DatosPortes = DatosPortes & TransformaComasPuntos(Sql) & "|"
        RSalb.Close
    End If
    '
    'Marcar Albaranes que se van a Facturar
    '----------------------------------------
    Sql = "select scaalb.codclien,fecenvio,NumAlbar,scaalb.nomclien from scaalb,sclien where scaalb.codclien=Sclien.codclien AND "
    Sql = Sql & cadWhere
    'de transporte y con fecha de envio
    Sql = Sql & " AND  tipalbaran=1 and not fecenvio is null order by codclien,fecenvio,numalbar"
    
    
    LblBar.Caption = "Leyendo datos"
    LblBar.Refresh
    RSalb.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Codclien1 = -1
    ClienConPortes = False
    cadW = ""
    T1 = Timer
    DoEvents
    While Not RSalb.EOF
        
            
        If RSalb!codClien <> Codclien1 Then
            If cadW <> "" Then
                If ClienConPortes Then TratarLineaPortes Comprobar, cadW, Sql, DatosPortes, Codclien1
            End If
            Espera 0.1
            ClienConPortes = False
            Codclien1 = RSalb!codClien
            Sql = RSalb!NomClien
            FecEnvio = RSalb!FecEnvio
            cadW = DevuelveDesdeBD(conAri, "AplicaPortesFactura", "sclien", "codclien", CStr(Codclien1))
            If cadW = "1" Then ClienConPortes = True
            cadW = " (slialb.codtipom='ALV' AND slialb.numalbar IN (" & RSalb!Numalbar
            If ClienConPortes Then
                LblBar.Caption = "Cliente: " & Format(RSalb!codClien, "000000") & " - " & RSalb!NomClien
            Else
                LblBar.Caption = RSalb!Numalbar
            End If
            LblBar.Refresh
        Else
            'mismo cliente. Comprobemos la fecha
            If FecEnvio <> RSalb!FecEnvio Then
                'Otra fecha de envio. Comprobemos portes hasta aqui
                If ClienConPortes Then TratarLineaPortes Comprobar, cadW, Sql, DatosPortes, Codclien1
            
                FecEnvio = RSalb!FecEnvio
                
                cadW = " (slialb.codtipom='ALV' AND slialb.numalbar IN (" & RSalb!Numalbar
                
            Else
                'Codclien y fechaenvio la misma
                'Todos igual. Metemos al select el albaran
                cadW = cadW & "," & RSalb!Numalbar
            End If
        
        End If
            

        RSalb.MoveNext
        If Timer - T1 > 4 Then
            Me.Refresh
            DoEvents
            T1 = Timer
        End If
        
    Wend
    RSalb.Close
    
        
    'comprobar la ultima Factura generada del blucle
    If cadW <> "" Then
        
        
                   
        If ClienConPortes Then
            LblBar.Caption = Sql
            TratarLineaPortes Comprobar, cadW, Sql, DatosPortes, Codclien1
        End If
       
        Espera 0.1
    End If
    
    If Comprobar Then
        Sql = "Select * from tmpsliped where codusu = " & vUsu.Codigo & " ORDER BY ampliaci,nomartic"
        RSalb.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Sql = ""
        While Not RSalb.EOF
            Sql = Sql & "1"
            RSalb.MoveNext
        Wend
        RSalb.Close
        If Sql <> "" Then
            CadenaDesdeOtroForm = ""
            frmListado3.Opcion = 2
            frmListado3.Show vbModal
            If CadenaDesdeOtroForm <> "" Then CargaPorteTipoHerbelca = True
            
        Else
            CargaPorteTipoHerbelca = True
        End If
        
    Else
        'Si llega aqui es qu todo ha ido benne
        CargaPorteTipoHerbelca = True
    
    End If
    Espera 0.2
    
    
ETraspasoAlbFac:
    If Err.Number <> 0 Then MuestraError Err.Number, "Portes TIPO 2", Err.Description
    Set RSalb = Nothing
    
End Function


'DatosDeArticPortes:  nomartic|preciove|preciouc|  los precios sin comas
Private Function TratarLineaPortes(Comprobar As Boolean, CadWh As String, NomClien As String, DatosDeArticPortes As String, CodigoCliente As Long) As Boolean
Dim Aux As String
Dim RN As ADODB.Recordset
Dim J As Integer
Dim Llega As Boolean
Dim ImporteT As Currency

    'Comprobamos que no tiene una linea de portes
    TratarLineaPortes = False
    Set RN = New ADODB.Recordset
    
    
    'Seimpre comprobaremos
    
    Aux = CadWh & ")) AND codartic = " & DBSet(vParamAplic.ArtPortesN, "T")
    Aux = "Select * from slialb WHERE " & Aux
    RN.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Llega = False
    If Not RN.EOF Then
        
        'YA EXISTE UNA LINEA con portes
        If Comprobar Then
            Aux = "insert into `tmpsliped` (`codusu`,`numpedcl`,`numlinea`,`codalmac`,codartic,`nomartic`,codclien)  VALUES ("
            Aux = Aux & vUsu.Codigo & "," & RN!Numalbar & "," & RN!numlinea & "," & RN!codAlmac & "," & DBSet(RN!codtipom, "T")
            Aux = Aux & "," & DBSet(NomClien, "T") & "," & CodigoCliente & ")"
            ejecutar Aux, False
        Else
            Llega = True
        End If
    End If
    RN.Close
    
    If Comprobar Then Exit Function
        
    'Si ya tiene TB nos salimos
    If Llega Then Exit Function
        
        
        'Veamos si llega al minimo exigible
        Aux = "Select sum(importel) from slialb where " & CadWh & "))"
        RN.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Llega = False
        ImporteT = 0
        If Not RN.EOF Then
            If Not RN.EOF Then
                ImporteT = RN.Fields(0)
                If RN.Fields(0) >= vParamAplic.ImporteMinimo Then Llega = True
            End If
        End If
        RN.Close
        
        'OCTUBRE 2011
        'Si el IMPORTE es cero... NO le facturamos PORTES
        If ImporteT = 0 Then Llega = True
        
        'Si no llega
        If Not Llega Then
            'Vere si hay RESTO PEDIDOS
            Aux = Replace(CadWh, "slialb", "scaalb")
            Aux = "select numpedcl from scaalb where " & Aux & ")) group by 1"
            RN.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            Aux = ""
            While Not RN.EOF
                If DBLet(RN.Fields(0), "T") <> "" Then Aux = Aux & ", " & RN.Fields(0)
                RN.MoveNext
            Wend
            RN.Close
            
            
            If Aux <> "" Then
                'VERE LOS PEDIDOS pendientes, si es k lo tien
                Aux = Mid(Aux, 2)
                Aux = Trim(Aux)
                Aux = "select sum(importel) from sliped where numpedcl in (" & Aux & ")"
                RN.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If Not RN.EOF Then
                    If Not IsNull(RN.Fields(0)) Then ImporteT = ImporteT + RN.Fields(0)
                End If
                RN.Close
                If ImporteT >= vParamAplic.ImporteMinimo Then Llega = True
                
            End If
        End If
        
        If Not Llega Then
                'NO llega al minimo.
                'Sicota, cargar portes
            
                J = InStrRev(CadWh, ",")
                If J = 0 Then
                    'Solo hay un albaran
                    J = InStrRev(CadWh, "(") 'YA NO PUEDE SER CERO
                End If
                Aux = Mid(CadWh, J + 1)
                Aux = "Select numalbar,codalmac,max(numlinea) from slialb where codtipom='ALV' AND numalbar =" & Aux & " GROUP BY 1,2"
                RN.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                'NO PUEDE SER EOF
                Aux = "'ALV'," & RN!Numalbar & "," & DBLet(RN.Fields(2), "N") + 1 & "," & RN!codAlmac & "," & DBSet(vParamAplic.ArtPortesN, "T") & ","
                
                'codtipom,numalbar,numlinea,codalmac,codartic,nomartic,cantidad,numbultos,
                'precioar,dtoline1,dtoline2,importel,origpre,codproveX,
                'En cadparam tengo el precioar y en cadtitulo el nomartic
                Aux = Aux & DBSet(RecuperaValor(DatosDeArticPortes, 1), "T") & ",1,1," & RecuperaValor(DatosDeArticPortes, 2) & ",0,0," & RecuperaValor(DatosDeArticPortes, 2) & ",'A',0)"
                
                Aux = "codtipom,numalbar,numlinea,codalmac,codartic,nomartic,cantidad,numbultos,precioar,dtoline1,dtoline2,importel,origpre,codproveX) VALUES  (" & Aux
                Aux = "INSERT INTO slialb (" & Aux
                
                'Si da error que salga
                conn.Execute Aux
        End If
    
    Set RN = Nothing
    TratarLineaPortes = True
End Function

Private Sub CargaComboTipoCliente()
    CargarCombo_Tabla Me.cboTipocliente, "stipclien", "tipclien", "descclien"
    'Metemos una linea que sea "todos"
    cboTipocliente.AddItem "Todos"
    cboTipocliente.ItemData(cboTipocliente.NewIndex) = "-1"
    cboTipocliente.ListIndex = cboTipocliente.NewIndex
End Sub

Private Function DevuelveClientesPedidosPorTipo() As String
Dim Aux As String
On Error GoTo EDevuelveClientesPedidosPorTipo
    DevuelveClientesPedidosPorTipo = ""
    
    Set miRsAux = New ADODB.Recordset
    Aux = "Select distinct(scaped.codclien) FROM scaped ,sliped ,sclien WHERE scaped.numpedcl=sliped.numpedcl"
    Aux = Aux & " AND scaped.codclien=sclien.codclien AND tipclien = " & Me.cboTipocliente.ItemData(cboTipocliente.ListIndex)
    If cadSelect <> "" Then
        Aux = Aux & " AND " & cadSelect
        Aux = Replace(Aux, "{", "(")
        Aux = Replace(Aux, "}", ")")
    End If
    
    miRsAux.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Aux = ""
    While Not miRsAux.EOF
        Aux = Aux & ", " & miRsAux!codClien
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    If Aux <> "" Then Aux = Mid(Aux, 2) 'primera coma
    DevuelveClientesPedidosPorTipo = Aux
EDevuelveClientesPedidosPorTipo:
    If Err.Number <> 0 Then MuestraError Err.Number, "Obteniedo cliente por tipo"
    Set miRsAux = Nothing
End Function



Private Sub CargarComboFacturacion()
'### Combo Tipo Facturación
'Cargaremos el combo, o bien desde una tabla o con valores fijos o como
'se quiera, la cuestion es cargarlo
' El estilo del combo debe de ser 2 - Dropdown List
' Si queremos que este ordenado, o lo ordenamos por la sentencia sql
' o marcamos la opcion sorted del combo
'0-Factura Colectiva, 1-Factura x Albaran

    cboFacturacion.Clear
    cboFacturacion.AddItem "Factura Colectiva"
    cboFacturacion.ItemData(cboFacturacion.NewIndex) = 0

    cboFacturacion.AddItem "Factura x Albaran"
    cboFacturacion.ItemData(cboFacturacion.NewIndex) = 1


End Sub


Private Sub CargaListTipoFacturas()
Dim i As Integer
Dim Rectifica As Byte '0. Sin cargar  1.- Cargada
Dim Marcar As Boolean
    
    'Cosillas "a mano"
    'Cuales vienen marcadas por defecto
    Titulo = "FAV|FRT|"
    
    
    If InstalacionEsEulerTaxco Then
        'Leeremos los seleccionad
       
        LeeGuardaListFacturas True
        If ListTipoFact.Tag <> "" Then Titulo = ListTipoFact.Tag
    Else
        'Herbelca
        'If vParamAplic.AlmacenB > 90 Then
        If vParamAplic.NumeroInstalacion = vbHerbelca Then
            'MOstraremos tb las de mostrador
            Titulo = Titulo & "FMO|FTI|"
        Else
            'If vParamAplic.Frecuencias Then
            If vParamAplic.NumeroInstalacion = 3 Then
                'MANTENI  Y REPARACION
                Titulo = Titulo & "FAM|FAR|"
            Else
                
                'Si tiene tienda
                If Not vParamTPV Is Nothing Then
                    Titulo = Titulo & "FTI|FAI|"
                Else
                    Titulo = Titulo & "FAM|FAS|"
                End If
            End If
        End If
    End If
    
    Titulo = "SELECT codtipom,nomtipom,if(instr('" & Titulo & "',codtipom)>0,1,0) marcar"
    Titulo = Titulo & " , if (codtipom='FRT',1,0) fuerzaorden"
    Titulo = Titulo & " FROM stipom WHERE codtipom LIKE 'F%' ORDER BY fuerzaorden,marcar desc,1"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open Titulo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        
        ListTipoFact.AddItem miRsAux!codtipom & " - " & miRsAux!nomtipom
        Marcar = False
        If miRsAux!Marcar = 1 Then Marcar = True
       
        
        If Marcar Then ListTipoFact.Selected(i) = True
        
        i = i + 1
            
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    Set miRsAux = Nothing
    
    
    Titulo = ""
End Sub

'Private Sub txtCSB_LostFocus(Index As Integer)
'    If Index = 0 And Me.txtCSB(0).Text = "" Then PonerFocoBtn Me.cmdAceptarFac
'End Sub



Private Function LeeGuardaListFacturas(Leer As Boolean) As String
Dim NF As Integer
Dim NombreFich As String

    On Error GoTo eLeeGuardaListFacturas

    NombreFich = vUsu.Codigo Mod 1000
    NombreFich = Format(NombreFich, "0000")
    NombreFich = App.Path & "\" & NombreFich & "FiltrEst.xdf"
    
    If Leer Then
       'Lo guardara en el TAG
       'ListTipoFact
       ListTipoFact.Tag = ""
       If Dir(NombreFich, vbArchive) <> "" Then
            NF = FreeFile
            Open NombreFich For Input As NF
            NombreFich = ""
            Line Input #NF, NombreFich
            Close #NF
            ListTipoFact.Tag = NombreFich
       End If
    Else
        cadParam = ""
        For indCodigo = 0 To ListTipoFact.ListCount - 1
            If ListTipoFact.Selected(indCodigo) Then
                Titulo = Trim(Mid(ListTipoFact.List(indCodigo), 1, InStr(1, ListTipoFact.List(indCodigo), "-") - 1))
                cadParam = cadParam & Titulo & "|"
            End If
            
        Next
        If cadParam <> "" Then
            If cadParam <> ListTipoFact.Tag Then
                NF = FreeFile
                Open NombreFich For Output As NF
                Print #NF, cadParam
                Close #NF
            End If
        End If
    End If
    

    Exit Function
    
eLeeGuardaListFacturas:
    MuestraError Err.Number, Err.Description
End Function




