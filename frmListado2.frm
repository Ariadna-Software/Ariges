VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmListado2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "L"
   ClientHeight    =   10500
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14025
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10500
   ScaleWidth      =   14025
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameBenxAge2 
      Height          =   10215
      Left            =   720
      TabIndex        =   630
      Top             =   0
      Visible         =   0   'False
      Width           =   6255
      Begin VB.TextBox txtRuta 
         Height          =   285
         Index           =   1
         Left            =   1320
         TabIndex        =   643
         Text            =   "Text1"
         Top             =   6840
         Width           =   735
      End
      Begin VB.TextBox txtDescRuta 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   1
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   953
         Text            =   "Text1"
         Top             =   6840
         Width           =   3375
      End
      Begin VB.TextBox txtRuta 
         Height          =   285
         Index           =   0
         Left            =   1320
         TabIndex        =   642
         Text            =   "Text1"
         Top             =   6360
         Width           =   735
      End
      Begin VB.TextBox txtDescRuta 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   0
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   951
         Text            =   "Text1"
         Top             =   6360
         Width           =   3375
      End
      Begin VB.TextBox txtDescZona 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   4
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   946
         Text            =   "Text1"
         Top             =   5280
         Width           =   3375
      End
      Begin VB.TextBox txtZona 
         Height          =   285
         Index           =   4
         Left            =   1320
         TabIndex        =   640
         Text            =   "Text1"
         Top             =   5280
         Width           =   735
      End
      Begin VB.TextBox txtDescZona 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   5
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   945
         Text            =   "Text1"
         Top             =   5640
         Width           =   3375
      End
      Begin VB.TextBox txtZona 
         Height          =   285
         Index           =   5
         Left            =   1320
         TabIndex        =   641
         Text            =   "Text1"
         Top             =   5640
         Width           =   735
      End
      Begin VB.CheckBox chkBenAge 
         Caption         =   "Aplica descuento"
         Height          =   255
         Index           =   9
         Left            =   2400
         TabIndex        =   796
         Top             =   9120
         Width           =   2055
      End
      Begin VB.CheckBox chkBenAge 
         Caption         =   "Totalizar"
         Height          =   255
         Index           =   8
         Left            =   4560
         TabIndex        =   797
         Top             =   9120
         Width           =   1455
      End
      Begin VB.CheckBox chkBenAge 
         Caption         =   "Quitar proveedores"
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   795
         Top             =   9120
         Width           =   2175
      End
      Begin VB.CheckBox chkBenAge 
         Caption         =   "Presu."
         Height          =   255
         Index           =   6
         Left            =   5040
         TabIndex        =   653
         Top             =   8640
         Width           =   975
      End
      Begin VB.ComboBox cboCoste 
         Height          =   315
         Index           =   0
         ItemData        =   "frmListado2.frx":0000
         Left            =   4320
         List            =   "frmListado2.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   648
         Top             =   7560
         Width           =   1695
      End
      Begin VB.TextBox txtimporte 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   4800
         TabIndex        =   649
         Text            =   "Text1"
         Top             =   8040
         Width           =   1095
      End
      Begin VB.TextBox txtAlma 
         Height          =   285
         Index           =   4
         Left            =   1320
         TabIndex        =   637
         Text            =   "Text1"
         Top             =   3480
         Width           =   975
      End
      Begin VB.TextBox txtDescAlma 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   4
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   721
         Text            =   "Text5"
         Top             =   3480
         Width           =   3615
      End
      Begin VB.TextBox txtAlma 
         Height          =   285
         Index           =   3
         Left            =   1320
         TabIndex        =   636
         Text            =   "Text1"
         Top             =   3120
         Width           =   975
      End
      Begin VB.TextBox txtDescAlma 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   3
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   718
         Text            =   "Text5"
         Top             =   3120
         Width           =   3615
      End
      Begin VB.CheckBox chkBenAge 
         Caption         =   "Comparativo"
         Height          =   255
         Index           =   2
         Left            =   3600
         TabIndex        =   652
         Top             =   8640
         Width           =   1335
      End
      Begin VB.TextBox txtAnyo 
         Height          =   285
         Index           =   3
         Left            =   2760
         TabIndex        =   647
         Text            =   "Text1"
         Top             =   8040
         Width           =   735
      End
      Begin VB.TextBox txtAnyo 
         Height          =   285
         Index           =   2
         Left            =   1320
         TabIndex        =   646
         Text            =   "Text1"
         Top             =   8040
         Width           =   735
      End
      Begin VB.CheckBox chkBenAge 
         Caption         =   "detalla articulo"
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   651
         Top             =   8640
         Width           =   1575
      End
      Begin VB.CheckBox chkBenAge 
         Caption         =   "detalla familia"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   650
         Top             =   8640
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.CommandButton cmdBeneficioAgente 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3720
         TabIndex        =   654
         Top             =   9600
         Width           =   975
      End
      Begin VB.TextBox txtAnyo 
         Height          =   285
         Index           =   1
         Left            =   2760
         TabIndex        =   645
         Text            =   "Text1"
         Top             =   7560
         Width           =   735
      End
      Begin VB.TextBox txtAnyo 
         Height          =   285
         Index           =   0
         Left            =   1320
         TabIndex        =   644
         Text            =   "Text1"
         Top             =   7560
         Width           =   735
      End
      Begin VB.TextBox txtFamia 
         Height          =   285
         Index           =   11
         Left            =   1320
         TabIndex        =   639
         Text            =   "Text1"
         Top             =   4560
         Width           =   1095
      End
      Begin VB.TextBox txtDescFamia 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   11
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   667
         Text            =   "Text5"
         Top             =   4560
         Width           =   3495
      End
      Begin VB.TextBox txtFamia 
         Height          =   285
         Index           =   10
         Left            =   1320
         TabIndex        =   638
         Text            =   "Text1"
         Top             =   4200
         Width           =   1095
      End
      Begin VB.TextBox txtDescFamia 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   10
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   666
         Text            =   "Text5"
         Top             =   4200
         Width           =   3495
      End
      Begin VB.TextBox txtCodProve 
         Height          =   285
         Index           =   22
         Left            =   1320
         TabIndex        =   635
         Top             =   2400
         Width           =   1095
      End
      Begin VB.TextBox txtDescProve 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   22
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   664
         Text            =   "Text5"
         Top             =   2400
         Width           =   3495
      End
      Begin VB.TextBox txtCodProve 
         Height          =   285
         Index           =   21
         Left            =   1320
         TabIndex        =   634
         Top             =   2040
         Width           =   1095
      End
      Begin VB.TextBox txtDescProve 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   21
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   661
         Text            =   "Text5"
         Top             =   2040
         Width           =   3495
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   37
         Left            =   4920
         TabIndex        =   655
         Top             =   9600
         Width           =   1095
      End
      Begin VB.TextBox txtDescAgente 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   7
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   659
         Text            =   "Text1"
         Top             =   1320
         Width           =   3495
      End
      Begin VB.TextBox txtAgente 
         Height          =   285
         Index           =   7
         Left            =   1320
         TabIndex        =   633
         Text            =   "Text1"
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox txtDescAgente 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   6
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   656
         Text            =   "Text1"
         Top             =   960
         Width           =   3495
      End
      Begin VB.TextBox txtAgente 
         Height          =   285
         Index           =   6
         Left            =   1320
         TabIndex        =   632
         Text            =   "Text1"
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   198
         Left            =   360
         TabIndex        =   954
         Top             =   6840
         Width           =   465
      End
      Begin VB.Image imgRuta 
         Height          =   240
         Index           =   1
         Left            =   1080
         Picture         =   "frmListado2.frx":003D
         Top             =   6840
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   197
         Left            =   360
         TabIndex        =   952
         Top             =   6360
         Width           =   465
      End
      Begin VB.Image imgRuta 
         Height          =   240
         Index           =   0
         Left            =   1080
         Picture         =   "frmListado2.frx":013F
         Top             =   6360
         Width           =   240
      End
      Begin VB.Label lblDpto 
         Caption         =   "r"
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
         Left            =   120
         TabIndex        =   950
         Top             =   6120
         Width           =   1515
      End
      Begin VB.Image imgZona 
         Height          =   240
         Index           =   4
         Left            =   1080
         Picture         =   "frmListado2.frx":0241
         Top             =   5280
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   196
         Left            =   360
         TabIndex        =   949
         Top             =   5280
         Width           =   465
      End
      Begin VB.Label lblDpto 
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
         Index           =   71
         Left            =   120
         TabIndex        =   948
         Top             =   5040
         Width           =   420
      End
      Begin VB.Image imgZona 
         Height          =   240
         Index           =   5
         Left            =   1080
         Picture         =   "frmListado2.frx":0343
         Top             =   5640
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   195
         Left            =   360
         TabIndex        =   947
         Top             =   5640
         Width           =   465
      End
      Begin VB.Image imgayuda 
         Height          =   255
         Index           =   4
         Left            =   3360
         ToolTipText     =   "Listado beneficio"
         Top             =   9720
         Width           =   255
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Coste"
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
         Index           =   47
         Left            =   3720
         TabIndex        =   765
         Top             =   7560
         Width           =   480
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   147
         Left            =   360
         TabIndex        =   724
         Top             =   9720
         Width           =   2985
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "€ mín. prov."
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
         Index           =   43
         Left            =   3720
         TabIndex        =   723
         Top             =   8040
         Width           =   990
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   146
         Left            =   360
         TabIndex        =   722
         Top             =   3480
         Width           =   465
      End
      Begin VB.Image imgAlma 
         Height          =   240
         Index           =   4
         Left            =   1080
         Picture         =   "frmListado2.frx":0445
         Top             =   3480
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   145
         Left            =   360
         TabIndex        =   720
         Top             =   3120
         Width           =   465
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Almacén"
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
         Index           =   42
         Left            =   120
         TabIndex        =   719
         Top             =   2880
         Width           =   735
      End
      Begin VB.Image imgAlma 
         Height          =   240
         Index           =   3
         Left            =   1080
         Picture         =   "frmListado2.frx":0547
         Top             =   3120
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   136
         Left            =   2280
         TabIndex        =   699
         Top             =   8040
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   135
         Left            =   720
         TabIndex        =   698
         Top             =   8040
         Width           =   465
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Mes"
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
         TabIndex        =   697
         Top             =   8040
         Width           =   345
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
         Index           =   36
         Left            =   240
         TabIndex        =   673
         Top             =   7560
         Width           =   330
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   130
         Left            =   720
         TabIndex        =   672
         Top             =   7560
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   129
         Left            =   2280
         TabIndex        =   671
         Top             =   7560
         Width           =   465
      End
      Begin VB.Image imgFamilia 
         Height          =   240
         Index           =   11
         Left            =   1080
         Picture         =   "frmListado2.frx":0649
         Top             =   4560
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   128
         Left            =   360
         TabIndex        =   670
         Top             =   4200
         Width           =   465
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
         Index           =   35
         Left            =   120
         TabIndex        =   669
         Top             =   3960
         Width           =   600
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   127
         Left            =   360
         TabIndex        =   668
         Top             =   4560
         Width           =   465
      End
      Begin VB.Image imgFamilia 
         Height          =   240
         Index           =   10
         Left            =   1080
         Picture         =   "frmListado2.frx":074B
         Top             =   4200
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   126
         Left            =   360
         TabIndex        =   665
         Top             =   2400
         Width           =   465
      End
      Begin VB.Image imgProveedor 
         Height          =   240
         Index           =   22
         Left            =   1080
         Picture         =   "frmListado2.frx":084D
         Top             =   2400
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   125
         Left            =   360
         TabIndex        =   663
         Top             =   2040
         Width           =   465
      End
      Begin VB.Image imgProveedor 
         Height          =   240
         Index           =   21
         Left            =   1080
         Picture         =   "frmListado2.frx":094F
         Top             =   2040
         Width           =   240
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
         Index           =   34
         Left            =   120
         TabIndex        =   662
         Top             =   1800
         Width           =   885
      End
      Begin VB.Image imgAgente 
         Height          =   240
         Index           =   7
         Left            =   1080
         Picture         =   "frmListado2.frx":0A51
         Top             =   1320
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   124
         Left            =   360
         TabIndex        =   660
         Top             =   1320
         Width           =   465
      End
      Begin VB.Image imgAgente 
         Height          =   240
         Index           =   6
         Left            =   1080
         Picture         =   "frmListado2.frx":0B53
         Top             =   960
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   123
         Left            =   360
         TabIndex        =   658
         Top             =   960
         Width           =   465
      End
      Begin VB.Label lblDpto 
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
         Index           =   46
         Left            =   120
         TabIndex        =   657
         Top             =   720
         Width           =   615
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
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
         Index           =   34
         Left            =   1200
         TabIndex        =   631
         Top             =   240
         Width           =   3975
      End
   End
   Begin VB.Frame FramePedxZon 
      Height          =   5415
      Left            =   1680
      TabIndex        =   393
      Top             =   0
      Visible         =   0   'False
      Width           =   6015
      Begin VB.CheckBox chkPedxZona 
         Caption         =   "Solo articulos con stock"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   418
         Top             =   4800
         Width           =   2055
      End
      Begin VB.CommandButton cmdPedxZona 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3360
         TabIndex        =   403
         Top             =   4800
         Width           =   1095
      End
      Begin VB.CheckBox chkPedxZona 
         Caption         =   "Pedidos con departamento"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   402
         Top             =   4440
         Width           =   2655
      End
      Begin VB.CheckBox chkPedxZona 
         Caption         =   "Clientes varios"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   401
         Top             =   4080
         Width           =   2055
      End
      Begin VB.TextBox txtZona 
         Height          =   285
         Index           =   1
         Left            =   1680
         TabIndex        =   400
         Text            =   "Text1"
         Top             =   3600
         Width           =   615
      End
      Begin VB.TextBox txtDescZona 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   1
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   416
         Text            =   "Text1"
         Top             =   3600
         Width           =   3375
      End
      Begin VB.TextBox txtZona 
         Height          =   285
         Index           =   0
         Left            =   1680
         TabIndex        =   399
         Text            =   "Text1"
         Top             =   3240
         Width           =   615
      End
      Begin VB.TextBox txtDescZona 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   0
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   413
         Text            =   "Text1"
         Top             =   3240
         Width           =   3375
      End
      Begin VB.TextBox txtAgente 
         Height          =   285
         Index           =   3
         Left            =   1680
         TabIndex        =   398
         Text            =   "Text1"
         Top             =   2520
         Width           =   615
      End
      Begin VB.TextBox txtDescAgente 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   3
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   411
         Text            =   "Text1"
         Top             =   2520
         Width           =   3375
      End
      Begin VB.TextBox txtAgente 
         Height          =   285
         Index           =   2
         Left            =   1680
         TabIndex        =   397
         Text            =   "Text1"
         Top             =   2160
         Width           =   615
      End
      Begin VB.TextBox txtDescAgente 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   2
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   408
         Text            =   "Text1"
         Top             =   2160
         Width           =   3375
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   33
         Left            =   4200
         TabIndex        =   396
         Text            =   "Text1"
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   32
         Left            =   1320
         TabIndex        =   395
         Text            =   "Text1"
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   26
         Left            =   4680
         TabIndex        =   404
         Top             =   4800
         Width           =   1095
      End
      Begin VB.Image imgayuda 
         Height          =   255
         Index           =   3
         Left            =   2760
         ToolTipText     =   "Stock pedido por zona"
         Top             =   4800
         Width           =   255
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   82
         Left            =   600
         TabIndex        =   417
         Top             =   3600
         Width           =   465
      End
      Begin VB.Image imgZona 
         Height          =   240
         Index           =   1
         Left            =   1320
         Picture         =   "frmListado2.frx":0C55
         Top             =   3600
         Width           =   240
      End
      Begin VB.Label lblDpto 
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
         Index           =   37
         Left            =   240
         TabIndex        =   415
         Top             =   3000
         Width           =   420
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   81
         Left            =   600
         TabIndex        =   414
         Top             =   3240
         Width           =   465
      End
      Begin VB.Image imgZona 
         Height          =   240
         Index           =   0
         Left            =   1320
         Picture         =   "frmListado2.frx":0D57
         Top             =   3240
         Width           =   240
      End
      Begin VB.Image imgAgente 
         Height          =   240
         Index           =   3
         Left            =   1320
         Picture         =   "frmListado2.frx":0E59
         Top             =   2520
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   80
         Left            =   600
         TabIndex        =   412
         Top             =   2520
         Width           =   465
      End
      Begin VB.Label lblDpto 
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
         Index           =   36
         Left            =   240
         TabIndex        =   410
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   79
         Left            =   600
         TabIndex        =   409
         Top             =   2160
         Width           =   465
      End
      Begin VB.Image imgAgente 
         Height          =   240
         Index           =   2
         Left            =   1320
         Picture         =   "frmListado2.frx":0F5B
         Top             =   2160
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   78
         Left            =   3360
         TabIndex        =   407
         Top             =   1245
         Width           =   465
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   33
         Left            =   3840
         Picture         =   "frmListado2.frx":105D
         Top             =   1215
         Width           =   240
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Fecha entrega"
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
         Index           =   35
         Left            =   240
         TabIndex        =   406
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   77
         Left            =   480
         TabIndex        =   405
         Top             =   1245
         Width           =   465
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   32
         Left            =   960
         Picture         =   "frmListado2.frx":10E8
         Top             =   1215
         Width           =   240
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Impresión pedidos por zona"
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
         Index           =   24
         Left            =   240
         TabIndex        =   394
         Top             =   240
         Width           =   5415
      End
   End
   Begin VB.Frame FramePropPedido 
      Height          =   7695
      Left            =   4680
      TabIndex        =   488
      Top             =   120
      Visible         =   0   'False
      Width           =   6735
      Begin VB.TextBox txtDescAlma 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   8
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   770
         Text            =   "Text5"
         Top             =   1920
         Width           =   3495
      End
      Begin VB.TextBox txtAlma 
         Height          =   285
         Index           =   8
         Left            =   1800
         TabIndex        =   493
         Text            =   "Text1"
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox txtAlma 
         Height          =   285
         Index           =   7
         Left            =   1800
         TabIndex        =   492
         Text            =   "Text1"
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox txtDescAlma 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   7
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   768
         Text            =   "Text5"
         Top             =   1560
         Width           =   3495
      End
      Begin VB.CheckBox chkPropPedido 
         Caption         =   "Mostrar referencias con texto auxiliar documentos "
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   505
         Top             =   6840
         Width           =   4215
      End
      Begin VB.TextBox txtAnyo 
         Height          =   285
         Index           =   5
         Left            =   5880
         TabIndex        =   502
         Text            =   "Text1"
         Top             =   6000
         Width           =   615
      End
      Begin VB.TextBox txtAnyo 
         Height          =   285
         Index           =   4
         Left            =   3120
         TabIndex        =   501
         Text            =   "Text1"
         Top             =   6000
         Width           =   735
      End
      Begin VB.CheckBox chkPropPedido 
         Caption         =   "Consumo  con departamentos"
         Height          =   195
         Index           =   1
         Left            =   3960
         TabIndex        =   504
         Top             =   6480
         Width           =   2535
      End
      Begin VB.CheckBox chkPropPedido 
         Caption         =   "Pedidos con departamentos"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   503
         Top             =   6480
         Width           =   2415
      End
      Begin VB.ComboBox cboProPed 
         Height          =   315
         Index           =   1
         ItemData        =   "frmListado2.frx":1173
         Left            =   4920
         List            =   "frmListado2.frx":1186
         Style           =   2  'Dropdown List
         TabIndex        =   500
         Top             =   5400
         Width           =   1575
      End
      Begin VB.ComboBox cboProPed 
         Height          =   315
         Index           =   0
         ItemData        =   "frmListado2.frx":11B8
         Left            =   1560
         List            =   "frmListado2.frx":11C2
         Style           =   2  'Dropdown List
         TabIndex        =   499
         Top             =   5400
         Width           =   1335
      End
      Begin VB.TextBox txtDescmarca 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   1
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   519
         Text            =   "Text5"
         Top             =   4680
         Width           =   3495
      End
      Begin VB.TextBox txtmarca 
         Height          =   285
         Index           =   1
         Left            =   1800
         TabIndex        =   498
         Text            =   "Text1"
         Top             =   4680
         Width           =   1095
      End
      Begin VB.TextBox txtDescmarca 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   0
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   516
         Text            =   "Text5"
         Top             =   4320
         Width           =   3495
      End
      Begin VB.TextBox txtmarca 
         Height          =   285
         Index           =   0
         Left            =   1800
         TabIndex        =   497
         Text            =   "Text1"
         Top             =   4320
         Width           =   1095
      End
      Begin VB.TextBox txtDescAlma 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   0
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   515
         Text            =   "Text5"
         Top             =   840
         Width           =   3495
      End
      Begin VB.TextBox txtAlma 
         Height          =   285
         Index           =   0
         Left            =   1800
         TabIndex        =   491
         Text            =   "Text1"
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtFamia 
         Height          =   285
         Index           =   3
         Left            =   1800
         TabIndex        =   496
         Text            =   "Text1"
         Top             =   3600
         Width           =   1095
      End
      Begin VB.TextBox txtDescFamia 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   3
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   512
         Text            =   "Text5"
         Top             =   3600
         Width           =   3495
      End
      Begin VB.TextBox txtFamia 
         Height          =   285
         Index           =   2
         Left            =   1800
         TabIndex        =   495
         Text            =   "Text1"
         Top             =   3240
         Width           =   1095
      End
      Begin VB.TextBox txtDescFamia 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   2
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   509
         Text            =   "Text5"
         Top             =   3240
         Width           =   3495
      End
      Begin VB.CommandButton cmdPropuestaPedido 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4080
         TabIndex        =   506
         Top             =   7200
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   32
         Left            =   5400
         TabIndex        =   507
         Top             =   7200
         Width           =   1095
      End
      Begin VB.TextBox txtCodProve 
         Height          =   285
         Index           =   17
         Left            =   1800
         TabIndex        =   494
         Top             =   2520
         Width           =   1095
      End
      Begin VB.TextBox txtDescProve 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   17
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   490
         Text            =   "Text5"
         Top             =   2520
         Width           =   3495
      End
      Begin VB.Image imgAlma 
         Height          =   240
         Index           =   8
         Left            =   1560
         Picture         =   "frmListado2.frx":11DC
         Top             =   1920
         Width           =   240
      End
      Begin VB.Image imgAlma 
         Height          =   240
         Index           =   7
         Left            =   1560
         Picture         =   "frmListado2.frx":12DE
         Top             =   1560
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Consolidar con almacén"
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
         Index           =   49
         Left            =   240
         TabIndex        =   767
         Top             =   1320
         Width           =   1995
      End
      Begin VB.Label Label3 
         Caption         =   "Porcentaje mismo cliente"
         Height          =   195
         Index           =   157
         Left            =   3960
         TabIndex        =   764
         Top             =   6000
         Width           =   1905
      End
      Begin VB.Image imgayuda 
         Height          =   255
         Index           =   2
         Left            =   3000
         Top             =   5400
         Width           =   255
      End
      Begin VB.Label Label3 
         Caption         =   "Mínimo albaranes sin indicar proveedor"
         Height          =   195
         Index           =   137
         Left            =   240
         TabIndex        =   700
         Top             =   6000
         Width           =   2985
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   100
         Left            =   240
         TabIndex        =   523
         Top             =   7320
         Width           =   3465
      End
      Begin VB.Label Label4 
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
         Index           =   22
         Left            =   3840
         TabIndex        =   522
         Top             =   5400
         Width           =   780
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Rotación"
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
         Index           =   21
         Left            =   240
         TabIndex        =   521
         Top             =   5400
         Width           =   750
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   99
         Left            =   840
         TabIndex        =   520
         Top             =   4680
         Width           =   465
      End
      Begin VB.Image imgMarca 
         Height          =   240
         Index           =   1
         Left            =   1560
         Picture         =   "frmListado2.frx":13E0
         Top             =   4680
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   98
         Left            =   840
         TabIndex        =   518
         Top             =   4320
         Width           =   465
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Marca"
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
         Index           =   20
         Left            =   240
         TabIndex        =   517
         Top             =   4080
         Width           =   525
      End
      Begin VB.Image imgMarca 
         Height          =   240
         Index           =   0
         Left            =   1560
         Picture         =   "frmListado2.frx":14E2
         Top             =   4320
         Width           =   240
      End
      Begin VB.Image imgAlma 
         Height          =   240
         Index           =   0
         Left            =   1560
         Picture         =   "frmListado2.frx":15E4
         Top             =   840
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Almacén"
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
         Left            =   240
         TabIndex        =   514
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   97
         Left            =   840
         TabIndex        =   513
         Top             =   3600
         Width           =   465
      End
      Begin VB.Image imgFamilia 
         Height          =   240
         Index           =   3
         Left            =   1560
         Picture         =   "frmListado2.frx":16E6
         Top             =   3600
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
         TabIndex        =   511
         Top             =   3000
         Width           =   600
      End
      Begin VB.Image imgFamilia 
         Height          =   240
         Index           =   2
         Left            =   1560
         Picture         =   "frmListado2.frx":17E8
         Top             =   3240
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   96
         Left            =   840
         TabIndex        =   510
         Top             =   3240
         Width           =   465
      End
      Begin VB.Image imgProveedor 
         Height          =   240
         Index           =   17
         Left            =   1560
         Picture         =   "frmListado2.frx":18EA
         Top             =   2520
         Width           =   240
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
         Index           =   17
         Left            =   240
         TabIndex        =   508
         Top             =   2520
         Width           =   885
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Propuesta de pedido"
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
         Index           =   29
         Left            =   1680
         TabIndex        =   489
         Top             =   240
         Width           =   3525
      End
   End
   Begin VB.Frame FrameCliPot 
      Height          =   5175
      Left            =   5160
      TabIndex        =   850
      Top             =   240
      Visible         =   0   'False
      Width           =   6255
      Begin VB.TextBox txtNumero 
         Height          =   285
         Index           =   3
         Left            =   1800
         TabIndex        =   854
         Text            =   "Text1"
         Top             =   3720
         Width           =   1215
      End
      Begin VB.TextBox txtNumero 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   2
         Left            =   1800
         TabIndex        =   853
         Text            =   "Text1"
         Top             =   3120
         Width           =   975
      End
      Begin VB.CommandButton cmdCrearCliente 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3720
         TabIndex        =   855
         Top             =   4560
         Width           =   975
      End
      Begin VB.TextBox txtTextoNoEditable 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   0
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   863
         Text            =   "Text1"
         Top             =   1080
         Width           =   3375
      End
      Begin VB.TextBox txtForpa 
         Height          =   285
         Index           =   3
         Left            =   1800
         TabIndex        =   852
         Text            =   "Text1"
         Top             =   2280
         Width           =   615
      End
      Begin VB.TextBox txtDescForpa 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   3
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   861
         Text            =   "Text1"
         Top             =   2280
         Width           =   3375
      End
      Begin VB.TextBox txtAgente 
         Height          =   285
         Index           =   10
         Left            =   1800
         TabIndex        =   851
         Text            =   "Text1"
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox txtDescAgente 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   10
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   859
         Text            =   "Text1"
         Top             =   1680
         Width           =   3375
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   47
         Left            =   4920
         TabIndex        =   856
         Top             =   4560
         Width           =   1095
      End
      Begin VB.Image imgIdClienteLibre 
         Height          =   240
         Left            =   2880
         Picture         =   "frmListado2.frx":19EC
         Top             =   3120
         Width           =   240
      End
      Begin VB.Image imgayuda 
         Height          =   240
         Index           =   7
         Left            =   3240
         ToolTipText     =   "Paso a cliente desde potenciales"
         Top             =   4680
         Width           =   360
      End
      Begin VB.Label Label8 
         Caption         =   "Label8"
         Height          =   495
         Left            =   3120
         TabIndex        =   944
         Top             =   3720
         Width           =   2655
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Cta. contabilidad"
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
         Index           =   70
         Left            =   240
         TabIndex        =   943
         Top             =   3720
         Width           =   1410
      End
      Begin VB.Shape Shape1 
         Height          =   1455
         Left            =   120
         Top             =   2880
         Width           =   5775
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Código de cliente"
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
         Index           =   69
         Left            =   240
         TabIndex        =   942
         Top             =   3120
         Width           =   1440
      End
      Begin VB.Image imgForPa 
         Height          =   240
         Index           =   3
         Left            =   1560
         Picture         =   "frmListado2.frx":1AEE
         Top             =   2280
         Width           =   240
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Forma de pago"
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
         Index           =   63
         Left            =   240
         TabIndex        =   862
         Top             =   2280
         Width           =   1260
      End
      Begin VB.Label lblDpto 
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
         Index           =   62
         Left            =   240
         TabIndex        =   860
         Top             =   1680
         Width           =   615
      End
      Begin VB.Image imgAgente 
         Height          =   240
         Index           =   10
         Left            =   1560
         Picture         =   "frmListado2.frx":1BF0
         Top             =   1680
         Width           =   240
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "NOMBRE"
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
         Left            =   240
         TabIndex        =   858
         Top             =   1080
         Width           =   690
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Crear cliente"
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
         Index           =   40
         Left            =   600
         TabIndex        =   857
         Top             =   360
         Width           =   4485
      End
   End
   Begin VB.Frame FrameResvtaAgente 
      Height          =   6135
      Left            =   1800
      TabIndex        =   599
      Top             =   0
      Visible         =   0   'False
      Width           =   6015
      Begin VB.CheckBox chkResVtaAgen 
         Caption         =   "Visitador"
         Height          =   255
         Index           =   4
         Left            =   1800
         TabIndex        =   941
         Top             =   4680
         Width           =   1455
      End
      Begin VB.CheckBox chkResVtaAgen 
         Caption         =   "Fact. Rectificativas"
         Height          =   255
         Index           =   3
         Left            =   3600
         TabIndex        =   612
         Top             =   4200
         Width           =   1935
      End
      Begin VB.CheckBox chkResVtaAgen 
         Caption         =   "Portes"
         Height          =   255
         Index           =   2
         Left            =   1800
         TabIndex        =   611
         Top             =   4200
         Width           =   1455
      End
      Begin VB.CommandButton cmdResVtaAgente 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3480
         TabIndex        =   615
         Top             =   5640
         Width           =   975
      End
      Begin VB.OptionButton optVtaAgen 
         Caption         =   "Marca"
         Height          =   255
         Index           =   1
         Left            =   2280
         TabIndex        =   614
         Top             =   5280
         Width           =   975
      End
      Begin VB.OptionButton optVtaAgen 
         Caption         =   "Agente"
         Height          =   255
         Index           =   0
         Left            =   1320
         TabIndex        =   613
         Top             =   5280
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.CheckBox chkResVtaAgen 
         Caption         =   "Presupuestos"
         Height          =   255
         Index           =   1
         Left            =   3600
         TabIndex        =   610
         Top             =   3720
         Width           =   1455
      End
      Begin VB.CheckBox chkResVtaAgen 
         Caption         =   "Albaranes"
         Height          =   255
         Index           =   0
         Left            =   1800
         TabIndex        =   609
         Top             =   3720
         Width           =   1455
      End
      Begin VB.TextBox txtAgente 
         Height          =   285
         Index           =   5
         Left            =   1800
         TabIndex        =   608
         Text            =   "Text1"
         Top             =   3240
         Width           =   615
      End
      Begin VB.TextBox txtDescAgente 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   5
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   626
         Text            =   "Text1"
         Top             =   3240
         Width           =   3375
      End
      Begin VB.TextBox txtAgente 
         Height          =   285
         Index           =   4
         Left            =   1800
         TabIndex        =   607
         Text            =   "Text1"
         Top             =   2880
         Width           =   615
      End
      Begin VB.TextBox txtDescAgente 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   4
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   623
         Text            =   "Text1"
         Top             =   2880
         Width           =   3375
      End
      Begin VB.TextBox txtmarca 
         Height          =   285
         Index           =   5
         Left            =   1680
         TabIndex        =   606
         Text            =   "Text1"
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox txtDescmarca 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   5
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   621
         Text            =   "Text5"
         Top             =   2040
         Width           =   3375
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   40
         Left            =   4080
         TabIndex        =   604
         Text            =   "Text1"
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   39
         Left            =   1680
         TabIndex        =   603
         Text            =   "Text1"
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   36
         Left            =   4680
         TabIndex        =   616
         Top             =   5640
         Width           =   1095
      End
      Begin VB.TextBox txtmarca 
         Height          =   285
         Index           =   4
         Left            =   1680
         TabIndex        =   605
         Text            =   "Text1"
         Top             =   1680
         Width           =   855
      End
      Begin VB.TextBox txtDescmarca 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   4
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   600
         Text            =   "Text5"
         Top             =   1680
         Width           =   3375
      End
      Begin VB.Image imgayuda 
         Height          =   240
         Index           =   1
         Left            =   1440
         ToolTipText     =   "Ventas por agente"
         Top             =   3720
         Width           =   240
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Ordenado"
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
         Index           =   45
         Left            =   240
         TabIndex        =   629
         Top             =   5280
         Width           =   825
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   122
         Left            =   240
         TabIndex        =   628
         Top             =   5640
         Width           =   3225
      End
      Begin VB.Image imgAgente 
         Height          =   240
         Index           =   5
         Left            =   1440
         Picture         =   "frmListado2.frx":1CF2
         Top             =   3240
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   121
         Left            =   720
         TabIndex        =   627
         Top             =   3240
         Width           =   465
      End
      Begin VB.Label lblDpto 
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
         Index           =   44
         Left            =   120
         TabIndex        =   625
         Top             =   2520
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   120
         Left            =   720
         TabIndex        =   624
         Top             =   2880
         Width           =   465
      End
      Begin VB.Image imgAgente 
         Height          =   240
         Index           =   4
         Left            =   1440
         Picture         =   "frmListado2.frx":1DF4
         Top             =   2880
         Width           =   240
      End
      Begin VB.Image imgMarca 
         Height          =   240
         Index           =   5
         Left            =   1320
         Picture         =   "frmListado2.frx":1EF6
         Top             =   2040
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   119
         Left            =   720
         TabIndex        =   622
         Top             =   2040
         Width           =   465
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
         Index           =   33
         Left            =   120
         TabIndex        =   620
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   118
         Left            =   3120
         TabIndex        =   619
         Top             =   1005
         Width           =   465
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   40
         Left            =   3840
         Picture         =   "frmListado2.frx":1FF8
         Top             =   975
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   115
         Left            =   720
         TabIndex        =   618
         Top             =   1005
         Width           =   465
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   39
         Left            =   1440
         Picture         =   "frmListado2.frx":2083
         Top             =   975
         Width           =   240
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Resumen ventas por agente"
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
         Index           =   33
         Left            =   1200
         TabIndex        =   617
         Top             =   240
         Width           =   3975
      End
      Begin VB.Image imgMarca 
         Height          =   240
         Index           =   4
         Left            =   1320
         Picture         =   "frmListado2.frx":210E
         Top             =   1680
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Marca"
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
         Index           =   31
         Left            =   120
         TabIndex        =   602
         Top             =   1440
         Width           =   525
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   114
         Left            =   720
         TabIndex        =   601
         Top             =   1680
         Width           =   465
      End
   End
   Begin VB.Frame FrameVEntasAgente 
      Height          =   5055
      Left            =   9240
      TabIndex        =   270
      Top             =   2880
      Visible         =   0   'False
      Width           =   6135
      Begin VB.TextBox txtDescForpa 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   2
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   848
         Text            =   "Text1"
         Top             =   3240
         Width           =   3375
      End
      Begin VB.TextBox txtForpa 
         Height          =   285
         Index           =   2
         Left            =   1560
         TabIndex        =   276
         Text            =   "Text1"
         Top             =   3240
         Width           =   615
      End
      Begin VB.TextBox txtDescForpa 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   1
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   845
         Text            =   "Text1"
         Top             =   2880
         Width           =   3375
      End
      Begin VB.TextBox txtForpa 
         Height          =   285
         Index           =   1
         Left            =   1560
         TabIndex        =   275
         Text            =   "Text1"
         Top             =   2880
         Width           =   615
      End
      Begin VB.CheckBox chkAgentes 
         Caption         =   "Rectificativas"
         Height          =   255
         Index           =   2
         Left            =   4080
         TabIndex        =   279
         Top             =   3840
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CheckBox chkAgentes 
         Caption         =   "Presupuestos"
         Height          =   255
         Index           =   1
         Left            =   2160
         TabIndex        =   278
         Top             =   3840
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CheckBox chkAgentes 
         Caption         =   "Facturas"
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   277
         Top             =   3840
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CommandButton cmdAgentes 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3480
         TabIndex        =   280
         Top             =   4440
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   16
         Left            =   4680
         TabIndex        =   281
         Top             =   4440
         Width           =   975
      End
      Begin VB.TextBox txtDescAgente 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   1
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   289
         Text            =   "Text1"
         Top             =   2160
         Width           =   3375
      End
      Begin VB.TextBox txtAgente 
         Height          =   285
         Index           =   1
         Left            =   1560
         TabIndex        =   274
         Text            =   "Text1"
         Top             =   2160
         Width           =   615
      End
      Begin VB.TextBox txtDescAgente 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   0
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   286
         Text            =   "Text1"
         Top             =   1800
         Width           =   3375
      End
      Begin VB.TextBox txtAgente 
         Height          =   285
         Index           =   0
         Left            =   1560
         TabIndex        =   273
         Text            =   "Text1"
         Top             =   1800
         Width           =   615
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   25
         Left            =   4320
         TabIndex        =   272
         Text            =   "Text1"
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   24
         Left            =   1920
         TabIndex        =   271
         Text            =   "Text1"
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   174
         Left            =   480
         TabIndex        =   849
         Top             =   3240
         Width           =   465
      End
      Begin VB.Image imgForPa 
         Height          =   240
         Index           =   2
         Left            =   1200
         Picture         =   "frmListado2.frx":2210
         Top             =   3240
         Width           =   240
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Forma de pago"
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
         Index           =   60
         Left            =   120
         TabIndex        =   847
         Top             =   2520
         Width           =   1260
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   173
         Left            =   480
         TabIndex        =   846
         Top             =   2880
         Width           =   465
      End
      Begin VB.Image imgForPa 
         Height          =   240
         Index           =   1
         Left            =   1200
         Picture         =   "frmListado2.frx":2312
         Top             =   2880
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   57
         Left            =   480
         TabIndex        =   290
         Top             =   2160
         Width           =   465
      End
      Begin VB.Image imgAgente 
         Height          =   240
         Index           =   0
         Left            =   1200
         Picture         =   "frmListado2.frx":2414
         Top             =   1800
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   56
         Left            =   480
         TabIndex        =   288
         Top             =   1800
         Width           =   465
      End
      Begin VB.Label lblDpto 
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
         Index           =   24
         Left            =   120
         TabIndex        =   287
         Top             =   1440
         Width           =   615
      End
      Begin VB.Image imgAgente 
         Height          =   240
         Index           =   1
         Left            =   1200
         Picture         =   "frmListado2.frx":2516
         Top             =   2160
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   25
         Left            =   4080
         Picture         =   "frmListado2.frx":2618
         Top             =   960
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   55
         Left            =   3480
         TabIndex        =   285
         Top             =   1005
         Width           =   465
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Fecha factura"
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
         Index           =   23
         Left            =   120
         TabIndex        =   284
         Top             =   720
         Width           =   1155
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   54
         Left            =   960
         TabIndex        =   283
         Top             =   1005
         Width           =   465
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   24
         Left            =   1680
         Picture         =   "frmListado2.frx":26A3
         Top             =   975
         Width           =   240
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Ventas por agente"
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
         Index           =   14
         Left            =   1320
         TabIndex        =   282
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.Frame FrameHorasTrabajadasEuler 
      Height          =   5655
      Left            =   5040
      TabIndex        =   816
      Top             =   240
      Visible         =   0   'False
      Width           =   6615
      Begin VB.CheckBox chkInformeProd 
         Caption         =   "Albaran venta"
         Height          =   195
         Index           =   4
         Left            =   4920
         TabIndex        =   940
         Top             =   3480
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.OptionButton optInfProd 
         Caption         =   "Nº Trabajo"
         Height          =   195
         Index           =   1
         Left            =   2040
         TabIndex        =   828
         Top             =   4680
         Width           =   1215
      End
      Begin VB.OptionButton optInfProd 
         Caption         =   "Trabajador"
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   827
         Top             =   4680
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.CheckBox chkInformeProd 
         Caption         =   "Producción"
         Height          =   195
         Index           =   3
         Left            =   360
         TabIndex        =   826
         Top             =   3840
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.CheckBox chkInformeProd 
         Caption         =   "Orden trabajo"
         Height          =   195
         Index           =   2
         Left            =   3360
         TabIndex        =   825
         Top             =   3480
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.CheckBox chkInformeProd 
         Caption         =   "Trabajo exterior"
         Height          =   195
         Index           =   1
         Left            =   1680
         TabIndex        =   824
         Top             =   3480
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CheckBox chkInformeProd 
         Caption         =   "Reparación"
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   823
         Top             =   3480
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.TextBox txtNumAlbar 
         Height          =   285
         Index           =   7
         Left            =   4320
         TabIndex        =   822
         Top             =   2760
         Width           =   1095
      End
      Begin VB.TextBox txtNumAlbar 
         Height          =   285
         Index           =   6
         Left            =   1560
         TabIndex        =   821
         Top             =   2760
         Width           =   1095
      End
      Begin VB.TextBox txtTrab 
         Height          =   285
         Index           =   8
         Left            =   1560
         TabIndex        =   820
         Top             =   1880
         Width           =   855
      End
      Begin VB.TextBox txtDescTra 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   8
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   837
         Text            =   "Text1"
         Top             =   1880
         Width           =   3495
      End
      Begin VB.TextBox txtTrab 
         Height          =   285
         Index           =   7
         Left            =   1560
         TabIndex        =   819
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox txtDescTra 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   7
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   835
         Text            =   "Text1"
         Top             =   1560
         Width           =   3495
      End
      Begin VB.CommandButton cmdInformeProductividad 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   4080
         TabIndex        =   829
         Top             =   4920
         Width           =   1095
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   48
         Left            =   4080
         TabIndex        =   818
         Text            =   "Text1"
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   47
         Left            =   1440
         TabIndex        =   817
         Text            =   "Text1"
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   46
         Left            =   5400
         TabIndex        =   830
         Top             =   4920
         Width           =   975
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Ordenado"
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
         Index           =   59
         Left            =   240
         TabIndex        =   844
         Top             =   4320
         Width           =   825
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Tipo trabajo"
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
         Index           =   58
         Left            =   240
         TabIndex        =   843
         Top             =   3120
         Width           =   1035
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   172
         Left            =   3600
         TabIndex        =   842
         Top             =   2760
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   171
         Left            =   600
         TabIndex        =   841
         Top             =   2760
         Width           =   465
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Nº Doc."
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
         Left            =   240
         TabIndex        =   840
         Top             =   2400
         Width           =   600
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
         Index           =   53
         Left            =   600
         TabIndex        =   839
         Top             =   1920
         Width           =   420
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   170
         Left            =   600
         TabIndex        =   838
         Top             =   1560
         Width           =   465
      End
      Begin VB.Image imgTecnico 
         Height          =   240
         Index           =   8
         Left            =   1200
         Picture         =   "frmListado2.frx":272E
         Top             =   1880
         Width           =   240
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Técnico"
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
         Index           =   56
         Left            =   240
         TabIndex        =   836
         Top             =   1320
         Width           =   645
      End
      Begin VB.Image imgTecnico 
         Height          =   240
         Index           =   7
         Left            =   1200
         Picture         =   "frmListado2.frx":2830
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
         Height          =   195
         Index           =   52
         Left            =   3360
         TabIndex        =   834
         Top             =   960
         Width           =   420
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   48
         Left            =   3840
         Picture         =   "frmListado2.frx":2932
         Top             =   960
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   47
         Left            =   1200
         Picture         =   "frmListado2.frx":29BD
         Top             =   960
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   169
         Left            =   600
         TabIndex        =   833
         Top             =   960
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
         Index           =   55
         Left            =   240
         TabIndex        =   832
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Informe productividad"
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
         Index           =   39
         Left            =   600
         TabIndex        =   831
         Top             =   240
         Width           =   4785
      End
   End
   Begin VB.Frame FrameBenClien 
      Height          =   6495
      Left            =   3000
      TabIndex        =   725
      Top             =   0
      Visible         =   0   'False
      Width           =   6135
      Begin VB.CheckBox chkBenAge 
         Caption         =   "Aplica descuento"
         Height          =   255
         Index           =   10
         Left            =   3360
         TabIndex        =   743
         Top             =   4800
         Width           =   1815
      End
      Begin VB.ComboBox cboCoste 
         Height          =   315
         Index           =   1
         ItemData        =   "frmListado2.frx":2A48
         Left            =   1440
         List            =   "frmListado2.frx":2A55
         Style           =   2  'Dropdown List
         TabIndex        =   742
         Top             =   4800
         Width           =   1695
      End
      Begin VB.CheckBox chkBenAge 
         Caption         =   "Ordenado por cliente"
         Height          =   255
         Index           =   5
         Left            =   4080
         TabIndex        =   746
         Top             =   5400
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.CheckBox chkBenAge 
         Caption         =   "detalla artículo"
         Height          =   255
         Index           =   4
         Left            =   2280
         TabIndex        =   745
         Top             =   5400
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CheckBox chkBenAge 
         Caption         =   "detalla marca"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   744
         Top             =   5400
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   41
         Left            =   4800
         TabIndex        =   748
         Top             =   5880
         Width           =   1095
      End
      Begin VB.CommandButton cmdbeneClien 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3600
         TabIndex        =   747
         Top             =   5880
         Width           =   975
      End
      Begin VB.TextBox txtDescClie 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   10
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   758
         Text            =   "Text1"
         Top             =   2400
         Width           =   3495
      End
      Begin VB.TextBox txtCliente 
         Height          =   285
         Index           =   10
         Left            =   1440
         TabIndex        =   735
         Text            =   "Text1"
         Top             =   2400
         Width           =   975
      End
      Begin VB.TextBox txtDescClie 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   9
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   755
         Text            =   "Text1"
         Top             =   2040
         Width           =   3495
      End
      Begin VB.TextBox txtCliente 
         Height          =   285
         Index           =   9
         Left            =   1440
         TabIndex        =   734
         Text            =   "Text1"
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox txtmarca 
         Height          =   285
         Index           =   7
         Left            =   1440
         TabIndex        =   737
         Text            =   "Text1"
         Top             =   3360
         Width           =   975
      End
      Begin VB.TextBox txtDescmarca 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   7
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   753
         Text            =   "Text5"
         Top             =   3360
         Width           =   3495
      End
      Begin VB.TextBox txtmarca 
         Height          =   285
         Index           =   6
         Left            =   1440
         TabIndex        =   736
         Text            =   "Text1"
         Top             =   3000
         Width           =   975
      End
      Begin VB.TextBox txtDescmarca 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   6
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   750
         Text            =   "Text5"
         Top             =   3000
         Width           =   3495
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   44
         Left            =   4200
         TabIndex        =   733
         Text            =   "Text1"
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   43
         Left            =   1440
         TabIndex        =   732
         Text            =   "Text1"
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox txtDescAlma 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   6
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   730
         Text            =   "Text5"
         Top             =   4320
         Width           =   3495
      End
      Begin VB.TextBox txtAlma 
         Height          =   285
         Index           =   6
         Left            =   1440
         TabIndex        =   741
         Text            =   "Text1"
         Top             =   4320
         Width           =   975
      End
      Begin VB.TextBox txtDescAlma 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   5
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   727
         Text            =   "Text5"
         Top             =   3960
         Width           =   3495
      End
      Begin VB.TextBox txtAlma 
         Height          =   285
         Index           =   5
         Left            =   1440
         TabIndex        =   739
         Text            =   "Text1"
         Top             =   3960
         Width           =   975
      End
      Begin VB.Image imgayuda 
         Height          =   255
         Index           =   6
         Left            =   3240
         ToolTipText     =   "Listado beneficio por cliente"
         Top             =   5880
         Width           =   255
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Coste"
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
         Index           =   48
         Left            =   120
         TabIndex        =   766
         Top             =   4800
         Width           =   480
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   156
         Left            =   240
         TabIndex        =   760
         Top             =   5880
         Width           =   2505
      End
      Begin VB.Image imgCliente 
         Height          =   240
         Index           =   10
         Left            =   1080
         Picture         =   "frmListado2.frx":2A85
         Top             =   2400
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   155
         Left            =   360
         TabIndex        =   759
         Top             =   2400
         Width           =   465
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
         Index           =   51
         Left            =   120
         TabIndex        =   757
         Top             =   1800
         Width           =   585
      End
      Begin VB.Image imgCliente 
         Height          =   240
         Index           =   9
         Left            =   1080
         Picture         =   "frmListado2.frx":2B87
         Top             =   2040
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   154
         Left            =   360
         TabIndex        =   756
         Top             =   2040
         Width           =   465
      End
      Begin VB.Image imgMarca 
         Height          =   240
         Index           =   7
         Left            =   1080
         Picture         =   "frmListado2.frx":2C89
         Top             =   3360
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   153
         Left            =   360
         TabIndex        =   754
         Top             =   3360
         Width           =   465
      End
      Begin VB.Image imgMarca 
         Height          =   240
         Index           =   6
         Left            =   1080
         Picture         =   "frmListado2.frx":2D8B
         Top             =   3000
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Marca"
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
         Index           =   46
         Left            =   120
         TabIndex        =   752
         Top             =   2760
         Width           =   525
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   152
         Left            =   360
         TabIndex        =   751
         Top             =   3000
         Width           =   465
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   44
         Left            =   3840
         Picture         =   "frmListado2.frx":2E8D
         Top             =   1335
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   151
         Left            =   360
         TabIndex        =   749
         Top             =   1365
         Width           =   465
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   43
         Left            =   1080
         Picture         =   "frmListado2.frx":2F18
         Top             =   1320
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   150
         Left            =   3240
         TabIndex        =   740
         Top             =   1365
         Width           =   465
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
         Index           =   45
         Left            =   120
         TabIndex        =   738
         Top             =   1080
         Width           =   495
      End
      Begin VB.Image imgAlma 
         Height          =   240
         Index           =   6
         Left            =   1080
         Picture         =   "frmListado2.frx":2FA3
         Top             =   4320
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   149
         Left            =   360
         TabIndex        =   731
         Top             =   4320
         Width           =   465
      End
      Begin VB.Image imgAlma 
         Height          =   240
         Index           =   5
         Left            =   1080
         Picture         =   "frmListado2.frx":30A5
         Top             =   3960
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Almacén"
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
         Index           =   44
         Left            =   120
         TabIndex        =   729
         Top             =   3720
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   148
         Left            =   360
         TabIndex        =   728
         Top             =   3960
         Width           =   465
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Beneficio por cliente"
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
         Index           =   36
         Left            =   1545
         TabIndex        =   726
         Top             =   480
         Width           =   2925
      End
   End
   Begin VB.Frame FrameBeneMarcaAgeProv 
      Height          =   6855
      Left            =   6600
      TabIndex        =   864
      Top             =   120
      Visible         =   0   'False
      Width           =   6255
      Begin VB.CheckBox chkBeneMarcaAgen 
         Caption         =   "Aplica descuento"
         Height          =   195
         Index           =   1
         Left            =   4440
         TabIndex        =   875
         Top             =   5640
         Width           =   1695
      End
      Begin VB.CommandButton cmdBeneMarcaAgen 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3720
         TabIndex        =   876
         Top             =   6240
         Width           =   1095
      End
      Begin VB.CheckBox chkBeneMarcaAgen 
         Caption         =   "Detalla artículo"
         Height          =   195
         Index           =   0
         Left            =   2760
         TabIndex        =   874
         Top             =   5640
         Width           =   1575
      End
      Begin VB.ComboBox cboCoste 
         Height          =   315
         Index           =   2
         ItemData        =   "frmListado2.frx":31A7
         Left            =   720
         List            =   "frmListado2.frx":31B4
         Style           =   2  'Dropdown List
         TabIndex        =   873
         Top             =   5520
         Width           =   1695
      End
      Begin VB.TextBox txtDescProve 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   25
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   895
         Text            =   "Text5"
         Top             =   4680
         Width           =   3495
      End
      Begin VB.TextBox txtCodProve 
         Height          =   285
         Index           =   25
         Left            =   1560
         TabIndex        =   872
         Top             =   4680
         Width           =   975
      End
      Begin VB.TextBox txtDescProve 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   24
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   892
         Text            =   "Text5"
         Top             =   4320
         Width           =   3495
      End
      Begin VB.TextBox txtCodProve 
         Height          =   285
         Index           =   24
         Left            =   1560
         TabIndex        =   871
         Top             =   4320
         Width           =   975
      End
      Begin VB.TextBox txtAgente 
         Height          =   285
         Index           =   12
         Left            =   1560
         TabIndex        =   870
         Text            =   "Text1"
         Top             =   3480
         Width           =   615
      End
      Begin VB.TextBox txtDescAgente 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   12
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   890
         Text            =   "Text1"
         Top             =   3480
         Width           =   3375
      End
      Begin VB.TextBox txtAgente 
         Height          =   285
         Index           =   11
         Left            =   1560
         TabIndex        =   869
         Text            =   "Text1"
         Top             =   3120
         Width           =   615
      End
      Begin VB.TextBox txtDescAgente 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   11
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   887
         Text            =   "Text1"
         Top             =   3120
         Width           =   3375
      End
      Begin VB.TextBox txtmarca 
         Height          =   285
         Index           =   9
         Left            =   1560
         TabIndex        =   868
         Text            =   "Text1"
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox txtDescmarca 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   9
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   885
         Text            =   "Text5"
         Top             =   2280
         Width           =   3255
      End
      Begin VB.TextBox txtmarca 
         Height          =   285
         Index           =   8
         Left            =   1560
         TabIndex        =   867
         Text            =   "Text1"
         Top             =   1920
         Width           =   855
      End
      Begin VB.TextBox txtDescmarca 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   8
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   882
         Text            =   "Text5"
         Top             =   1920
         Width           =   3255
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   50
         Left            =   4440
         TabIndex        =   866
         Text            =   "Text1"
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   49
         Left            =   1560
         TabIndex        =   865
         Text            =   "Text1"
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   48
         Left            =   4920
         TabIndex        =   877
         Top             =   6240
         Width           =   1095
      End
      Begin VB.Image imgayuda 
         Height          =   255
         Index           =   5
         Left            =   3360
         ToolTipText     =   "Listado beneficio marca-agente-proveedor"
         Top             =   6240
         Width           =   255
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   183
         Left            =   360
         TabIndex        =   898
         Top             =   6360
         Width           =   2985
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Coste"
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
         Index           =   56
         Left            =   360
         TabIndex        =   897
         Top             =   5280
         Width           =   480
      End
      Begin VB.Image imgProveedor 
         Height          =   240
         Index           =   25
         Left            =   1200
         Picture         =   "frmListado2.frx":31E4
         Top             =   4680
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   182
         Left            =   600
         TabIndex        =   896
         Top             =   4680
         Width           =   465
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
         Index           =   55
         Left            =   240
         TabIndex        =   894
         Top             =   4080
         Width           =   885
      End
      Begin VB.Image imgProveedor 
         Height          =   240
         Index           =   24
         Left            =   1200
         Picture         =   "frmListado2.frx":32E6
         Top             =   4320
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   181
         Left            =   600
         TabIndex        =   893
         Top             =   4320
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   180
         Left            =   600
         TabIndex        =   891
         Top             =   3480
         Width           =   465
      End
      Begin VB.Image imgAgente 
         Height          =   240
         Index           =   12
         Left            =   1200
         Picture         =   "frmListado2.frx":33E8
         Top             =   3480
         Width           =   240
      End
      Begin VB.Label lblDpto 
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
         Index           =   65
         Left            =   240
         TabIndex        =   889
         Top             =   2880
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   179
         Left            =   600
         TabIndex        =   888
         Top             =   3120
         Width           =   465
      End
      Begin VB.Image imgAgente 
         Height          =   240
         Index           =   11
         Left            =   1200
         Picture         =   "frmListado2.frx":34EA
         Top             =   3120
         Width           =   240
      End
      Begin VB.Image imgMarca 
         Height          =   240
         Index           =   9
         Left            =   1200
         Picture         =   "frmListado2.frx":35EC
         Top             =   2280
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   178
         Left            =   600
         TabIndex        =   886
         Top             =   2280
         Width           =   465
      End
      Begin VB.Image imgMarca 
         Height          =   240
         Index           =   8
         Left            =   1200
         Picture         =   "frmListado2.frx":36EE
         Top             =   1920
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Marca"
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
         Index           =   54
         Left            =   240
         TabIndex        =   884
         Top             =   1680
         Width           =   525
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   177
         Left            =   600
         TabIndex        =   883
         Top             =   1920
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   176
         Left            =   3600
         TabIndex        =   881
         Top             =   1245
         Width           =   465
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   50
         Left            =   4080
         Picture         =   "frmListado2.frx":37F0
         Top             =   1200
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   49
         Left            =   1200
         Picture         =   "frmListado2.frx":387B
         Top             =   1215
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   175
         Left            =   600
         TabIndex        =   880
         Top             =   1245
         Width           =   465
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Fecha factura"
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
         Index           =   64
         Left            =   240
         TabIndex        =   879
         Top             =   960
         Width           =   1155
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Beneficio Marca, Agente, Proveedor"
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
         Index           =   41
         Left            =   600
         TabIndex        =   878
         Top             =   360
         Width           =   5115
      End
   End
   Begin VB.Frame FrameMarcaFamilia 
      Height          =   6135
      Left            =   3000
      TabIndex        =   899
      Top             =   600
      Visible         =   0   'False
      Width           =   6495
      Begin VB.CommandButton cmdMarcaFamilia 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3960
         TabIndex        =   911
         Top             =   5520
         Width           =   1095
      End
      Begin VB.Frame FrameProveedor1 
         Height          =   975
         Left            =   120
         TabIndex        =   934
         Top             =   3480
         Width           =   6135
         Begin VB.TextBox txtDescProve 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   27
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   938
            Text            =   "Text5"
            Top             =   600
            Width           =   3495
         End
         Begin VB.TextBox txtCodProve 
            Height          =   285
            Index           =   27
            Left            =   1440
            TabIndex        =   909
            Top             =   600
            Width           =   975
         End
         Begin VB.TextBox txtDescProve 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   26
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   936
            Text            =   "Text5"
            Top             =   240
            Width           =   3495
         End
         Begin VB.TextBox txtCodProve 
            Height          =   285
            Index           =   26
            Left            =   1440
            TabIndex        =   908
            Top             =   240
            Width           =   975
         End
         Begin VB.Image imgProveedor 
            Height          =   240
            Index           =   27
            Left            =   1200
            Picture         =   "frmListado2.frx":3906
            Top             =   600
            Width           =   240
         End
         Begin VB.Label Label3 
            Caption         =   "Hasta"
            Height          =   195
            Index           =   194
            Left            =   480
            TabIndex        =   939
            Top             =   600
            Width           =   465
         End
         Begin VB.Image imgProveedor 
            Height          =   240
            Index           =   26
            Left            =   1200
            Picture         =   "frmListado2.frx":3A08
            Top             =   240
            Width           =   240
         End
         Begin VB.Label Label3 
            Caption         =   "Desde"
            Height          =   195
            Index           =   193
            Left            =   480
            TabIndex        =   937
            Top             =   240
            Width           =   465
         End
         Begin VB.Label lblDpto 
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
            Index           =   68
            Left            =   0
            TabIndex        =   935
            Top             =   0
            Width           =   885
         End
      End
      Begin VB.CheckBox chkMarcaFamilia 
         Caption         =   "Detalla articulo"
         Height          =   195
         Index           =   0
         Left            =   480
         TabIndex        =   910
         Top             =   4800
         Width           =   1935
      End
      Begin VB.TextBox txtDescFamia 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   17
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   930
         Text            =   "Text5"
         Top             =   3000
         Width           =   3495
      End
      Begin VB.TextBox txtFamia 
         Height          =   285
         Index           =   17
         Left            =   1560
         TabIndex        =   905
         Text            =   "Text1"
         Top             =   3000
         Width           =   855
      End
      Begin VB.TextBox txtDescFamia 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   16
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   927
         Text            =   "Text5"
         Top             =   2640
         Width           =   3495
      End
      Begin VB.TextBox txtFamia 
         Height          =   285
         Index           =   16
         Left            =   1560
         TabIndex        =   904
         Text            =   "Text1"
         Top             =   2640
         Width           =   855
      End
      Begin VB.Frame FrameAgente1 
         Height          =   975
         Left            =   120
         TabIndex        =   923
         Top             =   3480
         Width           =   6135
         Begin VB.TextBox txtAgente 
            Height          =   285
            Index           =   14
            Left            =   1440
            TabIndex        =   907
            Text            =   "Text1"
            Top             =   600
            Width           =   615
         End
         Begin VB.TextBox txtDescAgente 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   14
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   932
            Text            =   "Text1"
            Top             =   600
            Width           =   3495
         End
         Begin VB.TextBox txtAgente 
            Height          =   285
            Index           =   13
            Left            =   1440
            TabIndex        =   906
            Text            =   "Text1"
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox txtDescAgente 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   13
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   924
            Text            =   "Text1"
            Top             =   240
            Width           =   3495
         End
         Begin VB.Label Label3 
            Caption         =   "Hasta"
            Height          =   195
            Index           =   192
            Left            =   480
            TabIndex        =   933
            Top             =   600
            Width           =   465
         End
         Begin VB.Image imgAgente 
            Height          =   240
            Index           =   14
            Left            =   1200
            Picture         =   "frmListado2.frx":3B0A
            Top             =   600
            Width           =   240
         End
         Begin VB.Label lblDpto 
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
            Index           =   67
            Left            =   0
            TabIndex        =   926
            Top             =   0
            Width           =   615
         End
         Begin VB.Label Label3 
            Caption         =   "Desde"
            Height          =   195
            Index           =   189
            Left            =   480
            TabIndex        =   925
            Top             =   240
            Width           =   465
         End
         Begin VB.Image imgAgente 
            Height          =   240
            Index           =   13
            Left            =   1200
            Picture         =   "frmListado2.frx":3C0C
            Top             =   240
            Width           =   240
         End
      End
      Begin VB.TextBox txtDescmarca 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   11
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   920
         Text            =   "Text5"
         Top             =   2040
         Width           =   3495
      End
      Begin VB.TextBox txtmarca 
         Height          =   285
         Index           =   11
         Left            =   1560
         TabIndex        =   903
         Text            =   "Text1"
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox txtDescmarca 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   10
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   917
         Text            =   "Text5"
         Top             =   1680
         Width           =   3495
      End
      Begin VB.TextBox txtmarca 
         Height          =   285
         Index           =   10
         Left            =   1560
         TabIndex        =   902
         Text            =   "Text1"
         Top             =   1680
         Width           =   855
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   52
         Left            =   4560
         TabIndex        =   901
         Text            =   "Text1"
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   51
         Left            =   1560
         TabIndex        =   900
         Text            =   "Text1"
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   49
         Left            =   5160
         TabIndex        =   912
         Top             =   5520
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   191
         Left            =   600
         TabIndex        =   931
         Top             =   3000
         Width           =   465
      End
      Begin VB.Image imgFamilia 
         Height          =   240
         Index           =   17
         Left            =   1320
         Picture         =   "frmListado2.frx":3D0E
         Top             =   3000
         Width           =   240
      End
      Begin VB.Image imgFamilia 
         Height          =   240
         Index           =   16
         Left            =   1320
         Picture         =   "frmListado2.frx":3E10
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
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   58
         Left            =   120
         TabIndex        =   929
         Top             =   2400
         Width           =   600
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   190
         Left            =   600
         TabIndex        =   928
         Top             =   2640
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   188
         Left            =   240
         TabIndex        =   922
         Top             =   5640
         Width           =   3465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   187
         Left            =   600
         TabIndex        =   921
         Top             =   2040
         Width           =   465
      End
      Begin VB.Image imgMarca 
         Height          =   240
         Index           =   11
         Left            =   1200
         Picture         =   "frmListado2.frx":3F12
         Top             =   2040
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   186
         Left            =   600
         TabIndex        =   919
         Top             =   1680
         Width           =   465
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Marca"
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
         Left            =   120
         TabIndex        =   918
         Top             =   1440
         Width           =   525
      End
      Begin VB.Image imgMarca 
         Height          =   240
         Index           =   10
         Left            =   1200
         Picture         =   "frmListado2.frx":4014
         Top             =   1680
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   52
         Left            =   4200
         Picture         =   "frmListado2.frx":4116
         Top             =   960
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   185
         Left            =   3720
         TabIndex        =   916
         Top             =   1005
         Width           =   465
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Fecha factura"
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
         TabIndex        =   915
         Top             =   720
         Width           =   1155
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   184
         Left            =   600
         TabIndex        =   914
         Top             =   1005
         Width           =   465
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   51
         Left            =   1200
         Picture         =   "frmListado2.frx":41A1
         Top             =   975
         Width           =   240
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "ssss"
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
         Height          =   360
         Index           =   42
         Left            =   600
         TabIndex        =   913
         Top             =   240
         Width           =   5115
      End
   End
   Begin VB.Frame FrEstadisticasReparacionTecnico 
      Height          =   3495
      Left            =   120
      TabIndex        =   31
      Top             =   0
      Width           =   5895
      Begin VB.TextBox txtDescTra 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   0
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   36
         Text            =   "Text1"
         Top             =   1320
         Width           =   3495
      End
      Begin VB.TextBox txtTrab 
         Height          =   285
         Index           =   0
         Left            =   1080
         TabIndex        =   17
         Top             =   1320
         Width           =   855
      End
      Begin VB.CommandButton cmdEstadisticaReparacionTecnico 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3480
         TabIndex        =   20
         Top             =   2760
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   2
         Left            =   4680
         TabIndex        =   21
         Top             =   2760
         Width           =   975
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   3
         Left            =   4200
         TabIndex        =   19
         Text            =   "Text1"
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   2
         Left            =   1680
         TabIndex        =   18
         Text            =   "Text1"
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label3 
         Height          =   195
         Index           =   63
         Left            =   240
         TabIndex        =   320
         Top             =   2880
         Width           =   2865
      End
      Begin VB.Image imgTecnico 
         Height          =   240
         Index           =   0
         Left            =   720
         Picture         =   "frmListado2.frx":422C
         Top             =   1320
         Width           =   240
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Técnico"
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
         TabIndex        =   37
         Top             =   1080
         Width           =   645
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   3
         Left            =   3960
         Picture         =   "frmListado2.frx":432E
         Top             =   2040
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   3
         Left            =   3360
         TabIndex        =   35
         Top             =   2040
         Width           =   465
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   2
         Left            =   1440
         Picture         =   "frmListado2.frx":43B9
         Top             =   2040
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   2
         Left            =   840
         TabIndex        =   34
         Top             =   2040
         Width           =   465
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Fecha albarán"
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
         Left            =   240
         TabIndex        =   33
         Top             =   1800
         Width           =   1185
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Estadísticas reparación técnico"
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
         Index           =   1
         Left            =   240
         TabIndex        =   32
         Top             =   360
         Width           =   5415
      End
   End
   Begin VB.Frame FrameCambioEnFrecuencias 
      Height          =   3975
      Left            =   6480
      TabIndex        =   798
      Top             =   240
      Visible         =   0   'False
      Width           =   6855
      Begin VB.CommandButton cmdCambiClienteFreq 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   4080
         TabIndex        =   803
         Top             =   3240
         Width           =   975
      End
      Begin VB.TextBox txtDpto 
         Height          =   285
         Index           =   3
         Left            =   1560
         TabIndex        =   802
         Text            =   "Text1"
         Top             =   2640
         Width           =   735
      End
      Begin VB.TextBox txtDescDpto 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   3
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   814
         Text            =   "Text1"
         Top             =   2640
         Width           =   3975
      End
      Begin VB.TextBox txtDpto 
         Height          =   285
         Index           =   2
         Left            =   1560
         TabIndex        =   800
         Text            =   "Text1"
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox txtDescDpto 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   2
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   812
         Text            =   "Text1"
         Top             =   1320
         Width           =   3975
      End
      Begin VB.TextBox txtCliente 
         Height          =   285
         Index           =   12
         Left            =   1560
         TabIndex        =   801
         Text            =   "Text1"
         Top             =   2280
         Width           =   1215
      End
      Begin VB.TextBox txtDescClie 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   12
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   810
         Text            =   "Text1"
         Top             =   2280
         Width           =   3495
      End
      Begin VB.TextBox txtCliente 
         Height          =   285
         Index           =   11
         Left            =   1560
         TabIndex        =   799
         Text            =   "Text1"
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtDescClie 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   11
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   808
         Text            =   "Text1"
         Top             =   960
         Width           =   3495
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   45
         Left            =   5280
         TabIndex        =   804
         Top             =   3240
         Width           =   975
      End
      Begin VB.Image imgDpto 
         Height          =   240
         Index           =   3
         Left            =   1320
         Picture         =   "frmListado2.frx":4444
         Top             =   2640
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Departamento"
         Height          =   195
         Index           =   168
         Left            =   120
         TabIndex        =   815
         Top             =   2640
         Width           =   1065
      End
      Begin VB.Image imgDpto 
         Height          =   240
         Index           =   2
         Left            =   1320
         Picture         =   "frmListado2.frx":4546
         Top             =   1320
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Departamento"
         Height          =   195
         Index           =   167
         Left            =   120
         TabIndex        =   813
         Top             =   1320
         Width           =   825
      End
      Begin VB.Label Label3 
         Caption         =   "Cliente"
         Height          =   195
         Index           =   166
         Left            =   120
         TabIndex        =   811
         Top             =   2280
         Width           =   480
      End
      Begin VB.Image imgCliente 
         Height          =   240
         Index           =   12
         Left            =   1320
         Picture         =   "frmListado2.frx":4648
         Top             =   2280
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Cliente"
         Height          =   195
         Index           =   165
         Left            =   120
         TabIndex        =   809
         Top             =   960
         Width           =   600
      End
      Begin VB.Image imgCliente 
         Height          =   240
         Index           =   11
         Left            =   1320
         Picture         =   "frmListado2.frx":474A
         Top             =   960
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Destino"
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
         Index           =   51
         Left            =   120
         TabIndex        =   807
         Top             =   1920
         Width           =   645
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Frecuencias: Cambio cliente/Dpto"
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
         Index           =   38
         Left            =   960
         TabIndex        =   806
         Top             =   240
         Width           =   4785
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Origen"
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
         Index           =   50
         Left            =   120
         TabIndex        =   805
         Top             =   720
         Width           =   555
      End
   End
   Begin VB.Frame FrListadoReparaciones 
      Height          =   4335
      Left            =   120
      TabIndex        =   10
      Top             =   0
      Width           =   6495
      Begin VB.OptionButton optReparaciones 
         Caption         =   "Fecha entrada"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   3360
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton optReparaciones 
         Caption         =   "Fecha albarán"
         Height          =   375
         Index           =   1
         Left            =   1920
         TabIndex        =   8
         Top             =   3360
         Width           =   1815
      End
      Begin VB.CommandButton cmdReparaEfect 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   4080
         TabIndex        =   9
         Top             =   3720
         Width           =   975
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   1
         Left            =   4320
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   3000
         Width           =   1335
      End
      Begin VB.TextBox txtDpto 
         Height          =   285
         Index           =   1
         Left            =   1680
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox txtDescDpto 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   1
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   22
         Text            =   "Text1"
         Top             =   2280
         Width           =   3615
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   1
         Left            =   5280
         TabIndex        =   16
         Top             =   3720
         Width           =   975
      End
      Begin VB.TextBox txtDescDpto 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   0
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   1920
         Width           =   3615
      End
      Begin VB.TextBox txtDpto 
         Height          =   285
         Index           =   0
         Left            =   1680
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   1920
         Width           =   855
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   0
         Left            =   1680
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   3000
         Width           =   1335
      End
      Begin VB.TextBox txtCliente 
         Height          =   285
         Index           =   1
         Left            =   1680
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox txtDescClie 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   1
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   1200
         Width           =   3495
      End
      Begin VB.TextBox txtCliente 
         Height          =   285
         Index           =   0
         Left            =   1680
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtDescClie 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   0
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   840
         Width           =   3495
      End
      Begin VB.Label Label3 
         Height          =   195
         Index           =   158
         Left            =   240
         TabIndex        =   769
         Top             =   3840
         Width           =   3225
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   1
         Left            =   4080
         Picture         =   "frmListado2.frx":484C
         Top             =   3000
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
         Index           =   2
         Left            =   240
         TabIndex        =   30
         Top             =   2760
         Width           =   1575
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
         Index           =   1
         Left            =   240
         TabIndex        =   29
         Top             =   600
         Width           =   585
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
         Index           =   1
         Left            =   3600
         TabIndex        =   28
         Top             =   3000
         Width           =   420
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   1
         Left            =   840
         TabIndex        =   27
         Top             =   3000
         Width           =   465
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
         Index           =   0
         Left            =   840
         TabIndex        =   26
         Top             =   2280
         Width           =   420
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   0
         Left            =   840
         TabIndex        =   25
         Top             =   1920
         Width           =   465
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
         Index           =   23
         Left            =   840
         TabIndex        =   24
         Top             =   1200
         Width           =   420
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   42
         Left            =   840
         TabIndex        =   23
         Top             =   840
         Width           =   465
      End
      Begin VB.Image imgDpto 
         Height          =   240
         Index           =   1
         Left            =   1440
         Picture         =   "frmListado2.frx":48D7
         Top             =   2280
         Width           =   240
      End
      Begin VB.Label lblDpto 
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
         Index           =   0
         Left            =   240
         TabIndex        =   15
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Image imgDpto 
         Height          =   240
         Index           =   0
         Left            =   1440
         Picture         =   "frmListado2.frx":49D9
         Top             =   1920
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   1440
         Picture         =   "frmListado2.frx":4ADB
         Top             =   3000
         Width           =   240
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Reparaciones efectuadas"
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
         Index           =   0
         Left            =   240
         TabIndex        =   13
         Top             =   120
         Width           =   5895
      End
      Begin VB.Image imgCliente 
         Height          =   240
         Index           =   1
         Left            =   1440
         Picture         =   "frmListado2.frx":4B66
         Top             =   1200
         Width           =   240
      End
      Begin VB.Image imgCliente 
         Height          =   240
         Index           =   0
         Left            =   1440
         Picture         =   "frmListado2.frx":4C68
         Top             =   840
         Width           =   240
      End
   End
   Begin VB.Frame FrameSituAlbaranes 
      Height          =   5055
      Left            =   0
      TabIndex        =   358
      Top             =   0
      Visible         =   0   'False
      Width           =   6855
      Begin VB.CommandButton cmdSituAlbaran 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4080
         TabIndex        =   364
         Top             =   4200
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   23
         Left            =   5400
         TabIndex        =   365
         Top             =   4200
         Width           =   1095
      End
      Begin VB.ListBox List1 
         Height          =   1410
         Left            =   360
         Style           =   1  'Checkbox
         TabIndex        =   363
         Top             =   3120
         Width           =   3135
      End
      Begin VB.TextBox txtDescClie 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   6
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   373
         Text            =   "Text1"
         Top             =   2280
         Width           =   3495
      End
      Begin VB.TextBox txtCliente 
         Height          =   285
         Index           =   6
         Left            =   1320
         TabIndex        =   362
         Text            =   "Text1"
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox txtDescClie 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   5
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   370
         Text            =   "Text1"
         Top             =   1920
         Width           =   3495
      End
      Begin VB.TextBox txtCliente 
         Height          =   285
         Index           =   5
         Left            =   1320
         TabIndex        =   361
         Text            =   "Text1"
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   30
         Left            =   4680
         TabIndex        =   360
         Text            =   "Text1"
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   29
         Left            =   1560
         TabIndex        =   359
         Text            =   "Text1"
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   3
         Left            =   1320
         Picture         =   "frmListado2.frx":4D6A
         ToolTipText     =   "Quitar al haber"
         Top             =   2880
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   2
         Left            =   1680
         Picture         =   "frmListado2.frx":4EB4
         ToolTipText     =   "Puntear al haber"
         Top             =   2880
         Width           =   240
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Albaranes"
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
         Index           =   32
         Left            =   360
         TabIndex        =   375
         Top             =   2880
         Width           =   855
      End
      Begin VB.Image imgCliente 
         Height          =   240
         Index           =   6
         Left            =   1080
         Picture         =   "frmListado2.frx":4FFE
         Top             =   2280
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   75
         Left            =   360
         TabIndex        =   374
         Top             =   2280
         Width           =   465
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
         Index           =   31
         Left            =   240
         TabIndex        =   372
         Top             =   1560
         Width           =   585
      End
      Begin VB.Image imgCliente 
         Height          =   240
         Index           =   5
         Left            =   1080
         Picture         =   "frmListado2.frx":5100
         Top             =   1920
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   74
         Left            =   360
         TabIndex        =   371
         Top             =   1920
         Width           =   465
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   30
         Left            =   4440
         Picture         =   "frmListado2.frx":5202
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   73
         Left            =   3720
         TabIndex        =   369
         Top             =   1125
         Width           =   465
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   29
         Left            =   1320
         Picture         =   "frmListado2.frx":528D
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   72
         Left            =   600
         TabIndex        =   368
         Top             =   1125
         Width           =   465
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Fecha "
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
         Left            =   240
         TabIndex        =   367
         Top             =   840
         Width           =   1140
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Informe situación albaranes"
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
         Index           =   21
         Left            =   720
         TabIndex        =   366
         Top             =   360
         Width           =   5415
      End
   End
   Begin VB.Frame FraCambPrecTar 
      Height          =   6735
      Left            =   360
      TabIndex        =   551
      Top             =   240
      Visible         =   0   'False
      Width           =   6975
      Begin VB.OptionButton optSituaArt 
         Caption         =   "Obsoletos"
         Height          =   255
         Index           =   1
         Left            =   2160
         TabIndex        =   771
         Top             =   5160
         Width           =   1335
      End
      Begin VB.CommandButton cmdCambiPrecio 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3960
         TabIndex        =   563
         Top             =   6120
         Width           =   1215
      End
      Begin VB.OptionButton optSituaArt 
         Caption         =   "Caducados"
         Height          =   255
         Index           =   3
         Left            =   5280
         TabIndex        =   562
         Top             =   5160
         Width           =   1455
      End
      Begin VB.OptionButton optSituaArt 
         Caption         =   "Bloqueados"
         Height          =   255
         Index           =   2
         Left            =   3720
         TabIndex        =   561
         Top             =   5160
         Width           =   1335
      End
      Begin VB.TextBox txtDescTarifa 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   0
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   578
         Text            =   "Text5"
         Top             =   1680
         Width           =   3495
      End
      Begin VB.TextBox txtTarifa 
         Height          =   285
         Index           =   0
         Left            =   1920
         TabIndex        =   554
         Text            =   "Text1"
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox txtDescArticulo 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   9
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   576
         Text            =   "Text5"
         Top             =   4560
         Width           =   3495
      End
      Begin VB.TextBox txtArticulo 
         Height          =   285
         Index           =   9
         Left            =   1560
         MaxLength       =   16
         TabIndex        =   559
         Top             =   4560
         Width           =   1455
      End
      Begin VB.OptionButton optSituaArt 
         Caption         =   "Todos"
         Height          =   255
         Index           =   0
         Left            =   1080
         TabIndex        =   560
         Top             =   5160
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.TextBox txtDescArticulo 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   8
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   573
         Text            =   "Text5"
         Top             =   4200
         Width           =   3495
      End
      Begin VB.TextBox txtArticulo 
         Height          =   285
         Index           =   8
         Left            =   1560
         MaxLength       =   16
         TabIndex        =   558
         Top             =   4200
         Width           =   1455
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   38
         Left            =   1920
         TabIndex        =   553
         Text            =   "Text1"
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   34
         Left            =   5400
         TabIndex        =   564
         Top             =   6120
         Width           =   1095
      End
      Begin VB.TextBox txtDescFamia 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   7
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   570
         Text            =   "Text5"
         Top             =   3600
         Width           =   3495
      End
      Begin VB.TextBox txtFamia 
         Height          =   285
         Index           =   7
         Left            =   1920
         TabIndex        =   557
         Text            =   "Text1"
         Top             =   3600
         Width           =   1095
      End
      Begin VB.TextBox txtDescFamia 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   6
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   567
         Text            =   "Text5"
         Top             =   3240
         Width           =   3495
      End
      Begin VB.TextBox txtFamia 
         Height          =   285
         Index           =   6
         Left            =   1920
         TabIndex        =   556
         Text            =   "Text1"
         Top             =   3240
         Width           =   1095
      End
      Begin VB.TextBox txtDescProve 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   20
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   565
         Text            =   "Text5"
         Top             =   2400
         Width           =   3495
      End
      Begin VB.TextBox txtCodProve 
         Height          =   285
         Index           =   20
         Left            =   1920
         TabIndex        =   555
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Tarifa"
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
         Index           =   43
         Left            =   360
         TabIndex        =   579
         Top             =   1680
         Width           =   495
      End
      Begin VB.Image imgTarifa 
         Height          =   240
         Index           =   0
         Left            =   1680
         Picture         =   "frmListado2.frx":5318
         Top             =   1680
         Width           =   240
      End
      Begin VB.Image imgArticulo 
         Height          =   240
         Index           =   9
         Left            =   1320
         Picture         =   "frmListado2.frx":541A
         Top             =   4560
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   111
         Left            =   720
         TabIndex        =   577
         Top             =   4560
         Width           =   465
      End
      Begin VB.Image imgArticulo 
         Height          =   240
         Index           =   8
         Left            =   1320
         Picture         =   "frmListado2.frx":551C
         Top             =   4200
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Articulo"
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
         Index           =   29
         Left            =   360
         TabIndex        =   575
         Top             =   3960
         Width           =   660
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   110
         Left            =   720
         TabIndex        =   574
         Top             =   4200
         Width           =   465
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   38
         Left            =   1560
         Picture         =   "frmListado2.frx":561E
         Top             =   1200
         Width           =   240
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Fecha  cambio"
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
         Index           =   42
         Left            =   360
         TabIndex        =   572
         Top             =   1200
         Width           =   1200
      End
      Begin VB.Image imgFamilia 
         Height          =   240
         Index           =   7
         Left            =   1680
         Picture         =   "frmListado2.frx":56A9
         Top             =   3600
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   109
         Left            =   720
         TabIndex        =   571
         Top             =   3600
         Width           =   465
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
         Index           =   28
         Left            =   360
         TabIndex        =   569
         Top             =   3000
         Width           =   600
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   108
         Left            =   720
         TabIndex        =   568
         Top             =   3240
         Width           =   465
      End
      Begin VB.Image imgFamilia 
         Height          =   240
         Index           =   6
         Left            =   1680
         Picture         =   "frmListado2.frx":57AB
         Top             =   3240
         Width           =   240
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
         Index           =   27
         Left            =   360
         TabIndex        =   566
         Top             =   2400
         Width           =   885
      End
      Begin VB.Image imgProveedor 
         Height          =   240
         Index           =   20
         Left            =   1680
         Picture         =   "frmListado2.frx":58AD
         Top             =   2400
         Width           =   240
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
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
         Index           =   31
         Left            =   480
         TabIndex        =   552
         Top             =   360
         Width           =   5685
      End
   End
   Begin VB.Frame FrameControlAlbaranes 
      Height          =   5175
      Left            =   1800
      TabIndex        =   772
      Top             =   1080
      Visible         =   0   'False
      Width           =   6495
      Begin VB.CommandButton cmdControlAlbaranes 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3960
         TabIndex        =   779
         Top             =   4320
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   42
         Left            =   5280
         TabIndex        =   780
         Top             =   4320
         Width           =   975
      End
      Begin VB.TextBox txtEnvio 
         Height          =   285
         Index           =   1
         Left            =   1560
         TabIndex        =   776
         Text            =   "Text1"
         Top             =   2520
         Width           =   1095
      End
      Begin VB.TextBox txtDescEnvio 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   1
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   792
         Text            =   "Text5"
         Top             =   2520
         Width           =   3495
      End
      Begin VB.TextBox txtEnvio 
         Height          =   285
         Index           =   0
         Left            =   1560
         TabIndex        =   775
         Text            =   "Text1"
         Top             =   2040
         Width           =   1095
      End
      Begin VB.TextBox txtDescEnvio 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   0
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   789
         Text            =   "Text5"
         Top             =   2040
         Width           =   3495
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   46
         Left            =   3960
         TabIndex        =   774
         Text            =   "Text1"
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox txtZona 
         Height          =   285
         Index           =   3
         Left            =   1560
         TabIndex        =   778
         Text            =   "Text1"
         Top             =   3720
         Width           =   1095
      End
      Begin VB.TextBox txtDescZona 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   3
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   785
         Text            =   "Text5"
         Top             =   3720
         Width           =   3495
      End
      Begin VB.TextBox txtZona 
         Height          =   285
         Index           =   2
         Left            =   1560
         TabIndex        =   777
         Text            =   "Text1"
         Top             =   3240
         Width           =   1095
      End
      Begin VB.TextBox txtDescZona 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   2
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   783
         Text            =   "Text5"
         Top             =   3240
         Width           =   3495
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   45
         Left            =   1560
         TabIndex        =   773
         Text            =   "Text1"
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label lblDpto 
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
         Index           =   54
         Left            =   240
         TabIndex        =   794
         Top             =   2880
         Width           =   420
      End
      Begin VB.Image imgEnvio 
         Height          =   240
         Index           =   1
         Left            =   1320
         Picture         =   "frmListado2.frx":59AF
         Top             =   2520
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   164
         Left            =   600
         TabIndex        =   793
         Top             =   2520
         Width           =   465
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Código de  envio"
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
         Index           =   53
         Left            =   240
         TabIndex        =   791
         Top             =   1680
         Width           =   1380
      End
      Begin VB.Image imgEnvio 
         Height          =   240
         Index           =   0
         Left            =   1320
         Picture         =   "frmListado2.frx":5AB1
         Top             =   2040
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   163
         Left            =   600
         TabIndex        =   790
         Top             =   2040
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   162
         Left            =   3000
         TabIndex        =   788
         Top             =   1200
         Width           =   465
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   46
         Left            =   3600
         Picture         =   "frmListado2.frx":5BB3
         Top             =   1200
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   161
         Left            =   600
         TabIndex        =   787
         Top             =   1200
         Width           =   465
      End
      Begin VB.Image imgZona 
         Height          =   240
         Index           =   3
         Left            =   1320
         Picture         =   "frmListado2.frx":5C3E
         Top             =   3720
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   160
         Left            =   600
         TabIndex        =   786
         Top             =   3720
         Width           =   465
      End
      Begin VB.Image imgZona 
         Height          =   240
         Index           =   2
         Left            =   1320
         Picture         =   "frmListado2.frx":5D40
         Top             =   3240
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   159
         Left            =   600
         TabIndex        =   784
         Top             =   3240
         Width           =   465
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Fecha  envio"
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
         Left            =   240
         TabIndex        =   782
         Top             =   840
         Width           =   1050
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   45
         Left            =   1200
         Picture         =   "frmListado2.frx":5E42
         Top             =   1200
         Width           =   240
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
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
         Index           =   37
         Left            =   120
         TabIndex        =   781
         Top             =   240
         Width           =   6225
      End
   End
   Begin VB.Frame FrProveedorxVenta 
      Height          =   5055
      Left            =   120
      TabIndex        =   81
      Top             =   0
      Width           =   12375
      Begin VB.CheckBox chkVtaxProv 
         Caption         =   "Mostrar clientes"
         Height          =   255
         Index           =   3
         Left            =   4440
         TabIndex        =   100
         Top             =   4560
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox txtimporte 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   7800
         TabIndex        =   90
         Text            =   "Text1"
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox txtAlma 
         Height          =   285
         Index           =   2
         Left            =   1080
         TabIndex        =   85
         Text            =   "Text1"
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox txtDescAlma 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   2
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   715
         Text            =   "Text5"
         Top             =   1920
         Width           =   3495
      End
      Begin VB.TextBox txtAlma 
         Height          =   285
         Index           =   1
         Left            =   1080
         TabIndex        =   84
         Text            =   "Text1"
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox txtDescAlma 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   1
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   712
         Text            =   "Text5"
         Top             =   1560
         Width           =   3495
      End
      Begin VB.CheckBox chkVtaxProv 
         Caption         =   "Fam. comparativo"
         Height          =   255
         Index           =   2
         Left            =   1320
         TabIndex        =   98
         Top             =   4560
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CheckBox chkVtaxProv 
         Caption         =   "Detalla"
         Height          =   255
         Index           =   1
         Left            =   3240
         TabIndex        =   99
         Top             =   4560
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CheckBox chkVtaxProv 
         Caption         =   "Agente"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   97
         Top             =   4560
         Width           =   1335
      End
      Begin VB.TextBox txtAgente 
         Height          =   285
         Index           =   9
         Left            =   7440
         TabIndex        =   96
         Text            =   "Text1"
         Top             =   4080
         Width           =   735
      End
      Begin VB.TextBox txtDescAgente 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   9
         Left            =   8280
         Locked          =   -1  'True
         TabIndex        =   709
         Text            =   "Text1"
         Top             =   4080
         Width           =   3855
      End
      Begin VB.TextBox txtAgente 
         Height          =   285
         Index           =   8
         Left            =   7440
         TabIndex        =   95
         Text            =   "Text1"
         Top             =   3720
         Width           =   735
      End
      Begin VB.TextBox txtDescAgente 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   8
         Left            =   8280
         Locked          =   -1  'True
         TabIndex        =   706
         Text            =   "Text1"
         Top             =   3720
         Width           =   3855
      End
      Begin VB.TextBox txtDescFamia 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   15
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   704
         Text            =   "Text5"
         Top             =   4080
         Width           =   3495
      End
      Begin VB.TextBox txtFamia 
         Height          =   285
         Index           =   15
         Left            =   1080
         TabIndex        =   89
         Text            =   "Text1"
         Top             =   4080
         Width           =   1335
      End
      Begin VB.TextBox txtDescFamia 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   14
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   701
         Text            =   "Text5"
         Top             =   3720
         Width           =   3495
      End
      Begin VB.TextBox txtFamia 
         Height          =   285
         Index           =   14
         Left            =   1080
         TabIndex        =   88
         Text            =   "Text1"
         Top             =   3720
         Width           =   1335
      End
      Begin VB.TextBox txtDescProve 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   1
         Left            =   8640
         Locked          =   -1  'True
         TabIndex        =   120
         Text            =   "Text5"
         Top             =   3000
         Width           =   3495
      End
      Begin VB.TextBox txtCodProve 
         Height          =   285
         Index           =   1
         Left            =   7200
         TabIndex        =   94
         Text            =   "Text1"
         Top             =   3000
         Width           =   1455
      End
      Begin VB.TextBox txtDescProve 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   0
         Left            =   8640
         Locked          =   -1  'True
         TabIndex        =   117
         Text            =   "Text5"
         Top             =   2640
         Width           =   3495
      End
      Begin VB.TextBox txtCodProve 
         Height          =   285
         Index           =   0
         Left            =   7200
         TabIndex        =   93
         Text            =   "Text1"
         Top             =   2640
         Width           =   1455
      End
      Begin VB.TextBox txtArticulo 
         Height          =   285
         Index           =   2
         Left            =   1080
         MaxLength       =   16
         TabIndex        =   87
         Top             =   3000
         Width           =   1455
      End
      Begin VB.TextBox txtDescArticulo 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   2
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   116
         Text            =   "Text5"
         Top             =   3000
         Width           =   3495
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   10
         Left            =   4560
         TabIndex        =   83
         Text            =   "Text1"
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton cmdVentaxProv 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   9960
         TabIndex        =   101
         Top             =   4560
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   6
         Left            =   11040
         TabIndex        =   102
         Top             =   4560
         Width           =   975
      End
      Begin VB.TextBox txtArticulo 
         Height          =   285
         Index           =   1
         Left            =   1080
         MaxLength       =   16
         TabIndex        =   86
         Top             =   2640
         Width           =   1455
      End
      Begin VB.TextBox txtDescArticulo 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   1
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   109
         Text            =   "Text5"
         Top             =   2640
         Width           =   3495
      End
      Begin VB.TextBox txtCliente 
         Height          =   285
         Index           =   4
         Left            =   7200
         TabIndex        =   92
         Text            =   "Text1"
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox txtDescClie 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   4
         Left            =   8640
         Locked          =   -1  'True
         TabIndex        =   108
         Text            =   "Text1"
         Top             =   1920
         Width           =   3495
      End
      Begin VB.TextBox txtCliente 
         Height          =   285
         Index           =   3
         Left            =   7200
         TabIndex        =   91
         Text            =   "Text1"
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox txtDescClie 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   3
         Left            =   8640
         Locked          =   -1  'True
         TabIndex        =   106
         Text            =   "Text1"
         Top             =   1560
         Width           =   3495
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   9
         Left            =   2280
         TabIndex        =   82
         Text            =   "Text1"
         Top             =   840
         Width           =   1095
      End
      Begin VB.Image imgayuda 
         Height          =   255
         Index           =   0
         Left            =   9000
         ToolTipText     =   "Ventas proveedor"
         Top             =   840
         Width           =   255
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Importe mínimo"
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
         Index           =   50
         Left            =   6240
         TabIndex        =   717
         Top             =   840
         Width           =   1380
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   144
         Left            =   240
         TabIndex        =   716
         Top             =   1920
         Width           =   465
      End
      Begin VB.Image imgAlma 
         Height          =   240
         Index           =   2
         Left            =   840
         Picture         =   "frmListado2.frx":5ECD
         Top             =   1920
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   143
         Left            =   240
         TabIndex        =   714
         Top             =   1560
         Width           =   465
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Almacén"
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
         Index           =   41
         Left            =   120
         TabIndex        =   713
         Top             =   1300
         Width           =   735
      End
      Begin VB.Image imgAlma 
         Height          =   240
         Index           =   1
         Left            =   840
         Picture         =   "frmListado2.frx":5FCF
         Top             =   1560
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   ".."
         Height          =   195
         Index           =   142
         Left            =   6240
         TabIndex        =   711
         Top             =   4560
         Width           =   3465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   141
         Left            =   6480
         TabIndex        =   710
         Top             =   4080
         Width           =   465
      End
      Begin VB.Image imgAgente 
         Height          =   240
         Index           =   9
         Left            =   7200
         Picture         =   "frmListado2.frx":60D1
         Top             =   4080
         Width           =   240
      End
      Begin VB.Label lblDpto 
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
         Index           =   49
         Left            =   6240
         TabIndex        =   708
         Top             =   3420
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   140
         Left            =   6480
         TabIndex        =   707
         Top             =   3750
         Width           =   465
      End
      Begin VB.Image imgAgente 
         Height          =   240
         Index           =   8
         Left            =   7200
         Picture         =   "frmListado2.frx":61D3
         Top             =   3720
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   139
         Left            =   240
         TabIndex        =   705
         Top             =   4080
         Width           =   465
      End
      Begin VB.Image imgFamilia 
         Height          =   240
         Index           =   15
         Left            =   840
         Picture         =   "frmListado2.frx":62D5
         Top             =   4080
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   138
         Left            =   240
         TabIndex        =   703
         Top             =   3720
         Width           =   465
      End
      Begin VB.Image imgFamilia 
         Height          =   240
         Index           =   14
         Left            =   840
         Picture         =   "frmListado2.frx":63D7
         Top             =   3720
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
         Index           =   40
         Left            =   120
         TabIndex        =   702
         Top             =   3420
         Width           =   600
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   18
         Left            =   6480
         TabIndex        =   121
         Top             =   3000
         Width           =   465
      End
      Begin VB.Image imgProveedor 
         Height          =   240
         Index           =   1
         Left            =   6960
         Picture         =   "frmListado2.frx":64D9
         Top             =   3000
         Width           =   240
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
         Index           =   4
         Left            =   6240
         TabIndex        =   119
         Top             =   2400
         Width           =   885
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   17
         Left            =   6480
         TabIndex        =   118
         Top             =   2640
         Width           =   465
      End
      Begin VB.Image imgProveedor 
         Height          =   240
         Index           =   0
         Left            =   6960
         Picture         =   "frmListado2.frx":65DB
         Top             =   2640
         Width           =   240
      End
      Begin VB.Image imgArticulo 
         Height          =   240
         Index           =   2
         Left            =   840
         Picture         =   "frmListado2.frx":66DD
         Top             =   3000
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   16
         Left            =   240
         TabIndex        =   115
         Top             =   3000
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   15
         Left            =   240
         TabIndex        =   114
         Top             =   2640
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   14
         Left            =   6360
         TabIndex        =   113
         Top             =   1920
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   13
         Left            =   6360
         TabIndex        =   112
         Top             =   1560
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   12
         Left            =   3840
         TabIndex        =   111
         Top             =   840
         Width           =   465
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   10
         Left            =   4320
         Picture         =   "frmListado2.frx":67DF
         Top             =   840
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Articulo"
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
         TabIndex        =   110
         Top             =   2400
         Width           =   660
      End
      Begin VB.Image imgArticulo 
         Height          =   240
         Index           =   1
         Left            =   840
         Picture         =   "frmListado2.frx":686A
         Top             =   2640
         Width           =   240
      End
      Begin VB.Image imgCliente 
         Height          =   240
         Index           =   4
         Left            =   6960
         Picture         =   "frmListado2.frx":696C
         Top             =   1920
         Width           =   240
      End
      Begin VB.Image imgCliente 
         Height          =   240
         Index           =   3
         Left            =   6960
         Picture         =   "frmListado2.frx":6A6E
         Top             =   1560
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
         Index           =   11
         Left            =   6240
         TabIndex        =   107
         Top             =   1320
         Width           =   585
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Listado ventas por  proveedor"
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
         Index           =   5
         Left            =   3360
         TabIndex        =   105
         Top             =   120
         Width           =   5415
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Fecha factura"
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
         Left            =   120
         TabIndex        =   104
         Top             =   840
         Width           =   1155
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   11
         Left            =   1560
         TabIndex        =   103
         Top             =   840
         Width           =   465
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   9
         Left            =   2040
         Picture         =   "frmListado2.frx":6B70
         Top             =   840
         Width           =   240
      End
   End
   Begin VB.Frame FrameListTrabajadores 
      Height          =   2535
      Left            =   3240
      TabIndex        =   291
      Top             =   480
      Width           =   5895
      Begin VB.TextBox txtTrab 
         Height          =   285
         Index           =   4
         Left            =   1080
         TabIndex        =   298
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox txtDescTra 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   4
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   297
         Text            =   "Text1"
         Top             =   1320
         Width           =   3495
      End
      Begin VB.TextBox txtTrab 
         Height          =   285
         Index           =   3
         Left            =   1080
         TabIndex        =   295
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox txtDescTra 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   3
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   294
         Text            =   "Text1"
         Top             =   840
         Width           =   3495
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   17
         Left            =   4560
         TabIndex        =   293
         Top             =   1920
         Width           =   975
      End
      Begin VB.CommandButton cmdListTrabja 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3480
         TabIndex        =   292
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   59
         Left            =   120
         TabIndex        =   300
         Top             =   1320
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   58
         Left            =   120
         TabIndex        =   299
         Top             =   840
         Width           =   465
      End
      Begin VB.Image imgTecnico 
         Height          =   240
         Index           =   4
         Left            =   720
         Picture         =   "frmListado2.frx":6BFB
         Top             =   1320
         Width           =   240
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Listado trabajadores"
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
         Left            =   120
         TabIndex        =   296
         Top             =   240
         Width           =   5415
      End
      Begin VB.Image imgTecnico 
         Height          =   240
         Index           =   3
         Left            =   720
         Picture         =   "frmListado2.frx":6CFD
         Top             =   840
         Width           =   240
      End
   End
   Begin VB.Frame FrameFacturarCliente 
      Height          =   3015
      Left            =   840
      TabIndex        =   383
      Top             =   0
      Visible         =   0   'False
      Width           =   6615
      Begin VB.CheckBox chkFacturarCliente 
         Caption         =   "Imprimir facturas generadas"
         Height          =   255
         Left            =   1680
         TabIndex        =   392
         Top             =   2040
         Value           =   1  'Checked
         Width           =   2775
      End
      Begin VB.CommandButton cmdFacturarCli 
         Caption         =   "Facturar"
         Height          =   375
         Left            =   4080
         TabIndex        =   387
         Top             =   2400
         Width           =   1095
      End
      Begin VB.TextBox txtBancoPr 
         Height          =   285
         Index           =   2
         Left            =   2040
         TabIndex        =   386
         Text            =   "Text1"
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox txtDescBancoPr 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   2
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   390
         Text            =   "Text5"
         Top             =   1560
         Width           =   3615
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   31
         Left            =   2040
         TabIndex        =   385
         Text            =   "Text1"
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   25
         Left            =   5280
         TabIndex        =   388
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Banco"
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
         Left            =   240
         TabIndex        =   391
         Top             =   1560
         Width           =   510
      End
      Begin VB.Image imgBancoPr 
         Height          =   240
         Index           =   2
         Left            =   1680
         Picture         =   "frmListado2.frx":6DFF
         Top             =   1560
         Width           =   240
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Fecha factura"
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
         Index           =   33
         Left            =   240
         TabIndex        =   389
         Top             =   1080
         Width           =   1155
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   31
         Left            =   1680
         Picture         =   "frmListado2.frx":6F01
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Facturación cliente"
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
         Index           =   23
         Left            =   240
         TabIndex        =   384
         Top             =   360
         Width           =   5415
      End
   End
   Begin VB.Frame FrImprimirFac 
      Height          =   4575
      Left            =   120
      TabIndex        =   172
      Top             =   0
      Visible         =   0   'False
      Width           =   6495
      Begin VB.CommandButton cmdImprimirFac 
         Caption         =   "Imprimir"
         Height          =   375
         Left            =   3600
         TabIndex        =   179
         Top             =   4080
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   9
         Left            =   5040
         TabIndex        =   180
         Top             =   4080
         Width           =   1215
      End
      Begin VB.TextBox txtCodProve 
         Height          =   285
         Index           =   7
         Left            =   1200
         TabIndex        =   178
         Top             =   3240
         Width           =   1455
      End
      Begin VB.TextBox txtDescProve 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   7
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   182
         Text            =   "Text5"
         Top             =   3240
         Width           =   3495
      End
      Begin VB.TextBox txtCodProve 
         Height          =   285
         Index           =   6
         Left            =   1200
         TabIndex        =   177
         Text            =   "Text1"
         Top             =   2880
         Width           =   1455
      End
      Begin VB.TextBox txtDescProve 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   6
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   181
         Text            =   "Text5"
         Top             =   2880
         Width           =   3495
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   16
         Left            =   4560
         TabIndex        =   174
         Text            =   "Text1"
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   15
         Left            =   1920
         TabIndex        =   173
         Text            =   "Text1"
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtNumAlbar 
         Height          =   285
         Index           =   3
         Left            =   4560
         TabIndex        =   176
         Text            =   "Text1"
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox txtNumAlbar 
         Height          =   285
         Index           =   2
         Left            =   1920
         TabIndex        =   175
         Text            =   "Text1"
         Top             =   1995
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Label1"
         Height          =   375
         Left            =   120
         TabIndex        =   193
         Top             =   3720
         Width           =   6135
      End
      Begin VB.Image imgProveedor 
         Height          =   240
         Index           =   7
         Left            =   840
         Picture         =   "frmListado2.frx":6F8C
         Top             =   3240
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   34
         Left            =   240
         TabIndex        =   192
         Top             =   2880
         Width           =   465
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
         Index           =   8
         Left            =   120
         TabIndex        =   191
         Top             =   2520
         Width           =   885
      End
      Begin VB.Image imgProveedor 
         Height          =   240
         Index           =   6
         Left            =   840
         Picture         =   "frmListado2.frx":708E
         Top             =   2880
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   33
         Left            =   240
         TabIndex        =   190
         Top             =   3240
         Width           =   465
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   16
         Left            =   4320
         Picture         =   "frmListado2.frx":7190
         Top             =   1095
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   32
         Left            =   3720
         TabIndex        =   189
         Top             =   1125
         Width           =   465
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   15
         Left            =   1680
         Picture         =   "frmListado2.frx":721B
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   31
         Left            =   960
         TabIndex        =   188
         Top             =   1125
         Width           =   465
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Fecha factura"
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
         Left            =   120
         TabIndex        =   187
         Top             =   720
         Width           =   1155
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Imprimir facturas proveedores"
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
         Index           =   8
         Left            =   240
         TabIndex        =   186
         Top             =   240
         Width           =   5415
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Num. factura"
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
         Index           =   15
         Left            =   120
         TabIndex        =   185
         Top             =   1560
         Width           =   1080
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   30
         Left            =   960
         TabIndex        =   184
         Top             =   2040
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   29
         Left            =   3720
         TabIndex        =   183
         Top             =   2040
         Width           =   465
      End
   End
   Begin VB.Frame FrameTraza 
      Height          =   4935
      Left            =   120
      TabIndex        =   242
      Top             =   0
      Width           =   6735
      Begin VB.TextBox txtArticulo 
         Height          =   285
         Index           =   5
         Left            =   1440
         MaxLength       =   16
         TabIndex        =   247
         Top             =   1800
         Width           =   1455
      End
      Begin VB.CommandButton cmdTraza 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3720
         TabIndex        =   252
         Top             =   4320
         Width           =   1335
      End
      Begin VB.TextBox txtDescArticulo 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   5
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   261
         Text            =   "Text5"
         Top             =   1800
         Width           =   3495
      End
      Begin VB.TextBox txtArticulo 
         Height          =   285
         Index           =   4
         Left            =   1440
         MaxLength       =   16
         TabIndex        =   246
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox txtDescArticulo 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   4
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   260
         Text            =   "Text5"
         Top             =   1320
         Width           =   3495
      End
      Begin VB.TextBox txtDescProve 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   11
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   258
         Text            =   "Text5"
         Top             =   3120
         Width           =   3495
      End
      Begin VB.TextBox txtCodProve 
         Height          =   285
         Index           =   11
         Left            =   1440
         TabIndex        =   249
         Top             =   3120
         Width           =   1455
      End
      Begin VB.TextBox txtDescProve 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   10
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   256
         Text            =   "Text5"
         Top             =   2760
         Width           =   3495
      End
      Begin VB.TextBox txtCodProve 
         Height          =   285
         Index           =   10
         Left            =   1440
         TabIndex        =   248
         Top             =   2760
         Width           =   1455
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   15
         Left            =   5160
         TabIndex        =   253
         Top             =   4320
         Width           =   1215
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   23
         Left            =   5160
         TabIndex        =   251
         Text            =   "Text1"
         Top             =   3600
         Width           =   1215
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   22
         Left            =   2040
         TabIndex        =   250
         Text            =   "Text1"
         Top             =   3600
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   53
         Left            =   480
         TabIndex        =   263
         Top             =   1800
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   52
         Left            =   480
         TabIndex        =   262
         Top             =   1320
         Width           =   465
      End
      Begin VB.Image imgArticulo 
         Height          =   240
         Index           =   5
         Left            =   1080
         Picture         =   "frmListado2.frx":72A6
         Top             =   1800
         Width           =   240
      End
      Begin VB.Image imgArticulo 
         Height          =   240
         Index           =   4
         Left            =   1080
         Picture         =   "frmListado2.frx":73A8
         Top             =   1320
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   51
         Left            =   480
         TabIndex        =   259
         Top             =   3120
         Width           =   465
      End
      Begin VB.Image imgProveedor 
         Height          =   240
         Index           =   11
         Left            =   1080
         Picture         =   "frmListado2.frx":74AA
         Top             =   3120
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   50
         Left            =   480
         TabIndex        =   257
         Top             =   2760
         Width           =   465
      End
      Begin VB.Image imgProveedor 
         Height          =   240
         Index           =   10
         Left            =   1080
         Picture         =   "frmListado2.frx":75AC
         Top             =   2760
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   23
         Left            =   4920
         Picture         =   "frmListado2.frx":76AE
         Top             =   3600
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   49
         Left            =   4320
         TabIndex        =   255
         Top             =   3645
         Width           =   465
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   22
         Left            =   1800
         Picture         =   "frmListado2.frx":7739
         Top             =   3600
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   48
         Left            =   1200
         TabIndex        =   254
         Top             =   3645
         Width           =   465
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Artículo"
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
         Height          =   375
         Index           =   13
         Left            =   240
         TabIndex        =   245
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Albaran proveedor"
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
         Height          =   375
         Index           =   12
         Left            =   240
         TabIndex        =   244
         Top             =   2400
         Width           =   2055
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Trazabilidad albaranes"
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
         Index           =   11
         Left            =   240
         TabIndex        =   243
         Top             =   240
         Width           =   5415
      End
   End
   Begin VB.Frame FrLiqCambioPrecios 
      Height          =   5055
      Left            =   120
      TabIndex        =   122
      Top             =   0
      Visible         =   0   'False
      Width           =   6495
      Begin VB.CommandButton cmdCambiarImporte 
         Caption         =   "Cambiar"
         Height          =   375
         Left            =   3600
         TabIndex        =   130
         Top             =   4440
         Width           =   1215
      End
      Begin VB.TextBox txtimporte 
         Height          =   285
         Index           =   0
         Left            =   1200
         TabIndex        =   129
         Text            =   "Text1"
         Top             =   3720
         Width           =   1455
      End
      Begin VB.TextBox txtDescArticulo 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   3
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   140
         Text            =   "Text5"
         Top             =   3120
         Width           =   3495
      End
      Begin VB.TextBox txtArticulo 
         Height          =   285
         Index           =   3
         Left            =   1200
         MaxLength       =   16
         TabIndex        =   128
         Top             =   3120
         Width           =   1455
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   7
         Left            =   5040
         TabIndex        =   131
         Top             =   4440
         Width           =   1215
      End
      Begin VB.TextBox txtCodProve 
         Height          =   285
         Index           =   3
         Left            =   1200
         TabIndex        =   127
         Top             =   2280
         Width           =   1455
      End
      Begin VB.TextBox txtDescProve 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   3
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   136
         Text            =   "Text5"
         Top             =   2280
         Width           =   3495
      End
      Begin VB.TextBox txtCodProve 
         Height          =   285
         Index           =   2
         Left            =   1200
         TabIndex        =   126
         Text            =   "Text1"
         Top             =   1920
         Width           =   1455
      End
      Begin VB.TextBox txtDescProve 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   2
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   135
         Text            =   "Text5"
         Top             =   1920
         Width           =   3495
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   12
         Left            =   4560
         TabIndex        =   125
         Text            =   "Text1"
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   11
         Left            =   1920
         TabIndex        =   124
         Text            =   "Text1"
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label lblLiqu 
         Caption         =   "Label1"
         Height          =   375
         Left            =   120
         TabIndex        =   143
         Top             =   4320
         Width           =   3135
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Importe"
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
         TabIndex        =   142
         Top             =   3720
         Width           =   705
      End
      Begin VB.Image imgArticulo 
         Height          =   240
         Index           =   3
         Left            =   840
         Picture         =   "frmListado2.frx":77C4
         Top             =   3120
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Articulo"
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
         TabIndex        =   141
         Top             =   2760
         Width           =   660
      End
      Begin VB.Image imgProveedor 
         Height          =   240
         Index           =   3
         Left            =   840
         Picture         =   "frmListado2.frx":78C6
         Top             =   2280
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   22
         Left            =   240
         TabIndex        =   139
         Top             =   1920
         Width           =   465
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
         Left            =   120
         TabIndex        =   138
         Top             =   1560
         Width           =   885
      End
      Begin VB.Image imgProveedor 
         Height          =   240
         Index           =   2
         Left            =   840
         Picture         =   "frmListado2.frx":79C8
         Top             =   1920
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   21
         Left            =   240
         TabIndex        =   137
         Top             =   2280
         Width           =   465
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   12
         Left            =   4320
         Picture         =   "frmListado2.frx":7ACA
         Top             =   1095
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   20
         Left            =   3720
         TabIndex        =   134
         Top             =   1125
         Width           =   465
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   11
         Left            =   1680
         Picture         =   "frmListado2.frx":7B55
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   19
         Left            =   960
         TabIndex        =   133
         Top             =   1125
         Width           =   465
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Fecha albaran"
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
         TabIndex        =   132
         Top             =   720
         Width           =   1185
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Cambio precios albaranes proveedor"
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
         Index           =   6
         Left            =   240
         TabIndex        =   123
         Top             =   240
         Width           =   5415
      End
   End
   Begin VB.Frame FrameRiesgo 
      Height          =   2775
      Left            =   120
      TabIndex        =   482
      Top             =   0
      Visible         =   0   'False
      Width           =   6375
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   31
         Left            =   4800
         TabIndex        =   485
         Top             =   2160
         Width           =   1215
      End
      Begin VB.CommandButton cmdRiesgo 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3240
         TabIndex        =   484
         Top             =   2160
         Width           =   1335
      End
      Begin MSComctlLib.ProgressBar pb2 
         Height          =   495
         Left            =   360
         TabIndex        =   486
         Top             =   1320
         Visible         =   0   'False
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   873
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label3 
         Height          =   195
         Index           =   95
         Left            =   360
         TabIndex        =   487
         Top             =   960
         Width           =   5505
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Calculo de riesgo"
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
         Index           =   28
         Left            =   240
         TabIndex        =   483
         Top             =   480
         Width           =   2505
      End
   End
   Begin VB.Frame FrFacturaRecargas 
      Height          =   6015
      Left            =   120
      TabIndex        =   59
      Top             =   0
      Visible         =   0   'False
      Width           =   6015
      Begin VB.TextBox txtBancoPr 
         Height          =   285
         Index           =   1
         Left            =   720
         TabIndex        =   66
         Text            =   "Text1"
         Top             =   4560
         Width           =   855
      End
      Begin VB.TextBox txtDescBancoPr 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   1
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   202
         Text            =   "Text5"
         Top             =   4560
         Width           =   4095
      End
      Begin MSComctlLib.ProgressBar pb1 
         Height          =   375
         Left            =   240
         TabIndex        =   79
         Top             =   5400
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   8
         Left            =   1200
         TabIndex        =   77
         Text            =   "Text1"
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox txtDescArticulo 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   0
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   75
         Text            =   "Text5"
         Top             =   3840
         Width           =   3495
      End
      Begin VB.TextBox txtArticulo 
         Height          =   285
         Index           =   0
         Left            =   720
         MaxLength       =   16
         TabIndex        =   65
         Top             =   3840
         Width           =   1455
      End
      Begin VB.TextBox txtRecargaMov 
         Height          =   285
         Index           =   1
         Left            =   5040
         MaxLength       =   1
         TabIndex        =   63
         Text            =   "Text1"
         Top             =   1200
         Width           =   375
      End
      Begin VB.TextBox txtDescClie 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   2
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   72
         Text            =   "Text1"
         Top             =   3000
         Width           =   3495
      End
      Begin VB.TextBox txtCliente 
         Height          =   285
         Index           =   2
         Left            =   720
         TabIndex        =   64
         Text            =   "Text1"
         Top             =   3000
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   5
         Left            =   4800
         TabIndex        =   68
         Top             =   5400
         Width           =   975
      End
      Begin VB.CommandButton cmdFacturaMov 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3600
         TabIndex        =   67
         Top             =   5400
         Width           =   975
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   7
         Left            =   3240
         TabIndex        =   62
         Text            =   "Text1"
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   6
         Left            =   1200
         TabIndex        =   61
         Text            =   "Text1"
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Banco propio"
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
         Left            =   120
         TabIndex        =   203
         Top             =   4320
         Width           =   1095
      End
      Begin VB.Image imgBancoPr 
         Height          =   240
         Index           =   1
         Left            =   1320
         Picture         =   "frmListado2.frx":7BE0
         Top             =   4200
         Width           =   240
      End
      Begin VB.Label lblIndicadorT 
         Caption         =   "Label1"
         Height          =   375
         Left            =   240
         TabIndex        =   80
         Top             =   5040
         Width           =   3615
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Fecha factura"
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
         Left            =   120
         TabIndex        =   78
         Top             =   1800
         Width           =   1155
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   8
         Left            =   840
         Picture         =   "frmListado2.frx":7CE2
         Top             =   2175
         Width           =   240
      End
      Begin VB.Image imgArticulo 
         Height          =   240
         Index           =   0
         Left            =   840
         Picture         =   "frmListado2.frx":7D6D
         Top             =   3600
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Articulo"
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
         TabIndex        =   76
         Top             =   3600
         Width           =   660
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
         Index           =   8
         Left            =   120
         TabIndex        =   74
         Top             =   2760
         Width           =   585
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Tipo"
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
         Left            =   5040
         TabIndex        =   73
         Top             =   840
         Width           =   360
      End
      Begin VB.Image imgCliente 
         Height          =   240
         Index           =   2
         Left            =   840
         Picture         =   "frmListado2.frx":7E6F
         Top             =   2760
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   7
         Left            =   3000
         Picture         =   "frmListado2.frx":7F71
         Top             =   1222
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   10
         Left            =   2520
         TabIndex        =   71
         Top             =   1245
         Width           =   465
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   6
         Left            =   840
         Picture         =   "frmListado2.frx":7FFC
         Top             =   1222
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   9
         Left            =   360
         TabIndex        =   70
         Top             =   1245
         Width           =   465
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Fecha recarga"
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
         TabIndex        =   69
         Top             =   840
         Width           =   1200
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Facturación  recargas moviles"
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
         Index           =   4
         Left            =   240
         TabIndex        =   60
         Top             =   240
         Width           =   5415
      End
   End
   Begin VB.Frame FrRepGaranProv 
      Height          =   3855
      Left            =   4680
      TabIndex        =   465
      Top             =   120
      Visible         =   0   'False
      Width           =   6615
      Begin VB.CommandButton cmdRepGaranProve 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   4080
         TabIndex        =   471
         Top             =   3240
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   30
         Left            =   5400
         TabIndex        =   472
         Top             =   3240
         Width           =   975
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   37
         Left            =   4440
         TabIndex        =   470
         Text            =   "Text1"
         Top             =   2520
         Width           =   1215
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   36
         Left            =   1680
         TabIndex        =   469
         Text            =   "Text1"
         Top             =   2520
         Width           =   1215
      End
      Begin VB.TextBox txtCodProve 
         Height          =   285
         Index           =   16
         Left            =   1680
         TabIndex        =   468
         Text            =   "Text1"
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox txtDescProve 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   16
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   476
         Text            =   "Text5"
         Top             =   1680
         Width           =   3495
      End
      Begin VB.TextBox txtCodProve 
         Height          =   285
         Index           =   15
         Left            =   1680
         TabIndex        =   467
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox txtDescProve 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   15
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   473
         Text            =   "Text5"
         Top             =   1320
         Width           =   3495
      End
      Begin VB.Label Label3 
         Height          =   195
         Index           =   94
         Left            =   120
         TabIndex        =   481
         Top             =   3360
         Width           =   3705
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   93
         Left            =   3480
         TabIndex        =   480
         Top             =   2520
         Width           =   465
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   37
         Left            =   4080
         Picture         =   "frmListado2.frx":8087
         Top             =   2520
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   92
         Left            =   720
         TabIndex        =   479
         Top             =   2520
         Width           =   465
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Fecha albarán"
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
         Index           =   41
         Left            =   120
         TabIndex        =   478
         Top             =   2160
         Width           =   1185
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   36
         Left            =   1320
         Picture         =   "frmListado2.frx":8112
         Top             =   2520
         Width           =   240
      End
      Begin VB.Image imgProveedor 
         Height          =   240
         Index           =   16
         Left            =   1440
         Picture         =   "frmListado2.frx":819D
         Top             =   1680
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   91
         Left            =   600
         TabIndex        =   477
         Top             =   1680
         Width           =   465
      End
      Begin VB.Image imgProveedor 
         Height          =   240
         Index           =   15
         Left            =   1440
         Picture         =   "frmListado2.frx":829F
         Top             =   1320
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   90
         Left            =   600
         TabIndex        =   475
         Top             =   1320
         Width           =   465
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
         Index           =   16
         Left            =   120
         TabIndex        =   474
         Top             =   1080
         Width           =   885
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Reparaciones en garantía proveedor"
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
         Index           =   27
         Left            =   600
         TabIndex        =   466
         Top             =   480
         Width           =   5145
      End
   End
   Begin VB.Frame FrameReimpAlb 
      Height          =   3615
      Left            =   840
      TabIndex        =   419
      Top             =   120
      Visible         =   0   'False
      Width           =   6615
      Begin VB.OptionButton optAlbTrans 
         Caption         =   "Listado albaranes "
         Height          =   255
         Index           =   1
         Left            =   3600
         TabIndex        =   425
         Top             =   2400
         Width           =   2535
      End
      Begin VB.OptionButton optAlbTrans 
         Caption         =   "Imprime albaranes"
         Height          =   255
         Index           =   0
         Left            =   720
         TabIndex        =   424
         Top             =   2400
         Value           =   -1  'True
         Width           =   2055
      End
      Begin VB.CommandButton cmdImpAlbRut 
         Caption         =   "Imprimir"
         Height          =   375
         Left            =   3720
         TabIndex        =   427
         Top             =   3000
         Width           =   1095
      End
      Begin VB.CheckBox chkImpAlbRut 
         Caption         =   "Ya impresos"
         Height          =   255
         Index           =   0
         Left            =   720
         TabIndex        =   426
         Top             =   2760
         Width           =   2415
      End
      Begin VB.TextBox txtDescClie 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   8
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   433
         Text            =   "Text1"
         Top             =   1920
         Width           =   3495
      End
      Begin VB.TextBox txtCliente 
         Height          =   285
         Index           =   8
         Left            =   1560
         TabIndex        =   423
         Text            =   "Text1"
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox txtDescClie 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   7
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   430
         Text            =   "Text1"
         Top             =   1560
         Width           =   3495
      End
      Begin VB.TextBox txtCliente 
         Height          =   285
         Index           =   7
         Left            =   1560
         TabIndex        =   422
         Text            =   "Text1"
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   27
         Left            =   5040
         TabIndex        =   428
         Top             =   3000
         Width           =   1095
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   35
         Left            =   1920
         TabIndex        =   421
         Text            =   "Text1"
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label lblIndicAlb 
         Caption         =   "Label8"
         Height          =   255
         Left            =   240
         TabIndex        =   435
         Top             =   3120
         Width           =   3255
      End
      Begin VB.Image imgCliente 
         Height          =   240
         Index           =   8
         Left            =   1320
         Picture         =   "frmListado2.frx":83A1
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
         Index           =   12
         Left            =   720
         TabIndex        =   434
         Top             =   1920
         Width           =   420
      End
      Begin VB.Image imgCliente 
         Height          =   240
         Index           =   7
         Left            =   1320
         Picture         =   "frmListado2.frx":84A3
         Top             =   1560
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   83
         Left            =   720
         TabIndex        =   432
         Top             =   1560
         Width           =   465
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
         Index           =   39
         Left            =   120
         TabIndex        =   431
         Top             =   1320
         Width           =   585
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   35
         Left            =   1560
         Picture         =   "frmListado2.frx":85A5
         Top             =   960
         Width           =   240
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Fecha albarán"
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
         Index           =   38
         Left            =   120
         TabIndex        =   429
         Top             =   960
         Width           =   1185
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Albaranes con transporte"
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
         Index           =   25
         Left            =   120
         TabIndex        =   420
         Top             =   360
         Width           =   6015
      End
   End
   Begin VB.Frame FrameFrecuencia 
      Height          =   2175
      Left            =   3360
      TabIndex        =   376
      Top             =   120
      Width           =   6015
      Begin VB.CheckBox chkFrecu 
         Caption         =   "Legal"
         Height          =   255
         Left            =   3720
         TabIndex        =   382
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox texto 
         Height          =   280
         Index           =   5
         Left            =   360
         MaxLength       =   80
         TabIndex        =   380
         Top             =   1080
         Width           =   2205
      End
      Begin VB.CommandButton cmdFrecuencia 
         Caption         =   "Modificar"
         Height          =   375
         Left            =   3360
         TabIndex        =   379
         Top             =   1680
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   24
         Left            =   4680
         TabIndex        =   377
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Expediente"
         Height          =   195
         Index           =   76
         Left            =   360
         TabIndex        =   381
         Top             =   840
         Width           =   1065
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Cambiar expediente frecuencias"
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
         Index           =   22
         Left            =   240
         TabIndex        =   378
         Top             =   240
         Width           =   5415
      End
   End
   Begin VB.Frame frameContabTickets 
      Height          =   3495
      Left            =   120
      TabIndex        =   225
      Top             =   0
      Width           =   6255
      Begin VB.Frame FrameTapa 
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   120
         TabIndex        =   235
         Top             =   1440
         Width           =   6015
         Begin VB.TextBox txtTrab 
            Height          =   285
            Index           =   2
            Left            =   1680
            TabIndex        =   240
            Top             =   840
            Width           =   855
         End
         Begin VB.TextBox txtDescTra 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   2
            Left            =   2640
            Locked          =   -1  'True
            TabIndex        =   239
            Text            =   "Text1"
            Top             =   840
            Width           =   3255
         End
         Begin VB.OptionButton optTick 
            Caption         =   "Diario"
            Height          =   255
            Index           =   0
            Left            =   1560
            TabIndex        =   237
            Top             =   240
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton optTick 
            Caption         =   "Mensual"
            Height          =   255
            Index           =   1
            Left            =   3120
            TabIndex        =   236
            Top             =   240
            Width           =   1095
         End
         Begin VB.Image imgTecnico 
            Height          =   240
            Index           =   2
            Left            =   1320
            Picture         =   "frmListado2.frx":8630
            Top             =   840
            Width           =   240
         End
         Begin VB.Label lblDpto 
            AutoSize        =   -1  'True
            Caption         =   "Trabajador: "
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
            TabIndex        =   241
            Top             =   840
            Width           =   1035
         End
         Begin VB.Label lblDpto 
            AutoSize        =   -1  'True
            Caption         =   "Agrupa por: "
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
            Height          =   375
            Index           =   21
            Left            =   120
            TabIndex        =   238
            Top             =   240
            Width           =   1035
         End
      End
      Begin VB.CommandButton cmdContabTicket 
         Caption         =   "Contabilizar"
         Height          =   375
         Left            =   3720
         TabIndex        =   228
         Top             =   2880
         Width           =   1095
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   21
         Left            =   4560
         TabIndex        =   227
         Text            =   "Text1"
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   20
         Left            =   1920
         TabIndex        =   226
         Text            =   "Text1"
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   13
         Left            =   5040
         TabIndex        =   229
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label Label5 
         Height          =   255
         Left            =   240
         TabIndex        =   234
         Top             =   2880
         Width           =   3375
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Fecha "
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
         Index           =   20
         Left            =   240
         TabIndex        =   233
         Top             =   720
         Width           =   540
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   47
         Left            =   1080
         TabIndex        =   232
         Top             =   1125
         Width           =   465
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   21
         Left            =   4320
         Picture         =   "frmListado2.frx":8732
         Top             =   1095
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   20
         Left            =   1680
         Picture         =   "frmListado2.frx":87BD
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   46
         Left            =   3840
         TabIndex        =   231
         Top             =   1125
         Width           =   465
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "tickets agrupados"
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
         Index           =   10
         Left            =   600
         TabIndex        =   230
         Top             =   240
         Width           =   5415
      End
   End
   Begin VB.Frame FrameOtrasOfertas 
      Height          =   4455
      Left            =   120
      TabIndex        =   336
      Top             =   120
      Visible         =   0   'False
      Width           =   8415
      Begin VB.CommandButton cmdAceptarOfertas 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   5520
         TabIndex        =   340
         Top             =   3960
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   21
         Left            =   6960
         TabIndex        =   339
         Top             =   3960
         Width           =   1095
      End
      Begin MSComctlLib.ListView lw1 
         Height          =   2775
         Left            =   240
         TabIndex        =   338
         Top             =   960
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   4895
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Num. ofer"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Fecha"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "F. Entrega"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Di"
            Object.Width           =   6703
         EndProperty
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   1
         Left            =   600
         Picture         =   "frmListado2.frx":8848
         ToolTipText     =   "Puntear al haber"
         Top             =   3840
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   0
         Left            =   240
         Picture         =   "frmListado2.frx":8992
         ToolTipText     =   "Quitar al haber"
         Top             =   3840
         Width           =   240
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
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
         Index           =   27
         Left            =   240
         TabIndex        =   341
         Top             =   720
         Width           =   7725
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Otras ofertas"
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
         Index           =   19
         Left            =   2640
         TabIndex        =   337
         Top             =   240
         Width           =   2925
      End
   End
   Begin VB.Frame FrameListadoPlantillas 
      Height          =   3975
      Left            =   240
      TabIndex        =   321
      Top             =   0
      Visible         =   0   'False
      Width           =   6615
      Begin VB.CommandButton cmdImprPlatil 
         Caption         =   "Imprimir"
         Height          =   375
         Left            =   3840
         TabIndex        =   326
         Top             =   3120
         Width           =   975
      End
      Begin VB.TextBox txtDescGrupoP 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   1
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   335
         Text            =   "Text5"
         Top             =   2640
         Width           =   3495
      End
      Begin VB.TextBox txtDescGrupoP 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   0
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   0
         Text            =   "Text5"
         Top             =   2280
         Width           =   3495
      End
      Begin VB.TextBox txtGrupoPlan 
         Height          =   285
         Index           =   1
         Left            =   1680
         TabIndex        =   325
         Text            =   "Text1"
         Top             =   2640
         Width           =   855
      End
      Begin VB.TextBox txtGrupoPlan 
         Height          =   285
         Index           =   0
         Left            =   1680
         TabIndex        =   324
         Text            =   "Text1"
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox txtNumero 
         Height          =   285
         Index           =   1
         Left            =   4440
         TabIndex        =   323
         Text            =   "Text1"
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox txtNumero 
         Height          =   285
         Index           =   0
         Left            =   1680
         TabIndex        =   322
         Text            =   "Text1"
         Top             =   1320
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   20
         Left            =   5040
         TabIndex        =   327
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Grupo plantilla"
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
         Index           =   26
         Left            =   240
         TabIndex        =   334
         Top             =   1920
         Width           =   1230
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   67
         Left            =   840
         TabIndex        =   333
         Top             =   2280
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   66
         Left            =   840
         TabIndex        =   332
         Top             =   2640
         Width           =   465
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Plantilla"
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
         Index           =   25
         Left            =   240
         TabIndex        =   331
         Top             =   960
         Width           =   675
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   65
         Left            =   840
         TabIndex        =   330
         Top             =   1320
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   64
         Left            =   3840
         TabIndex        =   329
         Top             =   1320
         Width           =   465
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Listado plantillas ofertas"
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
         Left            =   1320
         TabIndex        =   328
         Top             =   360
         Width           =   3645
      End
   End
   Begin VB.Frame FrameAlbaProv 
      Height          =   4095
      Left            =   0
      TabIndex        =   204
      Top             =   0
      Width           =   6615
      Begin VB.TextBox txtCodProve 
         Height          =   285
         Index           =   9
         Left            =   1200
         TabIndex        =   215
         Text            =   "Text1"
         Top             =   2760
         Width           =   1455
      End
      Begin VB.TextBox txtDescProve 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   9
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   223
         Text            =   "Text5"
         Top             =   2760
         Width           =   3495
      End
      Begin VB.TextBox txtCodProve 
         Height          =   285
         Index           =   8
         Left            =   1200
         TabIndex        =   214
         Top             =   2400
         Width           =   1455
      End
      Begin VB.TextBox txtDescProve 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   8
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   219
         Text            =   "Text5"
         Top             =   2400
         Width           =   3495
      End
      Begin VB.TextBox txtNumAlbar 
         Height          =   285
         Index           =   5
         Left            =   4920
         TabIndex        =   213
         Text            =   "Text1"
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox txtNumAlbar 
         Height          =   285
         Index           =   4
         Left            =   1920
         TabIndex        =   210
         Text            =   "Text1"
         Top             =   1635
         Width           =   1215
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   19
         Left            =   4920
         TabIndex        =   208
         Text            =   "Text1"
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   18
         Left            =   1920
         TabIndex        =   205
         Text            =   "Text1"
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton cmdAlbaranProv 
         Caption         =   "Imprimir"
         Height          =   375
         Left            =   4320
         TabIndex        =   216
         Top             =   3480
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   10
         Left            =   5400
         TabIndex        =   217
         Top             =   3480
         Width           =   975
      End
      Begin VB.Image imgProveedor 
         Height          =   240
         Index           =   9
         Left            =   840
         Picture         =   "frmListado2.frx":8ADC
         Top             =   2760
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   45
         Left            =   240
         TabIndex        =   224
         Top             =   2760
         Width           =   465
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Imprimir albarán proveedor"
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
         Index           =   9
         Left            =   720
         TabIndex        =   222
         Top             =   120
         Width           =   5415
      End
      Begin VB.Image imgProveedor 
         Height          =   240
         Index           =   8
         Left            =   840
         Picture         =   "frmListado2.frx":8BDE
         Top             =   2400
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   44
         Left            =   240
         TabIndex        =   221
         Top             =   2400
         Width           =   465
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
         Index           =   11
         Left            =   120
         TabIndex        =   220
         Top             =   2040
         Width           =   885
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   43
         Left            =   3960
         TabIndex        =   218
         Top             =   1725
         Width           =   465
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Num. albaran"
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
         Left            =   120
         TabIndex        =   212
         Top             =   1320
         Width           =   1110
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   41
         Left            =   960
         TabIndex        =   211
         Top             =   1680
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   40
         Left            =   3960
         TabIndex        =   209
         Top             =   885
         Width           =   465
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   19
         Left            =   4680
         Picture         =   "frmListado2.frx":8CE0
         Top             =   840
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   18
         Left            =   1680
         Picture         =   "frmListado2.frx":8D6B
         Top             =   855
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   39
         Left            =   960
         TabIndex        =   207
         Top             =   885
         Width           =   465
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Fecha albaran"
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
         TabIndex        =   206
         Top             =   600
         Width           =   1185
      End
   End
   Begin VB.Frame FrameCopiaPrecios 
      Height          =   5655
      Left            =   120
      TabIndex        =   436
      Top             =   120
      Width           =   6855
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   34
         Left            =   1800
         TabIndex        =   446
         Text            =   "Text1"
         Top             =   4680
         Width           =   1215
      End
      Begin VB.TextBox txtDescFamia 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   1
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   463
         Text            =   "Text5"
         Top             =   3000
         Width           =   3495
      End
      Begin VB.TextBox txtFamia 
         Height          =   285
         Index           =   1
         Left            =   1800
         TabIndex        =   443
         Text            =   "Text1"
         Top             =   3000
         Width           =   1095
      End
      Begin VB.TextBox txtDescFamia 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   0
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   458
         Text            =   "Text5"
         Top             =   2640
         Width           =   3495
      End
      Begin VB.TextBox txtFamia 
         Height          =   285
         Index           =   0
         Left            =   1800
         TabIndex        =   442
         Text            =   "Text1"
         Top             =   2640
         Width           =   1095
      End
      Begin VB.TextBox txtDescArticulo 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   7
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   456
         Text            =   "Text5"
         Top             =   4080
         Width           =   3495
      End
      Begin VB.TextBox txtArticulo 
         Height          =   285
         Index           =   7
         Left            =   1800
         MaxLength       =   16
         TabIndex        =   445
         Top             =   4080
         Width           =   1455
      End
      Begin VB.TextBox txtDescArticulo 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   6
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   454
         Text            =   "Text5"
         Top             =   3720
         Width           =   3495
      End
      Begin VB.TextBox txtArticulo 
         Height          =   285
         Index           =   6
         Left            =   1800
         MaxLength       =   16
         TabIndex        =   444
         Top             =   3720
         Width           =   1455
      End
      Begin VB.TextBox txtDescProve 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   14
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   452
         Text            =   "Text5"
         Top             =   1920
         Width           =   3495
      End
      Begin VB.TextBox txtCodProve 
         Height          =   285
         Index           =   14
         Left            =   1800
         TabIndex        =   441
         Text            =   "Text1"
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox txtDescProve 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   13
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   449
         Text            =   "Text5"
         Top             =   1560
         Width           =   3495
      End
      Begin VB.TextBox txtCodProve 
         Height          =   285
         Index           =   13
         Left            =   1800
         TabIndex        =   440
         Top             =   1560
         Width           =   1095
      End
      Begin VB.OptionButton optCopiaPrecio 
         Caption         =   "Desde venta"
         Enabled         =   0   'False
         Height          =   255
         Index           =   1
         Left            =   4200
         TabIndex        =   439
         Top             =   960
         Width           =   2055
      End
      Begin VB.OptionButton optCopiaPrecio 
         Caption         =   "Desde compra"
         Enabled         =   0   'False
         Height          =   255
         Index           =   0
         Left            =   1200
         TabIndex        =   438
         Top             =   960
         Value           =   -1  'True
         Width           =   2055
      End
      Begin VB.CommandButton cmdP 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   4200
         TabIndex        =   447
         Top             =   5040
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   28
         Left            =   5400
         TabIndex        =   448
         Top             =   5040
         Width           =   1095
      End
      Begin VB.Label Label3 
         Height          =   195
         Index           =   107
         Left            =   240
         TabIndex        =   580
         Top             =   5160
         Width           =   3225
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   34
         Left            =   1560
         Picture         =   "frmListado2.frx":8DF6
         Top             =   4680
         Width           =   240
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Fecha cambio"
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
         Index           =   40
         Left            =   240
         TabIndex        =   464
         Top             =   4680
         Width           =   1155
      End
      Begin VB.Image imgFamilia 
         Height          =   240
         Index           =   1
         Left            =   1560
         Picture         =   "frmListado2.frx":8E81
         Top             =   3000
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   89
         Left            =   600
         TabIndex        =   462
         Top             =   3000
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   88
         Left            =   600
         TabIndex        =   461
         Top             =   2640
         Width           =   465
      End
      Begin VB.Image imgFamilia 
         Height          =   240
         Index           =   0
         Left            =   1560
         Picture         =   "frmListado2.frx":8F83
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
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   15
         Left            =   240
         TabIndex        =   460
         Top             =   2400
         Width           =   600
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
         Index           =   14
         Left            =   240
         TabIndex        =   459
         Top             =   3480
         Width           =   660
      End
      Begin VB.Image imgArticulo 
         Height          =   240
         Index           =   7
         Left            =   1560
         Picture         =   "frmListado2.frx":9085
         Top             =   4080
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   87
         Left            =   600
         TabIndex        =   457
         Top             =   4080
         Width           =   465
      End
      Begin VB.Image imgArticulo 
         Height          =   240
         Index           =   6
         Left            =   1560
         Picture         =   "frmListado2.frx":9187
         Top             =   3720
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   86
         Left            =   600
         TabIndex        =   455
         Top             =   3720
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   85
         Left            =   600
         TabIndex        =   453
         Top             =   1920
         Width           =   465
      End
      Begin VB.Image imgProveedor 
         Height          =   240
         Index           =   14
         Left            =   1560
         Picture         =   "frmListado2.frx":9289
         Top             =   1920
         Width           =   240
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
         Index           =   13
         Left            =   240
         TabIndex        =   451
         Top             =   1320
         Width           =   885
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   84
         Left            =   600
         TabIndex        =   450
         Top             =   1560
         Width           =   465
      End
      Begin VB.Image imgProveedor 
         Height          =   240
         Index           =   13
         Left            =   1560
         Picture         =   "frmListado2.frx":938B
         Top             =   1560
         Width           =   240
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Copia precios"
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
         Index           =   26
         Left            =   2145
         TabIndex        =   437
         Top             =   360
         Width           =   1965
      End
   End
   Begin VB.Frame FrGeneraFactLiq 
      Height          =   6855
      Left            =   120
      TabIndex        =   144
      Top             =   120
      Visible         =   0   'False
      Width           =   6495
      Begin VB.TextBox txtDescForpa 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   0
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   200
         Text            =   "Text1"
         Top             =   5640
         Width           =   3615
      End
      Begin VB.TextBox txtForpa 
         Height          =   285
         Index           =   0
         Left            =   1680
         TabIndex        =   156
         Top             =   5640
         Width           =   855
      End
      Begin VB.TextBox txtTrab 
         Height          =   285
         Index           =   1
         Left            =   1680
         TabIndex        =   155
         Top             =   5160
         Width           =   855
      End
      Begin VB.TextBox txtDescTra 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   1
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   198
         Text            =   "Text1"
         Top             =   5160
         Width           =   3615
      End
      Begin VB.CheckBox chkFacturPorv 
         Caption         =   "Tesoreria"
         Height          =   195
         Index           =   1
         Left            =   4680
         TabIndex        =   153
         Top             =   4245
         Width           =   1095
      End
      Begin VB.CheckBox chkFacturPorv 
         Caption         =   "Marcar Contabilizar"
         Height          =   195
         Index           =   0
         Left            =   4680
         TabIndex        =   152
         Top             =   3840
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox txtDescBancoPr 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   0
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   195
         Text            =   "Text5"
         Top             =   4680
         Width           =   3615
      End
      Begin VB.TextBox txtBancoPr 
         Height          =   285
         Index           =   0
         Left            =   1680
         TabIndex        =   154
         Text            =   "Text1"
         Top             =   4680
         Width           =   855
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   17
         Left            =   1680
         TabIndex        =   151
         Text            =   "Text1"
         Top             =   4200
         Width           =   1215
      End
      Begin VB.TextBox txtNumAlbar 
         Height          =   285
         Index           =   1
         Left            =   4560
         TabIndex        =   148
         Text            =   "Text1"
         Top             =   1995
         Width           =   1215
      End
      Begin VB.TextBox txtNumAlbar 
         Height          =   285
         Index           =   0
         Left            =   1920
         TabIndex        =   147
         Text            =   "Text1"
         Top             =   1995
         Width           =   1215
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   14
         Left            =   4680
         TabIndex        =   146
         Text            =   "Text1"
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   13
         Left            =   1920
         TabIndex        =   145
         Text            =   "Text1"
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtDescProve 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   5
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   160
         Text            =   "Text5"
         Top             =   3240
         Width           =   3495
      End
      Begin VB.TextBox txtCodProve 
         Height          =   285
         Index           =   5
         Left            =   1200
         TabIndex        =   150
         Text            =   "Text1"
         Top             =   3240
         Width           =   1455
      End
      Begin VB.TextBox txtDescProve 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   4
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   159
         Text            =   "Text5"
         Top             =   2880
         Width           =   3495
      End
      Begin VB.TextBox txtCodProve 
         Height          =   285
         Index           =   4
         Left            =   1200
         TabIndex        =   149
         Top             =   2880
         Width           =   1455
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   8
         Left            =   5040
         TabIndex        =   158
         Top             =   6240
         Width           =   1215
      End
      Begin VB.CommandButton cmdFacProv 
         Caption         =   "Generar"
         Height          =   375
         Left            =   3600
         TabIndex        =   157
         Top             =   6240
         Width           =   1215
      End
      Begin VB.Image imgForPa 
         Height          =   240
         Index           =   0
         Left            =   1320
         Picture         =   "frmListado2.frx":948D
         Top             =   5640
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Forma pago"
         Height          =   195
         Index           =   38
         Left            =   240
         TabIndex        =   201
         Top             =   5640
         Width           =   945
      End
      Begin VB.Label Label3 
         Caption         =   "Operador"
         Height          =   195
         Index           =   37
         Left            =   240
         TabIndex        =   199
         Top             =   5160
         Width           =   945
      End
      Begin VB.Image imgTecnico 
         Height          =   240
         Index           =   1
         Left            =   1320
         Picture         =   "frmListado2.frx":958F
         Top             =   5160
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Banco propio"
         Height          =   195
         Index           =   36
         Left            =   240
         TabIndex        =   197
         Top             =   4680
         Width           =   945
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha"
         Height          =   195
         Index           =   35
         Left            =   240
         TabIndex        =   196
         Top             =   4200
         Width           =   465
      End
      Begin VB.Image imgBancoPr 
         Height          =   240
         Index           =   0
         Left            =   1320
         Picture         =   "frmListado2.frx":9691
         Top             =   4680
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   17
         Left            =   1320
         Picture         =   "frmListado2.frx":9793
         Top             =   4200
         Width           =   240
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Datos facturación"
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
         Index           =   17
         Left            =   120
         TabIndex        =   194
         Top             =   3840
         Width           =   1500
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000080&
         BorderWidth     =   3
         X1              =   120
         X2              =   6240
         Y1              =   3720
         Y2              =   3720
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   28
         Left            =   3720
         TabIndex        =   171
         Top             =   2040
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   27
         Left            =   960
         TabIndex        =   170
         Top             =   2040
         Width           =   465
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Num. albaran"
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
         Left            =   120
         TabIndex        =   169
         Top             =   1560
         Width           =   1110
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Generar facturas Liq. proveedores"
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
         Index           =   7
         Left            =   240
         TabIndex        =   168
         Top             =   240
         Width           =   5415
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Fecha albaran"
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
         Left            =   120
         TabIndex        =   167
         Top             =   720
         Width           =   1185
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   26
         Left            =   960
         TabIndex        =   166
         Top             =   1125
         Width           =   465
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   14
         Left            =   4440
         Picture         =   "frmListado2.frx":981E
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   25
         Left            =   3720
         TabIndex        =   165
         Top             =   1125
         Width           =   465
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   13
         Left            =   1680
         Picture         =   "frmListado2.frx":98A9
         Top             =   1095
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   24
         Left            =   240
         TabIndex        =   164
         Top             =   3240
         Width           =   465
      End
      Begin VB.Image imgProveedor 
         Height          =   240
         Index           =   5
         Left            =   840
         Picture         =   "frmListado2.frx":9934
         Top             =   3240
         Width           =   240
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
         Index           =   10
         Left            =   120
         TabIndex        =   163
         Top             =   2520
         Width           =   885
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   23
         Left            =   240
         TabIndex        =   162
         Top             =   2880
         Width           =   465
      End
      Begin VB.Image imgProveedor 
         Height          =   240
         Index           =   4
         Left            =   840
         Picture         =   "frmListado2.frx":9A36
         Top             =   2880
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   375
         Left            =   120
         TabIndex        =   161
         Top             =   6360
         Width           =   3375
      End
   End
   Begin VB.Frame FrameLlamadas 
      Height          =   3975
      Left            =   360
      TabIndex        =   342
      Top             =   0
      Visible         =   0   'False
      Width           =   5895
      Begin VB.TextBox txtDescTra 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   6
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   355
         Text            =   "Text1"
         Top             =   2640
         Width           =   3495
      End
      Begin VB.TextBox txtTrab 
         Height          =   285
         Index           =   6
         Left            =   1200
         TabIndex        =   347
         Top             =   2640
         Width           =   855
      End
      Begin VB.TextBox txtDescTra 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   5
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   353
         Text            =   "Text1"
         Top             =   2280
         Width           =   3495
      End
      Begin VB.TextBox txtTrab 
         Height          =   285
         Index           =   5
         Left            =   1200
         TabIndex        =   346
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   28
         Left            =   3960
         TabIndex        =   345
         Text            =   "Text1"
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   27
         Left            =   1320
         TabIndex        =   344
         Text            =   "Text1"
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton cmdLlamadas 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3240
         TabIndex        =   348
         Top             =   3120
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   22
         Left            =   4440
         TabIndex        =   349
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Trabajadores"
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
         Index           =   29
         Left            =   120
         TabIndex        =   357
         Top             =   1920
         Width           =   1140
      End
      Begin VB.Image imgTecnico 
         Height          =   240
         Index           =   6
         Left            =   840
         Picture         =   "frmListado2.frx":9B38
         Top             =   2640
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   71
         Left            =   240
         TabIndex        =   356
         Top             =   2640
         Width           =   465
      End
      Begin VB.Image imgTecnico 
         Height          =   240
         Index           =   5
         Left            =   840
         Picture         =   "frmListado2.frx":9C3A
         Top             =   2280
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   70
         Left            =   240
         TabIndex        =   354
         Top             =   2280
         Width           =   465
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Fecha "
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
         Left            =   120
         TabIndex        =   352
         Top             =   960
         Width           =   540
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   28
         Left            =   3600
         Picture         =   "frmListado2.frx":9D3C
         Top             =   1320
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   69
         Left            =   3000
         TabIndex        =   351
         Top             =   1365
         Width           =   465
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   27
         Left            =   960
         Picture         =   "frmListado2.frx":9DC7
         Top             =   1320
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   68
         Left            =   360
         TabIndex        =   350
         Top             =   1320
         Width           =   465
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Listado Llamadas"
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
         Index           =   20
         Left            =   1320
         TabIndex        =   343
         Top             =   360
         Width           =   2925
      End
   End
   Begin VB.Frame FrameRecargaMov 
      Height          =   3375
      Left            =   0
      TabIndex        =   44
      Top             =   0
      Visible         =   0   'False
      Width           =   6495
      Begin VB.TextBox txtRecargaMov 
         Height          =   285
         Index           =   0
         Left            =   5640
         MaxLength       =   1
         TabIndex        =   49
         Text            =   "Text1"
         Top             =   1995
         Width           =   375
      End
      Begin VB.ComboBox cmbRecargaMov 
         Height          =   315
         Index           =   1
         ItemData        =   "frmListado2.frx":9E52
         Left            =   3840
         List            =   "frmListado2.frx":9E5F
         Style           =   2  'Dropdown List
         TabIndex        =   48
         Top             =   1980
         Width           =   975
      End
      Begin VB.ComboBox cmbRecargaMov 
         Height          =   315
         Index           =   0
         ItemData        =   "frmListado2.frx":9E72
         Left            =   1800
         List            =   "frmListado2.frx":9E7F
         Style           =   2  'Dropdown List
         TabIndex        =   47
         Top             =   1980
         Width           =   975
      End
      Begin VB.CommandButton cmdRecargasMov 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3960
         TabIndex        =   50
         Top             =   2760
         Width           =   975
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   5
         Left            =   4680
         TabIndex        =   46
         Text            =   "Text1"
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   4
         Left            =   1800
         TabIndex        =   45
         Text            =   "Text1"
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   4
         Left            =   5160
         TabIndex        =   52
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo"
         Height          =   195
         Index           =   8
         Left            =   5160
         TabIndex        =   58
         Top             =   2040
         Width           =   465
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Cobradas"
         Height          =   195
         Index           =   7
         Left            =   2880
         TabIndex        =   57
         Top             =   2040
         Width           =   825
      End
      Begin VB.Label Label3 
         Caption         =   "Facturadas"
         Height          =   195
         Index           =   6
         Left            =   840
         TabIndex        =   56
         Top             =   2040
         Width           =   825
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Informe recargas moviles"
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
         Index           =   3
         Left            =   600
         TabIndex        =   55
         Top             =   360
         Width           =   5415
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   5
         Left            =   3840
         TabIndex        =   54
         Top             =   1200
         Width           =   465
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   5
         Left            =   4440
         Picture         =   "frmListado2.frx":9E92
         Top             =   1200
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
         Index           =   5
         Left            =   240
         TabIndex        =   53
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   4
         Left            =   960
         TabIndex        =   51
         Top             =   1200
         Width           =   465
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   4
         Left            =   1560
         Picture         =   "frmListado2.frx":9F1D
         Top             =   1200
         Width           =   240
      End
   End
   Begin VB.Frame FrameCambioProve 
      Height          =   2415
      Left            =   0
      TabIndex        =   301
      Top             =   0
      Visible         =   0   'False
      Width           =   6255
      Begin VB.CommandButton Command1 
         Caption         =   "Cambiar"
         Height          =   375
         Left            =   3240
         TabIndex        =   304
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   18
         Left            =   4920
         TabIndex        =   305
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox txtCodProve 
         Height          =   285
         Index           =   12
         Left            =   1440
         TabIndex        =   303
         Text            =   "Text1"
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox txtDescProve 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   12
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   306
         Text            =   "Text5"
         Top             =   1080
         Width           =   3495
      End
      Begin VB.Image imgProveedor 
         Height          =   240
         Index           =   12
         Left            =   1200
         Picture         =   "frmListado2.frx":9FA8
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Proveedor"
         Height          =   195
         Index           =   60
         Left            =   240
         TabIndex        =   307
         Top             =   1080
         Width           =   1065
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Cambio proveedor albarán compra"
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
         Index           =   16
         Left            =   480
         TabIndex        =   302
         Top             =   360
         Width           =   5415
      End
   End
   Begin VB.Frame FrameMultibase 
      Height          =   5295
      Left            =   120
      TabIndex        =   38
      Top             =   120
      Width           =   5775
      Begin VB.ComboBox cboRoot 
         Height          =   315
         ItemData        =   "frmListado2.frx":A0AA
         Left            =   120
         List            =   "frmListado2.frx":A0B7
         Style           =   2  'Dropdown List
         TabIndex        =   763
         Top             =   4800
         Width           =   1815
      End
      Begin VB.Frame FrameErrorRestore 
         Height          =   4215
         Left            =   120
         TabIndex        =   761
         Top             =   480
         Visible         =   0   'False
         Width           =   5415
         Begin MSComctlLib.TreeView TreeView1 
            Height          =   3855
            Left            =   120
            TabIndex        =   762
            Top             =   240
            Width           =   4575
            _ExtentX        =   8070
            _ExtentY        =   6800
            _Version        =   393217
            LabelEdit       =   1
            Style           =   7
            Checkboxes      =   -1  'True
            Appearance      =   1
         End
         Begin VB.Image imgCheck 
            Height          =   240
            Index           =   5
            Left            =   4920
            Picture         =   "frmListado2.frx":A0E3
            ToolTipText     =   "Todos"
            Top             =   600
            Width           =   240
         End
         Begin VB.Image imgCheck 
            Height          =   240
            Index           =   4
            Left            =   4920
            Picture         =   "frmListado2.frx":A22D
            ToolTipText     =   "Quitar seleccion"
            Top             =   240
            Width           =   240
         End
      End
      Begin VB.CommandButton cmdMultibase2 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3240
         TabIndex        =   269
         Top             =   4800
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Frame FrameTablas 
         Height          =   3375
         Left            =   120
         TabIndex        =   264
         Top             =   960
         Visible         =   0   'False
         Width           =   5295
         Begin VB.ComboBox cboCampos 
            Height          =   315
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   268
            Top             =   1560
            Width           =   2895
         End
         Begin VB.ComboBox cboTablas 
            Height          =   315
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   266
            Top             =   480
            Width           =   2895
         End
         Begin VB.Label Label7 
            Caption         =   "TABLAS"
            Height          =   255
            Left            =   240
            TabIndex        =   267
            Top             =   1560
            Width           =   3255
         End
         Begin VB.Label Label6 
            Caption         =   "TABLAS"
            Height          =   255
            Left            =   240
            TabIndex        =   265
            Top             =   480
            Width           =   3255
         End
      End
      Begin VB.ListBox lstMultibase 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3300
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   42
         Top             =   960
         Width           =   5295
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   3
         Left            =   4440
         TabIndex        =   40
         Top             =   4800
         Width           =   975
      End
      Begin VB.CommandButton cmdMultibase 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3240
         TabIndex        =   39
         Top             =   4800
         Width           =   975
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Revisar caracteres especiales"
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
         Index           =   2
         Left            =   360
         TabIndex        =   41
         Top             =   120
         Width           =   4815
      End
      Begin VB.Label lblMultibase 
         Height          =   375
         Left            =   120
         TabIndex        =   43
         Top             =   4800
         Width           =   2895
      End
   End
   Begin VB.Frame FrameDtoCompra 
      Height          =   5175
      Left            =   120
      TabIndex        =   524
      Top             =   0
      Visible         =   0   'False
      Width           =   6855
      Begin VB.CheckBox chkDtoCompra 
         Caption         =   "Salto pag. proveedor"
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   533
         Top             =   4560
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.CheckBox chkDtoCompra 
         Caption         =   "Solo con rappel"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   532
         Top             =   4560
         Width           =   1815
      End
      Begin VB.CommandButton cmdDtoProve 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   4200
         TabIndex        =   534
         Top             =   4560
         Width           =   1095
      End
      Begin VB.TextBox txtmarca 
         Height          =   285
         Index           =   3
         Left            =   1800
         TabIndex        =   531
         Text            =   "Text1"
         Top             =   3960
         Width           =   1095
      End
      Begin VB.TextBox txtDescmarca 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   3
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   549
         Text            =   "Text5"
         Top             =   3960
         Width           =   3495
      End
      Begin VB.TextBox txtmarca 
         Height          =   285
         Index           =   2
         Left            =   1800
         TabIndex        =   530
         Text            =   "Text1"
         Top             =   3600
         Width           =   1095
      End
      Begin VB.TextBox txtDescmarca 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   2
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   546
         Text            =   "Text5"
         Top             =   3600
         Width           =   3495
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   33
         Left            =   5400
         TabIndex        =   535
         Top             =   4560
         Width           =   1095
      End
      Begin VB.TextBox txtFamia 
         Height          =   285
         Index           =   5
         Left            =   1800
         TabIndex        =   529
         Text            =   "Text1"
         Top             =   2880
         Width           =   1095
      End
      Begin VB.TextBox txtDescFamia 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   5
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   544
         Text            =   "Text5"
         Top             =   2880
         Width           =   3495
      End
      Begin VB.TextBox txtFamia 
         Height          =   285
         Index           =   4
         Left            =   1800
         TabIndex        =   528
         Text            =   "Text1"
         Top             =   2520
         Width           =   1095
      End
      Begin VB.TextBox txtDescFamia 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   4
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   542
         Text            =   "Text5"
         Top             =   2520
         Width           =   3495
      End
      Begin VB.TextBox txtCodProve 
         Height          =   285
         Index           =   19
         Left            =   1800
         TabIndex        =   527
         Text            =   "Text1"
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox txtDescProve 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   19
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   539
         Text            =   "Text5"
         Top             =   1800
         Width           =   3495
      End
      Begin VB.TextBox txtCodProve 
         Height          =   285
         Index           =   18
         Left            =   1800
         TabIndex        =   526
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox txtDescProve 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   18
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   536
         Text            =   "Text5"
         Top             =   1440
         Width           =   3495
      End
      Begin VB.Image imgMarca 
         Height          =   240
         Index           =   3
         Left            =   1560
         Picture         =   "frmListado2.frx":A377
         Top             =   3960
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   106
         Left            =   600
         TabIndex        =   550
         Top             =   3960
         Width           =   465
      End
      Begin VB.Image imgMarca 
         Height          =   240
         Index           =   2
         Left            =   1560
         Picture         =   "frmListado2.frx":A479
         Top             =   3600
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Marca"
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
         Index           =   26
         Left            =   240
         TabIndex        =   548
         Top             =   3360
         Width           =   525
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   105
         Left            =   600
         TabIndex        =   547
         Top             =   3600
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   104
         Left            =   600
         TabIndex        =   545
         Top             =   2880
         Width           =   465
      End
      Begin VB.Image imgFamilia 
         Height          =   240
         Index           =   5
         Left            =   1560
         Picture         =   "frmListado2.frx":A57B
         Top             =   2880
         Width           =   240
      End
      Begin VB.Image imgFamilia 
         Height          =   240
         Index           =   4
         Left            =   1560
         Picture         =   "frmListado2.frx":A67D
         Top             =   2520
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   103
         Left            =   600
         TabIndex        =   543
         Top             =   2520
         Width           =   465
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
         Index           =   25
         Left            =   240
         TabIndex        =   541
         Top             =   2280
         Width           =   600
      End
      Begin VB.Image imgProveedor 
         Height          =   240
         Index           =   19
         Left            =   1560
         Picture         =   "frmListado2.frx":A77F
         Top             =   1800
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   102
         Left            =   600
         TabIndex        =   540
         Top             =   1800
         Width           =   465
      End
      Begin VB.Image imgProveedor 
         Height          =   240
         Index           =   18
         Left            =   1560
         Picture         =   "frmListado2.frx":A881
         Top             =   1440
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   101
         Left            =   600
         TabIndex        =   538
         Top             =   1440
         Width           =   465
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
         Index           =   24
         Left            =   240
         TabIndex        =   537
         Top             =   1200
         Width           =   885
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Listados de descuentos proveedor"
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
         Index           =   30
         Left            =   480
         TabIndex        =   525
         Top             =   480
         Width           =   5085
      End
   End
   Begin VB.Frame FramePromociones 
      Height          =   6015
      Left            =   4680
      TabIndex        =   674
      Top             =   0
      Visible         =   0   'False
      Width           =   6735
      Begin VB.CommandButton cmdACtualizaPromo 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3960
         TabIndex        =   683
         Top             =   5400
         Width           =   1095
      End
      Begin VB.CommandButton cmdCambioPromo 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3840
         TabIndex        =   682
         Top             =   5400
         Width           =   1095
      End
      Begin VB.TextBox txtFamia 
         Height          =   285
         Index           =   13
         Left            =   1680
         TabIndex        =   681
         Text            =   "Text1"
         Top             =   3960
         Width           =   1095
      End
      Begin VB.TextBox txtDescFamia 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   13
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   693
         Text            =   "Text5"
         Top             =   3960
         Width           =   3495
      End
      Begin VB.TextBox txtFamia 
         Height          =   285
         Index           =   12
         Left            =   1680
         TabIndex        =   680
         Text            =   "Text1"
         Top             =   3600
         Width           =   1095
      End
      Begin VB.TextBox txtDescFamia 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   12
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   692
         Text            =   "Text5"
         Top             =   3600
         Width           =   3495
      End
      Begin VB.TextBox txtCodProve 
         Height          =   285
         Index           =   23
         Left            =   1680
         TabIndex        =   679
         Top             =   2880
         Width           =   1095
      End
      Begin VB.TextBox txtDescProve 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   23
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   690
         Text            =   "Text5"
         Top             =   2880
         Width           =   3495
      End
      Begin VB.TextBox txtTarifa 
         Height          =   285
         Index           =   1
         Left            =   1680
         TabIndex        =   678
         Text            =   "Text1"
         Top             =   2280
         Width           =   975
      End
      Begin VB.TextBox txtDescTarifa 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   1
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   688
         Text            =   "Text5"
         Top             =   2280
         Width           =   3495
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   42
         Left            =   1680
         TabIndex        =   677
         Text            =   "Text1"
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   41
         Left            =   1680
         TabIndex        =   676
         Text            =   "Text1"
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   38
         Left            =   5160
         TabIndex        =   684
         Top             =   5400
         Width           =   1095
      End
      Begin VB.Image imgFamilia 
         Height          =   240
         Index           =   13
         Left            =   1440
         Picture         =   "frmListado2.frx":A983
         Top             =   3960
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   134
         Left            =   480
         TabIndex        =   696
         Top             =   3600
         Width           =   465
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
         Index           =   38
         Left            =   120
         TabIndex        =   695
         Top             =   3360
         Width           =   600
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   133
         Left            =   480
         TabIndex        =   694
         Top             =   3960
         Width           =   465
      End
      Begin VB.Image imgFamilia 
         Height          =   240
         Index           =   12
         Left            =   1440
         Picture         =   "frmListado2.frx":AA85
         Top             =   3600
         Width           =   240
      End
      Begin VB.Image imgProveedor 
         Height          =   240
         Index           =   23
         Left            =   1440
         Picture         =   "frmListado2.frx":AB87
         Top             =   2880
         Width           =   240
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
         Index           =   37
         Left            =   120
         TabIndex        =   691
         Top             =   2880
         Width           =   885
      End
      Begin VB.Image imgTarifa 
         Height          =   240
         Index           =   1
         Left            =   1440
         Picture         =   "frmListado2.frx":AC89
         Top             =   2280
         Width           =   240
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Tarifa"
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
         Index           =   48
         Left            =   120
         TabIndex        =   689
         Top             =   2280
         Width           =   495
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   42
         Left            =   1320
         Picture         =   "frmListado2.frx":AD8B
         Top             =   1680
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   132
         Left            =   720
         TabIndex        =   687
         Top             =   1320
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   131
         Left            =   720
         TabIndex        =   686
         Top             =   1680
         Width           =   465
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Fecha  promoción"
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
         Index           =   47
         Left            =   120
         TabIndex        =   685
         Top             =   960
         Width           =   1485
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   41
         Left            =   1320
         Picture         =   "frmListado2.frx":AE16
         Top             =   1320
         Width           =   240
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "l"
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
         Index           =   35
         Left            =   480
         TabIndex        =   675
         Top             =   360
         Width           =   5085
      End
   End
   Begin VB.Frame FrameDtosActiv 
      Height          =   4335
      Left            =   3960
      TabIndex        =   581
      Top             =   0
      Visible         =   0   'False
      Width           =   6975
      Begin VB.TextBox txtDescFamia 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   9
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   598
         Text            =   "Text5"
         Top             =   2880
         Width           =   3495
      End
      Begin VB.TextBox txtFamia 
         Height          =   285
         Index           =   9
         Left            =   1920
         TabIndex        =   585
         Text            =   "Text1"
         Top             =   2880
         Width           =   1095
      End
      Begin VB.CommandButton cmdDtoActiv 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   4200
         TabIndex        =   586
         Top             =   3720
         Width           =   975
      End
      Begin VB.TextBox txtDescActiv 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   1
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   596
         Text            =   "Text5"
         Top             =   1680
         Width           =   3615
      End
      Begin VB.TextBox txtcodactiv 
         Height          =   285
         Index           =   1
         Left            =   1920
         TabIndex        =   583
         Text            =   "Text1"
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox txtDescActiv 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   0
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   593
         Text            =   "Text5"
         Top             =   1320
         Width           =   3615
      End
      Begin VB.TextBox txtcodactiv 
         Height          =   285
         Index           =   0
         Left            =   1920
         TabIndex        =   582
         Text            =   "Text1"
         Top             =   1320
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   35
         Left            =   5400
         TabIndex        =   588
         Top             =   3720
         Width           =   1095
      End
      Begin VB.TextBox txtFamia 
         Height          =   285
         Index           =   8
         Left            =   1920
         TabIndex        =   584
         Text            =   "Text1"
         Top             =   2400
         Width           =   1095
      End
      Begin VB.TextBox txtDescFamia 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   8
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   587
         Text            =   "Text5"
         Top             =   2400
         Width           =   3495
      End
      Begin VB.Image imgFamilia 
         Height          =   240
         Index           =   9
         Left            =   1680
         Picture         =   "frmListado2.frx":AEA1
         Top             =   2880
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   117
         Left            =   720
         TabIndex        =   597
         Top             =   1680
         Width           =   465
      End
      Begin VB.Image imgActividad 
         Height          =   240
         Index           =   1
         Left            =   1680
         Picture         =   "frmListado2.frx":AFA3
         Top             =   1680
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   116
         Left            =   720
         TabIndex        =   595
         Top             =   1320
         Width           =   465
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
         Index           =   32
         Left            =   360
         TabIndex        =   594
         Top             =   960
         Width           =   795
      End
      Begin VB.Image imgActividad 
         Height          =   240
         Index           =   0
         Left            =   1680
         Picture         =   "frmListado2.frx":B0A5
         Top             =   1320
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   113
         Left            =   720
         TabIndex        =   592
         Top             =   2880
         Width           =   465
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Descuentos actividad"
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
         Index           =   32
         Left            =   2040
         TabIndex        =   591
         Top             =   360
         Width           =   3015
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
         Index           =   30
         Left            =   360
         TabIndex        =   590
         Top             =   2160
         Width           =   600
      End
      Begin VB.Image imgFamilia 
         Height          =   240
         Index           =   8
         Left            =   1680
         Picture         =   "frmListado2.frx":B1A7
         Top             =   2400
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   112
         Left            =   720
         TabIndex        =   589
         Top             =   2400
         Width           =   465
      End
   End
   Begin VB.Frame FrameCerrarAviso 
      Height          =   4215
      Left            =   120
      TabIndex        =   308
      Top             =   120
      Visible         =   0   'False
      Width           =   7935
      Begin VB.CommandButton cmdGenAlbRep 
         Caption         =   "Gen Albaran"
         Height          =   375
         Left            =   5160
         TabIndex        =   316
         Top             =   3600
         Width           =   1215
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   26
         Left            =   1920
         TabIndex        =   310
         Text            =   "Text1"
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox texto 
         Height          =   280
         Index           =   3
         Left            =   240
         MaxLength       =   80
         TabIndex        =   314
         Tag             =   "Observación 1|T|S|||scappr|observa1||N|"
         Top             =   2640
         Width           =   7485
      End
      Begin VB.TextBox texto 
         Height          =   280
         Index           =   2
         Left            =   240
         MaxLength       =   80
         TabIndex        =   313
         Tag             =   "Observación 1|T|S|||scappr|observa1||N|"
         Top             =   2280
         Width           =   7485
      End
      Begin VB.TextBox texto 
         Height          =   280
         Index           =   1
         Left            =   240
         MaxLength       =   80
         TabIndex        =   312
         Tag             =   "Observación 1|T|S|||scappr|observa1||N|"
         Top             =   1920
         Width           =   7485
      End
      Begin VB.TextBox texto 
         Height          =   280
         Index           =   0
         Left            =   240
         MaxLength       =   80
         TabIndex        =   311
         Tag             =   "Observación 1|T|S|||scappr|observa1||N|"
         Top             =   1560
         Width           =   7485
      End
      Begin VB.TextBox texto 
         Height          =   280
         Index           =   4
         Left            =   240
         MaxLength       =   80
         TabIndex        =   315
         Tag             =   "Observación 1|T|S|||scappr|observa1||N|"
         Top             =   3000
         Width           =   7485
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   19
         Left            =   6600
         TabIndex        =   317
         Top             =   3600
         Width           =   1095
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   26
         Left            =   1560
         Picture         =   "frmListado2.frx":B2A9
         Top             =   840
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha albaran"
         Height          =   195
         Index           =   62
         Left            =   240
         TabIndex        =   319
         Top             =   840
         Width           =   1185
      End
      Begin VB.Label Label3 
         Caption         =   "Observaciones"
         Height          =   195
         Index           =   61
         Left            =   240
         TabIndex        =   318
         Top             =   1320
         Width           =   1065
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Cerrar aviso"
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
         Index           =   17
         Left            =   960
         TabIndex        =   309
         Top             =   120
         Width           =   5415
      End
   End
End
Attribute VB_Name = "frmListado2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public Opcion As Integer
    '1      .- Listado reparaciones efectuadas
    '2      .- Reparaciones tecnico
    
    '3      .- Revision carcteres multibase
    '4      .- Listado recargas telefonia movil
    '5      .- Facturacion de recargas
    
    '6      .- Listado de TRAZA por codprove en ventas.   ENERO 2008
    
    '       LIQUIDACION PROVEEDORES. Socios tipo TERRASANA
    '7      .- Cambio precio articulos
    '8      .- Generar facturas
    '9      .- Imprimir facturas proveedores (socios)
    '10     .-   "      ALBARANES   "           "
    
    
    '13     .- Generacion y facturacion de tickets agrupados
    '14     .- Listado del punto anterior
    
    '15     .- Listado trazabilidad albaranes
        
    '16     .- Ventas x agentes
    
    '17     .- Listado trabajadores . NO HACE DESDE HASTA
    
    '18     .- Cambio de proveedor en albaranes. Solicita el codprove
    
    '19     .- Cerrar aviso. Datos para crear albaran
    
    '20     .- Listado plantillas ofertas
    
    
    '21     .- Seleccionar otras ofertas del cliente
    '22     .- Listado llamadas
    
    '23     .- Listado situacion albaranes
    '24     .- Modificar expediente y legal en frecuencias
    
    '25     .- Datos para facturacion de cliente
    
    '26     .- Impresion pedidos por zona
        
    '27     .- REIMPRESION DE ALBARANES
    '28     .- Copiar precios compra-venta y venta-compra
    
    '30     .- Reparaciones en garantia de proveedor
    
    '31     .- Calculo del riesgo
    '32     .- Propuesta de pedido
    
    '33     .- Rp descuentos proveedor
    
    '34     .- Modificacion precios. Para pedir d/h tarifa, provee, familia tipo...
    '35     .- Dtos por activi marca fam
    
    '36     .- Resumen ventas agente
    
    '37     .- Beneficios por agente
    
    '38     .- Genera promociones.  D/H varios y lanzar frm
    '39     .-   actuali promo.   Comparte FRAME
    
    '40     .- Beneficios por proveedor. Igual que el de agente(COMPARTE FRAME)
    '41     .- Beneficio por cliente
    
                    
    '42     .- Listado control de albaranes  05/MAYO/2014
    '43     .-   "        "     "" facturados 15 Mayo
    
    '44     .- Cambio de precios lista proveedor  (igual que el de clientes)
    
    '45     .- Cambia departamento en frecencias (TEINSA)
    '46     .- Informe productividad
    '47     .- Cliente potencial. Pasar a cliente
    
    '48     .- beneficio Marca, Agente, Proveedor

    '49     .- ventas marca-familia
    '50     .- Compras marca-familia
    
Private IndiceImg As Integer
Private WithEvents frmCli As frmFacClientes3
Attribute frmCli.VB_VarHelpID = -1
Private WithEvents frmT As frmAdmTrabajadores
Attribute frmT.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmMtoArticulos As frmAlmArticu2
Attribute frmMtoArticulos.VB_VarHelpID = -1
Private WithEvents frmPr As frmComProveedores
Attribute frmPr.VB_VarHelpID = -1
Private WithEvents frmBaPr As frmFacBancosPropios
Attribute frmBaPr.VB_VarHelpID = -1
Private WithEvents frmFP As frmFacFormasPago
Attribute frmFP.VB_VarHelpID = -1
Private WithEvents frmAg As frmFacAgentesCom
Attribute frmAg.VB_VarHelpID = -1
Private WithEvents frmZ As frmFacZonas
Attribute frmZ.VB_VarHelpID = -1
Private WithEvents frmEn As frmFacFormasEnvio
Attribute frmEn.VB_VarHelpID = -1
Private WithEvents frmRut As frmFacRutas
Attribute frmRut.VB_VarHelpID = -1

Private PrimeraVez As Boolean




'---- Variables para el INFORME ----
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private cadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para el frmImprimir
Private cadNomRPT As String 'Nombre del informe a Imprimir
Private conSubRPT As Boolean 'Si el informe tiene subreports

'nuevo Febrero 2010
Private cadPDFrpt As String 'Nombre del informe a enviar por email
Private vImprimedirecto As Boolean '
Private vMultiInforme As Integer
'-----------------------------------





'Variables comunes a todos os botones aceptar
Dim cadFrom As String
Dim campo As String, devuelve As String
Dim codigo  As String
Dim ImpTot As Currency
Dim ImpTeo As Currency
Dim miSQL As String

Private Cadena_frmB As String
Private cadImpresion As String  'Facturacion




Private Sub cboCoste_KeyPress(Index As Integer, KeyAscii As Integer)
KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub cboProPed_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub cboRoot_Click()
    FrameErrorRestore.visible = cboRoot.ListIndex = 2
    FrameTablas.visible = cboRoot.ListIndex = 1
    cmdMultibase2.visible = cboRoot.ListIndex > 0
    cmdMultibase.visible = cboRoot.ListIndex = 0
    If cboRoot.ListIndex > 0 Then
        
        Screen.MousePointer = vbHourglass
        Me.lblMultibase.Caption = "               Cargando datos"
        Me.lblMultibase.Refresh
    
        If cboRoot.ListIndex = 1 Then
            If Me.cboTablas.ListCount = 0 Then CargaTablasCambio
            
        Else
            If Me.TreeView1.Nodes.Count = 0 Then CargaArbolTablas
        End If
        
        Screen.MousePointer = vbDefault
        Me.lblMultibase.Caption = ""
    End If
End Sub

Private Sub cboTablas_Click()
    cboCampos.Clear
    If cboTablas.ListIndex < 0 Then Exit Sub
    CargarCamposTabla
End Sub



Private Sub chkAgentes_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub chkBenAge_Click(Index As Integer)
    If Index = 5 Then
        If chkBenAge(5).Value = 1 Then
            chkBenAge(3).Caption = "detalla marca"
        Else
            chkBenAge(3).Caption = "detalla cliente"
        End If
    End If
End Sub

Private Sub chkBenAge_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub chkBeneMarcaAgen_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub chkDtoCompra_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub chkFacturPorv_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub



Private Sub chkImpAlbRut_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub chkInformeProd_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub chkMarcaFamilia_KeyPress(Index As Integer, KeyAscii As Integer)
     KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub chkPedxZona_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub chkPropPedido_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 3, False
End Sub

Private Sub chkResVtaAgen_Click(Index As Integer)
    If Index = 4 Then lblDpto(44).Caption = IIf(Me.chkResVtaAgen(4).Value = 1, "Visitador", "Agente")
End Sub

Private Sub chkResVtaAgen_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 3, False
End Sub



Private Sub chkVtaxProv_Click(Index As Integer)
    If Index = 0 Then chkVtaxProv(3).visible = chkVtaxProv(0).Value = 1
End Sub

Private Sub chkVtaxProv_KeyPress(Index As Integer, KeyAscii As Integer)
     KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub cmbRecargaMov_KeyPress(Index As Integer, KeyAscii As Integer)
     KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub cmdAceptarOfertas_Click()
    miSQL = ""
    For NumRegElim = 1 To lw1.ListItems.Count
        If lw1.ListItems(NumRegElim).Checked Then miSQL = miSQL & ", " & lw1.ListItems(NumRegElim).Text
    Next NumRegElim
    If miSQL = "" Then
        MsgBox "Selecciona alguna oferta", vbExclamation
        Exit Sub
    End If
    CadenaDesdeOtroForm = miSQL
    Unload Me
End Sub

Private Sub cmdACtualizaPromo_Click()
    'Vamos a actualizar los precios promo
    'de siguiente a actual
    campo = ""
    If Me.txtTarifa(1).Text = "" Then campo = campo & vbCrLf & "-Tarifa"
    If campo <> "" Then
        campo = "Campos obligatorios: " & campo
        MsgBox campo, vbExclamation
        PonerFoco txtFecha(41)
        Exit Sub
    End If
    
    InicializarVbles
    
    cadSelect = "{spromo.codlista} = " & txtTarifa(1).Text
    'Proveedor  18 19
    If txtCodProve(23).Text <> "" Or txtCodProve(23).Text <> "" Then
        'devuelve = "pDHProve=""Proveedor: "
        campo = "{sartic.codprove}"
        If Not PonerDesdeHasta(campo, "PRO", 23, 23, devuelve) Then Exit Sub
        
    End If
    
    
    'Familia  4 5
    If txtFamia(12).Text <> "" Or txtFamia(13).Text <> "" Then
        'devuelve = "pDHFamilia=""Familia: "
        campo = "{sartic.codfamia}"
        If Not PonerDesdeHasta(campo, "FAM", 12, 13, devuelve) Then Exit Sub
        
    End If
    
    cadSelect = QuitarCaracterACadena(cadSelect, "{")
    cadSelect = QuitarCaracterACadena(cadSelect, "}")
    
  
   'Insert into
    Set miRsAux = New ADODB.Recordset
    Screen.MousePointer = vbHourglass
    NumRegElim = 0
    campo = "select count(*) from "
    campo = campo & " spromo,sartic "
    campo = campo & " WHERE spromo.codartic=sartic.codartic AND " & cadSelect
    campo = campo & " AND not  fechain1  is null"
    miRsAux.Open campo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then NumRegElim = DBLet(miRsAux.Fields(0), "N")
    miRsAux.Close
    If NumRegElim = 0 Then
        MsgBox "No hay valores con esos parametros", vbInformation
    Else
        campo = "Se van a actualizar " & NumRegElim & " registro(s). ¿Continuar?"
        If MsgBox(campo, vbQuestion + vbYesNo) = vbNo Then NumRegElim = 0
    End If
    
    
    If NumRegElim > 0 Then
        campo = "select spromo.* from "
        campo = campo & " spromo,sartic "
        campo = campo & " WHERE spromo.codartic=sartic.codartic AND " & cadSelect
        campo = campo & " AND not  fechain1  is null" 'Vale, actualizo
        miRsAux.Open campo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            
            campo = "UPDATE spromo set "
            campo = campo & " fechaini = " & DBSet(miRsAux!fechain1, "F") & " , fechafin = " & DBSet(miRsAux!fechafi1, "F")
            campo = campo & " , precioac = " & DBSet(miRsAux!precionu, "N") & " , precioa1 = " & DBSet(miRsAux!precion1, "N") & ""
            campo = campo & " ,fechain1 = null , fechafi1 = null , precionu = null , precion1 = null"
            campo = campo & " WHERE spromo.codartic= " & DBSet(miRsAux!codArtic, "T") & " AND codlista = " & Me.txtTarifa(1).Text

             conn.Execute campo
            miRsAux.MoveNext
        Wend
        miRsAux.Close
    End If
    Set miRsAux = Nothing
    
    Screen.MousePointer = vbDefault
    
    If NumRegElim > 0 Then Unload Me
    
End Sub

Private Sub cmdAgentes_Click()
    
    If Me.chkAgentes(0).Value = 0 And Me.chkAgentes(1).Value = 0 Then
        MsgBox "Seleccione facturas", vbExclamation
        Exit Sub
    End If
    
    InicializarVbles
    
    'El nombre de la empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    'Si lleva articulo de portes, ese NO va a las lineas
    cadSelect = " 1 = 1"  'Para que no de error contar registros
    If vParamAplic.NumeroInstalacion <> 2 Then
        If vParamAplic.ArtPortesN <> "" Then campo = "{slifac.codartic} <> '" & vParamAplic.ArtPortesN & "'"
    End If
    cadFormula = cadSelect
   
    If txtFecha(24).Text <> "" Or txtFecha(25).Text <> "" Then
        devuelve = "vFechas=""Fecha: "
        campo = "{scafac.fecfactu}"
        If Not PonerDesdeHasta(campo, "F", 24, 25, devuelve) Then Exit Sub
        
    End If
    
    If txtAgente(0).Text <> "" Or txtAgente(1).Text <> "" Then
        devuelve = "vAgentes=""Agente: "
        campo = "{scafac.codagent}"
        If Not PonerDesdeHasta(campo, "AGT", 0, 1, devuelve) Then Exit Sub
    End If
     
      
    If txtForpa(1).Text <> "" Or txtForpa(2).Text <> "" Then
        devuelve = "vForpa=""Forma de pago: "
        campo = "{scafac.codforpa}"
        If Not PonerDesdeHasta(campo, "FOR", 1, 2, devuelve) Then Exit Sub
    End If
    
    'Enero 2013.  Puede entrar las rectificativas. Check para que salgan
    If Me.chkAgentes(2).Value = 1 Then
        miSQL = " 1=1 "  'para que los and de despues no den error
    Else
        'JULIO 2009
        'Las FRT no entran en el listado
        miSQL = " {scafac.codtipom} <> 'FRT'" 'NO ponemos las rectificativas
    End If
    
    If Me.chkAgentes(0).Value = 1 And Me.chkAgentes(1).Value = 1 Then
        'NO poenmos nada al select ya que pide las dos
            
    Else
        If Me.chkAgentes(0).Value = 1 Then
            miSQL = miSQL & " AND {scafac.codtipom} <> 'FAZ'"  'NO ponemos las "B"
        Else
            miSQL = miSQL & " AND {scafac.codtipom} = 'FAZ'"    'SOLO las B
        End If
    End If
    AnyadirAFormula cadFormula, miSQL
    cadSelect = cadSelect & " AND " & miSQL
    miSQL = ""
    'Montamos el select para los registros
    cadSelect = QuitarCaracterACadena(cadSelect, "{")
    cadSelect = QuitarCaracterACadena(cadSelect, "}")
    campo = "scafac.codtipom=slifac.codtipom and scafac.numfactu=slifac.numfactu and scafac.fecfactu=slifac.fecfactu AND "
    campo = "scafac,slifac WHERE " & campo & cadSelect
    
    
    If Not HayRegParaInforme(campo, "", True) Then
        MsgBox "No hay facturas con estos valores", vbExclamation
        Exit Sub
    End If
    
    'Report 61
    cadNomRPT = DevuelveDesdeBDNew(conAri, "scryst", "documrpt", "codcryst", "61", "N")
   
    
    LlamarImprimir False


    
End Sub

Private Sub cmdAlbaranProv_Click()

    InicializarVbles
    
    'Albaran socio
    If Not PonerParamRPT2(27, cadParam, numParam, cadNomRPT, vImprimedirecto, cadPDFrpt, vMultiInforme) Then Exit Sub
    
    
    
    cadSelect = "{sprove.tipprove}>=0"   'Antes ponia un tres: Estos proveedores son los REA o estimacion directa que luego
    cadFormula = "(" & cadSelect & ")"
    If txtFecha(18).Text <> "" Or txtFecha(19).Text <> "" Then
        campo = "{scaalp.fechaalb}"
        If Not PonerDesdeHasta(campo, "F", 18, 19, devuelve) Then Exit Sub
    End If
    
    If txtCodProve(8).Text <> "" Or txtCodProve(9).Text <> "" Then
        campo = "{scaalp.codprove}"
        If Not PonerDesdeHasta(campo, "PRO", 8, 9, devuelve) Then Exit Sub
    End If
     
    If txtNumAlbar(4).Text <> "" Or txtNumAlbar(5).Text <> "" Then
        campo = "{scaalp.numalbar}"
        If Not PonerDesdeHasta(campo, "ALP", 4, 5, "") Then Exit Sub
    End If
    
    
    'Montamos el select para los registros
    cadSelect = QuitarCaracterACadena(cadSelect, "{")
    cadSelect = QuitarCaracterACadena(cadSelect, "}")
    campo = "scaalp,sprove WHERE scaalp.codprove=sprove.codprove AND " & cadSelect
    
    
    If Not HayRegParaInforme(campo, "", True) Then
        MsgBox "No hay facturas con estos valores", vbExclamation
        Exit Sub
    End If
    
    'FALTA###
    LlamarImprimir False








    frmImprimir.Opcion = 2010
    frmImprimir.Show vbModal
End Sub

Private Sub cmdbeneClien_Click()
    
    
    InicializarVbles
    'El nombre de la empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    miSQL = ""
    
    If txtCliente(9).Text <> "" Or txtCliente(10).Text <> "" Then
        devuelve = "Cli: "
        campo = "{scafac.codclien}"
        If Not PonerDesdeHasta(campo, "CLI", 9, 10, devuelve) Then Exit Sub
        miSQL = devuelve
    End If
    
    'D/Hn almacen
    If txtAlma(5).Text <> "" Or txtAlma(6).Text <> "" Then
        codigo = ""
        campo = "{slifac.codalmac}"
        If Not PonerDesdeHasta(campo, "ALM", 5, 6, devuelve) Then Exit Sub
        
        'Cadena de arriba ph2
        If txtAlma(5).Text = txtAlma(6).Text Then
            devuelve = txtAlma(5).Text
        Else
            devuelve = txtAlma(5).Text & ".." & txtAlma(6).Text
        End If
        devuelve = "  Alm.[" & devuelve & "]"
        miSQL = Trim(miSQL & devuelve)
        
    End If
   
    
   
        
    cadParam = cadParam & "pdh1=""" & miSQL & """|"
    numParam = numParam + 1
    
    
    'Cadena para seleccion Desde y Hasta FECHA
    '--------------------------------------------
    miSQL = ""
    If txtFecha(43).Text <> "" Or txtFecha(44).Text <> "" Then
        devuelve = " Fecha: "
        campo = "{scafac.fecfactu}"
        If Not PonerDesdeHasta(campo, "F", 43, 44, devuelve) Then Exit Sub
        miSQL = devuelve & "   "
        'Si ha puesto mes
        
    End If
    
    If txtmarca(6).Text <> "" Or txtmarca(7).Text <> "" Then
        devuelve = "  Marca: "
        campo = "{sartic.codmarca}"
        If Not PonerDesdeHasta(campo, "MAR", 6, 7, devuelve) Then Exit Sub
        If Len(miSQL) > 60 Then devuelve = " Marca: " & txtmarca(6).Text & " - " & txtmarca(7).Text
        miSQL = Trim(miSQL & "  " & devuelve)
    End If
    
    miSQL = Trim(miSQL & "   (" & DevuelvePrecioCosteListado(1, False) & ")")
    
    If chkBenAge(10).Value = 1 Then miSQL = Trim(miSQL & "    Dto-Coste")
    
    
    
    cadParam = cadParam & "pdh2=""" & miSQL & """|"
    numParam = numParam + 1
    
    
    'Si detalla
    cadParam = cadParam & "DetallaFamilia=" & Abs(Me.chkBenAge(3).Value) & "|"  'Familia=Marca
    cadParam = cadParam & "DetallaArticulo=" & Abs(Me.chkBenAge(4).Value) & "|"
    numParam = numParam + 2
   
    Screen.MousePointer = vbHourglass
    benexClien
    Screen.MousePointer = vbDefault
        
    cadFormula = "{tmpinformes.codusu} = " & vUsu.codigo
    If chkBenAge(5).Value = 1 Then
        cadNomRPT = "rBenexCliMar.rpt"
    Else
        cadNomRPT = "rBenexMarCli.rpt"
    End If
     'QUITO los importe menores agrupados por proveedor
   ' If txtimporte(2).Text <> "" Then QuitarProveedoresImporteMenor
    
    Label3(156).Caption = ""
    campo = DevuelveDesdeBD(conAri, "count(*)", "tmpinformes", "codusu", CStr(vUsu.codigo))
    If campo = "" Then campo = "0"
    If Val(campo) = 0 Then
        'No existen datos
        MsgBox "No existen datos", vbExclamation
    Else
        LlamarImprimir False
    End If
    
    
    
    
    
End Sub

Private Sub cmdBeneficioAgente_Click()
Dim F As Date
Dim GenerarDatosEnTmp As Boolean
    
    'Si es comparativo debe indicar UN AÑO por lo menos
    If chkBenAge(2).Value = 1 Then
        If Opcion = 37 Then
            If txtAnyo(0).Text = "" Or txtAnyo(1).Text = "" Then
                MsgBox "Para el informe comparativo debe indicar el año", vbExclamation
                PonerFoco txtAnyo(0)
                Exit Sub
            End If
            If chkBenAge(1).Value = 1 Then chkBenAge(1).Value = 0
        End If
    End If
    
    'Si pone mes, el año debe ser el mismo
    If txtAnyo(2).Text <> "" Or txtAnyo(3).Text <> "" Then
        miSQL = ""
        'Si pone desde hasta mes
         If txtAnyo(2).Text <> "" And txtAnyo(3).Text <> "" Then
            If Val(txtAnyo(2).Text) > Val(txtAnyo(3).Text) Then miSQL = "Mes fin menor que el mes incio"
        End If
        
        'Si indica el mes , el año tiene que ser el mismo
        
        If txtAnyo(0).Text = "" Then
            miSQL = miSQL & vbCrLf & "Debe indicar el año"
        Else
            If txtAnyo(1).Text = "" Then
                miSQL = miSQL & vbCrLf & "Debe indicar UN único año"
            Else
                If txtAnyo(1).Text <> txtAnyo(0).Text Then miSQL = miSQL & vbCrLf & "El año debe ser el mismo"
            End If
        End If
        If miSQL <> "" Then
            PonerFoco txtAnyo(0)
            MsgBox miSQL, vbExclamation
            Exit Sub
        End If
    End If
    
    

    InicializarVbles
    'El nombre de la empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    miSQL = ""
    If txtAgente(6).Text <> "" Or txtAgente(7).Text <> "" Then
        devuelve = "Ag.:"
        campo = "{scafac.codagent}"
        If Not PonerDesdeHasta(campo, "AGT", 6, 7, devuelve) Then Exit Sub
        miSQL = miSQL & devuelve & "        "
    End If

    'Familia  4 5
    If txtFamia(10).Text <> "" Or txtFamia(11).Text <> "" Then
        devuelve = "Fam.: "
        campo = "{sartic.codfamia}"
        If Not PonerDesdeHasta(campo, "FAM", 10, 11, devuelve) Then Exit Sub
        'If miSQL <> "" Then miSQL = miSQL & vbCrLf
        If miSQL <> "" Then miSQL = miSQL & """ + chr(13) + """
        miSQL = miSQL & devuelve
    End If
    
    
    
    miSQL = Trim(miSQL & "   (" & DevuelvePrecioCosteListado(0, False) & ")")
    
    'Si ha marcado aplica dto
    'En el comparativo NO muestra costes, con lo cual, lo de AplicaDto solo es cuando no sea comparativo
    If Me.chkBenAge(2).Value = 0 Then   'No es compartaivo
        If Me.chkBenAge(9).Value = 1 Then miSQL = Trim(miSQL & "   Dto-Coste")
    End If
    
    If Me.chkBenAge(6).Value = 1 Then miSQL = miSQL & "  Presu"
    
    'Por si pone quitar proveedores
    miSQL = miSQL & "    @@@@@"
    cadParam = cadParam & "pdh1=""" & miSQL & """|"
    numParam = numParam + 1
    
    'Cadena para seleccion Desde y Hasta FECHA
    '--------------------------------------------
    'TRAMPA
    miSQL = ""
    txtFecha(11).Text = ""
    txtFecha(12).Text = ""
    
    If txtAnyo(0).Text <> "" Then txtFecha(11).Text = "01/01/" & txtAnyo(0).Text
    If txtAnyo(1).Text <> "" Then txtFecha(12).Text = "31/12/" & txtAnyo(1).Text
    If Trim(Me.txtAnyo(2).Text) <> "" Or Trim(txtAnyo(3).Text) <> "" Then
        'Ha puesto mes
        If txtAnyo(2).Text <> "" Then
            NumRegElim = DiasMes(CByte(txtAnyo(2).Text), CInt(txtAnyo(0).Text))
            NumRegElim = 1 'desde el UNO
            txtFecha(11).Text = NumRegElim & "/" & txtAnyo(2).Text & "/" & txtAnyo(0).Text
        End If
        If txtAnyo(3).Text <> "" Then
            NumRegElim = DiasMes(CByte(txtAnyo(3).Text), CInt(txtAnyo(0).Text))
            txtFecha(12).Text = NumRegElim & "/" & txtAnyo(3).Text & "/" & txtAnyo(0).Text
        End If
    End If
    If txtFecha(11).Text <> "" Or txtFecha(12).Text <> "" Then
        devuelve = " Fecha: "
        campo = "{scafac.fecfactu}"
        If Not PonerDesdeHasta(campo, "F", 11, 12, devuelve) Then Exit Sub
        
        'Si pone mes:
        If Me.txtAnyo(2).Text <> "" Or txtAnyo(3).Text <> "" Then
            devuelve = " Año: " & txtAnyo(0).Text & "     Mes: "
            If txtAnyo(2).Text <> "" Then devuelve = devuelve & " desde " & txtAnyo(2).Text
            If txtAnyo(3).Text <> "" Then devuelve = devuelve & " hasta " & txtAnyo(3).Text
        End If
        miSQL = devuelve & "   "
        'Si ha puesto mes
        
    End If
     
    
    If txtCodProve(21).Text <> "" Or txtCodProve(22).Text <> "" Then
        devuelve = "Prov: "
        campo = "{sartic.codprove}"
        If Not PonerDesdeHasta(campo, "PRO", 21, 22, devuelve) Then Exit Sub
        miSQL = miSQL & devuelve
    End If
    
    
    If Me.txtZona(4).Text <> "" Or txtZona(5).Text <> "" Then
        devuelve = "Zona:  "
        campo = "{sclien.codzonas}"
        If Not PonerDesdeHasta(campo, "ZON", 4, 5, devuelve) Then Exit Sub
        If Len(miSQL) > 50 Then
            miSQL = miSQL & """ + chr(13) + """
        Else
            miSQL = miSQL & "   "
        End If
        miSQL = Trim(miSQL & devuelve)
    End If
    
    
     If Me.txtRuta(1).Text <> "" Or txtRuta(1).Text <> "" Then
        devuelve = IIf(vParamAplic.NumeroInstalacion = 2, "Asociacion", "Ruta") & ":  "
        campo = "{sclien.codrutas}"
        If Not PonerDesdeHasta(campo, "RUT", 0, 1, devuelve) Then Exit Sub
        If Len(miSQL) > 35 Then
            miSQL = miSQL & """ + chr(13) + """
        Else
            miSQL = miSQL & "   "
        End If
        miSQL = Trim(miSQL & devuelve)
    End If
    
    
    
    
    'D/Hn almacen
    If txtAlma(3).Text <> "" Or txtAlma(4).Text <> "" Then
        codigo = ""
        
        If txtAlma(3).Text <> "" Then codigo = " AND {slifac.codalmac} >= " & txtAlma(3).Text
        If txtAlma(4).Text <> "" Then codigo = codigo & " AND {slifac.codalmac} <= " & txtAlma(4).Text
        codigo = Mid(codigo, 5) 'quito el primer and
        If cadSelect <> "" Then
            cadSelect = cadSelect & " AND "
            cadFormula = cadFormula & " AND "
        End If
        
        cadFormula = cadFormula & codigo
        codigo = Replace(codigo, "{", "")
        codigo = Replace(codigo, "}", "")
        cadSelect = cadSelect & codigo
    
        'Cadena de arriba ph1
        If txtAlma(3).Text = txtAlma(4).Text Then
            campo = txtAlma(3).Text
        Else
            campo = txtAlma(3).Text & ".." & txtAlma(4).Text
        End If
        campo = "Alm.[" & campo & "]"
        If miSQL <> "" Then
            'Si ha puesto importe fuerzo el salto de linea
            If txtImporte(2).Text <> "" Then
                miSQL = miSQL & """ + chr(13) + """
            Else
                miSQL = miSQL & "  "
            End If
        End If
            
        miSQL = miSQL & campo
        
    End If
    
    '19 Abril 2011
    'Manolo Belarte.
    'El importe NO es sobre el importe de la linea, SINO sobre el total del proveedor
    'Con lo cual pasan todos a TMPinformes
    If txtImporte(2).Text <> "" Then

        campo = "     Imp. mín: " & txtImporte(2).Text & "€ "
        miSQL = Trim(miSQL & campo)
    End If
    cadParam = cadParam & "pdh2=""" & miSQL & """|"
    numParam = numParam + 1
    
    
    
    'Articulos de varios   Mayo 2015
    codigo = "{sartic.artvario} =0 "   'que no sea de varios
    If cadSelect <> "" Then codigo = " AND " & codigo
    
    'En el compartivo NO ponemos la marca NO de varios
    If Opcion = 37 Then
        If chkBenAge(2).Value = 1 Then codigo = ""
    End If
    
    cadSelect = cadSelect & codigo
    cadFormula = cadFormula & codigo
    
    
        
    
    
    
    'Si detalla
    cadParam = cadParam & "DetallaFamilia=" & Abs(Me.chkBenAge(0).Value) & "|"
    cadParam = cadParam & "DetallaArticulo=" & Abs(Me.chkBenAge(1).Value) & "|"
    cadParam = cadParam & "Totaliza=" & Abs(Me.chkBenAge(8).Value) & "|"
    numParam = numParam + 3
    

    
    
    'el compartaivo de agente no llevara nunca lo de mes
    GenerarDatosEnTmp = True
    campo = "mes"
    If Opcion = 37 Then
        If chkBenAge(2).Value = 1 Then
            'Comparativo
            Screen.MousePointer = vbHourglass
            GenerarDatosEnTmp = False
            ComparativoAgentes
            Screen.MousePointer = vbDefault
            campo = "" 'este no ira por mes
            cadFormula = "{tmpinformes.codusu} = " & vUsu.codigo
            cadNomRPT = "rBenAgCompara"
        Else
            cadNomRPT = "rBenexAge"
        End If
        
    Else
        cadNomRPT = "rBenexProv"
    End If
    'If Me.txtAnyo(2).Text <> "" Or txtAnyo(3).Text <> "" Then cadNomRPT = cadNomRPT & "mes"
    If Me.txtAnyo(2).Text <> "" Or txtAnyo(3).Text <> "" Then cadNomRPT = cadNomRPT & campo
    cadNomRPT = cadNomRPT & ".rpt"
    
    
    If GenerarDatosEnTmp Then
        'Hay que insertar los datos en la tmp
        Screen.MousePointer = vbHourglass
        InsertarTmpBeneAgeProv
        Screen.MousePointer = vbDefault
    End If
    cadFormula = "{tmpinformes.codusu} = " & vUsu.codigo
    
     'QUITO los importe menores agrupados por proveedor
    If txtImporte(2).Text <> "" Then QuitarProveedoresImporteMenor
    
    'Mayo 2015
    
    chkBenAge(2).Tag = "" 'me guardare la selccion de proveedores

    If chkBenAge(7).Value = 1 Then
        'Quiataremos los proveedores que nos marcquem
        campo = DevuelveDesdeBD(conAri, "count(*)", "tmpinformes", "codusu", CStr(vUsu.codigo))
        If Val(campo) > 1 Then
            CadenaDesdeOtroForm = ""
            frmListado5.OtrosDatos = "select distinct codigo1 from tmpinformes where codusu=" & vUsu.codigo
            frmListado5.OpcionListado = 10
            frmListado5.Show vbModal
            If CadenaDesdeOtroForm <> "" Then
                Label3(147).Caption = "Eliminando proveedores"
                Label3(147).Refresh
                CadenaDesdeOtroForm = Mid(CadenaDesdeOtroForm, 2)
                chkBenAge(2).Tag = CadenaDesdeOtroForm
                miSQL = "DELETE FROM tmpinformes where codusu =" & vUsu.codigo
                miSQL = miSQL & " AND not codigo1 IN (" & CadenaDesdeOtroForm & ")"
                conn.Execute miSQL
                
                miSQL = "[**]"
            Else
                miSQL = ""
            End If
            cadParam = Replace(cadParam, "@@@@@", miSQL)  'Quito la cadena reservada en los desde hastas, en la que marcare que han quitado proveedores
        End If
    Else
        'No poner quitar proveedores
        'Quito la cadena reservada en los desde hastas
        cadParam = Replace(cadParam, "@@@@@", "")
    End If
    
    
    
    
    Label3(147).Caption = ""
    campo = DevuelveDesdeBD(conAri, "count(*)", "tmpinformes", "codusu", CStr(vUsu.codigo))
    If campo = "" Then campo = "0"
    If Val(campo) = 0 Then
        'No existen datos
        MsgBox "No existen datos", vbExclamation
    Else
        LlamarImprimir False
    End If
    
    
    
End Sub

Private Sub cmdBeneMarcaAgen_Click()
    
        
    InicializarVbles
    'El nombre de la empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    miSQL = ""
    
    If txtFecha(49).Text <> "" Or txtFecha(50).Text <> "" Then
        devuelve = " Fecha: "
        campo = "{scafac.fecfactu}"
        If Not PonerDesdeHasta(campo, "F", 49, 50, devuelve) Then Exit Sub
        miSQL = devuelve & "   "
        'Si ha puesto mes
        
    End If
    
    
    miSQL = Trim(miSQL & "   (" & DevuelvePrecioCosteListado(2, False) & ")")
        
    cadParam = cadParam & "pdh1=""" & miSQL & """|"
    numParam = numParam + 1
    
    
    'Cadena para seleccion Desde y Hasta FECHA
    '--------------------------------------------
    miSQL = ""
    
    If txtmarca(8).Text <> "" Or txtmarca(9).Text <> "" Then
        devuelve = "  Marca: "
        campo = "{sartic.codmarca}"
        If Not PonerDesdeHasta(campo, "MAR", 8, 9, devuelve) Then Exit Sub
        If Len(miSQL) > 60 Then devuelve = " Marca: " & txtmarca(8).Text & " - " & txtmarca(9).Text
        miSQL = Trim(miSQL & "  " & devuelve)
    End If
    If txtAgente(11).Text <> "" Or txtAgente(12).Text <> "" Then
        devuelve = " Agente: "
        campo = "{scafac.codagent}"
        If Not PonerDesdeHasta(campo, "AGT", 11, 12, devuelve) Then Exit Sub
        miSQL = Trim(miSQL & "  " & devuelve)
    End If
    
    
    If txtCodProve(24).Text <> "" Or txtCodProve(25).Text <> "" Then
        devuelve = " Proveedor: "
        campo = "{sartic.codprove}"
        If Not PonerDesdeHasta(campo, "PRO", 24, 25, devuelve) Then Exit Sub
        If miSQL <> "" Then miSQL = miSQL & """ + chr(13) + """
        miSQL = Trim(miSQL & "  " & devuelve)
    End If
    
    campo = ""
    If cadSelect <> "" Then campo = " AND "
    cadSelect = cadSelect & campo & " {sartic.artvario} =0 AND scafac.codtipom <> 'FAZ'"
    cadParam = cadParam & "pdh2=""" & miSQL & """|"
    numParam = numParam + 1
    
    
    'Si detalla
    cadParam = cadParam & "DetallaArticulo=" & Abs(Me.chkBeneMarcaAgen(0).Value) & "|"
    numParam = numParam + 1
   
    Screen.MousePointer = vbHourglass
    BenexMarcaAgenProv
    Screen.MousePointer = vbDefault
        
    cadFormula = "{tmpinformes.codusu} = " & vUsu.codigo

    cadNomRPT = "rBenexMarAgeProv.rpt"

     
    Label3(156).Caption = ""
    campo = DevuelveDesdeBD(conAri, "count(*)", "tmpinformes", "codusu", CStr(vUsu.codigo))
    If campo = "" Then campo = "0"
    If Val(campo) = 0 Then
        'No existen datos
        MsgBox "No existen datos", vbExclamation
    Else
        LlamarImprimir False
    End If
    

    
    
End Sub

Private Sub cmdCambiarImporte_Click()
Dim Fecha As Date
Dim vA As CArticulo
    cadSelect = ""
    'Comprobaciones
    If Me.txtImporte(0).Text = "" Then cadSelect = "      -Importe"
    If txtArticulo(3).Text = "" Or Me.txtDescArticulo(3).Text = "" Then cadSelect = cadSelect & vbCrLf & "     -Articulo"
    
    If cadSelect <> "" Then
        MsgBox "Campos obligatorios" & vbCrLf & cadSelect, vbExclamation
        Exit Sub
    End If
    
    
    InicializarVbles
    devuelve = ""
    
    
    'Cadena obligada. Los proveedores , el tipo tiene que ser el 3: REA
    cadSelect = " {slialp.codprove}=  {sprove.codprove}  AND {sprove.tipprove}= 3"
    
    'Cadena para seleccion Desde y Hasta FECHA
    '--------------------------------------------
    If txtFecha(11).Text <> "" Or txtFecha(12).Text <> "" Then
        campo = "{slialp.fechaalb}"
        If Not PonerDesdeHasta(campo, "F", 11, 12, devuelve) Then Exit Sub
    End If
    
    If txtCodProve(2).Text <> "" Or txtCodProve(3).Text <> "" Then
        campo = "{slialp.codprove}"
        If Not PonerDesdeHasta(campo, "PRO", 2, 3, devuelve) Then Exit Sub
    End If
    
    If cadSelect <> "" Then cadSelect = cadSelect & " AND "
    cadSelect = cadSelect & "  ({slialp.codartic} = '" & txtArticulo(3).Text & "')"
    
    'Vermos si hay registros
    
    cadSelect = QuitarCaracterACadena(cadSelect, "{")
    cadSelect = QuitarCaracterACadena(cadSelect, "}")
    'Pongo el oreder por comodidad
    cadSelect = cadSelect & " ORDER BY fechaalb, slialp.codprove"
    cadFrom = "Select count(*) from slialp,sprove  WHERE " & cadSelect
    
    
    IndiceImg = NumRegistros(cadFrom)
    If IndiceImg = 0 Then
        MsgBox "No hay datos con estos valores", vbExclamation
        Exit Sub
    Else
        cadFrom = "Hay " & IndiceImg & " registro(s) para actualizar el precio" & vbCrLf & _
            "Desea continuar con la actualizacion de precios?"
        
        If MsgBox(cadFrom, vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    End If
    
    cadFrom = "Select * from slialp,sprove WHERE " & cadSelect
    
    If Not BloqueoManual("LIQCMBPRE", "1") Then Exit Sub
    Screen.MousePointer = vbHourglass
    Set vA = New CArticulo
    If vA.LeerDatos(CStr(txtArticulo(3).Text)) Then
         
        Set miRsAux = New ADODB.Recordset
        cadSelect = "Select ultfecco from sartic where codartic = '" & DevNombreSQL(txtArticulo(3).Text) & "'"
        miRsAux.Open cadSelect, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Fecha = CDate("01/01/1900")
        If Not miRsAux.EOF Then
            If Not IsNull(miRsAux.Fields(0)) Then Fecha = miRsAux.Fields(0)
        End If
        miRsAux.Close
        numParam = 0   'Auqi tendre si ha cambiado la fecha o no
        
        ImpTeo = ImporteFormateado(txtImporte(0).Text)
        'Por si lo meto en una transaccion
        RealizarCambiosPreciosLiq Fecha
        
        'Si tengo que updatearl ultcompra
        If numParam = 1 Then vA.ActualizarUltFechaCompra CStr(Fecha), txtImporte(0).Text
        
        
       Me.lblLiqu.Caption = ""
       MsgBox "Proceso finalizado", vbExclamation
           
           
       'Para que no vuelvan a anzar el proceso
       txtArticulo(3).Text = ""
       txtDescArticulo(3).Text = ""
    End If
    Set vA = Nothing
    Set miRsAux = Nothing
    DesBloqueoManual "LIQCMBPRE"
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdCambiClienteFreq_Click()
    
    'Comprobaremos que ha puesto cliente
    If Me.txtCliente(11).Text = "" Or txtCliente(12).Text = "" Then
        MsgBox "Debe indicar el cliente", vbExclamation
        vMultiInforme = 12
        If Me.txtCliente(11).Text = "" Then vMultiInforme = 11
        PonerFoco txtCliente(vMultiInforme)
        Exit Sub
    End If
    
    
    'comprobamos si ha puesto departamento
    devuelve = ""
    campo = txtCliente(11) & " " & Me.txtDescClie(11).Text
    codigo = txtCliente(12) & " " & Me.txtDescClie(12).Text
    If Me.txtDpto(2).Text <> "" Then
        miSQL = "codclien = " & txtCliente(11).Text & " AND coddirec"
        miSQL = DevuelveDesdeBD(conAri, "nomdirec", "sdirec", miSQL, txtDpto(2))
        If miSQL = "" Then
            vMultiInforme = 2
            devuelve = "No existe el departamento " & txtDpto(2).Text & " para el cliente " & campo & vbCrLf
        Else
            Me.txtDescDpto(2).Text = miSQL
        End If
    End If
    If Me.txtDpto(3).Text <> "" Then
        miSQL = "codclien = " & txtCliente(12).Text & " AND coddirec"
        miSQL = DevuelveDesdeBD(conAri, "nomdirec", "sdirec", miSQL, txtDpto(3).Text)
        If miSQL = "" Then
            vMultiInforme = 3
            devuelve = devuelve & "No existe el departamento " & txtDpto(3).Text & " para el cliente " & codigo & vbCrLf
        Else
            Me.txtDescDpto(3).Text = miSQL
        End If
    End If
    
    If devuelve <> "" Then
        MsgBox devuelve, vbExclamation
        PonerFoco txtDpto(vMultiInforme)
        Exit Sub
    End If
    
    
        
    'OK, procederemos al cambio
    If Me.txtDpto(2).Text = Me.txtDpto(3).Text And Me.txtCliente(11).Text = Me.txtCliente(12).Text Then
        MsgBox "Mismo dato origen-destino", vbExclamation
        Exit Sub
    End If
    
    
    'vemos cuantos hay para cambiar
    cadFrom = ""
    If Me.txtDpto(2).Text <> "" Then cadFrom = "coddirec = " & txtDpto(2).Text & " AND "
    cadFrom = cadFrom & "codclien "
    cadFrom = DevuelveDesdeBD(conAri, "count(*)", "scafre", cadFrom, txtCliente(11).Text)
    
    If Val(cadFrom) = 0 Then
        MsgBox "Ninguna frecuencia para cambiar", vbExclamation
        Exit Sub
    End If
    
    'Ok . Preguntamos
          
    devuelve = "Va a cambiar las frecuencias :" & vbCrLf & vbCrLf & "ORIGEN: " & campo & vbCrLf
    If Me.txtDpto(2).Text <> "" Then devuelve = devuelve & "Departamento: " & Me.txtDpto(2).Text & " " & Me.txtDescDpto(2).Text & vbCrLf
    devuelve = devuelve & vbCrLf & vbCrLf & "DESTINO: " & codigo & vbCrLf
    If Me.txtDpto(3).Text <> "" Then devuelve = devuelve & "Departamento: " & Me.txtDpto(3).Text & " " & Me.txtDescDpto(3).Text & vbCrLf
    devuelve = devuelve & vbCrLf & "Total frecuencias: " & cadFrom
    devuelve = devuelve & vbCrLf & vbCrLf & "¿Contiuar?"
    If MsgBox(devuelve, vbQuestion + vbYesNo) = vbNo Then Exit Sub

    
   
    
    devuelve = "UPDATE scafre SET codclien = " & Me.txtCliente(12).Text
    If Me.txtDpto(3).Text <> "" Then devuelve = devuelve & " , coddirec = " & txtDpto(3).Text
    devuelve = devuelve & " WHERE codclien = " & Me.txtCliente(11).Text
    If Me.txtDpto(2).Text <> "" Then devuelve = devuelve & " and coddirec = " & txtDpto(2).Text
    
    If ejecutar(devuelve, True) Then
        CadenaDesdeOtroForm = "OK"
        Unload Me
    End If
    
    
End Sub

Private Sub cmdCambioPromo_Click()
    campo = ""
    If Me.txtFecha(41).Text = "" Then campo = campo & vbCrLf & "-Fecha inicio promocion"
    If Me.txtFecha(42).Text = "" Then campo = campo & vbCrLf & "-Fecha fin promocion"
    If Me.txtTarifa(1).Text = "" Then campo = campo & vbCrLf & "-Tarifa"
    If Me.txtCodProve(23).Text = "" Then campo = campo & vbCrLf & "-Proveedor"
    If campo <> "" Then
        campo = "Campos obligatorios: " & campo
        MsgBox campo, vbExclamation
        PonerFoco txtFecha(41)
        Exit Sub
    End If
    
    InicializarVbles
    
    cadSelect = "{spromo.codlista} = " & txtTarifa(1).Text
    'Proveedor  18 19
    If txtCodProve(23).Text <> "" Or txtCodProve(23).Text <> "" Then
        'devuelve = "pDHProve=""Proveedor: "
        campo = "{sartic.codprove}"
        If Not PonerDesdeHasta(campo, "PRO", 23, 23, devuelve) Then Exit Sub
        
    End If
    
    
    'Familia  4 5
    If txtFamia(12).Text <> "" Or txtFamia(13).Text <> "" Then
        'devuelve = "pDHFamilia=""Familia: "
        campo = "{sartic.codfamia}"
        If Not PonerDesdeHasta(campo, "FAM", 12, 13, devuelve) Then Exit Sub
        
    End If
    
    cadSelect = QuitarCaracterACadena(cadSelect, "{")
    cadSelect = QuitarCaracterACadena(cadSelect, "}")
    
    campo = "spromo,sartic "
    campo = campo & " WHERE spromo.codartic=sartic.codartic AND " & cadSelect
    
    
    Set miRsAux = New ADODB.Recordset
    Screen.MousePointer = vbHourglass
    
    'Compruebo que no hay datos con precionuevo
    devuelve = "Select count(*) from " & campo & " AND not fechain1 is null"
    NumRegElim = 0
    miRsAux.Open devuelve, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then NumRegElim = DBLet(miRsAux.Fields(0), "N")
    miRsAux.Close
    
    If NumRegElim > 0 Then
        MsgBox "Existen promociones sin actualizar", vbExclamation
        campo = ""
        NumRegElim = 0
    End If
   'Insert into
    If campo <> "" Then CargaTmpCambioPromo
    Screen.MousePointer = vbDefault
    
    If NumRegElim = 0 Then Exit Sub
    
    
    frmAlmCambPromo.Show vbModal
    Unload Me
End Sub



Private Sub CargaTmpCambioPromo()
On Error GoTo ECargaTmpCambioPromo

    conn.Execute "DELETE FROM tmpinformes where codusu = " & vUsu.codigo

    campo = "select precioac,precioa1,spromo.codartic,nomartic FROM " & campo
    miRsAux.Open campo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    campo = ""
    NumRegElim = 0
    While Not miRsAux.EOF
        NumRegElim = NumRegElim + 1
        'tmpinformes(codusu,codigo1,nombre1,nombre2,importe1,importe2,importe3,importe4)
        campo = campo & ", (" & vUsu.codigo & "," & NumRegElim & "," & DBSet(miRsAux!codArtic, "T") & ","
        campo = campo & DBSet(miRsAux!NomArtic, "T") & "," & DBSet(miRsAux!precioac, "N") & ","
        campo = campo & DBSet(miRsAux!precioa1, "N", "N") & ",0,0)"
        
        If (NumRegElim Mod 50) = 0 Then
            campo = Mid(campo, 2)
            campo = "INSERT INTO tmpinformes(codusu,codigo1,nombre1,nombre2,importeb1,importeb2,importeb3,importeb4) VALUES " & campo
            conn.Execute campo
            campo = ""
        End If
        
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If campo <> "" Then
        campo = Mid(campo, 2)
        campo = "INSERT INTO tmpinformes(codusu,codigo1,nombre1,nombre2,importeb1,importeb2,importeb3,importeb4) VALUES " & campo
        conn.Execute campo
    End If


    'En fecha1 y fecha2 pongo las fechas nuevas de promocion
    campo = "UpDATE tmpinformes set fecha1=" & DBSet(txtFecha(41).Text, "F") & ",  fecha2=" & DBSet(txtFecha(42).Text, "F")
    campo = campo & " ,campo1 = " & txtTarifa(1).Text
    campo = campo & " WHERE codusu = " & vUsu.codigo
    conn.Execute campo

ECargaTmpCambioPromo:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "CargaTmp"
    Else
        If NumRegElim = 0 Then MsgBox "No hay registros con estos valores", vbExclamation
    End If
    
End Sub


Private Sub cmdCambiPrecio_Click()
    
    'Trocito comun
    campo = ""
    If Me.txtFecha(38).Text = "" Then campo = campo & vbCrLf & "-Fecha"
    If Opcion = 34 And Me.txtTarifa(0).Text = "" Then campo = campo & vbCrLf & "-Tarifa"
    If Me.txtCodProve(20).Text = "" Then campo = campo & vbCrLf & "-Proveedor"
    If campo <> "" Then
        campo = "Campos obligatorios: " & campo
        MsgBox campo, vbExclamation
        Exit Sub
    End If
    InicializarVbles



    If Opcion = 34 Then
        CambioPreciosTarifasVenta
    Else
        CambioPreciosTarifasCOMPRAS
    End If
End Sub
    
    
Private Sub CambioPreciosTarifasVenta()
Dim OkDobleComprobacion As Boolean
Dim SelectPrees As String

    cadSelect = "{slista.codlista} = " & txtTarifa(0).Text
    SelectPrees = ""
    If txtCodProve(20).Text <> "" Or txtCodProve(20).Text <> "" Then
        'devuelve = "pDHProve=""Proveedor: "
        campo = "{sartic.codprove}"
        If Not PonerDesdeHasta(campo, "PRO", 20, 20, devuelve) Then Exit Sub
        SelectPrees = SelectPrees & " AND sartic.codprove =" & txtCodProve(20).Text
    End If
    
    
    'Familia  4 5
    If txtFamia(6).Text <> "" Or txtFamia(7).Text <> "" Then
        'devuelve = "pDHFamilia=""Familia: "
        campo = "{sartic.codfamia}"
        If Not PonerDesdeHasta(campo, "FAM", 6, 7, devuelve) Then Exit Sub
        If txtFamia(6).Text <> "" Then SelectPrees = SelectPrees & " AND sartic.codfamia >=" & txtFamia(6).Text
        If txtFamia(7).Text <> "" Then SelectPrees = SelectPrees & " AND sartic.codfamia <=" & txtFamia(7).Text
        
    End If
    
    
    'Marca 2 3
    
    If txtArticulo(8).Text <> "" Or txtArticulo(9).Text <> "" Then
        devuelve = " "
        campo = "{sartic.codartic}"
        If Not PonerDesdeHasta(campo, "ART", 8, 9, devuelve) Then Exit Sub
        If txtArticulo(8).Text <> "" Then SelectPrees = SelectPrees & " AND sartic.codfamia >=" & DBSet(txtArticulo(8).Text, "T")
        If txtArticulo(9).Text <> "" Then SelectPrees = SelectPrees & " AND sartic.codfamia <=" & DBSet(txtArticulo(9).Text, "T")
        
    End If
    
    If Me.optSituaArt(0).Value Then
        'TODOS. No hacemos nada
        campo = ""
    Else
        
        If Me.optSituaArt(1).Value Then
            campo = "1"
        
        ElseIf Me.optSituaArt(2).Value Then
            campo = "2"
        Else
            campo = "3"
        End If
        campo = " AND {sartic.codstatu} = " & campo
        SelectPrees = SelectPrees & " AND sartic.codstatu = " & campo
    End If
    cadSelect = cadSelect & campo
    cadSelect = QuitarCaracterACadena(cadSelect, "{")
    cadSelect = QuitarCaracterACadena(cadSelect, "}")
    
    campo = "slista,sartic "
    campo = campo & " WHERE slista.codartic=sartic.codartic AND " & cadSelect
    
    If Not HayRegParaInforme(campo, "", True) Then
        MsgBox "No hay registros con estos valores", vbExclamation
        Exit Sub
    End If
    
    
    
    'Ahora comprobare que ninugno tiene fechanue
    OkDobleComprobacion = True
    codigo = ""
    campo = "slista,sartic "
    campo = campo & " WHERE slista.codartic=sartic.codartic AND " & cadSelect
    campo = campo & " AND not fechanue is null"
    If HayRegParaInforme(campo, "", True) Then
        codigo = "Hay registros sin actualizar en lista de precios"
        OkDobleComprobacion = False
    End If
    
    
    '2017. Si hay actualizacion precios especiales tambien comprobaremos esto
    If vParamAplic.ActualizaPrecioEspecial Then
               
        campo = "sprees,sartic "
        campo = campo & " WHERE sprees.codartic=sartic.codartic AND "
        campo = campo & " not fechanue is null " & SelectPrees
        If HayRegParaInforme(campo, "", True) Then
            codigo = codigo & vbCrLf & "Hay registros sin actualizar en precios especiales"
            OkDobleComprobacion = False
        End If
    End If
    
    If Not OkDobleComprobacion Then
        MsgBox codigo, vbExclamation
        Exit Sub
    End If
    
    
    'Ponemos el tmpprecioac a cero
    campo = "UPDATE slista,sartic set tmpprecioac=0 WHERE "
    campo = campo & " slista.codartic=sartic.codartic AND " & cadSelect
    ejecutar campo, False
    
    frmAlmCambPrec.vFecha = CDate(txtFecha(38).Text)
    frmAlmCambPrec.parSelSQL = cadSelect
    frmAlmCambPrec.Ventas = True
    frmAlmCambPrec.Show vbModal
End Sub


Private Sub CambioPreciosTarifasCOMPRAS()
    


    campo = "{slispr.codprove}"
    If Not PonerDesdeHasta(campo, "PRO", 20, 20, devuelve) Then Exit Sub
    

    
    
    'Familia  4 5
    If txtFamia(6).Text <> "" Or txtFamia(7).Text <> "" Then
        'devuelve = "pDHFamilia=""Familia: "
        campo = "{sartic.codfamia}"
        If Not PonerDesdeHasta(campo, "FAM", 6, 7, devuelve) Then Exit Sub
        
    End If
    
    
    'Marca 2 3
    
    If txtArticulo(8).Text <> "" Or txtArticulo(9).Text <> "" Then
        devuelve = " "
        campo = "{sartic.codartic}"
        If Not PonerDesdeHasta(campo, "ART", 8, 9, devuelve) Then Exit Sub
        
    End If
    
    If Me.optSituaArt(0).Value Then
        'TODOS. No hacemos nada
        campo = ""
    Else
        
        If Me.optSituaArt(1).Value Then
            campo = "1"
        
        ElseIf Me.optSituaArt(2).Value Then
            campo = "2"
        Else
            campo = "3"
        End If
        campo = " AND {sartic.codstatu} = " & campo
    End If
    cadSelect = cadSelect & campo
    cadSelect = QuitarCaracterACadena(cadSelect, "{")
    cadSelect = QuitarCaracterACadena(cadSelect, "}")
    
    campo = "slispr,sartic "
    campo = campo & " WHERE slispr.codartic=sartic.codartic AND " & cadSelect
    
    If Not HayRegParaInforme(campo, "", True) Then
        MsgBox "No hay registros con estos valores", vbExclamation
        Exit Sub
    End If
    
    'Borro la temporal
    campo = "DELETE FROM tmpslipreu WHERE codusu = " & vUsu.codigo
    conn.Execute campo
    
    'Ahora comprobare que ninugno tiene fechanue
    campo = "slispr,sartic "
    campo = campo & " WHERE slispr.codartic=sartic.codartic AND " & cadSelect
    campo = campo & " AND not fechanue is null"
    If HayRegParaInforme(campo, "", True) Then
        MsgBox "Hay registros sin actualizar cambio precio", vbExclamation
        Exit Sub
    End If
    
    
    
    
    'Ponemos el tmpprecioac a cero
    campo = "INSERT INTO tmpslipreu(codusu,numofert,numlinea,codartic,nomartic,ampliaci,precioar) "
    campo = campo & "select " & vUsu.codigo & ",slispr.codprove,@rownum:=@rownum+1 AS rownum ,slispr.codartic,nomartic,format(precioac,4),0.00  "
    campo = campo & " FROM slispr,sartic ,(SELECT @rownum:=0) r WHERE "
    campo = campo & " slispr.codartic=sartic.codartic AND " & cadSelect
    If ejecutar(campo, False) Then
        
        'Tema estetico.  Ampliaci es texto y lelva un importe. Reemplzamos la coma po
        campo = "UPDATE tmpslipreu SET ampliaci=replace(ampliaci,'.',',') WHERE codusu = " & vUsu.codigo
        conn.Execute campo
        
        frmAlmCambPrec.vFecha = CDate(txtFecha(38).Text)
        frmAlmCambPrec.parSelSQL = "tmpslipreu.codusu = " & vUsu.codigo
        frmAlmCambPrec.Ventas = False
        
        frmAlmCambPrec.Show vbModal
    End If
End Sub




Private Sub cmdCancel_Click(Index As Integer)
    'Si estamos en calculo de riesgo, cancelar  puede parar el proceso para salir
    If Index = 31 Then
        If Opcion = 0 Then
            'Le ha dado a cancelar.
            If MsgBox("¿Desea parar el proceso?", vbQuestion + vbYesNo) = vbYes Then Opcion = 31
                
            Exit Sub
        End If
    End If
    Unload Me
End Sub

Private Sub cmdContabTicket_Click()
    If Opcion = 13 Then
        ContabilizarTickets
    Else
        Screen.MousePointer = vbHourglass
        ListadoTicketsAgrupados
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub ListadoTicketsAgrupados()

    'Meto el resume IVA en tmpnlotes
    Label5.Caption = "Obteniendo datos IVAs"
    Label5.Refresh
    devuelve = "DELETE FROM tmpinformes WHERE codusu = " & vUsu.codigo
    conn.Execute devuelve
    
    devuelve = " FROM    `sfactik` LEFT OUTER JOIN  `scafac` ON (`sfactik`.`numfactu`=`scafac`.`numfactu`) AND (`sfactik`.`fecfactu`=`scafac`.`fecfactu`) "
    cadFrom = ""
    If txtFecha(20).Text <> "" Then cadFrom = cadFrom & " AND fecfacFTG >='" & Format(txtFecha(20).Text, FormatoFecha) & "'"
    If txtFecha(21).Text <> "" Then cadFrom = cadFrom & " AND fecfacFTG <='" & Format(txtFecha(20).Text, FormatoFecha) & "'"
    If cadFrom <> "" Then devuelve = devuelve & " WHERE " & Mid(cadFrom, 5)
    devuelve = devuelve & " GROUP BY 1,2,3"
    Set miRsAux = New ADODB.Recordset
    NumRegElim = 1
    cadTitulo = "insert into `tmpinformes` (`codusu`,`codigo1`,`nombre1`,`nombre2`,`importe1`,`importe2`,`importe3`) VALUES (" & vUsu.codigo & ","
    For numParam = 1 To 3
        cadFrom = ",porciva" & numParam & " c1,sum(imporiv" & numParam & ") c2,sum(baseimp" & numParam & ") c3 "
        cadFrom = "SELECT numfacftg,fecfacftg" & cadFrom & devuelve
        miRsAux.Open cadFrom, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            ' importante. Formatear con 7 0's como minimo,para realizar el link en el informe
             cadFrom = NumRegElim & ",'" & Format(miRsAux!numfacftg, "0000000") & "','" & miRsAux!fecfacftg & "'"
             'Los importes
             If Not IsNull(miRsAux!C1) Then
                cadFrom = cadFrom & "," & TransformaComasPuntos(CStr(miRsAux!C1))
                cadFrom = cadFrom & "," & TransformaComasPuntos(CStr(miRsAux!C2))
                cadFrom = cadFrom & "," & TransformaComasPuntos(CStr(miRsAux!c3))
                conn.Execute cadTitulo & cadFrom & ")"
                NumRegElim = NumRegElim + 1
             End If
             miRsAux.MoveNext
        Wend
        miRsAux.Close
    Next
    Set miRsAux = Nothing
    Me.Refresh
    Label5.Caption = ""
    InicializarVbles
    If Not PonerParamRPT2(28, cadParam, numParam, cadNomRPT, vImprimedirecto, cadPDFrpt, vMultiInforme) Then Exit Sub
    
    '===================================================
    '============ PARAMETROS ===========================
    'Añadir el nombre de la Empresa como parametro
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    devuelve = ""
    If txtFecha(20).Text <> "" Or txtFecha(21).Text <> "" Then
        campo = "{sfactik.fecfacftg}"
        'Parametro Desde/Hasta Cliente
        devuelve = "pDHFecha=""Fecha " & devuelve & ": "
        If Not PonerDesdeHasta(campo, "F", 20, 21, devuelve) Then Exit Sub
    End If
    
    'OCTUBRE 2009
    'Al listado debe enivar para que enlaze que en scafac el tipom tiene que ser FTI
    campo = "{scafac.codtipom} = ""FTI"""
    AnyadirAFormula cadSelect, campo
    AnyadirAFormula cadFormula, campo

    'FALTA### comprobar
    '-------------------------------------
    cadFrom = "sfactik,scafac"
    devuelve = "`sfactik`.`numfactu`=`scafac`.`numfactu`) AND (`sfactik`.`fecfactu`=`scafac`.`fecfactu`) AND " & cadSelect
    If Not HayRegParaInforme(cadFrom, cadSelect) Then Exit Sub
    
    
    conSubRPT = False
    LlamarImprimir True
    Screen.MousePointer = vbDefault
End Sub

Private Sub ContabilizarTickets()
    'IdTrabajador
    'Es importante para las tablas de analitica. Es el que pasa el CC
    If txtTrab(2).Text = "" Then
        MsgBox "Introduza el trabajador que realiza la contabilización", vbExclamation
        Exit Sub
    End If
    
    
    
    'La fecha HASTA sera la fecha de factura para los
    If Me.optTick(1).Value Then
        'MENSUAL
        If txtFecha(21).Text = "" Then
            MsgBox "Debe poner la fecha ""hasta"". Será la fecha de factura ", vbExclamation
            Exit Sub
        End If
    End If
    
    
    'Compruebo si existe el tipo moviemiento
    campo = DevuelveDesdeBD(conAri, "nomtipom", "stipom", "codtipom", "FTG", "T")
    If campo = "" Then
        MsgBox "Falta definir el tipo de moviemiento: FTG", vbExclamation
        Exit Sub
    End If
    
    'Compruebo que no se ha quedado ningun FTG con anteriroridad
    campo = DevuelveDesdeBD(conAri, "numfactu", "scafac", "codtipom", "FTG", "T")
    If campo <> "" Then
        'EXISTE FTG sin haber sido borrado
        MsgBox "Existen FTG que no han sido borrados", vbExclamation
        Exit Sub
    End If
    
    
    
    
    If vEmpresa.TieneAnalitica Then
        cadFrom = ""  'cadena error
        
        
        'Falta configurar la forma de envio en empresa
        campo = DevuelveDesdeBD(conAri, "nomenvio", "senvio", "codenvio", vParamAplic.PorDefecto_Envio)
        If campo = "" Then cadFrom = "- Forma de envio en los parametros de la aplicacion" & vbCrLf
        
        'Comprobar que existen todos los centros de coste en los datos a facturar
        campo = DevuelveDesdeBD(conAri, "codccost", "straba", "codtraba", txtTrab(2).Text)
        If campo = "" Then
            cadFrom = cadFrom & "- Trabajador sin CC asignado: " & txtTrab(2).Text & vbCrLf
        Else
            'Tiene CC puesto. Veremos que existe en la conta
            devuelve = DevuelveDesdeBD(conConta, "nomccost", "cabccost", "codccost", campo, "T")
            If devuelve = "" Then cadFrom = cadFrom & "- Centro de coste del trabajador NO existe." & campo
        End If
        
        If cadFrom <> "" Then
            MsgBox "Falta configurar." & vbCrLf & cadFrom, vbExclamation
            Exit Sub
        End If
    End If
    
    InicializarVbles
    
    
    
    
    'Obtengo el que sera el ultimo registro insertado hasta ahora.
    'No hace falta. TODO proceso debe eliminar las facturas FTG
    'campo = SugerirCodigoSiguienteStr("scafac", "numfactu", "codtipom=""FTG""")
    'NumRegElim = Val(campo)
    
    
    
    cadSelect = " codtipom='FTG'"
    
    campo = "scafac.fecfactu"
    If txtFecha(20).Text <> "" Or txtFecha(21).Text <> "" Then
        If Not PonerDesdeHasta(campo, "F", 20, 21, devuelve) Then Exit Sub
    End If
                    
    'Compruebo si hay facturas FTG que no han sido contabilizadas
    If HayRegParaInforme("scafac", cadSelect, True) Then
        'Existen registros anterior pendientes de contabilizar
        MsgBox "Existen facturas FTG que no han sido contabilizadas"
    End If
                    
    
    'Compruebo que no hay FTI inferiores a la fecha
    If txtFecha(20).Text <> "" Then
        cadNomRPT = "codtipom = 'FTI' and intconta=0 and fecfactu<'" & Format(txtFecha(20).Text, FormatoFecha) & "'"
        If HayRegParaInforme("scafac", cadNomRPT, True) Then
            MsgBox "Existen Tickets pendientes de contabilizar con fecha inferior a la solicitada", vbExclamation
            Exit Sub
        End If
    End If
    
    
    
    campo = DevuelveDesdeBD(conAri, "codclien", "spatpvg", "codigo", "1", "N")
    If campo = "" Then
        MsgBox "No se ha encotrnado el cliente ""varios""", vbExclamation
        Exit Sub
    End If
    NumRegElim = Val(campo)
    
    
    
    
    'Monto la select de las facturas FTI
    cadSelect = " intconta = 0 and codtipom='FTI'"
    campo = "scafac.fecfactu"
    If txtFecha(20).Text <> "" Or txtFecha(21).Text <> "" Then
        If Not PonerDesdeHasta(campo, "F", 20, 21, devuelve) Then Exit Sub
    End If
    
    If Not HayRegParaInforme("scafac", cadSelect) Then Exit Sub
            
            
            
    'Si la contbilizacion es menusal, voy a ver si las fechas estan en el mismo mes
    'Si es mas de un mes NO dejo continuar
    If Me.optTick(1).Value Then
        Set miRsAux = New ADODB.Recordset
        miSQL = "Select distinct(fecfactu) from scafac WHERE " & cadSelect & " ORDER BY fecfactu"
        miRsAux.Open miSQL, conn, adOpenKeyset, adLockPessimistic, adCmdText
        numParam = 0
        If Not miRsAux.EOF Then
            miSQL = Format(miRsAux!FecFactu, "mmyyyy")
            miRsAux.MoveLast
            campo = Format(miRsAux!FecFactu, "mmyyyy")
            If miSQL <> campo Then
                MsgBox "Las fechas de los tickets a contabilizar NO son del mismo mes. " & miSQL & " " & campo, vbExclamation
                numParam = 1
            End If
        End If
        miRsAux.Close
        Set miRsAux = Nothing
        If numParam = 1 Then Exit Sub
    End If
    
    
    
    
    
    'Hay datos. Hago la pregunta
    campo = "Va a contabilizar los tickets agrupados. " & vbCrLf & "Se generará una factura "
    If Me.optTick(1).Value Then
        'Va a cojer un mes. Avisaremos que el periodo de facturacion es superior a un mes
        campo = campo & "con fecha: " & txtFecha(21).Text
    Else
        campo = campo & "por dia"
    End If
    campo = campo & vbCrLf & vbCrLf & "¿Continuar?"
    If MsgBox(campo, vbQuestion + vbYesNo) = vbNo Then Exit Sub
            
            
            
            
    'Si tiene registros hare la contabilizacion
    DesBloqueoManual ("GT")
    If Not BloqueoManual("GT", "1") Then
        MsgBox "Proceso inciado por otro usuario.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    Label5.Caption = "Inicio proceso facturacion/contabilizacion"
    
    
    Set miRsAux = New ADODB.Recordset
    
    
    HacerFacturaTICKETS
    
    Set miRsAux = Nothing
    
    'Liberamos el bloqueo
    DesBloqueoManual ("GT")

    Espera 0.5

    
End Sub

Private Sub cmdControlAlbaranes_Click()

    InicializarVbles
        'El nombre de la empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    
    
    If txtEnvio(0).Text <> "" Or txtEnvio(1).Text <> "" Then
        devuelve = "pDHEnvio=""Envio: "
        campo = "{scaalb.codenvio}"
        If Not PonerDesdeHasta(campo, "ENV", 0, 1, devuelve) Then Exit Sub
    End If
    
    miSQL = ""
    If txtZona(2).Text <> "" Or txtZona(3).Text <> "" Then
        devuelve = "Zona: "
        campo = "{scaalb.codzonas}"
        If Not PonerDesdeHasta(campo, "ZON", 2, 3, devuelve) Then Exit Sub
        miSQL = miSQL & devuelve
    End If
    
    If txtFecha(45).Text <> "" Or txtFecha(46).Text <> "" Then
        devuelve = "Fecha envio: "
        campo = "{scaalb.fecenvio}"
        If Not PonerDesdeHasta(campo, "F", 45, 46, devuelve) Then Exit Sub
        miSQL = Trim(miSQL & "      " & devuelve)
    End If
    
    
    devuelve = "pDHResto=""" & miSQL & """|"
    cadParam = cadParam & devuelve
    numParam = numParam + 1
    
    'SI es FACTURADOS, se ira a la tabla de scafac1,scafac
    ' si es albaranes a scaalb
    ' Al ser facturados no lleva visble los campos de ZONA
    If Opcion = 43 Then
        cadSelect = Replace(cadSelect, "{scaalb.", "{scafac1.")
        cadFormula = Replace(cadFormula, "{scaalb.", "{scafac1.")
    End If
    

    cadSelect = QuitarCaracterACadena(cadSelect, "{")
    cadSelect = QuitarCaracterACadena(cadSelect, "}")
     

    'campo = cadSelect
    If Opcion = 42 Then
        campo = " scaalb left join senvio on scaalb.codenvio=senvio.codenvio "
        campo = campo & " left join szonas on  scaalb.codzonas=szonas.codzonas"
        cadNomRPT = "rControlAlbaranes.rpt"
        cadTitulo = "Control albaranes"
    
    Else
        campo = " scafac left join scafac1 on scafac.codtipom=scafac1.codtipom "
        campo = campo & " AND scafac.numfactu=scafac1.numfactu AND scafac.fecfactu=scafac1.fecfactu"
        cadNomRPT = "rControlAlbaranesFac.rpt"
        cadTitulo = "Control albaranes facturados"
    End If
    If Not HayRegParaInforme(campo, cadSelect) Then Exit Sub
    
    
    
 
    
    LlamarImprimir True

End Sub

Private Sub cmdCrearCliente_Click()
    'YA NO SE HACE DESDE AQUI
    Exit Sub
    
    'Creara un cliente desde potenciales
    'De momento SOLO lo ha pedido Bacchus
    miSQL = ""
    If txtForpa(3).Text = "" Or txtAgente(10).Text = "" Then miSQL = "N"
    If txtDescForpa(3).Text = "" Or txtDescAgente(10).Text = "" Then miSQL = "N"
    If txtNumero(2).Text = "" Then miSQL = "N"
    If txtNumero(3).Text = "" Then miSQL = "N"
    If miSQL <> "" Then
        MsgBox "Campos obligatorios", vbExclamation
        Exit Sub
    End If
    
    
    miSQL = DevuelveDesdeBD(conAri, "nomclien", "sclien", "codclien", txtNumero(2).Text)
    If miSQL <> "" Then
        MsgBox "Ya existe el codigo de cliente: " & txtNumero(2).Text & " " & miSQL, vbExclamation
        Exit Sub
    End If
    
    
    'La cuenta contable en contabilidad.
    'codmacta
    If Len(txtNumero(3).Text) <> vEmpresa.DigitosUltimoNivel Then
        MsgBox "Longituda cuenta incorrecta", vbExclamation
        Exit Sub
    End If
    
    
    miSQL = DevuelveDesdeBD(conConta, "nommacta", "cuentas", "codmacta", txtNumero(3).Text, "T")
    If miSQL <> "" Then
        campo = "Ya existe cuenta en contabilidad: " & miSQL
    Else
        campo = "Se creará cuenta contable en conta"
    End If
    
    
    miSQL = Mid(txtTextoNoEditable(0).Text, InStr(1, txtTextoNoEditable(0).Text, "-") + 1)
    miSQL = "Va a dar de alta el cliente. " & vbCrLf & "Codigo: " & txtNumero(2).Text & " - " & miSQL & vbCrLf & vbCrLf
    miSQL = miSQL & campo & vbCrLf & vbCrLf & "¿Continuar?"
    
    If MsgBox(miSQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    conn.BeginTrans
    ConnConta.BeginTrans
    
    If crearCliente Then
        conn.CommitTrans
        ConnConta.CommitTrans
        miSQL = ""
    Else
        conn.RollbackTrans
        ConnConta.RollbackTrans
        miSQL = "NO"
    End If
    Screen.MousePointer = vbDefault
    
    If miSQL = "" Then Unload Me
        
End Sub

Private Sub cmdDtoActiv_Click()
    InicializarVbles
        'El nombre de la empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    
    'Proveedor  18 19
    If txtcodactiv(0).Text <> "" Or txtcodactiv(1).Text <> "" Then
        devuelve = "pDHActiv=""Actividad: "
        campo = "{sactivdtos.codactiv}"
        If Not PonerDesdeHasta(campo, "ACT", 0, 1, devuelve) Then Exit Sub
        
    End If
    
    
    'Familia  4 5
    If txtFamia(8).Text <> "" Or txtFamia(9).Text <> "" Then
        devuelve = "pDHFamilia=""Familia: "
        campo = "{sactivdtos.codfamia}"
        If Not PonerDesdeHasta(campo, "FAM", 8, 9, devuelve) Then Exit Sub
        
    End If
    
    
    'Septiembre 2013
    'Si el proveedor tiene la marca de que NO salen en los listados
    campo = " "
    If cadSelect <> "" Then campo = "AND "
    campo = campo & "({sprove.OcultarEnListDto}=0 )"
    cadSelect = cadSelect & campo
    cadFormula = cadFormula & campo

    cadSelect = QuitarCaracterACadena(cadSelect, "{")
    cadSelect = QuitarCaracterACadena(cadSelect, "}")
    

    'campo = cadSelect
    campo = "sactivdtos.codfamia=sfamia.codfamia and sprove.codprove=sfamia.codprove AND "
    campo = campo & cadSelect
    
    If Not HayRegParaInforme("sactivdtos,sfamia,sprove", campo, False) Then Exit Sub
    
    
    
    
    
    
    
    
    LlamarImprimir True
End Sub

Private Sub cmdDtoProve_Click()
    
    InicializarVbles
    
    
    'El nombre de la empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    cadFormula = cadSelect
    'Proveedor  18 19
    If txtCodProve(18).Text <> "" Or txtCodProve(19).Text <> "" Then
        devuelve = "pDHProve=""Proveedor: "
        campo = "{sdtomp.codprove}"
        If Not PonerDesdeHasta(campo, "PRO", 18, 19, devuelve) Then Exit Sub
        
    End If
    
    
    'Familia  4 5
    If txtFamia(4).Text <> "" Or txtFamia(5).Text <> "" Then
        devuelve = "pDHFamilia=""Familia: "
        campo = "{sdtomp.codfamia}"
        If Not PonerDesdeHasta(campo, "FAM", 4, 5, devuelve) Then Exit Sub
        
    End If
    
    
    'Marca 2 3
    
    If txtmarca(2).Text <> "" Or txtmarca(3).Text <> "" Then
        devuelve = "pDHMarca=""Trabajador: "
        campo = "{sdtomp.codmarca}"
        If Not PonerDesdeHasta(campo, "MAR", 2, 3, devuelve) Then Exit Sub
        
    End If
    
    
    'Solo con rappel
    If Me.chkDtoCompra(0).Value Then
        campo = ""
        If cadSelect <> "" Then campo = " AND "
        campo = campo & "({sdtomp.rap1}>0 or {sdtomp.rap2}>0 )"
        cadSelect = cadSelect & campo
        cadFormula = cadFormula & campo
    End If
    
    cadSelect = QuitarCaracterACadena(cadSelect, "{")
    cadSelect = QuitarCaracterACadena(cadSelect, "}")
    
    campo = "sdtomp "
    If cadSelect <> "" Then campo = campo & " WHERE " & cadSelect
    
    If Not HayRegParaInforme(campo, "", True) Then
        MsgBox "No hay registros con estos valores", vbExclamation
        Exit Sub
    End If
    
    'Salto de pagina por proveedor
    If Me.chkDtoCompra(1).Value = 1 Then
        cadParam = cadParam & "Salto=1|"
        numParam = numParam + 1
    End If
    
    
    LlamarImprimir False
End Sub

Private Sub cmdEstadisticaReparacionTecnico_Click()

    If Me.txtTrab(0).Text = "" Then
        MsgBox "Seleccione un técnico", vbExclamation
        Exit Sub
    End If
    
    'ES EL codtraba1
    cadSelect = "schrep.codtrab1 = " & txtTrab(0).Text

    'Ya tenemos el tecnico. Miramos las fechas
    If txtFecha(2).Text <> "" Or txtFecha(3).Text <> "" Then
        'Parametro Desde/Hasta Cliente
        
        'Marzo 2010
        'ANtes
        'campo = "schrep.fecrepar"
        campo = "schrep.fechaalb"
        
        
        If Not PonerDesdeHasta(campo, "F", 2, 3, devuelve) Then Exit Sub
        'Aqui lo añadiremos a  cadparam
        
    End If
    
    
    
    
    Screen.MousePointer = vbHourglass
   
    NumRegElim = 0
    Set miRsAux = New ADODB.Recordset
    'Aqui iremos grabanod los datos.
    'EstadisticaReparacionTecnico
    
    
    EstadisticaReparacionTecnicoNueva
    
    Set miRsAux = Nothing
    Label3(63).Caption = ""
    Screen.MousePointer = vbDefault
    
    If NumRegElim = 0 Then
        MsgBox "Ningun dato a mostrar", vbExclamation
        Exit Sub
    End If
    
    
    'Llegados aqui imprimiremos los registros
    InicializarVbles
    
    '===================================================
    '============ PARAMETROS ===========================
    'Añadir el nombre de la Empresa como parametro
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    cadParam = cadParam & "pDHFecha= "" Técnico: " & txtTrab(0).Text & " - " & Me.txtDescTra(0).Text & """|"
    numParam = 2
    campo = ""
    If txtFecha(2).Text <> "" Then campo = "     Desde " & txtFecha(2).Text
    If txtFecha(3).Text <> "" Then campo = campo & "      Hasta " & txtFecha(3).Text
    If campo <> "" Then
        numParam = 3
        campo = "pDHCliente= """ & Trim(campo) & """|"
        cadParam = cadParam & campo
    End If
    cadFormula = "{tmpnlotes.codusu}=" & vUsu.codigo

    cadNomRPT = "rRepEstadisticaTec.rpt"
    conSubRPT = False
    LlamarImprimir False
    
End Sub

'------------------------------------------------------------------
'            F A C T U R A S     P R O V E E      S O C I O S
'------------------------------------------------------------------
Private Sub cmdFacProv_Click()
Dim Conjunto As Collection
Dim TipoM As CTiposMov
    'Comprobaciones iniciales
    cadParam = ""
    If txtFecha(17).Text = "" Then cadParam = cadParam & "- fecha factura" & vbCrLf
    If txtBancoPr(0).Text = "" Then cadParam = cadParam & "- banco propio" & vbCrLf
    If txtForpa(0).Text = "" Then cadParam = cadParam & "- forma de pago" & vbCrLf
    If txtTrab(1).Text = "" Then cadParam = cadParam & "- trabajador" & vbCrLf

    devuelve = ""
    If vParamAplic.PorReten > 0 Then devuelve = "D"
    If vParamAplic.CtaReten = "" Xor devuelve = "" Then cadParam = cadParam & vbCrLf & "- Falta configurar cta retencion -  % retencion en parametros"
    If cadParam <> "" Then
        cadParam = "Campos requeridos: " & vbCrLf & vbCrLf & cadParam
        MsgBox cadParam, vbExclamation
        cadParam = ""
        Exit Sub
    End If
    
    'Tipo de moviemiento de facturas liqueidacion proveedores
    Set TipoM = New CTiposMov
    If Not TipoM.Leer("FLQ") Then  'tipo de movimiento FLQ
        MsgBox "No se puede continuar sin el tipo de moviemiento: FLQ", vbExclamation
        Exit Sub
    End If
    
    'Comprobaciones POSTERIORES ;)
    'Si la fecha esta en correctos
    'FALTA###
    
    
    
    'Cargo en ImpTeo el valor del porcentaje rea
    devuelve = DevuelveDesdeBDNew(conConta, "tiposiva", "porceiva", "codigiva", vParamAplic.IVA_REA)
    If devuelve = "" Then
        'ERROR con el tipo de IVA REA
        MsgBox "Tipo de IVA REA no configurado en parametros, o no existe", vbExclamation
        Exit Sub
    End If
    ImpTeo = CCur(devuelve)
    
    
    
    
    'Vamos a ver el conjunto de albaranes para pasar
    InicializarVbles
    devuelve = ""
    
    
    'Cadena obligada. Los proveedores , el tipo tiene que ser el 3: REA
    cadSelect = " {scaalp.codprove}=  {sprove.codprove}  AND {sprove.tipprove}= 3 "
    
    'Cadena para seleccion Desde y Hasta FECHA
    '--------------------------------------------
    If txtFecha(13).Text <> "" Or txtFecha(14).Text <> "" Then
        campo = "{scaalp.fechaalb}"
        If Not PonerDesdeHasta(campo, "F", 13, 14, devuelve) Then Exit Sub
    End If
    
    If txtCodProve(4).Text <> "" Or txtCodProve(5).Text <> "" Then
        campo = "{scaalp.codprove}"
        If Not PonerDesdeHasta(campo, "PRO", 4, 5, devuelve) Then Exit Sub
    End If
     
    If txtNumAlbar(0).Text <> "" Or txtNumAlbar(1).Text <> "" Then
        campo = "{scaalp.numalbar}"
        If Not PonerDesdeHasta(campo, "ALP", 0, 1, "") Then Exit Sub
    End If
    
    
    'Montamos el select para los registros
    cadSelect = QuitarCaracterACadena(cadSelect, "{")
    cadSelect = QuitarCaracterACadena(cadSelect, "}")
    cadSelect = " scaalp,sprove WHERE " & cadSelect
    
    
    
    If Not HayRegParaInforme(cadSelect, "", True) Then
        MsgBox "No hay albaranes para facturar con estos valores", vbExclamation
        Exit Sub
    Else
        'llegado aqui preguntamos si desea continuar
        cadFrom = "Seguro que desea continuar con la generacion de las facturas de liquidación?"
        If MsgBox(cadFrom, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    End If
    
    
    'Monto el SQL para saber que albaranes facturo
    Screen.MousePointer = vbHourglass
    cadFrom = "Select sprove.codprove,albaranxfactura FROM " & cadSelect & " GROUP by 1,2 ORDER BY 1"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open cadFrom, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Set Conjunto = New Collection
    While Not miRsAux.EOF
        Conjunto.Add miRsAux!Codprove & "|" & miRsAux!albaranxfactura & "|"
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    'AHora vamos a ir facturando los diversos proveedores
    For IndiceImg = 1 To Conjunto.Count
        'Facturamos al proveedor
        FacturarProveedor CLng(RecuperaValor(Conjunto.item(IndiceImg), 1)), Val(RecuperaValor(Conjunto.item(IndiceImg), 2)) = 1, TipoM
    Next IndiceImg
    
    Label1.Caption = ""
    Set TipoM = Nothing
    Set miRsAux = Nothing
    Screen.MousePointer = vbDefault
    MsgBox "Proceso finalizado", vbExclamation
End Sub


Private Sub FacturarProveedor(Codprove As Long, albaranxfactura As Boolean, ByRef Ctip As CTiposMov)
Dim vFactu As CFacturaCom
Dim vProve As CProveedor
Dim cad As String
Dim RA As ADODB.Recordset
Dim ColFacturar As Collection
Dim J As Integer



    Screen.MousePointer = vbHourglass
    On Error GoTo Error1
    
    Set vProve = New CProveedor
    
    'Tiene que ller los datos del proveedor
    If Not vProve.LeerDatos(CStr(Codprove)) Then
        Label1.Caption = "Error leyendo proveedor: " & Codprove
        Me.Refresh
        DoEvents
        Espera 1
        Exit Sub
    End If
    
    
    Label1.Caption = "ALbaranes a facturar proveedor :        " & vProve.Nombre
    Label1.Refresh

    cad = "Select scaalp.numalbar,scaalp.fechaalb FROM " & cadSelect & " AND scaalp.codprove = " & Codprove
    cad = cad & " ORDER BY scaalp.fechaalb,scaalp.numalbar"
    miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    cadFrom = "codprove = " & Codprove & " AND "
    Set ColFacturar = New Collection
    cadNomRPT = ""
    While Not miRsAux.EOF
        cad = "numalbar = '" & DevNombreSQL(miRsAux!NUmAlbar) & "' AND fechaalb = '" & Format(miRsAux!FechaAlb, FormatoFecha) & "'"
        If albaranxfactura Then
            cad = cadFrom & cad
            ColFacturar.Add cad
        Else
            If cadNomRPT <> "" Then cadNomRPT = cadNomRPT & " OR "
            cadNomRPT = cadNomRPT & "(" & cad & ")"
        End If
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If Not albaranxfactura Then
        cad = cadFrom & "(" & cadNomRPT & ")"
        ColFacturar.Add cad
    End If
    
    
    
   'AHORA YA TENGO EN Colfactuar el conjunto de labaraens y/o facturas
    For J = 1 To ColFacturar.Count
       'Pasar los Albaranes seleccionados con cadWHERE a una factura
        Set vFactu = New CFacturaCom
        vFactu.Proveedor = vProve.codigo
        vFactu.Numfactu = Ctip.Contador + 1
        vFactu.FecFactu = txtFecha(17).Text
        vFactu.FecRecep = txtFecha(17).Text
        vFactu.Trabajador = txtTrab(1).Text
        vFactu.BancoPr = txtBancoPr(0).Text
        
        vFactu.ForPago = txtForpa(0).Text
        vFactu.DtoPPago = 0
        vFactu.DtoGnral = 0

        'Si el proveedor tiene CTA BANCARIA se la asigno
        vFactu.CCC_Entidad = vProve.Banco
        vFactu.CCC_Oficina = vProve.Sucursal
        vFactu.CCC_CC = vProve.DigControl
        vFactu.CCC_CTa = vProve.CuentaBan
        vFactu.Iban = vProve.Iban
        
        
    
        'Obtengo los totales mediante el cadselect
        cad = "Select sum(importel) FROM slialp WHERE " & ColFacturar.item(J)
        miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not miRsAux.EOF Then
            ImpTot = DBLet(miRsAux.Fields(0), "N")
        End If
        miRsAux.Close
        
        
        vFactu.BrutoFac = ImpTot
        vFactu.BaseIVA1 = ImpTot
        vFactu.TipoIVA1 = vParamAplic.IVA_REA
        vFactu.PorceIVA1 = ImpTeo
        ImpTot = Round2((ImpTot * ImpTeo) / 100, 2)
        vFactu.ImpIVA1 = ImpTot
        ImpTot = vFactu.BrutoFac + ImpTot  'Base + IVA
        
        'Retencion
        vFactu.TipoRet = 1
        vFactu.PorRet = vParamAplic.PorReten
        vFactu.ImpRet2 = Round2((ImpTot * vFactu.PorRet) / 100, 2)
            
        
        vFactu.TotalFac = vFactu.BrutoFac + vFactu.ImpIVA1 - vFactu.ImpRet2
        

         'El select
         cad = ColFacturar.item(J)
         
         If Not vFactu.TraspasoAlbaranesAFactura(cad, (chkFacturPorv(1).Value = 1), (chkFacturPorv(0).Value = 1), True) Then
            'Para salir y finalizar el procesode facturacion de el proveedor
            cad = "Finalizacion de la facturacion para: " & vProve.Nombre & vbCrLf
            cad = cad & "Proceso: " & J & " / " & ColFacturar.Count & vbCrLf
            cad = cad & vbCrLf & "SQL: " & ColFacturar.item(J)
            MsgBox cad, vbExclamation
            J = ColFacturar.Count + 1  'Para que se salga
        Else
            'incremento el contador de facturas
            Ctip.IncrementarContador Ctip.TipoMovimiento
        End If
'        Set vFactu = Nothing
'

    Next J

    Screen.MousePointer = vbDefault

Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation

End Sub

Private Sub cmdFacturaMov_Click()
Dim AlbaranesGenerados As Collection
Dim MensajeError As String

    'Facuracion recargas moviles
    campo = ""
    If txtCliente(2).Text = "" Then campo = campo & " - Cliente" & vbCrLf
    If txtArticulo(0).Text = "" Then campo = campo & " - Articulo" & vbCrLf
    If txtBancoPr(1).Text = "" Or txtDescBancoPr(1).Text = "" Then campo = campo & " - Bancos propios" & vbCrLf
    
    
    If campo <> "" Then
        campo = "Campos requeridos : " & vbCrLf & campo
        MsgBox campo, vbExclamation
        Exit Sub
    End If
    
    'Alguna comprobacion mas
    
    If txtFecha(8).Text = "" Then
        MsgBox "Ponga la fecha de facturación", vbExclamation
        Exit Sub
    End If
    
    If vEmpresa.TieneAnalitica Then
        'Comprobar que existen todos los centros de coste en los datos a facturar
        'FALTA###
        
    End If
    
    InicializarVbles
    
    
    
    
    'Obtengo el que sera el ultimo registro insertado hasta ahora.
    campo = SugerirCodigoSiguienteStr("stelefonia", "id")
    NumRegElim = Val(campo)
    
    
    
    cadSelect = " id < " & NumRegElim & " AND Facturado = 0 "
    
    campo = "stelefonia.fecha"
    If txtFecha(6).Text <> "" Or txtFecha(7).Text <> "" Then
        If Not PonerDesdeHasta(campo, "F", 6, 7, devuelve) Then Exit Sub
    End If
    
    
    
            
                    
                    
                    
    'Compruebo si tiene registros
    If Not HayRegParaInforme("stelefonia", cadSelect) Then Exit Sub
            
            
            
    'Si tiene registros hare la contabilizacion
    
    DesBloqueoManual ("Telf")
    If Not BloqueoManual("Telf", "1") Then
        MsgBox "Proceso inciado por otro usuario.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    lblIndicadorT.Caption = "Inicio proceso facturacion"
    pb1.Value = 0
    
    Set AlbaranesGenerados = New Collection
    MensajeError = ""
    
    HacerFacturacionTelefonia AlbaranesGenerados, MensajeError
    
    If AlbaranesGenerados.Count > 0 Then
        
        If MensajeError <> "" Then
            campo = "Se generaron  " & AlbaranesGenerados.Count & " albaranes. "
            campo = campo & vbCrLf & vbCrLf & " ERROR GENERANDO ALBARANES" & vbCrLf & MensajeError
            MsgBox campo, vbInformation
        End If
        
        campo = ""
        For NumRegElim = 1 To AlbaranesGenerados.Count
            If campo <> "" Then campo = campo & ","
            campo = campo & AlbaranesGenerados.item(NumRegElim)
        Next NumRegElim
        campo = "scaalb.codtipom = 'ALV' AND scaalb.numalbar IN (" & campo & ")"
        miSQL = "SELECT scaalb.*,sclien.nomclien FROM scaalb INNER JOIN sclien ON scaalb.codclien=sclien.codclien "
        miSQL = miSQL & " WHERE " & campo
        
        TraspasoAlbaranesFacturas miSQL, campo, CStr(Now), txtBancoPr(1).Text, pb1, lblIndicadorT, True, "ALV", "", 1, True, False
        
    Else
        MensajeError = "No se ha generado ningun albaran" & MensajeError
        MsgBox MensajeError, vbExclamation
        
    End If
    'Liberamos el bloqueo
    DesBloqueoManual ("Telf")
    lblIndicadorT.Caption = "Proceso finalizado"
    Espera 0.3
End Sub

Private Sub cmdFacturarCli_Click()
    If txtFecha(31).Text = "" Or txtBancoPr(2).Text = "" Then
        MsgBox "Campos obligatorios", vbExclamation
        Exit Sub
    End If
        
    If MsgBox("¿Seguro que desa continuar con la facturación?", vbQuestion + vbYesNo) <> vbYes Then Exit Sub
    
    CadenaDesdeOtroForm = txtFecha(31).Text & "|" & txtBancoPr(2).Text & "|" & chkFacturarCliente.Value & "|"
    Unload Me
End Sub

Private Sub cmdFrecuencia_Click()
    If texto(5).Text = "" Then
        MsgBox "Ponga el expediente", vbExclamation
        Exit Sub
    End If
    
    miSQL = "¿Desea actualizar los cambios?"
    If MsgBox(miSQL, vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    CadenaDesdeOtroForm = texto(5).Text & "|" & Abs(Me.chkFrecu.Value) & "|"
    Unload Me
End Sub

Private Sub cmdGenAlbRep_Click()
    If txtFecha(26).Text = "" Then
        MsgBox "Ponga la fecha del albaran", vbExclamation
        Exit Sub
    End If
    CadenaDesdeOtroForm = txtFecha(26).Text & "|"
    For numParam = 0 To 4
        CadenaDesdeOtroForm = CadenaDesdeOtroForm & texto(numParam).Text & "|"
    Next
    numParam = 0
    Unload Me
End Sub

Private Sub cmdImprimirFac_Click()
    'Impresion de las facturas de proveedores
    'es decir , para casos de cooperativas en las cuales el socio
    'es el que nos emite la factura a nosotros (ej TERRASANA)
    
    
    
    
    
    InicializarVbles
    
    If Not PonerParamRPT2(26, cadParam, numParam, cadNomRPT, vImprimedirecto, cadPDFrpt, vMultiInforme) Then Exit Sub
    
    cadSelect = "{sprove.tipprove}=3"   'Estos proveedores son los REA que luego
    cadFormula = "(" & cadSelect & ")"                                    'les emitire SUS facturas
    If txtFecha(15).Text <> "" Or txtFecha(16).Text <> "" Then
        campo = "{scafpc.fecfactu}"
        If Not PonerDesdeHasta(campo, "F", 15, 16, devuelve) Then Exit Sub
    End If
    
    If txtCodProve(6).Text <> "" Or txtCodProve(7).Text <> "" Then
        campo = "{scafpc.codprove}"
        If Not PonerDesdeHasta(campo, "PRO", 6, 7, devuelve) Then Exit Sub
    End If
     
    If txtNumAlbar(2).Text <> "" Or txtNumAlbar(3).Text <> "" Then
        campo = "{scafpc.numfactu}"
        If Not PonerDesdeHasta(campo, "ALP", 2, 3, "") Then Exit Sub
    End If
    
    
    'Montamos el select para los registros
    cadSelect = QuitarCaracterACadena(cadSelect, "{")
    cadSelect = QuitarCaracterACadena(cadSelect, "}")
    campo = "scafpc,sprove WHERE scafpc.codprove=sprove.codprove AND " & cadSelect
    
    
    If Not HayRegParaInforme(campo, "", True) Then
        MsgBox "No hay facturas con estos valores", vbExclamation
        Exit Sub
    End If
    
    
    LlamarImprimir True
End Sub

Private Sub cmdImprPlatil_Click()
     
    InicializarVbles
    miSQL = ""
    devuelve = ""
    'Pasar nombre de la Empresa como parametro
    cadParam = "|pNomEmpre=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    '====================================================
    '================= FORMULA ==========================
    'Cadena para seleccion D/H CLIENTE
    '--------------------------------------------
     If txtNumero(0).Text <> "" Or txtNumero(1).Text <> "" Then
        campo = ""
        'Parametro Desde/Hasta Cliente
        If txtNumero(0).Text <> "" Then campo = "Desde " & txtNumero(0).Text
        If txtNumero(1).Text <> "" Then campo = campo & " hasta " & txtNumero(1).Text
        devuelve = "Plantilla: " & campo
        codigo = CadenaDesdeHastaBD(txtNumero(0).Text, txtNumero(1).Text, "{slipla.codplant}", "N")
        AnyadirAFormula cadFormula, codigo
        If Not AnyadirAFormula(cadSelect, codigo) Then Exit Sub
        
    End If
    If txtGrupoPlan(0).Text <> "" Or txtGrupoPlan(1).Text <> "" Then
        campo = ""
        'Parametro Desde/Hasta Cliente
        campo = ""
        'Parametro Desde/Hasta Cliente
        If txtGrupoPlan(0).Text <> "" Then campo = " Desde " & txtGrupoPlan(0).Text & " " & Me.txtDescGrupoP(0).Text
        If txtGrupoPlan(1).Text <> "" Then campo = campo & "  hasta " & txtGrupoPlan(1).Text & " " & Me.txtDescGrupoP(1).Text
        devuelve = devuelve & " Grupo: " & campo
        
        codigo = CadenaDesdeHastaBD(txtNumero(0).Text, txtNumero(1).Text, "{slipla.codgrupl}", "N")
        If Not AnyadirAFormula(cadSelect, codigo) Then Exit Sub
        
    End If
    codigo = "pDesde=""" & devuelve & """|"
    cadParam = cadParam & codigo
    numParam = numParam + 1
    cadNomRPT = "rFacPlantillas.rpt"
    LlamarImprimir False
    

End Sub

Private Sub cmdInformeProductividad_Click()
        
    InicializarVbles
        
    miSQL = ""
    If txtFecha(47).Text <> "" Or txtFecha(48).Text <> "" Then
        devuelve = "Fecha: "
        campo = "{sreloj.fecha}"
        If Not PonerDesdeHasta(campo, "F", 47, 48, devuelve) Then Exit Sub
        miSQL = miSQL & devuelve
    End If
        
    If txtTrab(7).Text <> "" Or txtTrab(8).Text <> "" Then
        devuelve = "    Trabajador. "
        campo = "{sreloj.codtraba}"
        If Not PonerDesdeHasta(campo, "TRA", 7, 8, devuelve) Then Exit Sub
        miSQL = Trim(miSQL & devuelve)
    End If
    
    If txtNumAlbar(6).Text <> "" Or txtNumAlbar(6).Text <> "" Then
        devuelve = "       Nº Documento "
        campo = "{sreloj.numalbar}"
        If Not PonerDesdeHasta(campo, "ALP", 6, 7, devuelve) Then Exit Sub
        If Len(miSQL) > 70 Then
            miSQL = miSQL & """ + chr(13) + """
            devuelve = Trim(devuelve)
        End If
        miSQL = Trim(miSQL & devuelve)
    End If
    
    campo = ""
    cadTitulo = ""
    NumRegElim = 0
    For numParam = 0 To 4  'EL de produccion no entra aqui
        If numParam <> 3 Then
            If Me.chkInformeProd(numParam).Value = 1 Then
                campo = campo & ", '" & RecuperaValor("ALR|ALE|ALO||ALV|", CInt(numParam) + 1) & "'"
                cadTitulo = cadTitulo & "- " & RecuperaValor("Reparación|T. Exterior|Orden Trabajo||Alb. venta|", CInt(numParam) + 1)
                cadPDFrpt = cadPDFrpt & numParam
                NumRegElim = NumRegElim + 1
            End If
        End If
    Next
    If campo <> "" Then
        campo = Mid(campo, 2)
        campo = "{sreloj.codtipom} IN [" & campo & "]"
    End If
    
    'numparam vale 3
    If Me.chkInformeProd(3).Value = 1 Then
        If campo <> "" Then campo = campo & " OR "
        campo = campo & " {sreloj.codtipom} is null"
        cadPDFrpt = cadPDFrpt & numParam
        cadTitulo = cadTitulo & "- Producción"
        NumRegElim = NumRegElim + 1
    End If
    If NumRegElim = 0 Then
        MsgBox "Seleccione algun tipo de documento", vbExclamation
        Exit Sub
    End If
    cadTitulo = "       Tipo trabajo: " & Mid(cadTitulo, 2) 'quito el primer guion
    
    If NumRegElim <> 4 Then
        'No ha seleccionado todos. Pong el SQL
        If Len(miSQL) > 70 Then
            If InStr(1, miSQL, "chr(13)") = 0 Then miSQL = miSQL & """ + chr(13) + """
        End If
        miSQL = Trim(miSQL & cadTitulo)
        
        campo = "(" & campo & ")"
        cadPDFrpt = Replace(campo, "[", "(")
        cadPDFrpt = Replace(cadPDFrpt, "]", ")")
        If cadSelect <> "" Then cadSelect = cadSelect & " AND "
        cadSelect = cadSelect & cadPDFrpt
        If cadFormula <> "" Then cadFormula = cadFormula & " AND "
        cadFormula = cadFormula & "(" & campo & ")"
        

        
   End If

    
    cadParam = "|pDH= """ & miSQL & """|"  'Familia=Marca
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = 2
    
    'Ver si hayb registros
    cadSelect = QuitarCaracterACadena(cadSelect, "{")
    cadSelect = QuitarCaracterACadena(cadSelect, "}")
    
    campo = "sreloj "
    If cadSelect <> "" Then campo = campo & " WHERE " & cadSelect
    
    If Not HayRegParaInforme(campo, "", True) Then
        MsgBox "No hay registros con estos valores", vbExclamation
        Exit Sub
    End If
    
    
    conn.Execute "Delete from tmpInformes WHERE codusu = " & vUsu.codigo
    
    
    '                   codusu,             codigo1,            campo2,     nombre1
    campo = "Select " & vUsu.codigo & ",@rownum:=@rownum+1 AS rownum,sreloj.codtraba,nomtraba,"
    '               ,   campo1,     nombre2,                        fecha
    campo = campo & "numalbar ,if(codtipom is null,'PROD',codtipom),Fecha,"
    'nombre3,
    campo = campo & "concat(DATE_FORMAT(HoraInicio, '%H:%i:%s'),' - ' ,if(HoraFin is null,'',"
    
    campo = campo & "if( date(horafin)=fecha,DATE_FORMAT(Horafin, '%H:%i:%s'),DATE_FORMAT(Horafin, '%H:%i:%s  (%d/%m) '))"
    campo = campo & ")),"
    'importe1
    'campo = campo & " if(horafin is null,0,Hour (timediff(horafin, horainicio)) + Round(Minute(timediff(horafin, horainicio)) / 60, 2))"
    campo = campo & " if(horafin is null,0,calculadas)"
    'Tarea
    campo = campo & ", concat(sreloj.codtipor,' ',coalesce(nomtipor,''))"
    
    campo = campo & " FROM (SELECT @rownum:=0) r,sreloj  left join straba on sreloj.codtraba=straba.codtraba"
    campo = campo & " left join stipor on sreloj.codtipor=stipor.codtipor"
    If cadSelect <> "" Then campo = campo & " WHERE " & cadSelect

    campo = "INSERT INTO tmpinformes(codusu,codigo1,campo2,nombre1,campo1,nombre2,fecha1,nombre3,importe1,obser) " & campo
    If ejecutar(campo, False) Then
        Espera 0.5
        campo = "update tmpinformes set importe2=floor(importe1),"
        campo = campo & "Importe3 = Round(Round(Importe1 - floor(Importe1), 2) * 100 * 0.6)"
        campo = campo & " Where CodUsu = " & vUsu.codigo & " And  Importe1 > 0"
        ejecutar campo, False
    
    
        'A imprimir
        cadTitulo = "Informe produccion"
        cadNomRPT = "rproductividad2.rpt"
        If Me.optInfProd(1).Value Then cadNomRPT = Replace(cadNomRPT, "2", "") 'quito el 2
        
        cadPDFrpt = cadNomRPT
        cadFormula = "{tmpinformes.codusu} = " & vUsu.codigo
        LlamarImprimir True
    End If
            

    
End Sub

Private Sub cmdListTrabja_Click()

    InicializarVbles
    
    If Not PonerParamRPT2(35, cadParam, numParam, cadNomRPT, vImprimedirecto, cadPDFrpt, vMultiInforme) Then Exit Sub
     
    If txtTrab(3).Text <> "" Or txtTrab(4).Text <> "" Then
        campo = "{straba.codtraba}"
        If Not PonerDesdeHasta(campo, "TRA", 3, 4, devuelve) Then Exit Sub
    End If
        
    LlamarImprimir True
End Sub

Private Sub cmdLlamadas_Click()


    InicializarVbles
    
    
    If Not PonerParamRPT2(41, cadParam, numParam, cadNomRPT, vImprimedirecto, cadPDFrpt, vMultiInforme) Then Exit Sub
    
    'El nombre de la empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    cadFormula = cadSelect
   
    If txtFecha(27).Text <> "" Or txtFecha(28).Text <> "" Then
        devuelve = "pDHFecha=""Fecha: "
        campo = "{sllama.feholla}"
        If Not PonerDesdeHasta(campo, "F", 27, 28, devuelve) Then Exit Sub
        
    End If
    
    If txtTrab(5).Text <> "" Or txtTrab(6).Text <> "" Then
        devuelve = "pdhTra=""Trabajador: "
        campo = "{sllama.codtraba}"
        If Not PonerDesdeHasta(campo, "TRA", 5, 6, devuelve) Then Exit Sub
        
    End If
    
    
    '
    cadSelect = QuitarCaracterACadena(cadSelect, "{")
    cadSelect = QuitarCaracterACadena(cadSelect, "}")
    
    campo = "sllama "
    If cadSelect <> "" Then campo = campo & " WHERE " & cadSelect
    
    If Not HayRegParaInforme(campo, "", True) Then
        MsgBox "No hay registros con estos valores", vbExclamation
        Exit Sub
    End If
    
    LlamarImprimir False





End Sub

Private Sub cmdMarcaFamilia_Click()



    'Poner desde hastas, Comun para compras ventas
       
    InicializarVbles
    'El nombre de la empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    
    
    
    miSQL = ""
    
    If txtFecha(51).Text <> "" Or txtFecha(52).Text <> "" Then
    
        devuelve = " Fecha: "
        If Opcion = 49 Then
            campo = "{scafac.fecfactu}"
        Else
            campo = "{scafpc.fecfactu}"
        End If
        If Not PonerDesdeHasta(campo, "F", 51, 52, devuelve) Then Exit Sub
        miSQL = devuelve & "   "
        'Si ha puesto mes
        
    End If
    
    
        
    cadParam = cadParam & "pdh1=""" & Trim(miSQL) & """|"
    numParam = numParam + 1
    
    
    'Cadena para seleccion Desde y Hasta FECHA
    '--------------------------------------------
    miSQL = ""
    
    If txtmarca(10).Text <> "" Or txtmarca(11).Text <> "" Then
        devuelve = "  Marca: "
        campo = "{sartic.codmarca}"
        If Not PonerDesdeHasta(campo, "MAR", 10, 11, devuelve) Then Exit Sub
        miSQL = Trim(miSQL & "  " & devuelve)
    End If
    
    If txtFamia(16).Text <> "" Or txtFamia(17).Text <> "" Then
        devuelve = "Familia: "
        campo = "{sartic.codfamia}"
        If Not PonerDesdeHasta(campo, "FAM", 16, 17, devuelve) Then Exit Sub
        If Len(miSQL) > 70 Then
            miSQL = miSQL & """ + chr(13) + """
        Else
            miSQL = miSQL & " "
        End If
        miSQL = Trim(miSQL & Trim(devuelve))
    End If
    
    
    
    
    
    If Opcion = 49 Then
        If txtAgente(13).Text <> "" Or txtAgente(14).Text <> "" Then
            devuelve = "Agente: "
            campo = "{scafac.codagent}"
            If Not PonerDesdeHasta(campo, "AGT", 13, 14, devuelve) Then Exit Sub
            If Len(miSQL) > 50 Then
                miSQL = miSQL & """ + chr(13) + """
            Else
                miSQL = miSQL & " "
            End If
            miSQL = Trim(miSQL & devuelve)
        End If
        
    Else
        If txtCodProve(26).Text <> "" Or txtCodProve(27).Text <> "" Then
            devuelve = " Proveedor: "
            campo = "{scafpc.codprove}"
            If Not PonerDesdeHasta(campo, "PRO", 26, 27, devuelve) Then Exit Sub
            If miSQL <> "" Then
                miSQL = miSQL & """ + chr(13) + """
                devuelve = Trim(devuelve)
            End If
            miSQL = Trim(miSQL & "  " & devuelve)
        End If
    End If
    
    cadParam = cadParam & "pdh2=""" & miSQL & """|"
    numParam = numParam + 1
    
    
    If Opcion = 49 Then
        campo = ""
        If cadSelect <> "" Then campo = " AND "
        cadSelect = cadSelect & campo & " {sartic.artvario} =0 "
        cadSelect = cadSelect & " AND scafac.codtipom <> 'FAZ'"
    Else
        'En el listado compras por familia, no separa artvario
    End If
     
    
    
    
    
    
    
    
    
    
    
    'Proceso generacion datos
    Screen.MousePointer = vbHourglass
    VentasMarcaFamilia
    Screen.MousePointer = vbDefault
    

    Label3(188).Caption = ""
    campo = DevuelveDesdeBD(conAri, "count(*)", "tmpinformes", "codusu", CStr(vUsu.codigo))
    If campo = "" Then campo = "0"
    If Val(campo) = 0 Then
        'No existen datos
        MsgBox "No existen datos", vbExclamation
    Else
        'Si hay datos en tmp
        cadNomRPT = IIf(Opcion = 49, "rVentasMarcaFamilia.rpt", "rComprasMarcaFamilia.rpt")
        cadFormula = "{tmpinformes.codusu} = " & vUsu.codigo
        cadParam = cadParam & "DetallaArticulo=" & Abs(Me.chkMarcaFamilia(0).Value) & "|"
        numParam = numParam + 1
        LlamarImprimir False
    End If
        
        
    
    
End Sub

Private Sub cmdMultibase_Click()
    'Revision caracteres multibase
    numParam = 0
    For NumRegElim = 1 To Me.lstMultibase.ListCount
        If Me.lstMultibase.Selected(CInt(NumRegElim - 1)) Then numParam = numParam + 1
    Next
    

    If numParam = 0 Then
        MsgBox "Seleccion alguna tabla para cambiar", vbExclamation
        Exit Sub
    End If
    
    If MsgBox("Este proceso puede durar mucho tiempo." & vbCrLf & "¿Desea continuar?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    Screen.MousePointer = vbHourglass
    Me.Tag = ""
    Set miRsAux = New ADODB.Recordset
    For numParam = 0 To Me.lstMultibase.ListCount - 1
        If Me.lstMultibase.Selected(CInt(numParam)) Then HacerCambiosMultibase CInt(numParam + 1)
    Next
    Me.lblMultibase.Caption = ""
    Set miRsAux = Nothing
    Screen.MousePointer = vbDefault
    If Me.Tag <> "" Then
        codigo = "Se han realizado los siguientes cambios:" & vbCrLf & vbCrLf & Me.Tag
        Me.Tag = ""
    Else
        codigo = "Proceso finalizado. No se efectuaron cambios"
    End If
    MsgBox codigo, vbInformation
End Sub

Private Sub cmdMultibase2_Click()
    If cboRoot.ListIndex = 1 Then
        If cboTablas.ListIndex < 0 Then
            MsgBox "Seleccione los campos", vbExclamation
            Exit Sub
        End If
    End If
    If MsgBox("Va a actualizar en los campos seleccionados. ¿Continuar?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    
    Screen.MousePointer = vbHourglass
    Set miRsAux = New ADODB.Recordset
    If vUsu.Nivel = 0 Then cboRoot.visible = False
    
    If cboRoot.ListIndex = 2 Then
        'Vamos a recuperar los carcateres incorrectos de un backup MAL recuperado
        UpdatearRestoreBakcup_
    Else
        UpdatearTablaRoot
    End If
    
    cadFrom = ""
    Me.lblMultibase.Caption = ""
    Set miRsAux = Nothing
    If vUsu.Nivel = 0 Then cboRoot.visible = True
    Screen.MousePointer = vbDefault
End Sub


Private Sub UpdatearRestoreBakcup_()
Dim i As Integer
Dim SQL As String

    For i = 1 To TreeView1.Nodes.Count
        If Not TreeView1.Nodes(i).Parent Is Nothing Then
            Me.lblMultibase.Caption = TreeView1.Nodes(i).Parent.Text
            Me.lblMultibase.Refresh
            If TreeView1.Nodes(i).Checked Then
                For numParam = 1 To 8
                    CarcateresRestores numParam, campo, devuelve
                    SQL = "UPDATE " & TreeView1.Nodes(i).Parent.Text & " SET "
                    SQL = SQL & TreeView1.Nodes(i) & " = REPLACE(" & TreeView1.Nodes(i) & ",'" & campo & "','" & devuelve & "') "
                    If Not ejecutar(SQL, False) Then Exit Sub
                Next numParam
            End If
        End If
    Next i
End Sub

Private Sub CarcateresRestores(Cual As Byte, C1 As String, C2 As String)
    Select Case Cual
    Case 1
        C1 = "Ã‘": C2 = "Ñ"

    Case 2
        C1 = "Ã±": C2 = "ñ"
    Case 3
        C1 = "Ã©": C2 = "é"
    
    Case 4
        C1 = "Ã­": C2 = "í"
    Case 5
        C1 = "Âº": C2 = "º"

    Case 6
        C1 = "Ã³": C2 = "ó"
    Case 7
        C1 = "Â±": C2 = "±"
    Case Else
        C1 = "Ã¡": C2 = "á"
    End Select





    
'
'select domclien,REPLACE(domclien,'Ã‘','Ñ') from sclien
'select domclien,REPLACE(domclien,'Ã±','ñ') from sclien
'select domclien,REPLACE(domclien,'Ã©','é') from sclien
'select domclien,REPLACE(domclien,'Ã­','í') from sclien
'select domclien,REPLACE(domclien,'Âº','º') from sclien
'select domclien,REPLACE(domclien,'Ã³','ó') from sclien
'select domclien,REPLACE(domclien,'Ã¡','á') from sclien
    
End Sub

Private Sub cmdP_Click()
Dim i As Integer

    If txtFecha(34).Text = "" Then
        MsgBox "Ponga la fecha de cambio", vbExclamation
        Exit Sub
    End If
    
    
    
    InicializarVbles
    
    'Me importara el cadselect al final
    If txtCodProve(13).Text <> "" Or txtCodProve(14).Text <> "" Then
        campo = "sartic.codprove"
        devuelve = ""
        If Not PonerDesdeHasta(campo, "PRO", 13, 14, devuelve) Then Exit Sub
    End If
    
    If txtFamia(0).Text <> "" Or txtFamia(1).Text <> "" Then
        campo = "sartic.codfamia"
        devuelve = ""
        If Not PonerDesdeHasta(campo, "FAM", 0, 1, devuelve) Then Exit Sub
    End If
    
    If txtArticulo(6).Text <> "" Or txtArticulo(7).Text <> "" Then
        campo = "sartic.codartic"
        devuelve = ""
        If Not PonerDesdeHasta(campo, "ART", 6, 7, devuelve) Then Exit Sub
    End If
    
    
    cadFrom = ""
    If Me.optCopiaPrecio(0).Value Then
        cadParam = "slispr"
    Else
        cadParam = "slista"
        cadFrom = " AND codlista = " & vParamAplic.CodTarifa
    End If
    If cadSelect <> "" Then cadSelect = " AND " & cadSelect
    
    
    'Lo meteremos en campo
    campo = ""
    campo = " l.codartic=sartic.codartic " & campo & cadFrom
   
    
    
    'En cadselect tengo el where.  Ahora lo completo con las tablas y  joins
    campo = campo & cadSelect & " AND fechanue = " & DBSet(txtFecha(34).Text, "F")
    If Not HayRegParaInforme("sartic," & cadParam & " l", campo, True) Then
        MsgBox "No hay precios a actualizar con estos valores", vbExclamation
        Exit Sub
   End If
   
        
   'Ahora voy a comprobar si en la listad DONDE voy a comprar hay arituclos con fecha de cambio.
   '

    cadFrom = ""
    If Me.optCopiaPrecio(0).Value Then
        cadParam = "slista"
        cadFrom = " AND codlista = " & vParamAplic.CodTarifa
    
    Else
        cadParam = "slispr"
    End If
   
    campo = " l.codartic=sartic.codartic "
    campo = campo & " AND fechanue>'1900-01-01'" 'Que tiene fecha cambio
    campo = campo & cadSelect & cadFrom
    If HayRegParaInforme("sartic," & cadParam & " l", campo, True) Then
        MsgBox "Hay precios pendientes de actualizar ", vbExclamation
        Exit Sub
    End If
    
    'Si actualizamos en slista (ventas), y tiene actualizar precio especial
    If Me.optCopiaPrecio(0).Value Then
        If vParamAplic.ActualizaPrecioEspecial Then
            campo = " l.codartic=sartic.codartic "
            campo = campo & " AND fechanue>'1900-01-01'" 'Que tiene fecha cambio
            campo = campo & cadSelect
             If HayRegParaInforme("sartic,sprees l", campo, True) Then
                MsgBox "Hay precios ESPECIALES pendientes de actualizar ", vbExclamation
                Exit Sub
            End If
        End If
    End If
    
    
    
    'Preguntamos de ciontinuar
    If MsgBox("Desea continuar con el proceso de actualización?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    Set miRsAux = New ADODB.Recordset
    CadenaDesdeOtroForm = Me.Opcion & ". " & cadSelect
    ActualizarPreciosVentaCompra
    Set miRsAux = Nothing
    Screen.MousePointer = vbDefault
    Label3(107).Caption = ""
    MsgBox "Proceso finalizado", vbExclamation
    
        Set LOG = New cLOG
        ' 13 Copia preicos
        LOG.Insertar 13, vUsu, CadenaDesdeOtroForm
        CadenaDesdeOtroForm = ""
        Set LOG = Nothing
    Unload Me
End Sub

Private Sub cmdPedxZona_Click()
Dim Aux As String

    InicializarVbles

    conn.Execute "DELETE from tmpsliped where codusu = " & vUsu.codigo
    conn.Execute "DELETE from tmpstockfec where codusu = " & vUsu.codigo
    
    devuelve = ""
    campo = "{scaped.fecentre}"
    If txtFecha(32).Text <> "" Or txtFecha(33).Text <> "" Then
        'Parametro Desde/Hasta Cliente
        If Not PonerDesdeHasta(campo, "F", 32, 33, devuelve) Then Exit Sub
    End If
    
    devuelve = ""
    campo = "{scaped.codagent}"
    If txtAgente(2).Text <> "" Or txtAgente(3).Text <> "" Then
        'Parametro Desde/Hasta Cliente
        If Not PonerDesdeHasta(campo, "AGT", 2, 3, devuelve) Then Exit Sub
    End If
    
    devuelve = ""
    campo = "{sclien.codzonas}"
    If txtZona(0).Text <> "" Or txtZona(1).Text <> "" Then
        'Parametro Desde/Hasta Cliente
        If Not PonerDesdeHasta(campo, "ZON", 0, 1, devuelve) Then Exit Sub
    End If
    
    
    'Si no marca solo van clientes NO VARIOS
    If Me.chkPedxZona(0).Value = 0 Then AnyadirAFormula cadSelect, " clivario = 0"
    
    
    If Me.chkPedxZona(1).Value = 0 Then AnyadirAFormula cadSelect, " scaped.coddirec is null"
    
    
    'Los que pone recoge cliente NO salen
    AnyadirAFormula cadSelect, " recogecl = 0"
    
    
    devuelve = "scaped.codclien=sclien.codclien AND scaped.numpedcl = sliped.numpedcl"
    If cadSelect <> "" Then devuelve = devuelve & " AND " & cadSelect
       

    If Not HayRegParaInforme("scaped,sclien,sliped", devuelve) Then Exit Sub
    
    'AHORA INSERTAMOS EN LA TMP parar poder precargar las zonas de andespues
    '`tmpsliped` (`codusu`,`numpedcl`,`numlinea`,`codalmac`,`codartic`,`nomartic`,`ampliaci`,`cantidad`,
    '`numbultos`,`importel`,`cantpedprov`,`fecpedprov`,`stockalm`,`stocktot`,`referart`,codclien
    

    campo = "SELECT " & vUsu.codigo & ",scaped.numpedcl,sliped.numlinea,sliped.codalmac,sliped.codartic,sliped.nomartic,sliped.ampliaci,"
    campo = campo & "sliped.cantidad,sliped.numbultos,sliped.importel,0,NULL,0,0,NULL,sclien.codclien,sclien.codzonas FROM scaped,sclien,sliped WHERE "
    campo = campo & devuelve
    campo = "insert into `tmpsliped` (`codusu`,`numpedcl`,`numlinea`,`codalmac`,`codartic`,`nomartic`,`ampliaci`,`cantidad`,`numbultos`,`importel`,`cantpedprov`,`fecpedprov`,`stockalm`,`stocktot`,`referart`,codclien,codzona) " & campo
    If Not ejecutar(campo, False) Then Exit Sub
    
    
    '***********************************************************************************************
    '
    'Ahora veremos aquellos que tienen direccion de envio. Con lo cual buscare la zona en la direnvio
    campo = "SELECT scaped.numpedcl,scaped.codclien,coddiren FROM scaped,sclien,sliped WHERE "
    campo = campo & devuelve
    campo = campo & " and coddiren>0  group by 1,2,3  "
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open campo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        campo = "codclien = " & miRsAux!codClien & " AND coddiren"
        campo = DevuelveDesdeBD(conAri, "codzona", "sdirenvio", campo, CStr(miRsAux!coddiren))
        
        'UPDATEAMOS tmp con la zona
        If campo <> "" Then
            campo = "UPDATE tmpsliped set codzona = " & campo & " WHERE codusu = " & vUsu.codigo
            campo = campo & " AND codclien = " & miRsAux!codClien & " AND numpedcl = " & miRsAux!numpedcl
            ejecutar campo, False
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If Me.chkPedxZona(1).Value = 1 Then
    'Ahora veremos aquellos que tienen departamento, direnvio null. Con lo cual buscare la zona en la coddirec
        campo = "SELECT scaped.numpedcl,scaped.codclien,coddirec FROM scaped,sclien,sliped WHERE "
        campo = campo & devuelve
        campo = campo & " and coddirec>0 and coddiren is null group by 1,2,3  "
        Set miRsAux = New ADODB.Recordset
        miRsAux.Open campo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            campo = "codclien = " & miRsAux!codClien & " AND coddirec"
            campo = DevuelveDesdeBD(conAri, "codzona", "sdirec", campo, CStr(miRsAux!CodDirec))
            
            'UPDATEAMOS tmp con la zona
            If campo <> "" Then
                campo = "UPDATE tmpsliped set codzona = " & campo & " WHERE codusu = " & vUsu.codigo
                campo = campo & " AND codclien = " & miRsAux!codClien & " AND numpedcl = " & miRsAux!numpedcl
                ejecutar campo, False
            End If
            miRsAux.MoveNext
        Wend
        miRsAux.Close
    End If
    
    Set miRsAux = Nothing
    
    'Abrimos el form para que seleccione pro zonas
    CadenaDesdeOtroForm = Abs(chkPedxZona(2).Value)
    frmVarios.Opcion = 5
    frmVarios.Show vbModal
End Sub

Private Sub cmdPropuestaPedido_Click()
Dim b As Boolean
    
    'Tiene que estar configurado en parametros
    If vParamAplic.Rot_ConsumMes1 = 0 Then
        MsgBox "Falta configurar en parmetros valores rotacion(Rot_ConsumMes1) ", vbExclamation
        Exit Sub
    End If
    
    'Almacen
    miSQL = ""
    
    If Me.txtAlma(0).Text = "" Then miSQL = miSQL & vbCrLf & "-Almacen"
    If Me.txtCodProve(17).Text = "" Then
        'Si no pone el proveedor debe poner minimo albaranes
        If txtAnyo(4).Text = "" Then miSQL = miSQL & vbCrLf & "-Minimo albaranes x meses sin proveedor"
    End If
    
    If Me.txtAlma(7).Text <> "" Then
        If Me.txtDescAlma(7).Text = "" Then
            miSQL = miSQL & vbCrLf & "-Error almacen consolidación"
        Else
            If Me.txtAlma(0).Text = Me.txtAlma(7).Text Then miSQL = miSQL = "Mismo almacen"
        End If
    End If
    
    
    'Marzo 2014
    'SEgundo ALMACEN Consolidado
    If Me.txtAlma(8).Text <> "" Then
        If Me.txtDescAlma(7).Text = "" Then
            miSQL = miSQL & vbCrLf & "-Indique el almacen consolidado(1)"
        Else
            If Me.txtDescAlma(8).Text = "" Then miSQL = miSQL & vbCrLf & "-Error almacen consolidación(2)"
            If Me.txtAlma(8).Text = Me.txtAlma(7).Text Then miSQL = miSQL & "Mismo almacen consolidado"
        End If
    End If
    
    
    If miSQL <> "" Then
        miSQL = "Campos obligatorios: " & miSQL
        MsgBox miSQL, vbExclamation
        If Me.txtAlma(0).Text = "" Then
            PonerFoco txtAlma(0)
        Else
            PonerFoco txtAnyo(4)
        End If
        Exit Sub
    End If
    
    
    
    Set miRsAux = New ADODB.Recordset
    Screen.MousePointer = vbHourglass
    
    
    b = GeneraInformepedidoProv
    
    
    
    
    Label3(100).Caption = ""
    Set miRsAux = Nothing
    Screen.MousePointer = vbDefault
    
    If b Then
        
        
        InicializarVbles
        'El nombre de la empresa
        cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
        numParam = numParam + 1
        cadParam = cadParam & "|Per1=""Ult. " & vParamAplic.Rot_ConsumMes1 & """|"
        cadParam = cadParam & "|Per2=""Ult. " & vParamAplic.Rot_ConsumMes2 & """|"
        numParam = numParam + 2
        cadParam = cadParam & "|pAlmacen=""" & Me.txtAlma(0).Text & "-" & Me.txtDescAlma(0).Text
        If Me.txtAlma(7).Text <> "" Then cadParam = cadParam & "   " & Me.txtAlma(7).Text & "-" & Me.txtDescAlma(7).Text
        If Me.txtAlma(8).Text <> "" Then cadParam = cadParam & "   " & Me.txtAlma(8).Text & "-" & Me.txtDescAlma(8).Text
        cadParam = cadParam & """|"
        numParam = numParam + 1
        
        cadParam = cadParam & "|Alma1=""Alm." & Me.txtAlma(0).Text & """|"
        cadParam = cadParam & "|Alma2=""Alm." & Me.txtAlma(7).Text & """|"
        cadParam = cadParam & "|Alma3=""Alm." & Me.txtAlma(8).Text & """|"
        numParam = numParam + 1
        'Valores
        miSQL = ""
        If txtAnyo(5).Text <> "" Then miSQL = miSQL & "    Cli: " & txtAnyo(5).Text & "%"
        If txtCodProve(17).Text <> "" Then miSQL = miSQL & "   Prov: " & txtCodProve(17).Text
    
        If Me.cboProPed(1).ListIndex > 0 Then miSQL = miSQL & "    Situacion: " & Me.cboProPed(1).List(Me.cboProPed(1).ListIndex)
        If Me.cboProPed(0).ListIndex > 0 Then miSQL = miSQL & "     " & Me.cboProPed(0).List(Me.cboProPed(1).ListIndex)
    
            
        cadSelect = ""
        If Me.txtFamia(2).Text <> "" Then cadSelect = cadSelect & " desde " & txtFamia(2).Text
        If Me.txtFamia(3).Text <> "" Then cadSelect = cadSelect & " hasta " & txtFamia(3).Text
        If cadSelect <> "" Then
            miSQL = miSQL & "   Familia: " & cadSelect
            cadSelect = ""
        End If
        
        If Me.txtmarca(0).Text <> "" Then cadSelect = cadSelect & " desde " & txtmarca(0).Text
        If Me.txtmarca(1).Text <> "" Then cadSelect = cadSelect & " hasta " & txtmarca(1).Text
        If cadSelect <> "" Then
            miSQL = miSQL & "   Marca: " & cadSelect
            cadSelect = ""
        End If
    
        cadParam = cadParam & "|Valores=""" & Trim(miSQL) & """|"
        numParam = numParam + 1
        
        
        'Abril 2012
        cadParam = cadParam & "mostrartxtAuxDoc=" & Abs(Me.chkPropPedido(2).Value) & "|"
        numParam = numParam + 1
        
        'Si lleva articulo de portes, ese NO va a las lineas
        cadSelect = "{tmppedprov.codusu} = " & vUsu.codigo
        cadFormula = cadSelect

        
        'If Not PonerParamRPT(28, cadParam, numParam, cadNomRPT, vImprimedirecto, cadPDFrpt) Then Exit Sub
        If txtAlma(7).Text = "" Then
            cadNomRPT = "rproped.rpt"
        Else
            cadNomRPT = "rpropedC.rpt"
        End If
        conSubRPT = False
        cadTitulo = "Propuesta pedido"
        LlamarImprimir False
    End If
    
End Sub

Private Sub cmdRecargasMov_Click()


    InicializarVbles
    
    '===================================================
    '============ PARAMETROS ===========================
    'Añadir el nombre de la Empresa como parametro
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    devuelve = ""
    campo = "{stelefonia.fecha}"
    If txtFecha(4).Text <> "" Or txtFecha(5).Text <> "" Then
        'Parametro Desde/Hasta Cliente
        devuelve = "pDHFecha=""Fecha " & devuelve & ": "
        If Not PonerDesdeHasta(campo, "F", 4, 5, devuelve) Then Exit Sub
    End If

    'Facturados
    devuelve = ""
    codigo = ""
    cadFrom = ""
    If Me.cmbRecargaMov(0).ListIndex > 0 Then
        If Me.cmbRecargaMov(0).ListIndex = 1 Then
            codigo = codigo & " Pendientes de facturar "
        Else
            codigo = codigo & "  Facturadas "
        End If
        campo = "({stelefonia.facturado} = " & cmbRecargaMov(0).ListIndex - 1 & ")"
        cadFrom = "facturado = " & cmbRecargaMov(0).ListIndex - 1
        devuelve = campo
    End If
    
    'Cobrado
    If Me.cmbRecargaMov(1).ListIndex > 0 Then
        If Me.cmbRecargaMov(1).ListIndex = 1 Then
            codigo = codigo & "     Pendientes de cobro "
        Else
            codigo = codigo & "     Cobradas "
        End If
        campo = "({stelefonia.cobrado} = " & cmbRecargaMov(1).ListIndex - 1 & ")"
        
        If devuelve <> "" Then
            devuelve = devuelve & " AND "
            cadFrom = cadFrom & " AND "
        End If
        cadFrom = cadFrom & "cobrado = " & cmbRecargaMov(1).ListIndex - 1
        devuelve = devuelve & campo
    End If
    
    
    'Tipo
    If txtRecargaMov(0).Text <> "" Then
        campo = "({stelefonia.tipo} = '" & txtRecargaMov(0).Text & "')"
        
        If devuelve <> "" Then
            devuelve = devuelve & " AND "
            cadFrom = cadFrom & " AND "
        End If
        cadFrom = cadFrom & "tipo = """ & txtRecargaMov(0).Text & """"
        devuelve = devuelve & campo
    End If
    
    If devuelve <> "" Then
        cadParam = cadParam & "pDHCliente= """ & Trim(codigo) & """|"
        numParam = numParam + 1
        If cadSelect <> "" Then cadSelect = cadSelect & " AND "
        cadSelect = cadSelect & cadFrom
        AnyadirAFormula cadFormula, devuelve
    End If
    
    
    
    
        
    
    
    If Not HayRegParaInforme("stelefonia", cadSelect) Then Exit Sub
    
    LlamarImprimir False

End Sub

Private Sub cmdReparaEfect_Click()
    InicializarVbles
    
    '===================================================
    '============ PARAMETROS ===========================
    'Añadir el nombre de la Empresa como parametro
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    '  Se habia colado esta linea      cmdRepGaranProve_Click
    
    codigo = "schrep"
    devuelve = ""
    'Cadena para seleccion CLIENTE
    '--------------------------------------------
    If txtCliente(0).Text <> "" Or txtCliente(1).Text <> "" Then
        campo = "{" & codigo & ".codclien}"
        'Parametro Desde/Hasta Cliente
        devuelve = "pDHCliente=""Cliente: "
        If Not PonerDesdeHasta(campo, "CLI", 0, 1, devuelve) Then Exit Sub
    End If

    
    'Cadena para seleccion DEPARTAMENTO
    '--------------------------------------------
    If txtDpto(0).Text <> "" Or txtDpto(1).Text <> "" Then
        campo = "{" & codigo & ".coddirec}"
        'Parametro Desde/Hasta Cliente
        devuelve = "pDHDpto=""Dpto: "
        If Not PonerDesdeHasta(campo, "DPT", 0, 1, devuelve) Then Exit Sub
    End If
    
    
    'Este trozo lo hace siempre
    If Me.optReparaciones(0).Value Then
        devuelve = "entrada"
        campo = "entre"
    Else
        devuelve = "reparación"
        campo = "repar"
        'AHora Marzo 2010
        campo = "haalb"  'fechaalb
    End If
    campo = "{" & codigo & ".fec" & campo & "}"
    cadParam = cadParam & "pOrden=" & campo & "|"
    numParam = numParam + 1
    
    If txtFecha(0).Text <> "" Or txtFecha(1).Text <> "" Then
        'Parametro Desde/Hasta Cliente
        devuelve = "pDHFecha=""Fecha " & devuelve & ": "
        If Not PonerDesdeHasta(campo, "F", 0, 1, devuelve) Then Exit Sub
    End If
    
    cadFrom = "schrep"
    If Not HayRegParaInforme(cadFrom, cadSelect) Then Exit Sub
    Screen.MousePointer = vbHourglass
    'Prepararo los datos
    codigo = "DELETE from tmpnlotes where codusu = " & vUsu.codigo
    conn.Execute codigo
    CargaImporteRealReparaciones
    Label3(158).Caption = ""
    
    'MOSTRAMOS EL INFORME
    If cadFormula <> "" Then cadFormula = cadFormula & " AND "
    cadFormula = cadFormula & "(isnull({tmpnlotes.codusu}) or {tmpnlotes.codusu}=" & vUsu.codigo & ")"
    
    conSubRPT = False
    LlamarImprimir False
    Screen.MousePointer = vbDefault
End Sub
Private Sub LlamarImprimir(PongoNombrePDF As Boolean, Optional NumeroDeCopias As Integer)
    If NumeroDeCopias = 0 Then NumeroDeCopias = 1
    With frmImprimir
        .FormulaSeleccion = cadFormula
        .OtrosParametros = cadParam
        .NumeroParametros = numParam
        .NumeroCopias = NumeroDeCopias
        .SoloImprimir = False
        .EnvioEMail = False
        .Opcion = 2000 + Opcion   '2000 mas la opcion de entrada
        .NombrePDF = ""
        If PongoNombrePDF Then .NombrePDF = cadPDFrpt
        .NombreRPT = cadNomRPT
        .ConSubInforme = conSubRPT
        .Show vbModal
    End With
End Sub

Private Sub cmdRepGaranProve_Click()
    Label3(94).Caption = "Incio proceso"
    Screen.MousePointer = vbHourglass
    HacerListadogarantiaProveedor
    Screen.MousePointer = vbDefault
    Label3(94).Caption = ""
    
    
     If NumRegElim = 0 Then
        MsgBox "Ningun dato a mostrar", vbExclamation
        Exit Sub
    End If
    
    
    'Llegados aqui imprimiremos los registros
    InicializarVbles
    
    '===================================================
    '============ PARAMETROS ===========================
    'Añadir el nombre de la Empresa como parametro
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = 1
    
    If txtCodProve(15).Text <> "" Or txtCodProve(16).Text <> "" Then
        devuelve = "pDHCliente= """
        campo = "{sartic.codprove}"
        If Not PonerDesdeHasta(campo, "PRO", 15, 16, devuelve) Then Exit Sub
        
    End If
    
    
    campo = ""
    If txtFecha(36).Text <> "" Then campo = "     Desde " & txtFecha(36).Text
    If txtFecha(37).Text <> "" Then campo = campo & "      Hasta " & txtFecha(37).Text
    If campo <> "" Then
        numParam = numParam + 1
        campo = "pDHFecha= """ & Trim(campo) & """|"
        cadParam = cadParam & campo
    End If
    cadFormula = "{tmpnlotes.codusu}=" & vUsu.codigo

    cadNomRPT = "rRepEstaGaranProv.rpt"
    conSubRPT = False
    LlamarImprimir False
    
    
End Sub



Private Function HacerListadogarantiaProveedor() As Boolean
Dim RT As ADODB.Recordset
Dim RAlb As ADODB.Recordset
Dim Donde As Byte
Dim C As Integer
Dim SQ As String


On Error GoTo EHacerListadogarantiaProveedor
    HacerListadogarantiaProveedor = False
    InicializarVbles
    cadTitulo = ""
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1

    'Comporobar si hay registros
    SQ = ""
    If txtFecha(36).Text <> "" Or txtFecha(37).Text <> "" Then
        campo = "{schrep.fechaalb}"
        If Not PonerDesdeHasta(campo, "F", 36, 37, "") Then Exit Function
        If txtFecha(36).Text <> "" Then cadTitulo = cadTitulo & " desde " & txtFecha(36).Text
        If txtFecha(37).Text <> "" Then cadTitulo = cadTitulo & " hasta " & txtFecha(37).Text
        If cadTitulo <> "" Then cadTitulo = "Fechas: " & cadTitulo
        SQ = cadSelect
    End If
    
    If txtCodProve(15).Text <> "" Or txtCodProve(16).Text <> "" Then
        
        campo = "{sartic.codprove}"
        If Not PonerDesdeHasta(campo, "PRO", 15, 16, devuelve) Then Exit Function
        cadTitulo = cadTitulo & devuelve
    End If
    
    If cadSelect <> "" Then
        cadSelect = Replace(cadSelect, "{", "(")
        cadSelect = Replace(cadSelect, "}", ")")
    End If
    
    If SQ <> "" Then
        SQ = Replace(SQ, "{", "(")
        SQ = Replace(SQ, "}", ")")
    End If
    
    
    
    Label3(94).Caption = "Preparando datos"
    Label3(94).Refresh
    miSQL = "DELETE from tmpnlotes where codusu = " & vUsu.codigo
    conn.Execute miSQL
    miSQL = "DELETE from tmpinformes where codusu = " & vUsu.codigo
    conn.Execute miSQL
    
    
    cadFormula = "insert into tmpinformes (codusu,codigo1,importe1,importe2,nombre1,nombre2,campo1,importeb1) VALUES (" & vUsu.codigo & ","
    cadFrom = "insert into tmpnlotes (codusu,codprove,numalbar,fechaalb,numlinea,nomartic) values (" & vUsu.codigo & ","
   
    Label3(94).Caption = "Leyendo reparaciones"
    Label3(94).Refresh
    
    miSQL = "select schrep.numalbar, schrep.fechaalb,sartic.codprove,1,sserie.codartic,schrep.nomartic,sserie.numserie"
    miSQL = miSQL & " from schrep ,sserie,sartic where schrep.numserie=sserie.numserie"
    miSQL = miSQL & " and schrep.codartic=sserie.codartic and  sartic.codartic=sserie.codartic "
   ' miSQL = miSQL & " and ((schrep.fechaalb >= '2001-11-01') "
   ' miSQL = miSQL & " and (schrep.fechaalb <= '2011-01-01')) and sartic.codprove>1"
    If cadSelect <> "" Then miSQL = miSQL & " AND " & cadSelect
    miSQL = miSQL & " and not fechavta  is null"
    miSQL = miSQL & " and schrep.fechaalb<DATE_ADD(fechavta, interval garantia day) group by 1,2"
   
    
    
    
    
    
    
    
    
    'HAremos la busqueda
    '-------------------------------------------
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    miSQL = ""
    cadTitulo = ""
    If miRsAux.EOF Then
        'MsgBox "Ningun dato generado", vbExclamation
        NumRegElim = 0
        miRsAux.Close
        GoTo EHacerListadogarantiaProveedor
    End If
    
    
    
    
    Label3(94).Caption = "Leyendo facturas"
    Label3(94).Refresh
    
    miSQL = "select slifac.numalbar,scafac1.fechaalb,sum(importel),sum(cantidad*preciove) from scafac1,slifac,sartic  where"
    miSQL = miSQL & " scafac1.codtipom=slifac.codtipom and scafac1.numfactu=slifac.numfactu and"
    miSQL = miSQL & " scafac1.FecFactu = slifac.FecFactu And scafac1.NumAlbar = slifac.NumAlbar And scafac1.codtipoa = slifac.codtipoa"
    miSQL = miSQL & " and slifac.codartic=sartic.codartic"
    If SQ <> "" Then
        codigo = Replace(SQ, "schrep", "scafac1")
        miSQL = miSQL & " AND " & codigo
    End If
    miSQL = miSQL & " group  by 1,2 order by 1,2"
    Set RT = New ADODB.Recordset
    RT.Open miSQL, conn, adOpenKeyset, adLockPessimistic, adCmdText
    
    
    
    
    
    'Los albaranaes
    Label3(94).Caption = "Leyendo albaranes"
    Label3(94).Refresh
    
    miSQL = "select slialb.numalbar,scaalb.fechaalb,sum(importel),sum(cantidad*preciove) from scaalb,slialb,sartic  where"
    miSQL = miSQL & " slialb.codtipom=scaalb.codtipom AND slialb.numalbar=scaalb.numalbar AND scaalb.codtipom='ALR' "
    miSQL = miSQL & " and slialb.codartic=sartic.codartic  "
    If SQ <> "" Then
        codigo = Replace(SQ, "schrep", "scaalb")
        miSQL = miSQL & " AND " & codigo
    End If
    miSQL = miSQL & " group  by 1,2 order by 1,2"
    Set RAlb = New ADODB.Recordset
    RAlb.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    
    NumRegElim = 0
    While Not miRsAux.EOF
        NumRegElim = NumRegElim + 1
        Label3(94).Caption = "Recorriendo datos " & NumRegElim
        Label3(94).Refresh
        

        'primero veo si esta en
        'Debug.Print miRsAux!FechaAlb & "  " & miRsAux!NumAlbar
        Donde = 0
        If SituarRsGarantiaProveedorAlbaran(RAlb, miRsAux!NUmAlbar, miRsAux!FechaAlb) Then
            ImpTot = DBLet(RAlb.Fields(2), "N")
            ImpTeo = DBLet(RAlb.Fields(3), "N")
            Donde = 1
        Else
            If SituarRsGarantiaProveedor(RT, miRsAux!NUmAlbar, miRsAux!FechaAlb) Then
                ImpTot = DBLet(RT.Fields(2), "N")
                ImpTeo = DBLet(RT.Fields(3), "N")
                Donde = 2
            Else
                ImpTot = 0
                ImpTeo = 0
            End If
        End If
        
        'INSERTAMOS
        'en tmpinformes
        codigo = "'" & DevNombreSQL(miRsAux!NomArtic) & "','" & DBLet(miRsAux!numSerie, "T") & "'," & Donde & "," & miRsAux!Codprove & ")"
        codigo = NumRegElim & "," & TransformaComasPuntos(CStr(ImpTot)) & "," & TransformaComasPuntos(CStr(ImpTeo)) & "," & codigo
        codigo = cadFormula & codigo
        conn.Execute codigo
        

        codigo = miRsAux!NUmAlbar & ",'" & Format(miRsAux!FechaAlb, FormatoFecha) & "',1,'')"
        codigo = NumRegElim & "," & codigo
        codigo = cadFrom & codigo
        conn.Execute codigo
                
        miRsAux.MoveNext
    Wend
    miRsAux.Close
        
    
    
    'Ultima fase, actualizar nombre proveedor
    Label3(94).Caption = "Leyendo proveedores"
    Label3(94).Refresh
    codigo = "Select importeb1 from tmpinformes where codusu = " & vUsu.codigo & " GROUP  BY  importeb1"
    miRsAux.Open codigo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        codigo = CStr(CLng(miRsAux!importeb1))
        Label3(94).Caption = "Proveedor: " & codigo
        Label3(94).Refresh
        
        miSQL = DevuelveDesdeBD(conAri, "nomprove", "sprove", "codprove", codigo, "N")
        codigo = "UPDATE tmpinformes set nombre3=" & DBSet(miSQL, "T") & " WHERE codusu = " & vUsu.codigo
        codigo = codigo & " AND importeb1 = " & DBSet(miRsAux!importeb1, "N")
        conn.Execute codigo
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    
EHacerListadogarantiaProveedor:
    If Err.Number <> 0 Then
        MuestraError Err.Number
        NumRegElim = 0
    End If
    Set miRsAux = Nothing
    Set RT = Nothing
    Set RAlb = Nothing
End Function



Private Function SituarRsGarantiaProveedor(ByRef R As ADODB.Recordset, NA As Long, FA As Date) As Boolean
Dim Salir As Boolean
    SituarRsGarantiaProveedor = False
    If R.EOF Then
        Exit Function
    Else
        Salir = False
        While Not Salir
            'Mismo numero de albaran
            If R!NUmAlbar = NA Then
                If R!FechaAlb = FA Then
                    'OK es este. NO lo muevas
                    SituarRsGarantiaProveedor = True
                    Salir = True
                End If
            Else
                If R!NUmAlbar > NA Then
                    'El numero de albaran es mayor que el que pedimos. Nos salimos sin mover
                    Salir = True
                Else
                    'Es menor. Que se mueva
                End If
            End If
            
            If Not Salir Then
                R.MoveNext
                If R.EOF Then Salir = True
            End If
        Wend
    End If
End Function



Private Function SituarRsGarantiaProveedorAlbaran(ByRef R As ADODB.Recordset, NA As Long, FA As Date) As Boolean
Dim Salir As Boolean
    SituarRsGarantiaProveedorAlbaran = False
    If R.EOF Then
        Exit Function
    Else
        Salir = False
        While Not Salir
            'Mismo numero de albaran
            If R!NUmAlbar = NA Then
                If R!FechaAlb = FA Then
                    'OK es este. NO lo muevas
                    SituarRsGarantiaProveedorAlbaran = True
                    Salir = True
                End If
            Else
                If R!NUmAlbar > NA Then
                    'El numero de albaran es mayor que el que pedimos. Nos salimos sin mover
                    Salir = True
                Else
                    'Es menor. Que se mueva
                End If
            End If
            
            If Not Salir Then
                R.MoveNext
                If R.EOF Then Salir = True
               
            End If
        Wend
    End If
End Function

Private Sub cmdResVtaAgente_Click()
Dim b As Boolean
    
    
    'Noviembre 2013
    'No puede marcar a la vez, presupuesto y rectificativas

    If Me.chkResVtaAgen(1).Value = 1 Then
        If Me.chkResVtaAgen(3).Value = 1 Then
            MsgBox "Si selecciona presupuestos no saldran ningun tipo mas", vbExclamation
            Exit Sub
        End If
    End If
    
    
    'Preparaamos
    conn.Execute "DELETE FROM tmpinformes WHERE codusu = " & vUsu.codigo
    

    NumRegElim = 0
    Screen.MousePointer = vbHourglass
    Set miRsAux = New ADODB.Recordset
    b = CargaDatosResumenVtaAgente
    Label3(122).Caption = ""
    Set miRsAux = Nothing
    Screen.MousePointer = vbDefault
    
    
    If b Then
    
            'Vamos a imprimir
    
             InicializarVbles
             
             'Pasar nombre de la Empresa como parametro
             cadParam = "|pNomEmpre=""" & vEmpresa.nomempre & """|"
             numParam = numParam + 1
             
              '====================================================
             '================= FORMULA ==========================
             'Cadena para seleccion D/H CLIENTE
             
             'Cadena para seleccion Desde y Hasta FECHA
             '--------------------------------------------
             miSQL = ""
             If txtFecha(39).Text <> "" Or txtFecha(40).Text <> "" Then
                 campo = "{scafac.fecfactu}"
                 devuelve = "Fechas: "
                 If Not PonerDesdeHasta(campo, "F", 39, 40, devuelve) Then Exit Sub
                 miSQL = devuelve
             End If
             
             'Seguna haya seleccionado los checks
                 
            If chkResVtaAgen(2).Value = 1 Then miSQL = miSQL & "    -Portes"
            If chkResVtaAgen(3).Value = 1 Then miSQL = miSQL & "    -Rectificativas"
            If chkResVtaAgen(0).Value = 1 Then miSQL = miSQL & "    -Albaranes"
            If chkResVtaAgen(1).Value = 1 Then miSQL = miSQL & "    -Presupuestos"
             
             
             
             
             
            cadParam = cadParam & "pDHFecha=""" & miSQL & """|"
            numParam = numParam + 1
             
            'Cadena para seleccion Desde y Hasta ARTICULO
            '--------------------------------------------
            miSQL = ""
            If txtmarca(4).Text <> "" Or txtmarca(5).Text <> "" Then
                 campo = "{slifac.Noooodartic}"
                 devuelve = "Mar:"
                 If Not PonerDesdeHasta(campo, "MAR", 4, 5, devuelve) Then Exit Sub
                 miSQL = devuelve & "  "
            End If
             
             
            

             
             
             If txtAgente(4).Text <> "" Or txtAgente(5).Text <> "" Then
                devuelve = "Agente:"
                If Me.chkResVtaAgen(4).Value Then devuelve = "Visitador:"
                campo = "{scafac.codagent}"
                If Not PonerDesdeHasta(campo, "AGT", 4, 5, devuelve) Then Exit Sub
                miSQL = miSQL & devuelve
            End If
            
             cadParam = cadParam & "pDHAgMar=""" & Trim(miSQL) & """|"
             numParam = numParam + 1
            If optVtaAgen(0).Value Then
                cadNomRPT = "rvtaagenmarcAg"   '"rvtaagenmarcAg.rpt"
            Else
                cadNomRPT = "rvtaagenmarcMa"     '"rvtaagenmarcMa.rpt"
            End If
            If Me.chkResVtaAgen(4).Value = 1 Then cadNomRPT = cadNomRPT & "visi"
            cadNomRPT = cadNomRPT & ".rpt"
            cadFormula = "{tmpinformes.codusu}=" & vUsu.codigo
            LlamarImprimir False
    End If

End Sub

Private Sub cmdRiesgo_Click()
    Opcion = 0 'Lleva cancelar parar parar el proceso
    Label3(95).Caption = "Preparando datos"
    Label3(95).Refresh
    miSQL = "DELETE FROM tmpinformes WHERE codusu = " & vUsu.codigo
    conn.Execute miSQL
    'Meto en tmp toso lso que voy a tratar
                                     '   codclien        limi    situ  tipoiva
    miSQL = "INSERT INTO tmpinformes (codusu,codigo1,nombre1,importe1,campo1,campo2) "
    miSQL = miSQL & "SELECT " & vUsu.codigo & ",codclien,nomclien,limcredi,codsitua,tipoiva "

    miSQL = miSQL & " from sclien where credipriv<9"
    
    
    
    
    conn.Execute miSQL
    
    'AHora vere cuantos hay, si es que hay
    miSQL = DevuelveDesdeBD(conAri, "count(*)", "tmpinformes", "codusu", CStr(vUsu.codigo))
    If miSQL = "" Then miSQL = "0"
    If Val(miSQL) = 0 Then
        MsgBox "Ningun cliente en operaciones aseguradas", vbExclamation
    Else
        pb2.Value = 0
        pb2.Max = CInt(miSQL)
        pb2.visible = True
        Set miRsAux = New ADODB.Recordset
        
        RecorrerRiesgo
        
        Set miRsAux = Nothing
        pb2.visible = False
        
    End If
    Label3(95).Caption = ""
    Opcion = 31
End Sub

Private Sub cmdSituAlbaran_Click()
Dim i As Integer
    InicializarVbles
    cadTitulo = ""
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1

    'Comporobar si hay registros
    If txtFecha(29).Text <> "" Or txtFecha(30).Text <> "" Then
        campo = "{scaalb.fechaalb}"
        If Not PonerDesdeHasta(campo, "F", 29, 30, "") Then Exit Sub
        If txtFecha(29).Text <> "" Then cadTitulo = cadTitulo & "desde " & txtFecha(29).Text
        If txtFecha(30).Text <> "" Then cadTitulo = cadTitulo & " hasta " & txtFecha(30).Text
        If cadTitulo <> "" Then cadTitulo = "Fechas: " & cadTitulo
    End If
    
    If txtCliente(5).Text <> "" Or txtCliente(6).Text <> "" Then
        campo = "{scaalb.codclien}"
        'Parametro Desde/Hasta Cliente
        devuelve = "pDHClien=""Cliente: "
        If Not PonerDesdeHasta(campo, "CLI", 5, 6, devuelve) Then Exit Sub
    End If
    
    devuelve = ""
    miSQL = ""
    IndiceImg = 0
    For i = 0 To List1.ListCount - 1
        If List1.Selected(i) Then
            IndiceImg = IndiceImg + 1
            NumRegElim = InStrRev(List1.List(i), "(")
            If NumRegElim = 0 Then
                MsgBox "No se ha encontrado (", vbExclamation
                Exit Sub
            End If
            campo = Mid(List1.List(i), NumRegElim + 1, 3)
            miSQL = miSQL & " - " & campo
            devuelve = devuelve & ", '" & campo & "'"
            
        End If
    Next i
    If devuelve = "" Then
        MsgBox "Seleccione algun tipo de albarán", vbExclamation
        Exit Sub
    End If
    
    If IndiceImg <> List1.ListCount Then
        If cadTitulo <> "" Then cadTitulo = cadTitulo & "        "
        miSQL = Mid(miSQL, 3)
        cadTitulo = cadTitulo & "Tipo albaran: " & miSQL
        
    End If
    miSQL = cadTitulo
    cadParam = cadParam & "pDHFecha=""" & miSQL & """|"
    
    devuelve = Mid(devuelve, 2)
    If cadSelect <> "" Then cadSelect = cadSelect & " AND "
    cadSelect = cadSelect & " (codtipom IN (" & devuelve & "))"
    If cadFormula <> "" Then cadFormula = cadFormula & " AND "
    cadFormula = cadFormula & "( {scaalb.codtipom} IN [" & devuelve & "])"
    
    'Pongo en campo el select
    
    If Not HayRegParaInforme("scaalb", cadSelect) Then Exit Sub
    
    
    
    miSQL = DevuelveDesdeBDNew(conAri, "scryst", "documrpt", "codcryst", "86", "N")
    If miSQL = "" Then miSQL = "rFacSituacionAlb.rpt"
    cadNomRPT = miSQL
    LlamarImprimir False

End Sub

Private Sub cmdTraza_Click()
    Screen.MousePointer = vbHourglass
    HacerInformeTrazabilidad
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdVentaxProv_Click()
Dim cad As String
Dim miAUX As String
    
    'Si es por agente no puede ser por familia comarati
    If Me.chkVtaxProv(0).Value = 1 And chkVtaxProv(2).Value = 1 Then
        MsgBox "Si es por agente, no puede marcar 'Familia comparativo'", vbExclamation
        chkVtaxProv(2).Value = 0
        Exit Sub
        
    End If
    
    If Me.chkVtaxProv(0).Value = 0 And chkVtaxProv(2).Value = 0 Then
        If Me.txtImporte(1).Text <> "" Then
            MsgBox "Importe minimo solo para comparativos o por agente", vbExclamation
            Exit Sub
        End If
    End If
    
    
    
    'Si es por agente, comparativo DEBE PONER desde hasta
    If Me.chkVtaxProv(0).Value = 1 Or chkVtaxProv(2).Value = 1 Then
        miAUX = ""
        If txtFecha(9).Text = "" Or txtFecha(10).Text = "" Then
            miAUX = "Debe indicar periodo"
        Else
            If Year(CDate(txtFecha(9).Text)) <> Year(CDate(txtFecha(10).Text)) Then miAUX = "Debe pertenecer al mismo año"
        End If
        If miAUX <> "" Then
            MsgBox miAUX, vbExclamation
            PonerFoco txtFecha(9)
            Exit Sub
        End If
    End If
    
    InicializarVbles
    
    'Pasar nombre de la Empresa como parametro
    cadParam = "|pNomEmpre=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    
    'Marzo 2011. Voy a agrupar D/H
    
    
    '====================================================
    '================= FORMULA ==========================
    'Cadena para seleccion D/H CLIENTE  y agente
    '--------------------------------------------
    miAUX = ""
    If txtCliente(3).Text <> "" Or txtCliente(4).Text <> "" Then
        campo = "{scafac.codclien}"
        'Parametro Desde/Hasta Cliente
        cad = " Cli:"
        If Not PonerDesdeHasta(campo, "CLI", 3, 4, cad) Then Exit Sub
        miAUX = cad
    End If
    If txtAgente(8).Text <> "" Or txtAgente(9).Text <> "" Then
        campo = "{scafac.codagent}"
        'Parametro Desde/Hasta Cliente
        cad = "      Ag:"
        If Not PonerDesdeHasta(campo, "AGT", 8, 9, cad) Then Exit Sub
        miAUX = Trim(miAUX & cad)
    End If
    cad = "pDHCliente=""" & miAUX & """|"
    cadParam = cadParam & cad
    numParam = numParam + 1
    
    'Cadena para seleccion Desde y Hasta FECHA
    '--------------------------------------------
    If txtFecha(9).Text <> "" Or txtFecha(10).Text <> "" Then
        campo = "{scafac.fecfactu}"
        cad = "pDHFecha=""Fecha Fact.: "
        If Not PonerDesdeHasta(campo, "F", 9, 10, cad) Then Exit Sub
    End If
    
    
    
    'Cadena para seleccion Desde y Hasta ARTICULO
    '--------------------------------------------
    miAUX = ""
    If txtArticulo(1).Text <> "" Or txtArticulo(2).Text <> "" Then
        campo = "{slifac.codartic}"
        cad = "Artículo: "
        If Not PonerDesdeHasta(campo, "ART", 1, 2, cad) Then Exit Sub
        miAUX = cad
    End If
    
    If Me.txtImporte(1).Text <> "" Then miAUX = miAUX & "        Importe  min: " & Me.txtImporte(1).Text
    cad = "pDHDpto=""" & Trim(miAUX) & """|"
    cadParam = cadParam & cad
    numParam = numParam + 1
    
    'Cadena para seleccion Desde y Hasta PROVEEDOR
    '---------------------------------------------
    'Agrupamos por proveedor
    '  pero cogiendolo de sartic.codprove
    miAUX = ""
    If txtCodProve(0).Text <> "" Or txtCodProve(1).Text <> "" Then
        campo = "{sartic.codprove}"
        cad = "Pro:"
        If Not PonerDesdeHasta(campo, "PRO", 0, 1, cad) Then Exit Sub
        miAUX = cad
    End If
    If txtFamia(14).Text <> "" Or txtFamia(15).Text <> "" Then
        campo = "{sartic.codfamia}"
        cad = "     Fam:"
        If Not PonerDesdeHasta(campo, "FAM", 14, 15, cad) Then Exit Sub
         miAUX = Trim(miAUX & cad)
    End If
    
    If txtAlma(1).Text <> "" Or txtAlma(2).Text <> "" Then
        campo = "{slifac.codalmac}"
        cad = "     Alm:"
        If Not PonerDesdeHasta(campo, "ALM", 1, 2, cad) Then Exit Sub
        miAUX = Trim(miAUX & cad)
    End If
    
    cad = "pDHPro=""" & miAUX & """|"
    cadParam = cadParam & cad
    numParam = numParam + 1
    'ANTES MARZO 2011  ahora es lo de arriba
    'If txtCodProve(0).Text <> "" Or txtCodProve(1).Text <> "" Then
    '    campo = "{slifac.codprovex}"
    '    cad = "pDHPro=""Proveedor: "
    '    If Not PonerDesdeHasta(campo, "PRO", 0, 1, cad) Then Exit Sub
    'End If

     
    'Si detalla o no
    cad = "detalla= " & chkVtaxProv(1).Value & "|"
    cadParam = cadParam & cad
    numParam = numParam + 1
     
    
    'Pongo en campo el select
    Screen.MousePointer = vbHourglass
    codigo = " sclien.codclien=scafac.codclien and slifac.codartic=sartic.codartic and scafac.codtipom=slifac.codtipom "
    codigo = " scafac.fecfactu = slifac.fecfactu AND scafac.numfactu=slifac.numfactu AND " & codigo
    cad = "scafac,slifac,sclien,sartic"
    If cadSelect <> "" Then codigo = codigo & " AND " & cadSelect
    campo = codigo
    If HayRegParaInforme(cad, codigo) Then
        Label3(142).Caption = "Leyendo BD"  'indicador
        Label3(142).Refresh
        If Me.chkVtaxProv(0).Value = 1 Then
            'Por AGENTE
            AgrupaVtasxProveedorxAgente
            If Me.chkVtaxProv(3).Value = 1 Then
                cadNomRPT = "rvtaxcodproveAgeArt.rpt"
            Else
                cadNomRPT = "rvtaxcodproveAge.rpt"
            End If
            cadFormula = "{tmpinformes.codusu} = " & vUsu.codigo
        Else
            If Me.chkVtaxProv(2).Value = 1 Then
                AgrupaVtasxProveedorxFamilia
                cadNomRPT = "rvtaxcodproveFam.rpt"
                cadFormula = "{tmpinformes.codusu} = " & vUsu.codigo
                
            Else
        
                Label3(142).Caption = "Leyendo familia"  'indicador
                Label3(142).Refresh
                Espera 0.2
                'por cliente
                cadNomRPT = "rvtaxcodprove.rpt"
            End If
        End If
        
        
        'Para los listados cmparativos y por agente
        codigo = ""
        If InStr(1, cadFormula, "tmpinformes") > 0 Then
            'Si ha puesto importe mimino puede que no existan datos al borrar los que no cumplen el minimo
            If Me.txtImporte(1).Text <> "" Then
                Label3(142).Caption = "Comprobar existe datos"  'indicador
                Label3(142).Refresh
        
                codigo = DevuelveDesdeBD(conAri, "count(*)", "tmpinformes", "codusu", CStr(vUsu.codigo))
                If codigo = "" Then codigo = "0"
                If Val(codigo) = 0 Then
                    MsgBox "No existen datos con estos parametros", vbExclamation
                Else
                    codigo = ""
                End If
            End If
        End If
        
        Label3(142).Caption = ""  'indicador
        Label3(142).Refresh
        
        If codigo = "" Then LlamarImprimir False
        
    End If
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub cmdVolver_Click()
    FrameErrorRestore.visible = False
End Sub

Private Sub Command1_Click()
    If txtCodProve(12).Text = "" Or Me.txtDescProve(12).Text = "" Then
        MsgBox "Seleccione el proveedor", vbExclamation
        Exit Sub
    End If
    
    
    
    
     'Compruebo si esta bloqueado el proveedor
    miSQL = DevuelveDesdeBDNew(conAri, "sprove", "codsitua", "codprove", txtCodProve(12).Text, "N")
    
    If Val(miSQL) > 0 Then
            devuelve = "tipositu"
            miSQL = DevuelveDesdeBDNew(conAri, "ssitua", "nomsitua", "codsitua", miSQL, "N", devuelve)
            
            
            If devuelve = "1" Then 'Cliente Bloqueado por Situación Especial.
                MsgBox UCase("Proveedor bloqueado por: ") & miSQL & "-" & devuelve, vbInformation, "Situación Especial del proveedor."
            Else
                MsgBox miSQL, vbInformation, "Situación Especial del proveedor."
            End If
            Exit Sub
    End If
    
    
    
    CadenaDesdeOtroForm = txtCodProve(12).Text
    Unload Me
End Sub

Private Sub cmdImpAlbRut_Click()
Dim N As Integer

    'Fecha
    If txtFecha(35).Text = "" Then
        'INDEIQUE UNA FECHA
        MsgBox "Indique la fecha", vbExclamation
        PonerFoco txtFecha(35)
        Exit Sub
    End If
    
    InicializarVbles
    '49: Albaran de transporte
    If optAlbTrans(0).Value Then
        If Not PonerParamRPT2(49, cadParam, numParam, cadNomRPT, vImprimedirecto, cadPDFrpt, vMultiInforme) Then Exit Sub
    Else
        cadNomRPT = "rAlbTraPend.rpt"  'nombre del listadito que salen los albaranes que hay
    End If
    
    'El nombre de la empresa
    cadParam = cadParam & "|pNomEmpre=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    'Codusu para el subreport
    cadParam = cadParam & "|vCodUsu=" & vUsu.codigo & "|"
    numParam = numParam + 1
   
    If txtFecha(35).Text <> "" Or txtFecha(35).Text <> "" Then
        campo = "{scaalb.fechaalb}"
        If Not PonerDesdeHasta(campo, "F", 35, 35, devuelve) Then Exit Sub
    End If

    'Cadena para seleccion CLIENTE
    '--------------------------------------------
    If txtCliente(7).Text <> "" Or txtCliente(8).Text <> "" Then
        campo = "{scaalb.codclien}"
        'Parametro Desde/Hasta Cliente
        devuelve = "pDHCliente=""Cliente: "
        If Not PonerDesdeHasta(campo, "CLI", 7, 8, devuelve) Then Exit Sub
    End If

    'Todos o solo reimpresos
    'Febrero 2012
    ' Belarte dice que ya NO hay que tener en cuenta esa marca.
    ' solo los de la fecha y que no esten impresos
    'AnyadirAFormula cadSelect, " tipAlbaran = 1"   'con transporte
    'AnyadirAFormula cadFormula, " {scaalb.tipAlbaran} = 1"   'con transporte
    
    
    
    AnyadirAFormula cadSelect, " scaalb.codtipom = 'ALV'"
    AnyadirAFormula cadFormula, " {scaalb.codtipom} = 'ALV'"   'con transporte
    
    If Me.chkImpAlbRut(0).Value = 0 Then
        AnyadirAFormula cadSelect, " albImpreso = 0"
        AnyadirAFormula cadFormula, " {scaalb.albImpreso} = 0"
    End If
        


    If Not HayRegParaInforme("scaalb", cadSelect) Then Exit Sub
    

    
    
    If Me.optAlbTrans(0).Value Then
        'Vere si imprime Todas las zonas
        CadenaDesdeOtroForm = cadSelect
        frmVarios.Opcion = 6
        frmVarios.Show vbModal
        If CadenaDesdeOtroForm = "NO" Then Exit Sub
        If CadenaDesdeOtroForm <> "" Then
            'Lleva las zonas que quiere imprimir
            'Las añadire al cadselect
            
            AnyadirAFormula cadSelect, " (scaalb.codzonas) IN (" & CadenaDesdeOtroForm & ")"
            AnyadirAFormula cadFormula, " {scaalb.codzonas} IN [" & CadenaDesdeOtroForm & "]"
        End If
    End If
    Screen.MousePointer = vbHourglass
    N = vParamAplic.NumCop_AlbaranRuta
    
    If CargarDatosImprimeAlbaranConTransporte Then
        If optAlbTrans(0).Value Then GenerearFicheroTxtAlbaranRuta
        LlamarImprimir False, N
    End If
    If optAlbTrans(0).Value Then
    
        
    
    
        'Si ha pulsado imprimir then
        If HaPulsadoElBotonDeImprimir Then
            'UPDATEAMOS scaalb para que no reimpimrpima los albaranes
            
            miSQL = "UPDATE scaalb SET albImpreso = 1 WHERE " & cadSelect
            ejecutar miSQL, False
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub



Private Sub Command2_Click()

End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        Select Case Opcion
        Case 0
            PonerFoco txtCliente(0)
        Case 1
            PonerFoco txtTrab(0)
        Case 6 To 10
            'En ambos listados lo primero es una fecha
            If Opcion = 6 Then
                numParam = 9
            ElseIf Opcion = 7 Then
                numParam = 11
            ElseIf Opcion = 8 Then
                numParam = 13  'liquidacion factura sprov
                txtFecha(17).Text = Format(Now, "dd/mm/yyyy")
            Else
                numParam = 8 + Opcion 'impresion facturas  index:17 y 18
            End If
            PonerFoco txtFecha(CInt(numParam))
            
        Case 13
            cadParam = ""
            'Poner el nombre del trabajador que esta conectado
            Me.txtTrab(2).Text = PonerTrabajadorConectado(cadParam)
            Me.txtDescTra(2).Text = cadParam
        
        Case 21
            If vParamAplic.TipoDtos Then
                lw1.ColumnHeaders(4).Text = "Departamento"
            Else
                lw1.ColumnHeaders(4).Text = "Direccion"
            End If
            CargarOtrasOfertas
        Case 23
            'Tipo albarens
            'Valores por defecto DESMARCADOS
            
        Case 24
            texto(5).Text = RecuperaValor(CadenaDesdeOtroForm, 1)
            miSQL = RecuperaValor(CadenaDesdeOtroForm, 2)
            Me.chkFrecu.Value = Abs(miSQL)
            CadenaDesdeOtroForm = "" 'Para que no devulev nada
        Case 25
            PonerFoco txtFecha(31)
        Case 26
            
            PonerFoco txtFecha(32)
        Case 32
            PonerFoco txtAlma(0)
        Case 34
            PonerFoco txtFecha(38)
        Case 36
            PonerFoco txtFecha(39)
        Case 37
            PonerFoco txtAgente(6)
        Case 38
            PonerFoco txtFecha(41)
        Case 41
            PonerFoco txtFecha(43)
        Case 47
            campo = DevuelveDesdeBD(conAri, "max(codclien)", "sclien", "1", "1")
            PonerIdPrevistoCliente Val(campo) + 1



            PonerFoco txtAgente(10)
        End Select
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub PonerIdPrevistoCliente(IdPrev As Long)
    txtNumero(2).Text = Format(IdPrev, "000000")
    numParam = vEmpresa.DigitosUltimoNivel - 2 'Menos el 43 del principio de la codmacta
    txtNumero(3).Text = "43" & Right(String(10, "0") & IdPrev, numParam)
    txtNumero_LostFocus 3
End Sub

Private Sub CargaIconosAyuda()
Dim Ima As Image
    On Error Resume Next 'mejor que no diera errores, pero bien, tampoco vamos a enfadarnos
    For Each Ima In Me.imgayuda
        Ima.Picture = frmPpal.imgListComun.ListImages(46).Picture
    Next
    Err.Clear
End Sub



Private Sub Form_Load()
Dim H As Integer
Dim W As Integer
Dim IndiceCancel As Integer

    Me.Icon = frmPpal.Icon
    PrimeraVez = True
    CargaIconosAyuda
    limpiar Me
    FrListadoReparaciones.visible = False
    FrEstadisticasReparacionTecnico.visible = False
    FrameMultibase.visible = False
    FrameRecargaMov.visible = False
    Me.FrFacturaRecargas.visible = False
    FrProveedorxVenta.visible = False
    FrLiqCambioPrecios.visible = False
    Me.FrGeneraFactLiq.visible = False
    Me.FrImprimirFac.visible = False
    FrameAlbaProv.visible = False
    frameContabTickets.visible = False
    Me.FrameTraza.visible = False
    Me.FrameVEntasAgente.visible = False
    FrameListTrabajadores.visible = False
    FrameCambioProve.visible = False
    FrameCerrarAviso.visible = False
    FrameListadoPlantillas.visible = False
    FrameOtrasOfertas.visible = False
    FrameLlamadas.visible = False
    Me.FrameSituAlbaranes.visible = False
    FrameFrecuencia.visible = False
    FrameFacturarCliente.visible = False
    FramePedxZon.visible = False
    FrameReimpAlb.visible = False
    FrameCopiaPrecios.visible = False
    FrRepGaranProv.visible = False
    FrameRiesgo.visible = False
    FramePropPedido.visible = False
    FrameDtoCompra.visible = False
    Me.FraCambPrecTar.visible = False
    FrameDtosActiv.visible = False
    FrameResvtaAgente.visible = False
    FramePromociones.visible = False
    FrameBenClien.visible = False
    FrameControlAlbaranes.visible = False
    FrameHorasTrabajadasEuler.visible = False
    FrameCliPot.visible = False
    FrameBeneMarcaAgeProv.visible = False
    FrameMarcaFamilia.visible = False
    Caption = "Listado"
    IndiceCancel = Opcion
    Select Case Opcion
    Case 1
        'Listado reparaciones efectuadas
        PonerFrameVisible FrListadoReparaciones, H, W
        Me.lblDpto(0).Caption = DevuelveTextoDepto(True)
        
        
        Label3(158).Caption = ""  'Indicador
        
        
    Case 2
        PonerFrameVisible Me.FrEstadisticasReparacionTecnico, H, W
        Label3(63).Caption = ""
        
    Case 3
        Caption = "MULTIBASE"
        PonerFrameVisible Me.FrameMultibase, H, W
        CargaListMultibase
        cboRoot.ListIndex = 0
        cboRoot.visible = vUsu.Nivel = 0
    Case 4
        'Informe recarga movil
        PonerFrameVisible FrameRecargaMov, H, W
        Me.cmbRecargaMov(0).ListIndex = 0
        Me.cmbRecargaMov(1).ListIndex = 0
        
    Case 5
        
        'Ene 2013
        'YA no se utliza
        
        'Facturacion recargas moviles
        Caption = "Facturación"
        PonerFrameVisible FrFacturaRecargas, H, W
        txtFecha(8).Text = Format(Now, "dd/mm/yyyy")
        lblIndicadorT.Caption = ""
        pb1.visible = False
        'Lo del articulo lo pongo visib
        'txtArticulo(0).Text = vParamAplic.CodarticTfnia
        txtArticulo_LostFocus 0
        txtArticulo(0).visible = False
        Me.txtDescArticulo(0).visible = False
        Me.imgArticulo(0).visible = False
        Label4(2).visible = False
        
    Case 6
        'Ventas por codprove
        'TRAZA enero 2008
        PonerFrameVisible FrProveedorxVenta, H, W
        Label3(142).Caption = ""
    Case 7
        lblLiqu.Caption = ""
        PonerFrameVisible FrLiqCambioPrecios, H, W
    Case 8
        Label1.Caption = ""
        PonerFrameVisible FrGeneraFactLiq, H, W
    Case 9
        Label2.Caption = ""
        PonerFrameVisible FrImprimirFac, H, W
    Case 10
        PonerFrameVisible FrameAlbaProv, H, W
        
        
        'CadenaDesdeOtroForm
         
         
        Me.txtNumAlbar(4).Text = RecuperaValor(CadenaDesdeOtroForm, 1)
        Me.txtNumAlbar(5).Text = Me.txtNumAlbar(4).Text
         
        Me.txtFecha(18).Text = RecuperaValor(CadenaDesdeOtroForm, 2)
        Me.txtFecha(19).Text = Me.txtFecha(18).Text
        
        Me.txtCodProve(8).Text = RecuperaValor(CadenaDesdeOtroForm, 3)
        Me.txtCodProve(9).Text = Me.txtCodProve(8).Text
        
        Me.txtDescProve(8).Text = RecuperaValor(CadenaDesdeOtroForm, 4)
        Me.txtDescProve(9).Text = Me.txtDescProve(8).Text
        
        
        CadenaDesdeOtroForm = ""
        
    Case 13, 14
        Caption = "Tickets agrupados"
        If Opcion = 13 Then
            lblTitulo(10).Caption = "Facturar " & lblTitulo(10).Caption
            cmdContabTicket.Caption = "Contabilizar"
        Else
            lblTitulo(10).Caption = "Listados " & lblTitulo(10).Caption
            cmdContabTicket.Caption = "Aceptar"
        End If
        Me.FrameTapa.visible = Opcion = 13
        PonerFrameVisible frameContabTickets, H, W
        IndiceCancel = 13
    Case 15
         
        PonerFrameVisible FrameTraza, H, W
        
    Case 16
        PonerFrameVisible FrameVEntasAgente, H, W
        
    Case 17
        PonerFrameVisible FrameListTrabajadores, H, W
        
        
    Case 18
        Caption = "Cambio"
        PonerFrameVisible FrameCambioProve, H, W
    
    Case 19
        Caption = "Generar albarán"
        PonerFrameVisible Me.FrameCerrarAviso, H, W
        txtFecha(26).Text = Format(Now, "dd/mm/yyyy")
    Case 20
        
        PonerFrameVisible Me.FrameListadoPlantillas, H, W
        
    Case 21
        Caption = "Seleccionar"
        'optras ofertas del cliente
        PonerFrameVisible Me.FrameOtrasOfertas, H, W
    Case 22
        
        PonerFrameVisible Me.FrameLlamadas, H, W
        
    Case 23
        PonerFrameVisible FrameSituAlbaranes, H, W
        CargaListMov
        
        
    Case 24
        PonerFrameVisible FrameFrecuencia, H, W
        
    Case 25
        PonerFrameVisible FrameFacturarCliente, H, W
        txtFecha(31).Text = Format(Now, "dd/mm/yyyy")
        
    Case 26
        PonerFrameVisible FramePedxZon, H, W
        
    Case 27
        lblIndicAlb.Caption = ""
        txtFecha(35).Text = Format(Now, "dd/mm/yyyy")
        PonerFrameVisible FrameReimpAlb, H, W
    Case 28
        PonerFrameVisible FrameCopiaPrecios, H, W
        If CadenaDesdeOtroForm = "V" Then
            optCopiaPrecio(1).Value = True
            CadenaDesdeOtroForm = ""
        Else
            optCopiaPrecio(0).Value = True
        End If
        
    Case 30
        PonerFrameVisible FrRepGaranProv, H, W
        txtFecha(36).Text = Format("01/01/" & Year(Now), "dd/mm/yyyy")
        
    Case 31
        Label3(95).Caption = ""
        PonerFrameVisible FrameRiesgo, H, W
        
        
    Case 32
        PonerFrameVisible FramePropPedido, H, W
        cboProPed(0).ListIndex = 1
        cboProPed(1).ListIndex = 1
        Label3(100).Caption = ""
        imgayuda(2).ToolTipText = lblTitulo(29).Caption
        Me.txtAnyo(5).Text = 70
    Case 33
        PonerFrameVisible FrameDtoCompra, H, W
    Case 34, 44
        'cambio de precios
        '34.  De ventas
        '44: compras
        PonerFrameVisible FraCambPrecTar, H, W
        
        
        Me.lblTitulo(31).Caption = "Cambio precios tarifas "
        If Opcion = 34 Then
            lblTitulo(31).Caption = lblTitulo(31).Caption & "(VENTAS)"
        Else
            lblTitulo(31).Caption = lblTitulo(31).Caption & "(COMPRAS)"
        End If
        
        
        'Solo ventas
        lblDpto(43).visible = Opcion = 34
        imgTarifa(0).visible = Opcion = 34
        txtTarifa(0).visible = Opcion = 34
        txtDescTarifa(0).visible = Opcion = 34
        
        
        IndiceCancel = 34
    Case 35
        PonerFrameVisible Me.FrameDtosActiv, H, W
        
    Case 36
        Label3(122).Caption = ""
        PonerFrameVisible Me.FrameResvtaAgente, H, W
    Case 37, 40
        
        lblTitulo(34).Caption = "Beneficios por "
        If Opcion = 37 Then
            lblTitulo(34).Caption = lblTitulo(34).Caption & "agente"
        Else
            lblTitulo(34).Caption = lblTitulo(34).Caption & "proveedor"
        End If
        lblDpto(72).Caption = IIf(vParamAplic.NumeroInstalacion = 2, "Asociacion", "Ruta")
        chkBenAge(2).visible = Opcion = 37   'comparativo
        cboCoste(0).ListIndex = 0
        Label3(147).Caption = ""
        PonerFrameVisible Me.FrameBenxAge2, H, W
        IndiceCancel = 37
    Case 38, 39
        If Opcion = 38 Then
             Me.lblTitulo(35).Caption = " Cambio precios promociones"
             lblDpto(47).Caption = "Nueva fecha promoción"
        Else
            Me.lblTitulo(35).Caption = "Actualizar precios promociones"
            lblDpto(47).Caption = "Fecha promoción"
            IndiceCancel = 38
        End If
        cmdACtualizaPromo.visible = Opcion = 39
        Me.cmdCambioPromo.visible = Opcion = 38
        PonerFrameVisible FramePromociones, H, W
        
    Case 41
        Label3(156).Caption = ""
        cboCoste(1).ListIndex = 0
        PonerFrameVisible FrameBenClien, H, W
        
    Case 42, 43
        IndiceCancel = 42
        PonerFrameVisible FrameControlAlbaranes, H, W
        
        
        conSubRPT = Opcion = 42
        
        miSQL = "Informes control de albaranes"
        If Not conSubRPT Then miSQL = miSQL & " facturados"
        lblTitulo(37) = miSQL
        
        
        'La zona no esta visible para FACTURADOS opc=43
        lblDpto(54).visible = conSubRPT
        For numParam = 0 To 1
            Label3(159 + numParam).visible = conSubRPT
            imgZona(2 + numParam).visible = conSubRPT
            txtZona(2 + numParam).visible = conSubRPT
            txtDescZona(2 + numParam).visible = conSubRPT
        Next
        
    Case 45
        PonerFrameVisible FrameCambioEnFrecuencias, H, W
        
    Case 46
        PonerFrameVisible FrameHorasTrabajadasEuler, H, W
    Case 47
        PonerFrameVisible FrameCliPot, H, W
        Me.txtTextoNoEditable(0).Text = CStr(CadenaDesdeOtroForm)
        CadenaDesdeOtroForm = "" 'Si lo crea lo cargara el codclien
        Label8.Caption = ""
       
        
        
    Case 48
        Label3(183).Caption = ""
        cboCoste(2).ListIndex = 0
        PonerFrameVisible FrameBeneMarcaAgeProv, H, W
    Case 49, 50
        IndiceCancel = 49
        lblTitulo(42).Caption = IIf(Opcion = 49, "Ventas", "Compras")
        lblTitulo(42).Caption = lblTitulo(42).Caption & " marca-familia"
        lblTitulo(42).ForeColor = IIf(Opcion = 49, &H800000, &HC00000)
        Me.FrameAgente1.visible = Opcion = 49
        FrameProveedor1.visible = Opcion = 50
        FrameProveedor1.BorderStyle = 0
        FrameAgente1.BorderStyle = 0
        Label3(188).Caption = ""
        PonerFrameVisible FrameMarcaFamilia, H, W
        
    End Select
    Me.Height = H + 150
    Me.Width = W
    Me.cmdCancel(IndiceCancel).Cancel = True
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set miRsAux = Nothing
    
    If Opcion = 23 Then
        
        
        cadFrom = ""
        For numParam = 0 To List1.ListCount - 1
            'Seleccinamos los que NO estan marcados
            If Not List1.Selected(numParam) Then
                campo = List1.List(numParam)
                vMultiInforme = InStrRev(campo, "(")
                If vMultiInforme > 0 Then
                    campo = Mid(campo, vMultiInforme + 1)
                    vMultiInforme = InStr(1, campo, ")")
                    If vMultiInforme > 0 Then
                        campo = Mid(campo, 1, vMultiInforme - 1)
                        cadFrom = cadFrom & campo & "|"
                    End If
                End If
            End If
        Next
        textoValueGuardar "situalb", cadFrom
    End If
    
End Sub

Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub frmAg_DatoSeleccionado(CadenaSeleccion As String)
    Cadena_frmB = CadenaSeleccion
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
    Cadena_frmB = CadenaDevuelta
    
End Sub

Private Sub frmBaPr_DatoSeleccionado(CadenaSeleccion As String)
    Me.txtBancoPr(IndiceImg).Text = RecuperaValor(CadenaSeleccion, 1)
    Me.txtDescBancoPr(IndiceImg).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmC_Selec(vFecha As Date)
    txtFecha(IndiceImg).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmCli_DatoSeleccionado(CadenaSeleccion As String)
    txtCliente(IndiceImg).Text = RecuperaValor(CadenaSeleccion, 1)
    txtDescClie(IndiceImg).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmEn_DatoSeleccionado(CadenaSeleccion As String)
    miSQL = CadenaSeleccion
End Sub

Private Sub frmFP_DatoSeleccionado(CadenaSeleccion As String)
    Me.txtForpa(IndiceImg).Text = RecuperaValor(CadenaSeleccion, 1)
    Me.txtDescForpa(IndiceImg).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoArticulos_DatoSeleccionado(CadenaSeleccion As String)
    Me.txtArticulo(IndiceImg).Text = RecuperaValor(CadenaSeleccion, 1)
    Me.txtDescArticulo(IndiceImg).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmPr_DatoSeleccionado(CadenaSeleccion As String)
    txtCodProve(IndiceImg) = RecuperaValor(CadenaSeleccion, 1)
    txtDescProve(IndiceImg) = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmRut_DatoSeleccionado(CadenaSeleccion As String)
    campo = CadenaSeleccion
End Sub

Private Sub frmT_DatoSeleccionado(CadenaSeleccion As String)
    txtTrab(IndiceImg).Text = RecuperaValor(CadenaSeleccion, 1)
    txtDescTra(IndiceImg).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub



Private Sub frmZ_DatoSeleccionado(CadenaSeleccion As String)
    txtZona(IndiceImg).Text = RecuperaValor(CadenaSeleccion, 1)
    txtDescZona(IndiceImg).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub



Private Sub imgActividad_Click(Index As Integer)
Cadena_frmB = ""
    Screen.MousePointer = vbHourglass
    Set frmB = New frmBuscaGrid
    frmB.vTitulo = "Activiadad"
    campo = "Codigo|sactiv|codactiv|N||20·"
    campo = campo & "descripcion|sactiv|nomactiv|T||45·"
    frmB.vCampos = campo
    frmB.vCargaFrame = False
    frmB.vDevuelve = "0|1|"
    frmB.vselElem = 1
    frmB.vConexionGrid = 1  'ODBC Ariges
    frmB.vTabla = "sactiv"
    frmB.vSQL = ""
    frmB.Show vbModal
    Set frmB = Nothing
    Screen.MousePointer = vbDefault
    If Cadena_frmB <> "" Then
        
        Me.txtcodactiv(Index).Text = RecuperaValor(Cadena_frmB, 1)
        Me.txtDescActiv(Index).Text = RecuperaValor(Cadena_frmB, 2)
       
    End If
End Sub

Private Sub imgAgente_Click(Index As Integer)
    Cadena_frmB = ""
    Set frmAg = New frmFacAgentesCom
    frmAg.DatosADevolverBusqueda = "0|1|"
    frmAg.Show vbModal
    Set frmAg = Nothing
    
     If Cadena_frmB <> "" Then
         
        txtAgente(Index).Text = RecuperaValor(Cadena_frmB, 1)
        txtDescAgente(Index).Text = RecuperaValor(Cadena_frmB, 2)
       
    End If
    
End Sub

Private Sub imgAlma_Click(Index As Integer)
    Cadena_frmB = ""
    Screen.MousePointer = vbHourglass
    Set frmB = New frmBuscaGrid
    frmB.vTitulo = "Almacén"
    campo = "Codigo|salmpr|codalmac|N||20·"
    campo = campo & "descripcion|salmpr|nomalmac|T||45·"
    frmB.vCampos = campo
    frmB.vCargaFrame = False
    frmB.vDevuelve = "0|1|"
    frmB.vselElem = 1
    frmB.vConexionGrid = 1  'ODBC Ariges
    frmB.vTabla = "salmpr"
    frmB.vSQL = ""
    frmB.Show vbModal
    Set frmB = Nothing
    Screen.MousePointer = vbDefault
    If Cadena_frmB <> "" Then
        
        Me.txtAlma(Index).Text = RecuperaValor(Cadena_frmB, 1)
        Me.txtDescAlma(Index).Text = RecuperaValor(Cadena_frmB, 2)
        If Index = 0 Then PonerFoco txtCodProve(17)
    End If
End Sub

Private Sub imgArticulo_Click(Index As Integer)
    IndiceImg = Index
    Set frmMtoArticulos = New frmAlmArticu2
    'frmMtoArticulos.DatosADevolverBusqueda3 = "@1@" 'Abrimos en Modo Busqueda
    frmMtoArticulos.DesdeTPV = False
    frmMtoArticulos.Show vbModal
    Set frmMtoArticulos = Nothing
End Sub

Private Sub imgayuda_Click(Index As Integer)
Dim Ayuda As String

    'Sera las ayuda. Tampoco queiero la biblia, pero,
    'si un "pelin" de ayuda no me vendria mal a mi, imaginemos a el cliente final
    Select Case Index
    Case 0
        Ayuda = "Importe minimo para listado por agentes o compartativos"
    Case 1
        Ayuda = Ayuda & "No sale articulo punto verde" & vbCrLf & vbCrLf
        Ayuda = Ayuda & "(En listado ventas agente NO salen los portes y NO salen las facturas rectificativas)" & vbCrLf
        
        Ayuda = Ayuda & vbCrLf & " - Visitador: El desde/hasta agente sera visitador. El informe incluye un subnivel nuevo " & vbCrLf
        
    Case 2
        Ayuda = "Campo NSAL es el numero de albaranes/facturas en los que se encuentra el articulo, en los ultimos " & vParamAplic.Rot_ConsumMes1 & " meses" & vbCrLf & vbCrLf
        Ayuda = Ayuda & "-Si indica el minimo de albaranes(SIN PONER PROVEEDOR), entonces sólo saldran aquellos cuyo resultado sea >=0" & vbCrLf & vbCrLf
        Ayuda = Ayuda & "-Si indica valor en ""% mismo cliente"" (que NO sea de VARIOS) marcará en el listado los articulos en los cuales " & vbCrLf
        Ayuda = Ayuda & "     las ventas, en un n-%. pertencen a un mismo cliente y mas de 1-Uds vendidas." & vbCrLf
        
        Ayuda = Ayuda & vbCrLf & "-Si consolida almacén, saldrán los datos agrupados para los dos almacenes"
        Ayuda = Ayuda & vbCrLf & "-Rotacion: Los articulos de varios saldran cuando tengan pedidos cliente"
    Case 3
        Ayuda = "Si selecciona solo articulos con stock:"
        Ayuda = Ayuda & vbCrLf & " - Quitara los articulos que no tengan la marca de control stock(y 0 en stock)"
        Ayuda = Ayuda & vbCrLf & " y si alguno de los articulos del pedido tiene stock muestra todo el pedido"
        Ayuda = Ayuda & vbCrLf & " También mostrará todo si alguno de los articulos es de varios"
        Ayuda = Ayuda & vbCrLf & " - Si tiene la marca de 'Servir completo' y alguna de las lineas no tiene para servir NO sale el pedido"
        
    Case 4
        Ayuda = "-No muestra las ventas de articulos varios excepto en el comparativo."
        Ayuda = Ayuda & vbCrLf & "-Marcar  'Aplica descuento', calculará el coste aplicando el valor de Dto Sin Cargo de descuentos proveedor"
    Case 5
        Ayuda = "-Marcar   'Aplica descuento', calculará el coste aplicando el valor de Dto Sin Cargo de descuentos proveedor"
    Case 6
        Ayuda = "- Incluye varios y presu"
        Ayuda = "-Marcar   'Aplica descuento', calculará el coste aplicando el valor de Dto Sin Cargo de descuentos proveedor"
    
    Case 7
        Ayuda = "Con los datos del cliente potencial lo insertará en la tabla de clientes."
        Ayuda = Ayuda & vbCrLf & "-Codigo:  Codigo que le va a asignar. "
        Ayuda = Ayuda & vbCrLf & " Los prismaticos permiten buscar un hueco a partir del codigo introducido."
        Ayuda = Ayuda & vbCrLf & "-Cuenta contable . Puede ya existir. Insertara con esa cuenta el cliente"
    End Select
    Ayuda = imgayuda(Index).ToolTipText & vbCrLf & String(47, "=") & vbCrLf & vbCrLf & Ayuda
    MsgBox Ayuda, vbInformation

End Sub

Private Sub imgBancoPr_Click(Index As Integer)
    IndiceImg = Index
    Set frmBaPr = New frmFacBancosPropios
    frmBaPr.DatosADevolverBusqueda = "1" 'Abrimos en Modo Busqueda
    frmBaPr.Show vbModal
    Set frmBaPr = Nothing
End Sub

Private Sub imgCheck_Click(Index As Integer)
Dim i As Integer

    If Index < 2 Then
        'Seleecionar otras ofertas
        For i = 1 To Me.lw1.ListItems.Count
            lw1.ListItems(i).Checked = Index = 1
        Next i
    ElseIf Index < 4 Then
        'Seleccionar Tipos albaran para listado situacion labaranes
        For i = 0 To List1.ListCount - 1
            List1.Selected(i) = Index = 2
        Next i
    ElseIf Index < 6 Then
        For i = 1 To TreeView1.Nodes.Count
            TreeView1.Nodes(i).Checked = Index = 5
        Next i
    End If
    
End Sub

Private Sub imgCliente_Click(Index As Integer)
    
    Screen.MousePointer = vbHourglass
    IndiceImg = Index
    Set frmCli = New frmFacClientes3
    frmCli.DatosADevolverBusqueda = "0|1|"
    frmCli.Show vbModal
    Set frmCli = Nothing
End Sub


Private Sub LanzaBusquedaDpto(Indice As Integer)
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        If Indice = 0 Then
            campo = Me.lblDpto(0).Caption
        Else
            campo = "Departamento"
            Indice = Indice + 9
        End If
        frmB.vTitulo = campo & " " & txtCliente(Indice).Text & " - " & txtDescClie(Indice).Text
        campo = "Cod.|sdirec|coddirec|N||20·"
        campo = campo & "Desc.|sdirec|nomdirec|T||40·"
        frmB.vCampos = campo
        frmB.vCargaFrame = False
        frmB.vDevuelve = "0|1|"
        frmB.vselElem = 1
        frmB.vConexionGrid = 1  'ODBC Ariges
        frmB.vTabla = "sdirec"
        frmB.vSQL = "codclien = " & txtCliente(Indice).Text
        frmB.Show vbModal
        Set frmB = Nothing
        Screen.MousePointer = vbDefault
End Sub


Private Sub imgDpto_Click(Index As Integer)
    
    
    Cadena_frmB = ""
    
    Select Case Index
    Case 0, 1
        'DPTO
        IndiceImg = Index
       
        If txtCliente(0).Text <> "" And txtCliente(0).Text = txtCliente(1).Text And txtDescClie(0).Text <> "" Then
            'OK
            LanzaBusquedaDpto 0
            
        Else
            MsgBox "Para poner el departamento cliente debe y el hasta  debe ser el mismo", vbExclamation
        End If
    Case 2, 3
        IndiceImg = Index
        If txtCliente(Index + 9).Text = "" Then
            MsgBox "Indique el cliente", vbExclamation
        Else
            LanzaBusquedaDpto Index
        End If
    End Select
    

    If Cadena_frmB <> "" Then
        txtDpto(IndiceImg).Text = RecuperaValor(Cadena_frmB, 1)
        txtDescDpto(IndiceImg) = RecuperaValor(Cadena_frmB, 2)
    End If

End Sub

Private Sub imgEnvio_Click(Index As Integer)
    miSQL = ""
    Set frmEn = New frmFacFormasEnvio
    frmEn.DatosADevolverBusqueda = "0|1|"
    frmEn.Show vbModal
    Set frmEn = Nothing
    If miSQL <> "" Then
    
    End If
End Sub

Private Sub imgFamilia_Click(Index As Integer)
    Cadena_frmB = ""
    Screen.MousePointer = vbHourglass
    Set frmB = New frmBuscaGrid
    frmB.vTitulo = "Familia"
    campo = "Codigo|sfamia|Codfamia|N||20·"
    campo = campo & "descripcion|sfamia|nomfamia|T||45·"
    frmB.vCampos = campo
    frmB.vCargaFrame = False
    frmB.vDevuelve = "0|1|"
    frmB.vselElem = 1
    frmB.vConexionGrid = 1  'ODBC Ariges
    frmB.vTabla = "sfamia"
    frmB.vSQL = ""
    frmB.Show vbModal
    Set frmB = Nothing
    Screen.MousePointer = vbDefault
    If Cadena_frmB <> "" Then
        
        Me.txtFamia(Index).Text = RecuperaValor(Cadena_frmB, 1)
        Me.txtDescFamia(Index).Text = RecuperaValor(Cadena_frmB, 2)
        If Index = 2 Then
            PonerFoco txtFamia(3)
        Else
            PonerFoco txtmarca(0)
        End If
    End If
End Sub

Private Sub imgFecha_Click(Index As Integer)
   IndiceImg = Index
   Set frmC = New frmCal
   frmC.Fecha = Now
   If txtFecha(IndiceImg).Text <> "" Then
        If IsDate(txtFecha(IndiceImg).Text) Then frmC.Fecha = CDate(txtFecha(IndiceImg).Text)
    End If
   frmC.Show vbModal
   Set frmC = Nothing
    
End Sub


Private Sub imgForPa_Click(Index As Integer)
    IndiceImg = Index
    Set frmFP = New frmFacFormasPago
    frmFP.DatosADevolverBusqueda = "0|1|"
    frmFP.Show vbModal
    Set frmFP = Nothing
End Sub

Private Sub imgIdClienteLibre_Click()
    
    Set miRsAux = New ADODB.Recordset
    If txtNumero(2).Text = "" Then
        miSQL = "0"
    Else
        miSQL = txtNumero(2).Text
    End If
    campo = " Where codClien > " & miSQL
    campo = "select codclien,@rownum:=@rownum+1 AS rownum from sclien, (SELECT @rownum:=" & miSQL & ") r" & campo
    miRsAux.Open campo, conn, adOpenKeyset, adLockReadOnly, adCmdText
    NumRegElim = -1
    While Not miRsAux.EOF
        
        If (miRsAux!codClien - miRsAux!rownum) > 0 Then
            NumRegElim = miRsAux!codClien - 1
            'Este es el codigo
            miRsAux.MoveLast
        Else
            'No hacemos nada
            NumRegElim = miRsAux!codClien + 1
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    If NumRegElim >= 0 Then PonerIdPrevistoCliente NumRegElim
    
    

End Sub

Private Sub imgMarca_Click(Index As Integer)
    Cadena_frmB = ""
    Screen.MousePointer = vbHourglass
    Set frmB = New frmBuscaGrid
    frmB.vTitulo = "Marcas"
    campo = "Codigo|smarca|codmarca|N||20·"
    campo = campo & "descripcion|smarca|nommarca|T||45·"
    frmB.vCampos = campo
    frmB.vCargaFrame = False
    frmB.vDevuelve = "0|1|"
    frmB.vselElem = 1
    frmB.vConexionGrid = 1  'ODBC Ariges
    frmB.vTabla = "smarca"
    frmB.vSQL = ""
    frmB.Show vbModal
    Set frmB = Nothing
    Screen.MousePointer = vbDefault
    If Cadena_frmB <> "" Then
        
        Me.txtmarca(Index).Text = RecuperaValor(Cadena_frmB, 1)
        Me.txtDescmarca(Index).Text = RecuperaValor(Cadena_frmB, 2)
        If Index = 0 Then
            PonerFoco txtmarca(1)
        Else
            PonerFocoBtn Me.cmdPropuestaPedido
        End If
    End If

End Sub

Private Sub imgProveedor_Click(Index As Integer)
    IndiceImg = Index
    Set frmPr = New frmComProveedores
    frmPr.DatosADevolverBusqueda = "0|1|"
    frmPr.Show vbModal
    Set frmPr = Nothing
End Sub

Private Sub imgRuta_Click(Index As Integer)
    campo = ""
    Set frmRut = New frmFacRutas
    frmRut.DatosADevolverBusqueda = "0|1"
    frmRut.DeConsulta = True
    frmRut.Show vbModal
    Set frmRut = Nothing
    If campo <> "" Then
        Me.txtRuta(Index).Text = RecuperaValor(campo, 1)
        Me.txtDescRuta(Index).Text = RecuperaValor(campo, 2)
    End If
End Sub

Private Sub imgTarifa_Click(Index As Integer)
            Cadena_frmB = ""
            Screen.MousePointer = vbHourglass
            Set frmB = New frmBuscaGrid
            frmB.vTitulo = "Tarifas"
            campo = "Codigo|starif|codlista|N||20·"
            campo = campo & "Nombre|startif|nomlista|T||40·"
            frmB.vCampos = campo
            frmB.vCargaFrame = False
            frmB.vDevuelve = "0|1|"
            frmB.vselElem = 1
            frmB.vConexionGrid = 1  'ODBC Ariges
            frmB.vTabla = "starif"
            frmB.vSQL = ""
            frmB.Show vbModal
            Set frmB = Nothing
            Screen.MousePointer = vbDefault
            If Cadena_frmB <> "" Then
                Me.txtTarifa(Index).Text = RecuperaValor(Cadena_frmB, 1)
                Me.txtDescTarifa(Index).Text = RecuperaValor(Cadena_frmB, 2)
            End If
End Sub

Private Sub imgTecnico_Click(Index As Integer)
    IndiceImg = Index
    If Index < 3 Then
        Set frmT = New frmAdmTrabajadores
        frmT.DatosADevolverBusqueda = "0|1|"
        frmT.Show vbModal
        Set frmT = Nothing

    Else
        'Listado trabajadores
            Cadena_frmB = ""
            Screen.MousePointer = vbHourglass
            Set frmB = New frmBuscaGrid
            frmB.vTitulo = "Trabajadores"
            campo = "Codigo|straba|codtraba|N||20·"
            campo = campo & "Nombre|straba|nomtraba|T||40·"
            campo = campo & "NIF|straba|niftraba|T||20·"
            frmB.vCampos = campo
            frmB.vCargaFrame = False
            frmB.vDevuelve = "0|1|"
            frmB.vselElem = 1
            frmB.vConexionGrid = 1  'ODBC Ariges
            frmB.vTabla = "straba"
            frmB.vSQL = ""
            frmB.Show vbModal
            Set frmB = Nothing
            Screen.MousePointer = vbDefault
            If Cadena_frmB <> "" Then
                Me.txtTrab(Index).Text = RecuperaValor(Cadena_frmB, 1)
                Me.txtDescTra(Index).Text = RecuperaValor(Cadena_frmB, 2)
            End If
    End If
End Sub

Private Sub imgZona_Click(Index As Integer)
    IndiceImg = Index
    Set frmZ = New frmFacZonas
    frmZ.DatosADevolverBusqueda = "0|1|"
    frmZ.Show vbModal
    Set frmZ = Nothing
End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub



Private Sub optAlbTrans_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub



Private Sub optCopiaPrecio_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub optInfProd_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub optReparaciones_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub







Private Sub optSituaArt_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub texto_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub TreeView1_NodeCheck(ByVal Node As MSComctlLib.Node)
Dim N As Node

    If Not Node.Child Is Nothing Then
        Set N = Node.Child
        While Not N Is Nothing
            N.Checked = N.Parent.Checked
            Set N = N.Next
        Wend
    End If
End Sub

Private Sub txtAgente_GotFocus(Index As Integer)
    ConseguirFoco txtAgente(Index), 3
End Sub

Private Sub txtAgente_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If (Shift And vbCtrlMask) > 0 Then
        'Caption = KeyCode
        If KeyCode = 65 Then imgAgente_Click Index
    End If
End Sub

Private Sub txtAgente_KeyPress(Index As Integer, KeyAscii As Integer)
     KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtAgente_LostFocus(Index As Integer)
    miSQL = ""
    txtAgente(Index).Text = Trim(txtAgente(Index).Text)
    If txtAgente(Index).Text <> "" Then
        If PonerFormatoEntero(txtAgente(Index)) Then
            
            miSQL = DevuelveDesdeBD(conAri, "nomagent", "sagent", "codagent", txtAgente(Index).Text)
            If miSQL = "" Then MsgBox "No existe el agente: " & txtAgente(Index).Text, vbExclamation
        End If
    End If
    Me.txtDescAgente(Index).Text = miSQL
    miSQL = ""
End Sub



Private Sub txtAlma_GotFocus(Index As Integer)
    ConseguirFoco txtAlma(Index), 3
End Sub



Private Sub txtAlma_KeyPress(Index As Integer, KeyAscii As Integer)
      KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtAlma_LostFocus(Index As Integer)
    txtAlma(Index).Text = Trim(txtAlma(Index).Text)
    codigo = ""
    miSQL = ""
    If txtAlma(Index).Text <> "" Then
        If IsNumeric(txtAlma(Index).Text) Then
            codigo = DevuelveDesdeBD(conAri, "nomalmac", "salmpr", "codalmac", txtAlma(Index).Text, "N")
            If codigo = "" Then miSQL = "El codigo no pertence a ningun almacén"
        Else
            miSQL = "Campo numerico"
        End If
    End If
     
    Me.txtDescAlma(Index).Text = codigo
    If miSQL <> "" Then
        MsgBox miSQL, vbExclamation
        If Index = 0 Then txtAlma(Index).Text = ""
        PonerFoco txtBancoPr(Index)
    End If
End Sub

Private Sub txtanyo_GotFocus(Index As Integer)
    ConseguirFoco txtAnyo(Index), 3
End Sub

Private Sub txtanyo_KeyPress(Index As Integer, KeyAscii As Integer)
     KEYpressGnral KeyAscii, 3, True
End Sub

Private Sub txtanyo_LostFocus(Index As Integer)
    txtAnyo(Index).Text = Trim(txtAnyo(Index).Text)
    miSQL = ""
    If txtAnyo(Index).Text <> "" Then
        If Not PonerFormatoEntero(txtAnyo(Index)) Then txtAnyo(Index).Text = ""
    End If
    
    If Index = 2 Or Index = 3 Then
        'SON MES
        If txtAnyo(Index).Text <> "" Then
            If Val(txtAnyo(Index).Text) < 1 Or Val(txtAnyo(Index).Text) > 12 Then
                MsgBox "Mes incorrecto", vbExclamation
                txtAnyo(Index).Text = ""
                PonerFoco txtAnyo(Index)
                
            End If
        End If
    End If
    
End Sub

Private Sub txtArticulo_GotFocus(Index As Integer)
    ConseguirFoco txtArticulo(Index), 3
End Sub

Private Sub txtArticulo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If (Shift And vbCtrlMask) > 0 Then
        'Caption = KeyCode
        If KeyCode = 65 Then imgArticulo_Click Index
    End If
End Sub

Private Sub txtArticulo_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
    
End Sub


Private Sub txtArticulo_LostFocus(Index As Integer)
Dim T As String
    
    txtArticulo(Index).Text = Trim(txtArticulo(Index).Text)
    If txtArticulo(Index).Text = "" Then
        'EN blanco
        txtDescArticulo(Index).Text = ""
        Exit Sub
    End If
    
    
    T = "codartic"
    codigo = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", txtArticulo(Index).Text, "T", T)
    If codigo = "" Then
        MsgBox "No existe el artículo : " & txtArticulo(Index).Text, vbExclamation
    Else
        txtArticulo(Index).Text = T
    End If
    Me.txtDescArticulo(Index).Text = codigo
    codigo = ""
    
End Sub



Private Sub txtBancoPr_GotFocus(Index As Integer)
    ConseguirFoco txtBancoPr(Index), 3
End Sub

Private Sub txtBancoPr_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtBancoPr_LostFocus(Index As Integer)
    txtBancoPr(Index).Text = Trim(txtBancoPr(Index).Text)
    codigo = ""
    miSQL = ""
    If txtBancoPr(Index).Text <> "" Then
        If IsNumeric(txtBancoPr(Index).Text) Then
            codigo = DevuelveDesdeBD(conAri, "nombanpr", "sbanpr", "codbanpr", txtBancoPr(Index).Text, "N")
            If codigo = "" Then miSQL = "El codigo no pertence a ningun banco propio"
        Else
            miSQL = "Campo numerico"
        End If
    End If
     
    Me.txtDescBancoPr(Index).Text = codigo
    If miSQL <> "" Then
        MsgBox miSQL, vbExclamation
        txtBancoPr(Index).Text = ""
        PonerFoco txtBancoPr(Index)
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
            txtCliente(Index).Text = Format(txtCliente(Index).Text, "00000")
            Descri = DevuelveDesdeBD(conAri, "nomclien", "sclien", "codclien", txtCliente(Index).Text, "N")
            If Descri = "" Then MsgBox "No existe el cliente : " & txtCliente(Index).Text, vbExclamation
           
        End If
        
        If Index = 11 Or Index = 12 Then
             If Descri = "" Then
                    txtCliente(Index).Text = ""
                    PonerFoco txtCliente(Index)
              Else
                     Me.txtDpto(Index - 9).Text = ""
                     Me.txtDescDpto(Index - 9).Text = ""
              End If
        End If
        
    End If
    Me.txtDescClie(Index).Text = Descri
   
        

    
End Sub


    

Private Sub txtcodactiv_GotFocus(Index As Integer)
    ConseguirFoco txtcodactiv(Index), 3
End Sub

Private Sub txtcodactiv_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtcodactiv_LostFocus(Index As Integer)
 txtcodactiv(Index).Text = Trim(txtcodactiv(Index).Text)
    codigo = ""
    miSQL = ""
    If txtcodactiv(Index).Text <> "" Then
        If IsNumeric(txtcodactiv(Index).Text) Then
            codigo = DevuelveDesdeBD(conAri, "nomactiv", "sactiv", "codactiv", txtcodactiv(Index).Text, "N")
        Else
            miSQL = "Campo numerico"
        End If
    End If
    Me.txtDescActiv(Index).Text = codigo
    If miSQL <> "" Then
        MsgBox miSQL, vbExclamation
        txtcodactiv(Index).Text = ""
        PonerFoco txtcodactiv(Index)
    End If
End Sub

Private Sub txtCodProve_GotFocus(Index As Integer)
    ConseguirFoco txtCodProve(Index), 3
End Sub

Private Sub txtCodProve_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If (Shift And vbCtrlMask) > 0 Then
        'Caption = KeyCode
        If KeyCode = 65 Then imgProveedor_Click Index
    End If
End Sub

Private Sub txtCodProve_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtCodProve_LostFocus(Index As Integer)
    txtCodProve(Index).Text = Trim(txtCodProve(Index).Text)
    codigo = ""
    miSQL = ""
    If txtCodProve(Index).Text <> "" Then
        If IsNumeric(txtCodProve(Index).Text) Then
            codigo = DevuelveDesdeBD(conAri, "nomprove", "sprove", "codprove", txtCodProve(Index).Text, "N")
            If codigo = "" Then
                If Index = 20 Or Index = 12 Then
                    'Codprove REQUERIDO
                    miSQL = "No existe proveedor"
                Else
                    MsgBox "El codigo no pertence a ningun proveedor", vbExclamation
                End If
            End If
        Else
            miSQL = "Campo numerico"
        End If
    End If
    Me.txtDescProve(Index).Text = codigo
    If miSQL <> "" Then
        MsgBox miSQL, vbExclamation
        txtCodProve(Index).Text = ""
        PonerFoco txtCodProve(Index)
    End If
End Sub





Private Sub txtEnvio_GotFocus(Index As Integer)
     ConseguirFoco txtEnvio(Index), 3
End Sub

Private Sub txtEnvio_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 65 Then imgEnvio_Click Index
End Sub

Private Sub txtEnvio_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtEnvio_LostFocus(Index As Integer)
    txtEnvio(Index).Text = Trim(txtEnvio(Index).Text)
    codigo = ""
    miSQL = ""
    campo = ""
    If txtEnvio(Index).Text <> "" Then
        If IsNumeric(txtEnvio(Index).Text) Then
            codigo = Format(txtEnvio(Index).Text, "000")
            miSQL = DevuelveDesdeBD(conAri, "nomenvio", "senvio", "codenvio", txtEnvio(Index).Text, "N")
            If miSQL = "" Then campo = "El codigo no pertence a ningun forma de envio"
                
        Else
            campo = "Campo numerico"
        End If
    End If
    
        
    txtDescEnvio(Index).Text = miSQL
    If campo <> "" Then MsgBox campo, vbExclamation
    txtEnvio(Index).Text = codigo
    If codigo = "" And campo <> "" Then PonerFoco txtEnvio(Index)
    
End Sub

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
    txtFamia(Index).Text = Trim(txtFamia(Index).Text)
    codigo = ""
    miSQL = ""
    If txtFamia(Index).Text <> "" Then
        If IsNumeric(txtFamia(Index).Text) Then
            codigo = DevuelveDesdeBD(conAri, "nomfamia", "sfamia", "codfamia", txtFamia(Index).Text, "N")
            If codigo = "" Then MsgBox "El codigo no pertence a ningun familia", vbExclamation
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

Private Sub txtForpa_GotFocus(Index As Integer)
    ConseguirFoco txtForpa(Index), 3
End Sub

Private Sub txtForpa_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtForpa_LostFocus(Index As Integer)
    txtForpa(Index).Text = Trim(txtForpa(Index).Text)
    codigo = ""
    miSQL = ""
    numParam = 0
    If txtForpa(Index).Text <> "" Then
        If IsNumeric(txtForpa(Index).Text) Then
            codigo = DevuelveDesdeBD(conAri, "nomforpa", "sforpa", "codforpa", txtForpa(Index).Text, "N")
            If codigo = "" Then miSQL = "El codigo no pertence a ningun forma de pago"
            numParam = 1
        Else
            miSQL = "Campo numerico"
        End If
    End If
    Me.txtDescForpa(Index).Text = codigo
    If miSQL <> "" Then
        
        
        If Index = 0 Then
            'Es obligado
            MsgBox miSQL, vbExclamation
            txtForpa(Index).Text = ""
            PonerFoco txtForpa(Index)
        Else
            If numParam = 0 Then
                MsgBox miSQL, vbExclamation
                txtForpa(Index).Text = ""
            End If
            txtDescForpa(Index).Text = ""
        End If
    End If
End Sub

Private Sub txtGrupoPlan_GotFocus(Index As Integer)
    ConseguirFoco txtGrupoPlan(Index), 3
End Sub

Private Sub txtGrupoPlan_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtGrupoPlan_LostFocus(Index As Integer)
    txtGrupoPlan(Index).Text = Trim(txtGrupoPlan(Index).Text)
    miSQL = ""
    If txtGrupoPlan(Index).Text <> "" Then
        If Not PonerFormatoEntero(txtGrupoPlan(Index)) Then
            txtGrupoPlan(Index).Text = ""
            'PonerFoco txtGrupoPlan(Index)
        Else
            miSQL = DevuelveDesdeBD(conAri, "nomgrupl", "sgrupl", "codgrupl", txtGrupoPlan(Index).Text)
            If miSQL = "" Then miSQL = "no existe"
        End If
    End If
    txtDescGrupoP(Index).Text = miSQL
End Sub

Private Sub txtImporte_GotFocus(Index As Integer)
    ConseguirFoco txtImporte(Index), 3
End Sub

Private Sub txtImporte_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtImporte_LostFocus(Index As Integer)
    txtImporte(Index).Text = Trim(txtImporte(Index).Text)
    If txtImporte(Index).Text = "" Then Exit Sub
    Select Case Index
    Case 0
    
        PonerFormatoDecimal txtImporte(Index), 2   'decimal 10,4  en formato decimal
    Case 1
        'El uno es obligado el campo
        If Not PonerFormatoDecimal(txtImporte(Index), 3) Then txtImporte(Index).Text = ""   'importe
        
    Case 2
        PonerFormatoDecimal txtImporte(Index), 1   '2 decimales
    End Select
End Sub




Private Sub txtmarca_GotFocus(Index As Integer)
    ConseguirFoco txtmarca(Index), 3
End Sub

Private Sub txtmarca_KeyPress(Index As Integer, KeyAscii As Integer)
     KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtmarca_LostFocus(Index As Integer)
    txtmarca(Index).Text = Trim(txtmarca(Index).Text)
    codigo = ""
    miSQL = ""

    If txtmarca(Index).Text <> "" Then
        If IsNumeric(txtmarca(Index).Text) Then
            codigo = DevuelveDesdeBD(conAri, "nommarca", "smarca", "codmarca", txtmarca(Index).Text, "N")
            If codigo = "" Then MsgBox "El código no pertence a ninguna marca", vbExclamation
            
        Else
            miSQL = "Campo numerico"
        End If
    End If
    
    Me.txtDescmarca(Index).Text = codigo
    If miSQL <> "" Then
        MsgBox miSQL, vbExclamation
        If Index < 2 Then
            txtmarca(Index).Text = ""
            PonerFoco txtmarca(Index)
        End If
    End If
End Sub

Private Sub txtNumAlbar_GotFocus(Index As Integer)
    ConseguirFoco txtNumAlbar(Index), 3
End Sub

Private Sub txtNumAlbar_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub



Private Sub txtNumero_GotFocus(Index As Integer)
    ConseguirFoco txtNumero(Index), 3
End Sub

Private Sub txtNumero_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtNumero_LostFocus(Index As Integer)
Dim Mal As Boolean
    txtNumero(Index).Text = Trim(txtNumero(Index).Text)
    If txtNumero(Index).Text = "" Then Exit Sub
    
    Mal = True
    If Not PonerFormatoEntero(txtNumero(Index)) Then
        'Mal = True
        
    Else
        If Index = 2 Then
            Mal = False  'OK
            miSQL = DevuelveDesdeBD(conAri, "nomclien", "sclien", "codclien", txtNumero(2).Text)
            If miSQL <> "" Then
                MsgBox "Ya existe el codigo de cliente: " & txtNumero(2).Text & " " & miSQL, vbExclamation
                'Mal = True
            Else
                PonerIdPrevistoCliente CLng(txtNumero(2).Text)
            End If
            
        ElseIf Index = 3 Then
            If Len(txtNumero(3).Text) <> vEmpresa.DigitosUltimoNivel Then
                MsgBox "Longituda de cuenta incorrecta" & txtNumero(3).Text, vbExclamation
            Else
                miSQL = DevuelveDesdeBD(conConta, "nommacta", "cuentas", "codmacta", txtNumero(3).Text, "T")
                If miSQL = "" Then miSQL = "*** Nueva cuenta contabilidad ***"
                Label8.Caption = miSQL
                Mal = False
            End If
        End If
    End If
    
    If Mal Then
        txtNumero(Index).Text = ""
        If Index = 2 Then txtNumero(3).Text = "": Label8.Caption = ""
        If Index = 3 Then Label8.Caption = ""
    End If
End Sub

Private Sub txtRecargaMov_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub



Private Sub txtRuta_GotFocus(Index As Integer)
    ConseguirFoco txtRuta(Index), 3
    
End Sub

Private Sub txtRuta_KeyPress(Index As Integer, KeyAscii As Integer)
   KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtRuta_LostFocus(Index As Integer)
    txtRuta(Index).Text = Trim(txtRuta(Index).Text)
    codigo = ""
    miSQL = ""

    If txtRuta(Index).Text <> "" Then
        If IsNumeric(txtRuta(Index).Text) Then
            codigo = DevuelveDesdeBD(conAri, "nomrutas", "srutas", "codrutas", txtRuta(Index).Text, "N")
            If codigo = "" Then miSQL = "El código no pertence a ningun registro"
        Else
            miSQL = "Campo numerico"
        End If
    End If
    
    Me.txtDescRuta(Index).Text = codigo
    If miSQL <> "" Then
        MsgBox miSQL, vbExclamation
        If Index = 0 Then
            txtRuta(Index).Text = ""
            PonerFoco txtRuta(Index)
        End If
    End If
End Sub

Private Sub txtTarifa_GotFocus(Index As Integer)
    ConseguirFoco txtTarifa(Index), 3
End Sub

Private Sub txtTarifa_KeyPress(Index As Integer, KeyAscii As Integer)
     KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtTarifa_LostFocus(Index As Integer)

    txtTarifa(Index).Text = Trim(txtTarifa(Index).Text)
    codigo = ""
    miSQL = ""

    If txtTarifa(Index).Text <> "" Then
        If IsNumeric(txtTarifa(Index).Text) Then
            codigo = DevuelveDesdeBD(conAri, "nomlista", "starif", "codlista", txtTarifa(Index).Text, "N")
            If codigo = "" Then miSQL = "El código no pertence a ninguna tarifa"
        Else
            miSQL = "Campo numerico"
        End If
    End If
    
    Me.txtDescTarifa(Index).Text = codigo
    If miSQL <> "" Then
        MsgBox miSQL, vbExclamation
        If Index = 0 Then
            txtTarifa(Index).Text = ""
            PonerFoco txtTarifa(Index)
        End If
    End If
End Sub

Private Sub txtTrab_GotFocus(Index As Integer)
    ConseguirFoco txtTrab(Index), 3
End Sub

Private Sub txtTrab_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtTrab_LostFocus(Index As Integer)


    txtTrab(Index).Text = Trim(txtTrab(Index).Text)
    codigo = ""
    miSQL = ""

    If txtTrab(Index).Text <> "" Then
        If IsNumeric(txtTrab(Index).Text) Then
            codigo = DevuelveDesdeBD(conAri, "nomtraba", "straba", "codtraba", txtTrab(Index).Text, "N")
            If codigo = "" Then miSQL = "El código no pertence a ningun trabajador"
        Else
            miSQL = "Campo numerico"
        End If
    End If
    
    Me.txtDescTra(Index).Text = codigo
    If miSQL <> "" Then
        MsgBox miSQL, vbExclamation
        If Index < 2 Then
            txtTrab(Index).Text = ""
            PonerFoco txtTrab(Index)
        End If
    End If
End Sub



Private Sub txtDpto_GotFocus(Index As Integer)
    ConseguirFoco txtDpto(Index), 3
End Sub

Private Sub txtDpto_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtDpto_LostFocus(Index As Integer)
Dim vC As CCliente
    'Si el cliente ES EL MISMO
    campo = ""
    txtDpto(Index).Text = Trim(txtDpto(Index).Text)
    If Index < 2 Then
        If txtDpto(Index).Text <> "" Then
             'Index=0 or 1.  Departamento sera puesto si, y solo si, el cliente es el mismo
             If txtCliente(0).Text <> "" And txtCliente(0).Text = txtCliente(1).Text And txtDescClie(0).Text <> "" Then
                 'PERFECTO, el cliente existe y es el mismo
                 Set vC = New CCliente
                 vC.codigo = txtCliente(0).Text
                 vC.DptoCliente txtDpto(Index).Text, campo
                 Set vC = Nothing
             Else
                 'Todavia no ha puesto el cliente
                 MsgBox "Para poner el departamento cliente debe y el hasta  debe ser el mismo", vbExclamation
                 txtDpto(Index).Text = ""
        
             End If
        End If
        Me.txtDescDpto(Index).Text = campo
    ElseIf Index < 4 Then
        
            If txtDpto(Index).Text <> "" Then
                If txtCliente(Index + 9).Text <> "" And txtDescClie(Index + 9).Text <> "" Then
                     'PERFECTO, el cliente existe y es el mismo
                     Set vC = New CCliente
                     vC.codigo = txtCliente(Index + 9).Text
                     vC.DptoCliente txtDpto(Index).Text, campo
                     Set vC = Nothing
                     If campo = "" Then PonerFoco txtDpto(Index)
                 Else
                     'Todavia no ha puesto el cliente
                     MsgBox "Debe poner el cliente debe ", vbExclamation
                     txtDpto(Index).Text = ""
                 End If
            End If
            Me.txtDescDpto(Index).Text = campo
            If campo = "" Then txtDpto(Index).Text = ""
                
            
        
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

'Dado un FRAME lo pone a true y lo situa en x:120 y:0 y devuelve lo que debe medir el form
Private Sub PonerFrameVisible(ByRef F As Frame, ByRef CH As Integer, CW As Integer)
    F.Top = 0
    F.Left = 120
    F.visible = True
    CH = F.Height + 420
    CW = F.Width + 240
End Sub





Private Sub InicializarVbles()
    cadFormula = ""
    cadSelect = ""
    cadParam = ""
    numParam = 0
'    cadTitulo = ""
'    cadNomRPT = ""
    conSubRPT = False
    vMultiInforme = 0
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
        If indD = 27 Or indH = 28 Then Subtipo = "FH"
    Case "CLI"
        'Cliente
        Set TDes = txtCliente(indD)
        Set THas = txtCliente(indH)
        Set DesD = txtDescClie(indD)
        Set DesH = txtDescClie(indH)
        Subtipo = "N"
    Case "DPT"
        'DEpartamento
        Set TDes = txtDpto(indD)
        Set THas = txtDpto(indH)
        Set DesD = txtDescDpto(indD)
        Set DesH = txtDescDpto(indH)
        Subtipo = "N"
        
    Case "ENV"
          
        Set TDes = Me.txtEnvio(indD)
        Set THas = txtEnvio(indH)
        Subtipo = "N"
         
        Set DesD = txtDescEnvio(indD)
        Set DesH = txtDescEnvio(indH)

        
        
    Case "PRO"
        Set TDes = txtCodProve(indD)
        Set THas = txtCodProve(indH)
        Set DesD = txtDescProve(indD)
        Set DesH = txtDescProve(indH)
        Subtipo = "N"
 
    Case "ART"

        Set TDes = txtArticulo(indD)
        Set THas = txtArticulo(indH)
        Set DesD = txtDescArticulo(indD)
        Set DesH = txtDescArticulo(indH)
        Subtipo = "T"
    Case "AGT"
        Set TDes = txtAgente(indD)
        Set THas = txtAgente(indH)
        Set DesD = txtDescAgente(indD)
        Set DesH = txtDescAgente(indH)
        Subtipo = "N"
      
    Case "ALP"
        'Numero albaran proveedores
         
        Set TDes = txtNumAlbar(indD)
        Set THas = txtNumAlbar(indH)
        Subtipo = "T"
 
    Case "TRA"
        'TRABAJADOR
         
        Set TDes = Me.txtTrab(indD)
        Set THas = txtTrab(indH)
        Subtipo = "N"
        If indD = 5 Then
            'llamadas
            Set DesD = txtDescTra(indD)
            Set DesH = txtDescTra(indH)
        End If
        
        
    Case "ZON"
        'ZONA
         
        Set TDes = Me.txtZona(indD)
        Set THas = txtZona(indH)
        Subtipo = "N"
        Set DesD = txtDescZona(indD)
        Set DesH = txtDescZona(indH)
    
        
    Case "FAM"
        'FAMILIA
         
        Set TDes = Me.txtFamia(indD)
        Set THas = txtFamia(indH)
        Subtipo = "N"
        Set DesD = txtDescFamia(indD)
        Set DesH = txtDescFamia(indH)
    
        
    Case "MAR"
    
        Set TDes = Me.txtmarca(indD)
        Set THas = txtmarca(indH)
        Subtipo = "N"
        Set DesD = txtDescmarca(indD)
        Set DesH = txtDescmarca(indH)
        
    Case "ACT"
        'ACTIVIADD
         
        Set TDes = Me.txtcodactiv(indD)
        Set THas = txtcodactiv(indH)
        Subtipo = "N"
        If indD = 5 Then
            'llamadas
            Set DesD = txtDescActiv(indD)
            Set DesH = txtDescActiv(indH)
        End If
    
    Case "FOR"
        'FORMA DE PAGO
         
        Set TDes = Me.txtForpa(indD)
        Set THas = txtForpa(indH)
        Subtipo = "N"
                 
        Set DesD = txtDescForpa(indD)
        Set DesH = txtDescForpa(indH)

    
    
    Case "ALM"
        'Almacen
        
    
        Set TDes = txtAlma(indD)
        Set THas = txtAlma(indH)
        Subtipo = "N"

        Set DesD = txtDescAlma(indD)
        Set DesH = txtDescAlma(indH)

    Case "RUT"
        'Almacen
        
    
        Set TDes = txtRuta(indD)
        Set THas = txtRuta(indH)
        Subtipo = "N"

        Set DesD = txtDescRuta(indD)
        Set DesH = txtDescRuta(indH)




    End Select
    
    devuelve = CadenaDesdeHasta(TDes.Text, THas.Text, campo, Subtipo)
    If devuelve = "Error" Then Exit Function
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
            cadParam = cadParam & AnyadirParametroDH(param, TDes, THas, DesD, DesH) & """|"
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

'Para las reparaciones. Carga el importe real y teorico.
Private Sub CargaImporteRealReparaciones()
'Dim ImpTot As Currency
'Dim ImpTeo As Currency
Dim miSQL As String
Dim RT As ADODB.Recordset

    'A partir de la reparacion , mirare en los albaranes, y de los albaranes ver el coste real de la reparacion y el teorico
    Set miRsAux = New ADODB.Recordset
    Set RT = New ADODB.Recordset
    
    'Meto el select para las
    If cadSelect <> "" Then
        codigo = Replace(cadSelect, "schrep.", "h.")
        codigo = " AND " & codigo
    Else
        codigo = ""
    End If
    

 
    
    'Montamos el select al reves
    
    
'    codigo = "s.codtipom=l.codtipom and s.codtipoa=l.codtipoa " & codigo
'    codigo = "s.fecfactu = l.fecfactu and s.numfactu=l.numfactu AND " & codigo
'    codigo = "s.codtipom=l.codtipom AND " & codigo
'    codigo = "sartic.codartic = l.codartic AND " & codigo
'    'codigo = "select l.*,s.fechaalb,preciove,h.numrepar,h.fecrepar from  schrep h,slifac l,scafac1 s,sartic where  h.numalbar = s.numalbar and h.fechaalb=s.fechaalb  AND " & codigo
'    codigo = "select l.*,s.fechaalb,preciove,h.numrepar,h.fecrepar from  schrep h,slifac l,scafac1 s,sartic where  h.numalbar = s.numalbar and h.fechaalb=s.fechaalb AND h.codtipom= s.codtipoa AND " & codigo
'    codigo = codigo & " ORDER BY s.numalbar ,s.fechaalb"
    
    'AHORA   FEBRERO 2014
    Label3(158).Caption = "Selecionar reparaciones"
    Label3(158).Refresh
    miSQL = "select numrepar,fecrepar,codtipom,numalbar,fechaalb from schrep h where " & Mid(codigo, 5)

    
    
    
    
    
    
    
    
    
    
    
    'EL ORDEN
    
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    codigo = ""
    NumRegElim = 1
    miSQL = ""
    cadTitulo = ""
    While Not miRsAux.EOF
        Label3(158).Caption = miRsAux!numrepar & " " & miRsAux!fecrepar
        Label3(158).Refresh
        
        
        miSQL = "Select sum(importel) tot,sum(round(preciove*cantidad,2)) *100 teo,sum(coalesce(l.preciouc,0.000)*cantidad) coste "
        miSQL = miSQL & " FROM slifac l,scafac1 s,sartic"
        
        miSQL = miSQL & " WHERE s.codtipom=l.codtipom and s.codtipoa=l.codtipoa "
        miSQL = miSQL & " AND s.fecfactu = l.fecfactu and s.numfactu=l.numfactu AND "
        miSQL = miSQL & " s.numalbar=l.numalbar AND sartic.codartic = l.codartic "
        miSQL = miSQL & " AND s.codtipoa=" & DBSet(miRsAux!codtipom, "T")
        miSQL = miSQL & " AND s.numalbar=" & miRsAux!NUmAlbar
        miSQL = miSQL & " AND s.fechaalb=" & DBSet(miRsAux!FechaAlb, "F")
        
        RT.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not RT.EOF Then
            If Not IsNull(RT!Tot) Or Not IsNull(RT!teo) Then
                NumRegElim = NumRegElim + 1
                miSQL = ", (" & vUsu.codigo & "," & miRsAux!numrepar & ",'" & Format(miRsAux!fecrepar, FormatoFecha) & "',0,0,"
                miSQL = miSQL & NumRegElim & "," & Val(DBLet(RT!teo, "N")) & "," & TransformaComasPuntos(CStr(RT!Tot))
                miSQL = miSQL & ",'" & DBLet(RT!coste, "N") & "')"
                cadTitulo = cadTitulo & miSQL
            End If
        Else
           ' S top
        End If
        RT.Close
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close


    If cadTitulo <> "" Then
        'El ultimo
        cadTitulo = Mid(cadTitulo, 2)
        
        miSQL = "insert into tmpnlotes (codusu,codprove,fechaalb,numalbar,nomartic,numlinea,codartic,cantidad,numlotes) VALUES " & cadTitulo
        conn.Execute miSQL
        cadTitulo = ""
    End If

   
    If cadSelect <> "" Then
        codigo = Replace(cadSelect, "schrep.", "h.")
        codigo = " AND " & codigo
    Else
        codigo = ""
    End If
    
    Label3(158).Caption = "Garantias"
    Label3(158).Refresh
    miSQL = "select numrepar,fecrepar,tieneman,h.codclien,m.nummante,"
    miSQL = miSQL & " mes01act,mes02act,mes03act,mes04act,mes05act,mes06act,mes07act,mes08act,mes09act,mes10act,mes11act,mes12act"
    miSQL = miSQL & " from schrep h,sserie s left join scaman m  on s.nummante=m.nummante and s.codclien=m.codclien"
    miSQL = miSQL & " where h.numserie=s.numserie and s.codartic=h.codartic "
    If codigo <> "" Then miSQL = miSQL & codigo
    
    'EL ORDEN
    IndiceImg = 12
    If txtFecha(1).Text <> "" Then IndiceImg = Month(CDate(txtFecha(1).Text))
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    cadTitulo = "" 'Para saber los mantenimientos que ya hemos sumado
    While Not miRsAux.EOF
        Label3(158).Caption = "Garantia " & miRsAux!numrepar & " " & miRsAux!fecrepar
        Label3(158).Refresh
        ImpTot = 0
        If miRsAux!TieneMan = 1 Then
            If IsNull(miRsAux!nummante) Then
                
                miSQL = DevuelveDesdeBD(conAri, "nomclien", "sclien", "codclien", miRsAux!codClien)
                miSQL = "Cliente: " & miRsAux!codClien & "   " & miSQL & vbCrLf
                miSQL = "Error grave." & vbCrLf & "Con mantenimiento y sin numero de mantenimiento" & vbCrLf & miSQL
                miSQL = miSQL & "Reparacion: " & miRsAux!numrepar & " " & miRsAux!fecrepar & vbCrLf
                'MsgBox miSQL, vbExclamation
            End If
            miSQL = Format(miRsAux!codClien, "000000") & DBLet(miRsAux!nummante, "T") & "|"
                
            If InStr(1, cadTitulo, miSQL) = 0 Then
                cadTitulo = cadTitulo & miSQL
                '--------------------------------------------------------------------
                'OK, TIENE MANTENIMIENTO
                'Ire recorriendo los importes desde mes01act hasta el mes hasta
                'Si la fecha es fin es nada, entonces hare tooodos
                For NumRegElim = 1 To IndiceImg
                    If Not IsNull(miRsAux.Fields(NumRegElim + 4)) Then ImpTot = ImpTot + miRsAux.Fields(NumRegElim + 4)
                Next
            Else
               ' St op
            End If
        End If
        If ImpTot <> 0 Then
            'UPDATEAMOS LA tmp
            miSQL = "UPDATE tmpnlotes set nomartic=" & CLng(ImpTot * 100) & " WHERE codusu = " & vUsu.codigo
            miSQL = miSQL & " AND codprove = " & miRsAux!numrepar & " AND fechaalb = '" & Format(miRsAux!fecrepar, FormatoFecha) & "' AND numalbar =0"
            conn.Execute miSQL
        End If
        '--------------
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    Set RT = Nothing
    
End Sub
    


Private Sub EstadisticaReparacionTecnicoNueva()
Dim RT As ADODB.Recordset
Dim OptimizarSelect As String
Dim RAlb As ADODB.Recordset
Dim C As Long
Dim EnAlbaranes As Boolean

    Label3(63).Caption = "Obteniendo reg. albaranes"
    Label3(63).Refresh
    

    'Preparamos las temporales
    codigo = "DELETE FROM tmpinformes WHERE codusu = " & vUsu.codigo
    conn.Execute codigo
    
    codigo = "DELETE FROM tmpnlotes WHERE codusu = " & vUsu.codigo
    conn.Execute codigo
    
    
    
    
    'LOS INSERTS PARA LAS TABLAS temporales                                         numserie
    cadFormula = "insert into tmpinformes (codusu,codigo1,importe1,importe2,nombre1,nombre2) VALUES (" & vUsu.codigo & ","
    cadFrom = "insert into tmpnlotes (codusu,codprove,numalbar,fechaalb,numlinea,nomartic) values (" & vUsu.codigo & ","
    
    
    'Optimizacion del select
    If cadSelect <> "" Then
        'Codigo = Replace(cadSelect, "schrep.", "h.")
        codigo = " WHERE " & cadSelect
    Else
        codigo = ""
    End If
    

    codigo = "Select distinct(codtipom) from schrep  " & codigo
    miRsAux.Open codigo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    OptimizarSelect = ""
    While Not miRsAux.EOF
        codigo = DBLet(miRsAux!codtipom, "T")
        If codigo <> "" Then OptimizarSelect = OptimizarSelect & " OR scafac1.codtipoa = '" & codigo & "'"
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    If OptimizarSelect <> "" Then
        OptimizarSelect = Mid(OptimizarSelect, 4) 'quito el preimer or
        OptimizarSelect = "(" & OptimizarSelect & ")"
    End If
    
    'Cargo el RS con todos los datos de los albarnes
    miSQL = "select scafac1.numalbar,scafac1.fechaalb,scafac1.codtipoa,sum(importel),sum(cantidad*preciove)"
    miSQL = miSQL & " from scafac1,slifac,sartic where  scafac1.codtipom =slifac.codtipom  and scafac1.numfactu  =slifac.numfactu"
    miSQL = miSQL & " and scafac1.fecfactu  =slifac.fecfactu"
    miSQL = miSQL & " and scafac1.codtipoa  =slifac.codtipoa  and scafac1.numalbar  =slifac.numalbar and sartic.codartic=slifac.codartic"
    
    
    cadNomRPT = ""
    If txtFecha(2).Text <> "" Then cadNomRPT = cadNomRPT & " AND fechaalb >='" & Format(txtFecha(2).Text, FormatoFecha) & "'"
    If txtFecha(3).Text <> "" Then cadNomRPT = cadNomRPT & " AND fechaalb <='" & Format(txtFecha(3).Text, FormatoFecha) & "'"
    If OptimizarSelect <> "" Then cadNomRPT = cadNomRPT & " AND " & OptimizarSelect
    

    miSQL = miSQL & cadNomRPT
    miSQL = miSQL & " group by scafac1.numalbar,scafac1.fechaalb,scafac1.codtipoa order by codtipoa,numalbar,fechaalb"

    'Cargamos las sumas en facturas
    miRsAux.Open miSQL, conn, adOpenKeyset, adLockPessimistic, adCmdText
    
    'Cargamos las sumas en albaranaes  ############
    miSQL = "Select scaalb.numalbar,scaalb.fechaalb,scaalb.codtipom,sum(importel),sum(cantidad*preciove)"
    miSQL = miSQL & " from scaalb,slialb,sartic where  scaalb.codtipom =slialb.codtipom  and scaalb.numalbar  =slialb.numalbar"
    miSQL = miSQL & " and sartic.codartic=slialb.codartic"
    cadNomRPT = Replace(cadNomRPT, " scafac1.codtipoa", "scaalb.codtipom")
    miSQL = miSQL & cadNomRPT
    miSQL = miSQL & " group by numalbar,fechaalb,codtipom order by codtipom,numalbar,fechaalb"

    Set RAlb = New ADODB.Recordset
    RAlb.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    
    
    
    
    
    
    
    
    
    
    'Cargamos el rS de la reparaciones
    If cadSelect <> "" Then
        'Codigo = Replace(cadSelect, "schrep.", "h.")
        codigo = " WHERE " & cadSelect
    Else
        codigo = ""
    End If
    

    codigo = " from schrep  " & codigo
    
    Set RT = New ADODB.Recordset
    RT.Open "select count(*)" & codigo, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    C = DBLet(RT.Fields(0), "N")
    RT.Close
    
    RT.Open "select * " & codigo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    codigo = ""
    While Not RT.EOF
        NumRegElim = NumRegElim + 1
        
        Label3(63).Caption = "Rep: " & RT!numrepar & "  (" & NumRegElim & "/" & C & ")"
        Label3(63).Refresh
        
        If IsNull(RT!codtipom) Or IsNull(RT!NUmAlbar) Or IsNull(RT!FechaAlb) Then
            ImpTeo = 0
            ImpTot = 0
        Else
            
            PonerIMportesAlbaranes RAlb, RT!codtipom, RT!NUmAlbar, RT!FechaAlb, ImpTot, ImpTeo, EnAlbaranes
        End If
        
 

            
                'INSERTAMOS
                'en tmpinformes
                codigo = "'" & DevNombreSQL(RT!NomArtic) & "','" & DBLet(RT!numSerie, "T") & "')"
                codigo = NumRegElim & "," & TransformaComasPuntos(CStr(ImpTot)) & "," & TransformaComasPuntos(CStr(ImpTeo)) & "," & codigo
                codigo = cadFormula & codigo
                conn.Execute codigo
                
                
                'en tmpnlotes
                '       numprove
                codigo = DBLet(RT!nummante, "T")
                If codigo <> "" Then
                    codigo = UCase(codigo)
                    'Debug.Print Codigo
                    If codigo = "S/MTO" Or codigo = "SIN ESPC." Then
                        codigo = "0"
                    Else
                        codigo = "1"
                    End If
                Else
                    codigo = "0"
                End If
                codigo = Abs(EnAlbaranes) & ",'" & Format(RT!fecrepar, FormatoFecha) & "'," & codigo & ",'" & Trim(DevNombreSQL(RT!NomClien)) & "')"
                codigo = NumRegElim & "," & codigo
                codigo = cadFrom & codigo
                conn.Execute codigo
                
                 
      
        
        RT.MoveNext
    Wend
    RT.Close
    RAlb.Close
    miRsAux.Close
  Set RT = Nothing
  Set RAlb = Nothing
 

    
    


End Sub



Private Sub PonerIMportesAlbaranes(ByRef RAlba As ADODB.Recordset, codtipom As String, alb As Long, fech As Date, ByRef Impt As Currency, ByRef ImpTeor As Currency, ByRef EnAlbaranes As Boolean)
Dim fin As Boolean
Dim Esta As Boolean

    Impt = 0
    ImpTeor = 0
    
    'Comprobamos en albaranes primero
    EnAlbaranes = True
    Esta = False
    If Not RAlba.EOF Then
        
        fin = False
        While Not fin
            If RAlba!codtipom = codtipom Then
                If RAlba!NUmAlbar = alb Then
                    If RAlba!FechaAlb = fech Then
                        'AQUI ESTA
                        fin = True
                        Esta = True
                        Impt = RAlba.Fields(3)
                        ImpTeor = RAlba.Fields(4)
                    End If
                    
                Else
                    'SI ha sobrepasado YA no esta
                    If RAlba!NUmAlbar > alb Then fin = True
                End If
            End If
            RAlba.MoveNext
            If RAlba.EOF Then fin = True
        Wend
    
        RAlba.MoveFirst
        If Esta Then Exit Sub  'Ya lo hemos encontrado
    End If
    
    
    EnAlbaranes = False
    If miRsAux.EOF Then Exit Sub
    fin = False
    While Not fin
        If miRsAux!codtipoa = codtipom Then
            If miRsAux!NUmAlbar = alb Then
                If miRsAux!FechaAlb = fech Then
                    'AQUI ESTA
                    fin = True
                    Impt = miRsAux.Fields(3)
                    ImpTeor = miRsAux.Fields(4)
                End If
            Else
                If miRsAux!NUmAlbar > alb Then miRsAux.MoveLast
            End If
        End If
        miRsAux.MoveNext
        If miRsAux.EOF Then fin = True
    Wend
    miRsAux.MoveFirst
End Sub

Private Sub EstadisticaReparacionTecnico()
    'Preparamos las temporales
    codigo = "DELETE FROM tmpinformes WHERE codusu = " & vUsu.codigo
    conn.Execute codigo
    
    codigo = "DELETE FROM tmpnlotes WHERE codusu = " & vUsu.codigo
    conn.Execute codigo
    
    

    
    'LOS INSERTS PARA LAS TABLAS temporales                                         numserie
    cadFormula = "insert into tmpinformes (codusu,codigo1,importe1,importe2,nombre1,nombre2) VALUES (" & vUsu.codigo & ","
    cadFrom = "insert into tmpnlotes (codusu,codprove,numalbar,fechaalb,numlinea,nomartic) values (" & vUsu.codigo & ","
    
    'Montamos el select al reves
    'PARA LAS FACTURAS
    If cadSelect <> "" Then
        codigo = Replace(cadSelect, "schrep.", "h.")
        codigo = " AND " & codigo
    Else
        codigo = ""
    End If
    codigo = " s.codtipom=l.codtipom and s.codtipoa=l.codtipoa " & codigo
    codigo = " s.fecfactu = l.fecfactu and s.numfactu=l.numfactu AND " & codigo
    codigo = " s.codtipom=l.codtipom AND " & codigo
    codigo = " sartic.codartic = l.codartic AND " & codigo
    codigo = " h.numserie=sserie.numserie AND h.codclien=sserie.codclien AND " & codigo
    codigo = " sclien.codclien = h.codclien AND " & codigo
    codigo = " where  h.numalbar = s.numalbar and h.fechaalb=s.fechaalb AND " & codigo
    'Las tablas
    codigo = " from schrep h,slifac l,scafac1 s,sclien , sserie,sartic" & codigo
    codigo = "select l.*,s.fechaalb,preciove,h.fecrepar,h.nomclien,tieneman,h.nomartic,h.numserie " & codigo
    'EL ORDEN
    codigo = codigo & " ORDER BY s.numalbar ,s.fechaalb"
    miRsAux.Open codigo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    codigo = ""
    miSQL = ""
    While Not miRsAux.EOF
    

    
        If codigo <> CStr(miRsAux!NUmAlbar & "|" & Format(miRsAux!FechaAlb, FormatoFecha)) Then
            If codigo <> "" Then
                NumRegElim = NumRegElim + 1
                'INSERTAMOS
                'en tmpinformes
                codigo = RecuperaValor(miSQL, 1)
                codigo = NumRegElim & "," & TransformaComasPuntos(CStr(ImpTot)) & "," & TransformaComasPuntos(CStr(ImpTeo)) & "," & codigo
                codigo = cadFormula & codigo
                conn.Execute codigo
                
                
                'en tmpnlotes
                '       numprove
                codigo = RecuperaValor(miSQL, 2)
                codigo = NumRegElim & "," & codigo
                codigo = cadFrom & codigo
                conn.Execute codigo
                
                
            End If
            'Meto dos datos enpipados
            miSQL = "'" & DevNombreSQL(miRsAux!NomArtic) & "','" & miRsAux!numSerie & "')"
            miSQL = miSQL & "|"
            miSQL = miSQL & "0,'" & Format(miRsAux!fecrepar, FormatoFecha) & "'," & miRsAux!TieneMan & ",'" & DevNombreSQL(miRsAux!NomClien) & "')|"
            codigo = miRsAux!NUmAlbar & "|" & Format(miRsAux!FechaAlb, FormatoFecha)
            ImpTot = 0
            ImpTeo = 0
        End If
        ImpTeo = ImpTeo + (miRsAux!PrecioVe * miRsAux!cantidad)
        ImpTot = miRsAux!ImporteL + ImpTot
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    If miSQL <> "" Then
        NumRegElim = NumRegElim + 1
        
        'en tmpinformes
        codigo = RecuperaValor(miSQL, 1)
        codigo = NumRegElim & "," & TransformaComasPuntos(CStr(ImpTot)) & "," & TransformaComasPuntos(CStr(ImpTeo)) & "," & codigo
        codigo = cadFormula & codigo
        conn.Execute codigo
        
        
        'en tmpnlotes
        '       numprove
        codigo = RecuperaValor(miSQL, 2)
        codigo = NumRegElim & "," & codigo
        codigo = cadFrom & codigo
        conn.Execute codigo
        
        
    End If
    

    
    'AHORA HAGO EL INSERT PARA LOS ALBARANES QUE NO HAN SIDO FACTURADOS
    'PARA LOS ALBARANES
    If cadSelect <> "" Then
        codigo = Replace(cadSelect, "schrep.", "h.")
        codigo = " AND " & codigo
    Else
        codigo = ""
    End If

    miSQL = "select l.*,preciove,tieneman,h.fechaalb,h.numserie,h.nomclien,fecrepar"
    miSQL = miSQL & " from schrep h,scaalb c,slialb l,sartic a,sserie s "
    miSQL = miSQL & " WHERE h.codtipom=c.codtipom and h.numalbar=c.numalbar and h.fechaalb=c.fechaalb and"
    miSQL = miSQL & " l.numalbar=c.numalbar and l.codtipom=c.codtipom and l.codartic=a.codartic"
    miSQL = miSQL & " and h.numserie=s.numserie and h.codclien =s.codclien" & codigo
    
    
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    codigo = ""
    miSQL = ""
    While Not miRsAux.EOF
        If codigo <> CStr(miRsAux!NUmAlbar & "|" & Format(miRsAux!FechaAlb, FormatoFecha)) Then
            If codigo <> "" Then
                NumRegElim = NumRegElim + 1
                'INSERTAMOS
                'en tmpinformes
                codigo = RecuperaValor(miSQL, 1)
                codigo = NumRegElim & "," & TransformaComasPuntos(CStr(ImpTot)) & "," & TransformaComasPuntos(CStr(ImpTeo)) & "," & codigo
                codigo = cadFormula & codigo
                conn.Execute codigo
                
                
                'en tmpnlotes
                '       numprove
                codigo = RecuperaValor(miSQL, 2)
                codigo = NumRegElim & "," & codigo
                codigo = cadFrom & codigo
                conn.Execute codigo
                
            End If
            'Meto dos datos enpipados
            miSQL = "'" & DevNombreSQL(miRsAux!NomArtic) & "','" & miRsAux!numSerie & "')"
            miSQL = miSQL & "|"
            miSQL = miSQL & "1,'" & Format(miRsAux!fecrepar, FormatoFecha) & "'," & miRsAux!TieneMan & ",'" & DevNombreSQL(miRsAux!NomClien) & "')|"
            codigo = miRsAux!NUmAlbar & "|" & Format(miRsAux!FechaAlb, FormatoFecha)
            ImpTot = 0
            ImpTeo = 0
        End If
        ImpTeo = ImpTeo + (miRsAux!PrecioVe * miRsAux!cantidad)
        ImpTot = miRsAux!ImporteL + ImpTot
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    If miSQL <> "" Then
        NumRegElim = NumRegElim + 1
        
        'en tmpinformes
        codigo = RecuperaValor(miSQL, 1)
        codigo = NumRegElim & "," & TransformaComasPuntos(CStr(ImpTot)) & "," & TransformaComasPuntos(CStr(ImpTeo)) & "," & codigo
        codigo = cadFormula & codigo
        conn.Execute codigo
        
        
        'en tmpnlotes
        '       numprove
        codigo = RecuperaValor(miSQL, 2)
        codigo = NumRegElim & "," & codigo
        codigo = cadFrom & codigo
        conn.Execute codigo
        
        
    End If
    


End Sub










'------------------------------------------------------------------
'------------------------------------------------------------------
'               M U L T I B A S E
'------------------------------------------------------------------
Private Sub CargaListMultibase()
    Me.lstMultibase.Clear
    miSQL = "Clientes|Proveedores|Trabajadores|Direcciones|"
    For numParam = 1 To 4
        Me.lstMultibase.AddItem RecuperaValor(miSQL, CInt(numParam))
    Next numParam
    'Como organiza informacion
    '         tabla  clave    campos a cambiar(empieza con coma) tipodatos clave.
    'Clientes
    miSQL = "sclien:codclien:,nomclien,nomcomer ,domclien ,codpobla ,pobclien,perclie1,perclie2:N|"
    miSQL = miSQL & "sprove:codprove:,nomprove,nomcomer ,domprove ,codpobla ,pobprove ,perprov1 ,perprov2:N|"
    miSQL = miSQL & "straba:codtraba:,nomtraba,domtraba,codpobla,pobtraba:N|"
    miSQL = miSQL & "sdirec:codclien,coddirec:,nomdirec ,domdirec ,pobdirec ,prodirec ,perdirec:N,N|"
        
End Sub


Private Sub HacerCambiosMultibase(numlinea As Integer)
Dim TotalReg As Long
Dim i As Integer
Dim J As Integer
Dim Claves As Integer
Dim Campos As Integer
Dim Cambios As Long
Dim T1 As Single
'Reutilizacion de variables
'cadTitulo cadNomRPT  conSubRPT

    On Error GoTo EHacerCambiosMultibase
    campo = lstMultibase.List(numlinea - 1)
    lblMultibase.Caption = "Preparando datos: " & campo
    
    lblMultibase.Refresh

    cadFormula = RecuperaValor(miSQL, numlinea)
    cadFormula = Replace(cadFormula, ":", "|")
    cadFormula = cadFormula & "|"  'Le añado el pipe final
    'Primero el conteo
    cadParam = "Select count(*) from " & RecuperaValor(cadFormula, 1)
    miRsAux.Open cadParam, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    TotalReg = 0
    If Not miRsAux.EOF Then
        If Not IsNull(miRsAux.Fields(0)) Then TotalReg = miRsAux.Fields(0)
    End If
    miRsAux.Close
    
    DoEvents
    If TotalReg = 0 Then
        lblMultibase.Caption = "Tabla vacia " & campo
        lblMultibase.Refresh
        Espera 1
    End If
    
    'Veamos cuantos campos hay que ver la conversion de campos, y las claves
    cadParam = RecuperaValor(cadFormula, 2)
    Claves = 1
    Cambios = 0
    While cadParam <> ""
        NumRegElim = InStr(1, cadParam, ",")
        If NumRegElim = 0 Then
            cadParam = ""
        Else
            Claves = Claves + 1
            cadParam = Mid(cadParam, NumRegElim + 1)
        End If
    Wend
    cadParam = RecuperaValor(cadFormula, 3)
    Campos = 0 'aqui cero pq empieza con la coma
    While cadParam <> ""
        NumRegElim = InStr(1, cadParam, ",")
        If NumRegElim = 0 Then
            cadParam = ""
        Else
            Campos = Campos + 1
            cadParam = Mid(cadParam, NumRegElim + 1)
        End If
    Wend
        

                            'claves                                 'campos cambiar
    cadParam = "SELECT " & RecuperaValor(cadFormula, 2) & RecuperaValor(cadFormula, 3)
    cadParam = cadParam & " FROM " & RecuperaValor(cadFormula, 1)
    miRsAux.Open cadParam, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    Cambios = 0
    T1 = Timer   'Para hacer doevents cada 3 segundos
    While Not miRsAux.EOF
        NumRegElim = NumRegElim + 1
        'Los labels
        lblMultibase.Caption = campo & " ( " & NumRegElim & " / " & TotalReg & " )"
        lblMultibase.Refresh
        If Timer - T1 > 3 Then
            DoEvents
            Me.Refresh
            Espera 0.2
            T1 = Timer
        End If
        
        cadSelect = "" 'LOS UPDATES
        For i = Claves To Campos
            If Not IsNull(miRsAux.Fields(i)) Then
                cadParam = miRsAux.Fields(i)  'Cojo el valor del field
                cadNomRPT = RevisaCaracterMultibase(cadParam)  'Obtengo la modificaicon por campos multibase
                If cadParam <> cadNomRPT Then
                    'HAY que modificar ya que son disitintos el de laBD y el calculado por el modulo de multibase
                    cadSelect = cadSelect & ", " & miRsAux.Fields(i).Name & " = '" & DevNombreSQL(cadNomRPT) & "'"
                End If
            End If
        Next
        'SI cadselect <>"" entonces hay que ejecutar SQL
        If cadSelect <> "" Then
            'Los campos claves van del 0 a claves -1
            cadParam = ""
            cadTitulo = RecuperaValor(cadFormula, 4) 'los tipos de datos
            cadTitulo = Replace(cadTitulo, ",", "|") & "|"
            For J = 0 To Claves - 1
                cadParam = cadParam & " AND " & miRsAux.Fields(J).Name & " = "
                codigo = RecuperaValor(cadTitulo, J + 1)

                Select Case codigo
                Case "F"
                    cadParam = cadParam & "'" & Format(miRsAux.Fields(i).Value, FormatoFecha) & "'"
                Case "T"
                    cadParam = cadParam & "'" & miRsAux.Fields(i).Value & "'"
                Case Else  'NUMERICO
                    cadParam = cadParam & miRsAux.Fields(J).Value
                End Select
            Next J
            
            
            'Acabas de montar el UPDATE
            cadTitulo = "UPDATE " & RecuperaValor(cadFormula, 1)
            cadSelect = Mid(cadSelect, 2)   'QUITO la coma
            cadParam = Mid(cadParam, 5)     'QUITO el primer AND
            cadTitulo = cadTitulo & " SET " & cadSelect & " WHERE " & cadParam
            conn.Execute cadTitulo
            Cambios = Cambios + 1
        End If
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    lblMultibase.Caption = "FIN " & campo
    lblMultibase.Refresh
    If Cambios > 0 Then Me.Tag = Me.Tag & vbCrLf & "   .- " & campo & " : " & Cambios
    Exit Sub
EHacerCambiosMultibase:
    MuestraError Err.Number
End Sub
'       fin mULTIBASE
'------------------------------------------------------------------'------------------------------------------------------------------


'------------------------------------------------------------------
'------------------------------------------------------------------
'
'       Facturacion de recargas de telefonia
'
'------------------------------------------------------------------
'------------------------------------------------------------------



Private Sub HacerFacturacionTelefonia(vAlbaranes As Collection, MenError As String)
Dim RT As ADODB.Recordset
Dim b As Boolean
Dim NumAlb As String
Dim Almacen As Integer



    'El proceso sera el siguiente:
    'Voy a agrupar por dia (podria ser por mes),trabajador
    'Y para cada uno de los resultados del recodset voy a generar un albaran.
    'Me guardare los albaranes generados y despues los facturare.
    'Para ello
    campo = "Select codtraba,count(*) as cantidad,sum(importe)as total from stelefonia WHERE " & cadSelect & " GROUP by codtraba"
    
    Set RT = New ADODB.Recordset
    RT.Open campo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not RT.EOF
        
        Almacen = DevuelveDesdeBD(conAri, "codalmac", "straba", "codtraba", CStr(RT!CodTraba), "N")
        
        conn.BeginTrans
        
        'Obtener el contador de Albaran (ALV).
        b = ObtenerContadorAlbaran(NumAlb)
        
        If b Then
            'Actualizar los stocks de todos los articulos comprados
            'Insertar movimiento en smoval
            'B = InsertarMovAlmacen(NumAlb)  ¿ FALTA### ?
    
            'Insertar en las tablas de Albaranes: scaalb, slialb
            'en el campo scafac1.numalbar guardamos el nº de ticket
            If b Then b = InsertarAlbaran(NumAlb, CStr(RT!CodTraba), 1, RT!cantidad, RT!total, MenError)
        
        End If



       
        If Not b Then
            conn.RollbackTrans
            RT.Close
            Set RT = Nothing
            Exit Sub
        Else
            vAlbaranes.Add CStr(NumAlb)
            conn.CommitTrans
            
            'Le pongo a facturado en la telefonia
            miSQL = "UPDATE stelefonia SET facturado = 1 WHERE " & cadSelect & " AND codtraba = " & RT!CodTraba
            conn.Execute miSQL
        End If
    
    
        RT.MoveNext
    Wend
    RT.Close
    


End Sub


Private Function ObtenerContadorAlbaran(NumAlb As String) As Boolean
Dim vTipoMov As CTiposMov
Dim Existe As Boolean

    On Error GoTo ErrConAlb

    Set vTipoMov = New CTiposMov
    If vTipoMov.Leer("ALV") Then
        Do
            NumAlb = vTipoMov.ConseguirContador("ALV")
            vTipoMov.IncrementarContador ("ALV")
            miSQL = "select count(*) from scaalb where codtipom='ALV' and numalbar=" & NumAlb
            Existe = (RegistrosAListar(miSQL) > 0)
        Loop Until Existe = False
        ObtenerContadorAlbaran = True
    Else
        ObtenerContadorAlbaran = False
    End If
    Set vTipoMov = Nothing
    Exit Function
    
ErrConAlb:
    ObtenerContadorAlbaran = False
    MuestraError Err.Number, "Obtener contador albaran", Err.Description
End Function

Private Function InsertarAlbaran(NumAlb As String, CodTraba As String, CodAlmc As Integer, cantidad As Currency, Importe As Currency, menErr As String) As Boolean
Dim b As Boolean
Dim vClien As CCliente
Dim SQL As String

    On Error GoTo EInsAlb



    'Cabecera de albaran
    '----------------------------------
    SQL = "INSERT INTO scaalb (codtipom,numalbar,fechaalb,factursn,codclien,nomclien,domclien,codpobla,pobclien,proclien,nifclien,telclien,"
    SQL = SQL & "coddirec,nomdirec,referenc,codtraba,codtrab1,codtrab2,codagent,codforpa,codenvio,dtoppago,dtognral,tipofact,"
    SQL = SQL & "observa01,observa02,observa03,observa04,observa05,numofert,fecofert,numpedcl,fecpedcl,fecentre,sementre,esticket,numtermi,numventa) "
                                                                    'Facturar   cliente
    SQL = SQL & " VALUES ('ALV'," & NumAlb & "," & DBSet(Now, "F") & ",1," & txtCliente(2).Text & ","
    
    'Obtenemos los datos del cliente
    Set vClien = New CCliente
    If vClien.Existe(txtCliente(2).Text) Then
        If vClien.LeerDatos(txtCliente(2).Text) Then
            SQL = SQL & DBSet(vClien.Nombre, "T", "N") & ", " & DBSet(vClien.Domicilio, "T", "N") & ","
            SQL = SQL & DBSet(vClien.CPostal, "T", "N") & ", " & DBSet(vClien.Poblacion, "T", "N") & "," & DBSet(vClien.Provincia, "T", "N") & ","
            SQL = SQL & DBSet(vClien.NIF, "T", "N") & "," & DBSet(vClien.TfnoClien, "T") & ","
            'coddirec,nomdirec,referenc a nulo
            SQL = SQL & "NULL,NULL,NULL,"
            
            SQL = SQL & CodTraba & "," & CodTraba & "," & CodTraba & "," 'trabajador
            '                              cod forpa
            SQL = SQL & vClien.Agente & ",1," & vClien.FEnvio & ",0,0," & vClien.TipoFactu & ","
            SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," 'observaciones
            SQL = SQL & ValorNulo & "," & ValorNulo & "," 'datos oferta: aqui guardamos nº venta
            'En los campos de datos del pedido guardamos los datos del ticket
            'SQL = SQL & NumTicket & "," & DBSet(RSVenta!fecventa, "F") & "," & ValorNulo & "," & ValorNulo & ",1," & DBSet(RSVenta!NumTermi, "N") & "," & DBSet(RSVenta!NumVenta, "N", "S") & ")" 'esticket=1, terminal
            SQL = SQL & "NULL,NULL," & ValorNulo & "," & ValorNulo & ",0,NULL,NULL)"
            b = vClien.ActualizaUltFecMovim(Now)
        Else
            b = False
        End If
    End If
    Set vClien = Nothing
    
    
    If b Then
        'Insertar Cabecera
'    MenError = "Error al insertar en la tabla Cabecera de Albaranes (scaalb )."
        conn.Execute SQL, , adCmdText
        
        'Lineas del albaran
        'Inserta en tabla "slialb" todas las lineas de venta
        SQL = "INSERT INTO slialb "
        SQL = SQL & "(codtipom, numalbar,numlinea, codalmac, codartic, nomartic, ampliaci, cantidad, precioar, "
        SQL = SQL & "dtoline1, dtoline2, importel, origpre) VALUES ("
        SQL = SQL & "'ALV'," & DBSet(NumAlb, "N") & ",1," & CodAlmc & ",'" & DevNombreSQL(txtArticulo(0).Text) & "','" & DevNombreSQL(txtDescArticulo(0).Text)
        SQL = SQL & "',NULL," & cantidad & "," & TransformaComasPuntos(CStr(Round(Importe / cantidad, 4))) & ",0,0," & TransformaComasPuntos(CStr(Importe)) & ",'')"
        'SQL = SQL & " FROM sliven WHERE " & Replace(cadSel, "scaven", "sliven")
        conn.Execute SQL, , adCmdText
    End If


    
    'Guardamos los valores identificativos de la factura generada
    'para imprimirla posteriormente
    If b Then cadImpresion = "{scaalb.codtipom}='ALV' and {scaalb.numalbar}=" & DBSet(NumAlb, "N")

EInsAlb:
    If Err.Number <> 0 Then
        menErr = "Insertando el Albaran: " & vbCrLf & Err.Description
        b = False
    End If
    InsertarAlbaran = b
End Function



'De momento no miro si tiene DTOs o no. Simplemente acltualizo precio y redondeo
'a dos decimales
Private Function RealizarCambiosPreciosLiq(ByRef FechaUltCompra As Date) As Boolean


    On Error GoTo ERealizarCambiosPreciosLiq
    RealizarCambiosPreciosLiq = False
    
    cadFormula = "UPDATE slialp Set precioar = " & TransformaComasPuntos(CStr(ImpTeo)) & " , importel = "
    cadTitulo = "UPDATE smoval SET impormov = "
    miRsAux.Open cadFrom, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        'Label
        Me.lblLiqu.Caption = miRsAux!NUmAlbar & " - " & miRsAux!FechaAlb & " : " & miRsAux!codArtic
        Me.lblLiqu.Refresh
        
        ImpTot = miRsAux!cantidad * ImpTeo
        ImpTot = Round2(ImpTot, 2)
        devuelve = TransformaComasPuntos(CStr(ImpTot)) & " WHERE numalbar = '" & DevNombreSQL(miRsAux!NUmAlbar) & "'"
        devuelve = devuelve & " And fechaalb = '" & Format(miRsAux!FechaAlb, FormatoFecha) & "' AND"
        devuelve = devuelve & " codprove = " & miRsAux!Codprove
        devuelve = devuelve & " AND numlinea = " & miRsAux!numlinea
        devuelve = cadFormula & devuelve
        conn.Execute devuelve
        
        'UPDATEO smoval
        devuelve = cadTitulo & TransformaComasPuntos(CStr(ImpTot))
        devuelve = devuelve & " WHERE detamovi = 'ALC' AND fechamov = '" & Format(miRsAux!FechaAlb, FormatoFecha) & "'"
        devuelve = devuelve & " AND codigope = " & miRsAux!Codprove & " AND document = '" & DevNombreSQL(miRsAux!NUmAlbar) & "'"
        devuelve = devuelve & " AND codartic = '" & DevNombreSQL(miRsAux!codArtic) & "' AND numlinea =" & miRsAux!numlinea
        conn.Execute devuelve
        
        'Si el albaran es masyor que la utlima fecha de compra entonces
        If miRsAux!FechaAlb > FechaUltCompra Then
            numParam = 1
            FechaUltCompra = miRsAux!FechaAlb
        End If
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close



    RealizarCambiosPreciosLiq = True
    Exit Function
ERealizarCambiosPreciosLiq:
    MuestraError Err.Number, Err.Description
End Function






'------------------------------------------------------------------
'------------------------------------------------------------------
'
'       Facturacion y contabilizacion de tickets
'       ========================================
'
'
'
'
'       Cuando esta la marca de contabilizar tickets agrupados, lo que haremos sera
'       a partir de los FTI crear las facturas agrupadas con el contador FTG "EN LA CONTABILIDAD"
'       en el ariges, en scafac, no creo ninguna factura
'       O bien una diaria o una mensual (dependera del parametro)
'
'
'       Insertaremos en una tabla los tckets que entran en cada factura
'----------------------------------------------------------------------
'----------------------------------------------------------------------
Private Sub HacerFacturaTICKETS()
Dim Cliente As Long
Dim b As Boolean
 
    
        'Si va agrupado por fecha, o no
        Label5.Caption = "Obteniendo facturas"
        Label5.Refresh
        
        
        '------------------------------------------------------------------------------
        '  LOG de acciones
        Set LOG = New cLOG
        miSQL = "Desde " & txtFecha(20).Text & " hasta " & txtFecha(21).Text & vbCrLf
        miSQL = miSQL & "Diario: " & CStr(Me.optTick(0).Value) & vbCrLf
        miSQL = miSQL & "Trabajador: " & txtTrab(2).Text & " " & Me.txtDescTra(2).Text
        LOG.Insertar 6, vUsu, miSQL
        miSQL = ""
        Set LOG = Nothing
        '-----------------------------------------------------------------------------
        
        
        
        
        If Me.optTick(1).Value Then
            'MENSUAL
            'JUNIO 2010
            'CODCLIEN
            miSQL = "Select codclien from scafac WHERE " & cadSelect & " GROUP by codclien ORDER BY codclien "
            miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            campo = ""
            
            While Not miRsAux.EOF
                campo = campo & miRsAux!codClien & "|"
                miRsAux.MoveNext
            Wend
            miRsAux.Close
            If campo = "" Then
                MsgBox "Error agrupando por cliente", vbExclamation
            Else
                While campo <> ""
                    numParam = InStr(1, campo, "|")
                    If numParam = 0 Then
                        campo = ""
                    Else
                        devuelve = Mid(campo, 1, numParam - 1)
                        Cliente = Val(devuelve)
                        campo = Mid(campo, numParam + 1)
                        devuelve = Format(txtFecha(21).Text, FormatoFecha)
                        b = ObtenerDatosTickets2(False, Cliente)
                        If Not b Then campo = "" 'para que salga
                    End If
                Wend
            End If
            
        Else
            'Veo las fechas
            'Y para cada fecha y cliente
            miSQL = "Select fecfactu,codclien from scafac WHERE " & cadSelect & " GROUP by fecfactu,codclien ORDER BY fecfactu,codclien "
            miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            campo = ""
            'JUNIO 2010
            'Por fecha Y CODCLIEN.
            While Not miRsAux.EOF
                'Los 12 primeros son para el codclien. Los siguientes para la fecha
                campo = campo & Mid(miRsAux!codClien & "            ", 1, 12) & Format(miRsAux.Fields!FecFactu, FormatoFecha) & "|"
                miRsAux.MoveNext
            Wend
            miRsAux.Close
            
            'Ya tengo todas las fechas que voy a tratar
            While campo <> ""
                   numParam = InStr(1, campo, "|")
                   If numParam = 0 Then
                        campo = ""
                   Else
                        'El cliente
                        Cliente = Val(Mid(campo, 1, 12))
                                          
                        devuelve = Mid(campo, 13, numParam - 13)
                        
                        campo = Mid(campo, numParam + 1)
                    
                        Label5.Caption = "Obteniendo facturas. Fec: " & devuelve & "  Cli: " & Cliente
                        Label5.Refresh
                    
                    
                        'CONTABILIZAMOS LA FACTURA ESA
                        b = ObtenerDatosTickets2(True, Cliente)
                        'Se ha producido un error.Salgo aunaque queden fecs por tratar
                        If Not b Then campo = ""
                            
                   End If
            Wend
            
        End If
        Set miRsAux = Nothing
            
            
        If b Then
            'AHORA LANZAREMOS A CONTABILIZAR FACTURAS de frmlistado
            Label5.Caption = "Contablizando FTGs"
            Label5.Refresh
            AbrirListado 248   'Contabilizacion de facturas TICKET AGRUPADAS
            
            
            Label5.Caption = "Comprobando contabilizacion"
            Label5.Refresh
            DoEvents
            
            
            'Aqui viene la fiesta. Vere si hay facturas FTG con intconta=0
            'Significara que han dado error al entrar en la conta
            Set miRsAux = New ADODB.Recordset
            miSQL = "Select numfactu from scafac where codtipom='FTG' And intconta=0"
            miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not miRsAux.EOF Then
                If Not IsNull(miRsAux!Numfactu) Then b = False
            End If
            miRsAux.Close
            Set miRsAux = Nothing
                       
            
        End If

            
        If Not b Then
            Screen.MousePointer = vbHourglass
            Label5.Caption = "Reestableciendo FTI. Paso 1"
            Label5.Refresh
            'HA IDO MAL
            'Vuelvo a poner los FTI que haya puesto a contabilizado, a 0
            
            
            'Dos pasos:
            'Primero ver que facturas FTG se han generado.
            'Las meto en la variable cadfrom
            
            Set miRsAux = New ADODB.Recordset
            miSQL = "Select numfactu,fecfactu from scafac where codtipom='FTG' And intconta=0"
            miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            cadFrom = ""
            While Not miRsAux.EOF
                cadFrom = cadFrom & " numfacftg =  " & miRsAux!Numfactu & " AND fecfacftg = '" & Format(miRsAux!FecFactu, FormatoFecha) & "'|"
                miRsAux.MoveNext
            Wend
            miRsAux.Close
            
            
            'Segundo.
            'Para cad factura FTG generada veo que FTI asoacidos tiene y los updateo
            Label5.Caption = "Reestableciendo FTI. Paso 2"
            Label5.Refresh
            miSQL = "UPDATE scafac SET intconta=0 WHERE codtipom='FTI' AND numfactu ="
            While cadFrom <> ""
                numParam = InStr(1, cadFrom, "|")
                If numParam = 0 Then
                    cadFrom = ""
                Else
                    devuelve = Mid(cadFrom, 1, numParam - 1)
                    cadFrom = Mid(cadFrom, numParam + 1)
                         
                    devuelve = "Select numfactu,fecfactu FROM sfactik where " & devuelve
                    miRsAux.Open devuelve, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                    While Not miRsAux.EOF
                
                        devuelve = miSQL & miRsAux!Numfactu & " AND fecfactu = '" & Format(miRsAux!FecFactu, FormatoFecha) & "'"
                        conn.Execute devuelve
                        miRsAux.MoveNext
                    Wend
                    miRsAux.Close
                End If
            Wend
       
            
                    
            Me.Refresh
            Espera 0.5
            Label5.Caption = "Eliminado asociaciones"
            Label5.Refresh
            
            'Si ha ido mal entonces borraremos tanto los FTG (proceso que se hace despues)
            'como en la tabla que asocia con los tickets
            ' REestablecer en contadores
            ' devuelve= MINIMO
            miSQL = "Select numfactu,fecfactu from scafac where codtipom='FTG' AND intconta=0"
            miSQL = miSQL & " GROUP BY numfactu,fecfactu ORDER BY numfactu"
            miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            devuelve = ""
            While Not miRsAux.EOF
                If devuelve = "" Then devuelve = miRsAux!Numfactu & "|'" & Format(miRsAux!FecFactu, FormatoFecha) & "|"
                miSQL = "DELETE from sfactik WHERE numfacftg=" & miRsAux!Numfactu & " AND fecfacFTG='" & Format(miRsAux!FecFactu, FormatoFecha) & "'"
                conn.Execute miSQL
                miRsAux.MoveNext
            Wend
            miRsAux.Close
            
            
            'Pong los contadores como estaban
            If devuelve <> "" Then
                NumRegElim = Val(RecuperaValor(devuelve, 1))
                'devuelve = RecuperaValor(devuelve, 1)
                miSQL = "UPDATE stipom SET contador = " & NumRegElim & " WHERE codtipom='FTG'"
                conn.Execute miSQL
            End If
        End If 'De si ha ido bien o mal
        
        'BORRAMOS todos los datos
        '------------------------------------------
        Label5.Caption = "Eliminando datos temporales de tablas scafac..."
        Label5.Refresh
        DoEvents
        
        miSQL = "DELETE  from slifac where codtipom='FTG'"
        conn.Execute miSQL
        
        

        'Habre metido una linea en scafac1
        miSQL = "DELETE  from scafac1 where codtipom='FTG'"
        conn.Execute miSQL

        
        miSQL = "DELETE  from scafac where codtipom='FTG'"
        conn.Execute miSQL
        
        
        
        'Si todo ha ido bien muestro un msg
        Label5.Caption = ""
        Label5.Refresh
        If b Then MsgBox "Proceso finalizado con éxito", vbInformation
        
         Screen.MousePointer = vbDefault
End Sub





Private Function ObtenerDatosTickets2(Diario As Boolean, Cliente As Long) As Boolean
Dim TiposIva As Byte
Dim vCl As CCliente
Dim vTM As CTiposMov

        

        'NUEVO JUNIO 2010
        'Agruparemos por mes o dia... Y CLIENTE!!!!
        Set vCl = New CCliente
        If Not vCl.LeerDatos(CStr(Cliente)) Then Exit Function
        
        On Error GoTo EObteniendoDatosTickets


        ObtenerDatosTickets2 = False



        'En la tabla tmpspla pondre todos los importes por tp iva
        conn.Execute "DELETE from tmpinformes where codusu = " & vUsu.codigo
        
        
        'Veo todos los importes y bases imponibles etc
        'Para no tener que hacer selects y demas me guardare que tipos de iva estoy tratatando
        '
        cadNomRPT = "|"
        TiposIva = 0
        For numParam = 1 To 3
            miSQL = "SELECT codigiv" & numParam & " tipodeiva,sum(baseimp" & numParam & ") labase,sum(imporiv" & numParam & ") importeiva FROM SCafac where "
            miSQL = miSQL & " intconta=0 and codtipom='FTI' "
            If Diario Then
                miSQL = miSQL & " AND fecfactu='" & devuelve & "' AND codclien = " & Cliente
            Else
                'MOdificacion 13 - Agosto - 2008
                'Si no pongo esto suma tooooodas las facturas FTI que no esten contabilizadas
                'Desde
                If txtFecha(20).Text <> "" Then miSQL = miSQL & " AND fecfactu>='" & Format(txtFecha(20).Text, FormatoFecha) & "'"
                'El campo HASTA es obligado
                miSQL = miSQL & " AND fecfactu<='" & Format(txtFecha(21).Text, FormatoFecha) & "'"
                miSQL = miSQL & " AND codclien = " & Cliente
            End If
            
            miSQL = miSQL & " group by 1 ORDER by tipodeiva"  'primero
            miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not miRsAux.EOF
                
                If Not IsNull(miRsAux!tipodeiva) Then
                    ImpTot = DBLet(miRsAux!labase, "N")
                    ImpTeo = DBLet(miRsAux!ImporteIva, "N")
                    miSQL = "|" & miRsAux!tipodeiva & "|"
                    
                    If InStr(1, cadNomRPT, miSQL) > 0 Then
                        'YA LO HE INSERTADO. UPDATEO
                        miSQL = "UPDATE tmpinformes SET importe1=importe1 + " & TransformaComasPuntos(CStr(ImpTot))
                        miSQL = miSQL & " ,importe2=importe2 + " & TransformaComasPuntos(CStr(ImpTeo))
                        miSQL = miSQL & " WHERE codusu = " & vUsu.codigo & " AND codigo1 = " & miRsAux!tipodeiva
                    Else
                        miSQL = "INSERT INTO `tmpinformes` (`codusu`,`codigo1`,`importe1`,importe2) values (" & vUsu.codigo & "," & miRsAux!tipodeiva
                        miSQL = miSQL & "," & TransformaComasPuntos(CStr(ImpTot)) & "," & TransformaComasPuntos(CStr(ImpTeo)) & ")"
                        TiposIva = TiposIva + 1
                        cadNomRPT = cadNomRPT & miRsAux!tipodeiva & "|"
                    End If
                    conn.Execute miSQL
                
                End If
                
                miRsAux.MoveNext
            Wend
            miRsAux.Close
        Next numParam
        
        If TiposIva > 3 Or cadNomRPT = "|" Then
            'ERROR  ERROR ERROR
            'ERROR en los tipos de iva. Hay mas de 3 o no hay ninguno
            If cadNomRPT = "" Then TiposIva = 0
            cadNomRPT = "Error en los tipos de IVA." & vbCrLf & "Total IVAS: " & TiposIva & vbCrLf & " Fec: " & devuelve
            MsgBox cadNomRPT, vbExclamation
            Exit Function
        End If
        InsertarUnaFacturaTicket2 vCl, vTM


          
    
    
        'Ahora, despues de crear la factura temporal FTG, insertare en la tabla
        'que lleva la relacion, numfactura, codticket
        miSQL = "INSERT INTO sfactik(`numfacFTG`,`fecfacFTG`,`numfactu`,`fecfactu`,`codtraba`)"
        miSQL = miSQL & " SELECT " & vTM.Contador & ",'" & devuelve & "',numfactu,fecfactu," & txtTrab(2).Text & " FROM scafac where "
        miSQL = miSQL & cadSelect
        If Diario Then miSQL = miSQL & " AND fecfactu='" & devuelve & "'"
        miSQL = miSQL & " AND codclien= " & Cliente
        conn.Execute miSQL
        
         
         vTM.IncrementarContador vTM.TipoMovimiento
        
         Set vTM = Nothing
    
        'Lo pongo a contabuilizado
        'Pongo la marca de contabilizado
        miSQL = "UPDATE scafac SET intconta = 1 WHERE " & cadSelect
        
        If Diario Then miSQL = miSQL & " AND fecfactu='" & devuelve & "'"
        miSQL = miSQL & " AND codclien= " & Cliente
        conn.Execute miSQL

        ObtenerDatosTickets2 = True
        
        Exit Function
EObteniendoDatosTickets:

    If Err.Number <> 0 Then
        MuestraError Err.Number, Err.Description & vbCrLf & miSQL
    End If
    Set vCl = Nothing
    Set vTM = Nothing
End Function


Private Sub InsertarUnaFacturaTicket2(ByRef vCli As CCliente, ByRef vTipoM As CTiposMov)
Dim TiposIva As Byte
    'No hay control de errores. Si salta, que vaya al sub ppal


        'Ya tengo las bases ivas para las facturas
        'Ahora creo la FTG para poder utilizar las funciones ya realizadas
        
        

        
             Set vTipoM = New CTiposMov
             vTipoM.Leer "FTG"
             vTipoM.ConseguirContador vTipoM.TipoMovimiento
             
             miSQL = "INSERT INTO `scafac` (`codtipom`,`numfactu`,`fecfactu`,`codclien`,`nomclien`,`domclien`,`codpobla`,"
             miSQL = miSQL & "`pobclien`,`proclien`,`nifclien`,`telclien`,`coddirec`,`nomdirec`,"
             miSQL = miSQL & "`codagent`,`codforpa`,`dtoppago`,`dtognral`,`codbanco`,`codsucur`,`digcontr`,`cuentaba`,"
             miSQL = miSQL & "`brutofac`,`impdtopp`,`impdtogr`,`intconta`,`totalfac`,"
             'LOS IVAS
             miSQL = miSQL & "`baseimp1`,`codigiv1`,`porciva1`,`imporiv1`,"
             miSQL = miSQL & "`baseimp2`,`codigiv2`,`porciva2`,`imporiv2`,"
             miSQL = miSQL & "`baseimp3`,`codigiv3`,`porciva3`,`imporiv3`)"
             
             'Cargo los ivas
             cadNomRPT = "Select codigo1,importe1,importe2 from tmpinformes where codusu = " & vUsu.codigo
             miRsAux.Open cadNomRPT, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
             cadNomRPT = ""
             TiposIva = 0
             ImpTot = 0
             ImpTeo = 0
             While Not miRsAux.EOF
                 TiposIva = TiposIva + 1
                 codigo = DevuelveDesdeBD(conConta, "porceiva", "tiposiva", "codigiva", miRsAux!Codigo1)
                 cadFrom = "," & TransformaComasPuntos(CStr(miRsAux!Importe1)) & "," & miRsAux!Codigo1 & "," & TransformaComasPuntos(codigo) & ","
                 cadFrom = cadFrom & TransformaComasPuntos(CStr(miRsAux!Importe2))
                 
                 'Meto en el select
                 cadNomRPT = cadNomRPT & cadFrom
                 
                 'ImpTot
                 ImpTot = ImpTot + miRsAux!Importe1
                 ImpTeo = ImpTeo + miRsAux!Importe2
                 miRsAux.MoveNext
             Wend
             miRsAux.Close
                 
                 
             'Si no tiene 3 tipos de ivas meter los null
             For numParam = TiposIva + 1 To 3
                 cadNomRPT = cadNomRPT & ",NULL,NULL,NULL,NULL"
             Next
             
             
             'Ahora relleno los datos que faltan
             'INSERT INTO `scafac` (`codtipom`,`numfactu`,`fecfactu`,`codclien`,`nomclien`,`domclien`,`codpobla`,"
             '`pobclien`,`proclien`,`nifclien`,`telclien`,`coddirec`,`nomdirec`,"
             '`codagent`,`codforpa`,`dtoppago`,`dtognral`,`codbanco`,`codsucur`,`digcontr`,`cuentaba`,"
             '`brutofac`,`impdtopp`,`impdtogr`,`intconta`,`totalfac`,"
                         
             cadFrom = " VALUES ('" & vTipoM.TipoMovimiento & "'," & vTipoM.Contador & ",'" & devuelve & "'," & vCli.codigo
             cadFrom = cadFrom & ",'" & vCli.Nombre & "','','0','','','0',NULL,NULL,NULL" '0: codpos y nif
             'Agente:
             cadFrom = cadFrom & "," & vCli.Agente & "," & vCli.ForPago & ",0,0,NULL,NULL,NULL,NULL,"
             'Bruto factra
             cadFrom = cadFrom & "" & TransformaComasPuntos(CStr(ImpTot)) & ",0,0,0," & TransformaComasPuntos(CStr(ImpTot + ImpTeo))
              
             miSQL = miSQL & cadFrom & cadNomRPT & ")"
             conn.Execute miSQL
             
            'Si lleva la analitica metere una linea en slifac1 que es desde donde,
            ' el proceso de contabilizacion cojera EL CODTRABA para obtener el CC
                
                miSQL = "insert into `scafac1` (`codtipom`,`numfactu`,`fecfactu`,codtipoa,numalbar,`codenvio`,`codtraba`,`codtrab1`,`codtrab2`)"
                miSQL = miSQL & " VALUES ('FTG'," & vTipoM.Contador & ",'" & devuelve & "','DAV','8',"  'Pongo tipoa y numalbar a piñon
                miSQL = miSQL & vParamAplic.PorDefecto_Envio & "," & txtTrab(2).Text & "," & txtTrab(2).Text & "," & txtTrab(2).Text & ")"
                conn.Execute miSQL
            
            
            
            


    

End Sub




'------------------------------------------------------------------
'------------------------------------------------------------------
'
'       Informe de trazabilidad
'       ========================================
'
'
'
'
'       A partir del desde /hasta mostraremos el informe que tiene la asociacion
'       entre albaranes de compra / venta
'
'
'       Hay datos tanto en albaranes como en facturas, con lo cual insertare sobre tmp
'
'
'----------------------------------------------------------------------
'----------------------------------------------------------------------
Private Sub HacerInformeTrazabilidad()

    
    
    InicializarVbles
    
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    
    
    If txtFecha(22).Text <> "" Or txtFecha(23).Text <> "" Then
        campo = "{slcomven.fechaalbc}"
        devuelve = "pDHFamilia=""Fecha: "
        If Not PonerDesdeHasta(campo, "F", 22, 23, devuelve) Then Exit Sub
    End If
    
    If txtCodProve(10).Text <> "" Or txtCodProve(11).Text <> "" Then
        campo = "{slcomven.codprovec}"
        devuelve = "pDHProveedor=""Proveedor: "
        If Not PonerDesdeHasta(campo, "PRO", 10, 11, devuelve) Then Exit Sub
    End If
     
    If txtArticulo(4).Text <> "" Or txtArticulo(5).Text <> "" Then
        campo = "{slcomven.codartic}"
        devuelve = "pDHArticulo=""Artículo: "
        If Not PonerDesdeHasta(campo, "ART", 4, 5, devuelve) Then Exit Sub
    End If
    
    
    'Montamos el select para los registros
    cadSelect = QuitarCaracterACadena(cadSelect, "{")
    cadSelect = QuitarCaracterACadena(cadSelect, "}")
    If cadSelect = "" Then cadSelect = " 1 = 1 "
    campo = "slcomven WHERE  " & cadSelect
    
    
    If Not HayRegParaInforme(campo, "", True) Then
        MsgBox "No hay facturas con estos valores", vbExclamation
        Exit Sub
    End If
    
    
    
    cadNomRPT = "rTraza.rpt"
    LlamarImprimir False
    
End Sub




Private Sub CargaTablasCambio()


    Set miRsAux = New ADODB.Recordset
    
    miRsAux.Open "show tables", conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Me.cboTablas.AddItem miRsAux.Fields(0)
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing


End Sub



Private Sub CargarCamposTabla()
'Dim Cad As String
'Dim Aux As String
Dim RS As ADODB.Recordset
Dim i As Integer
Dim TieneClaves As Boolean

    
    miSQL = "Select * from " & Me.cboTablas.List(cboTablas.ListIndex) & " LIMIT 1,1"
    Set RS = New ADODB.Recordset
    RS.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
 
        TieneClaves = False
        For i = 0 To RS.Fields.Count - 1
           
            
            
            'SOLO TEXTOS
            If RS.Fields(i).Type = 129 Or RS.Fields(i).Type = 200 Or RS.Fields(i).Type = adVarChar Then
    
       
  
                If RS.Fields(i).Properties(18).Value Then
                    'NO HACEMOS NADA. Es campo clave
                
                Else
                    cboCampos.AddItem RS.Fields(i).Name
                End If
                
            End If
            
            'Para saber si tiene claves
            If RS.Fields(i).Properties(18).Value Then TieneClaves = True
            
        Next i
        
        
        
    RS.Close
    Set RS = Nothing

    If cboCampos.ListCount > 0 And Not TieneClaves Then
        MsgBox "No tiene campos clave", vbInformation
        Me.cboCampos.Clear
    End If
End Sub




Private Sub UpdatearTablaRoot()
Dim i As Integer
Dim TienDatos As Boolean

    On Error GoTo EUpdatearTablaRoot
    
    devuelve = Me.cboTablas.List(cboTablas.ListIndex)
    miSQL = "Select " & Me.cboCampos.List(cboCampos.ListIndex) & "," & devuelve & ".* from " & devuelve

    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    cadFrom = ""
    miSQL = ""
    While Not miRsAux.EOF
        If IsNull(miRsAux.Fields(0)) Then
            Me.lblMultibase.Caption = ""
            Me.lblMultibase.Refresh
        Else
            miSQL = miRsAux.Fields(0)
            Me.lblMultibase.Caption = miSQL
            Me.lblMultibase.Refresh
            devuelve = RevisaCaracterMultibase(miSQL)
            
            If miSQL <> devuelve Then
                    'La clave
                    cadFrom = ""
                    For i = 0 To miRsAux.Fields.Count - 1
                        If miRsAux.Fields(i).Properties(18).Value Then
                            Select Case miRsAux.Fields(i).Type
                            Case 133
                                campo = CStr(miRsAux.Fields(i))
                                campo = "'" & Format(campo, "yyyy-mm-dd") & "'"
            
                            Case 135 'Fecha/Hora
                                campo = DBSet(miRsAux.Fields(i), "FH", "S")
                            'Numero normal, sin decimales
                            Case 2, 3, 16 To 19
                                campo = miRsAux.Fields(i)
                            Case 129, 200
                                campo = DBSet(miRsAux.Fields(i), "T")
                            Case Else
                                MsgBox "No tratado: " & miRsAux.Fields(i).Type, vbExclamation
                                Exit Sub
                                
                            End Select
                            cadFrom = cadFrom & " AND " & miRsAux.Fields(i).Name & " = " & campo
                        End If
                    Next i
                    cadFrom = Mid(cadFrom, 6)
                    devuelve = DevNombreSQL(devuelve)
                    miSQL = "UPDATE " & Me.cboTablas.List(cboTablas.ListIndex) & " SET " & Me.cboCampos.List(cboCampos.ListIndex)
                    miSQL = miSQL & " = '" & devuelve & "' WHERE " & cadFrom
                    conn.Execute miSQL
            End If 'DEl campo <>
        End If 'de ISNULL
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    'If miSQL <> "" Then
        MsgBox "Proceso finalizado", vbInformation
    'Else
    '    MsgBox "No hay registros", vbInformation
    'End If
    Exit Sub
EUpdatearTablaRoot:
    MuestraError Err.Number, Err.Description
End Sub


Private Sub CargarOtrasOfertas()
Dim IT 'As ListItem
    Me.lw1.ListItems.Clear
    lblDpto(27).Caption = miRsAux!NomClien
    While Not miRsAux.EOF
        Set IT = lw1.ListItems.Add()
        IT.Text = Format(miRsAux!NumOfert, "000000")
        IT.SubItems(1) = Format(miRsAux!fecofert, "dd/mm/yyyy")
        IT.SubItems(2) = Format(miRsAux!FecEntre, "dd/mm/yyyy")
        IT.SubItems(3) = DBLet(miRsAux!nomdirec, "T") & " "
        If Val(miRsAux!aceptado) = 0 Then
            IT.Checked = True
        Else
            IT.Checked = False
        End If
        miRsAux.MoveNext
    Wend
    
End Sub


Private Sub CargaListMov()
Dim R As ADODB.Recordset
    Set R = New ADODB.Recordset
    If Me.Opcion = 23 Then
    
    
        'Opciones guardadas
        textValueLeer "situalb", campo
        
    
        'Estoy cargando el list de las fras
        Me.List1.Clear
        miSQL = "select * from stipom where codtipom like 'AL%' AND codtipom<>'ALC' order by codtipom"
        R.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not R.EOF
            Me.List1.AddItem R!nomtipom & " (" & R!codtipom & ")"
            If InStr(1, campo, R!codtipom) > 0 Then
                List1.Selected(List1.NewIndex) = False
            Else
                List1.Selected(List1.NewIndex) = True
            End If
            R.MoveNext
        Wend
        R.Close
        
    End If
    Set R = Nothing
End Sub

Private Sub txtZona_GotFocus(Index As Integer)
    ConseguirFoco txtZona(Index), 3
End Sub

Private Sub txtZona_KeyPress(Index As Integer, KeyAscii As Integer)
     KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtZona_LostFocus(Index As Integer)
     miSQL = ""
     If txtZona(Index).Text <> "" Then
        If IsNumeric(txtZona(Index).Text) Then
            miSQL = DevuelveDesdeBD(conAri, "nomzonas", "szonas", "codzonas", txtZona(Index).Text, "N")
            If miSQL = "" Then miSQL = "El código no pertence a ninguna zona"
        Else
            MsgBox "Campo zona debe ser numérico", vbExclamation
            txtZona(Index).Text = ""
            PonerFoco txtZona(Index)
        End If
    End If
    Me.txtDescZona(Index).Text = miSQL
End Sub



Private Function CargarDatosImprimeAlbaranConTransporte() As Boolean
Dim Aux As String

    CargarDatosImprimeAlbaranConTransporte = False
    
    If optAlbTrans(1).Value Then
        'Solo quiere el listado de albaranes. NO quiere reimprimir los albaranes
        CargarDatosImprimeAlbaranConTransporte = True
        Exit Function
    End If
        
        
    'Para cada albaran pendiente de reeimprimir habra que ver si tiene resto de pedido
    'Si lo tiene cargaremos la tabla
    miSQL = "DELETE FROM tmpsliped WHERE codusu = " & vUsu.codigo
    conn.Execute miSQL
    
    If optAlbTrans(0).Value Then
        If chkImpAlbRut(0).Value = 0 Then
            'Para tener un temporal por si se va la luz
            miSQL = "DELETE FROM tmpnseries WHERE codusu =" & vUsu.codigo
            conn.Execute miSQL
            
            miSQL = "INSERT INTO tmpnseries (codusu ,numlinealb,numserie) "
            miSQL = miSQL & " select " & vUsu.codigo & " ,numalbar,fechaalb from scaalb where " & cadSelect
            conn.Execute miSQL
            
            
        End If
    End If
    
    
    
    '
    '**** linkamos POR codzona--> CODDIREN.  pARA NO CREAR MAS CAMPOS EN TMPSLIPED.. En codlamac llevare el coddiren
    '
    miSQL = "Select " & vUsu.codigo & ",scaped.numpedcl,numlinea,codartic,nomartic,cantidad,coddiren,codclien FROM scaped,sliped where scaped.numpedcl =sliped.numpedcl"
    miSQL = miSQL & " AND (scaped.numpedcl,fecpedcl) in "
    miSQL = miSQL & "( select numpedcl,fecpedcl from scaalb where " & cadSelect & ")"
    
    
    
    Set miRsAux = New ADODB.Recordset
    
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Aux = ""
    While Not miRsAux.EOF
        'miSQL = "INSERT INTO tmpsliped(codusu, numpedcl, numlinea, codartic, nomartic, cantidad,codzona,codclien) " & miSQL
        'caped.numpedcl,numlinea,codartic,nomartic,cantidad,coddiren,codclien
        miSQL = ", (" & vUsu.codigo & "," & miRsAux!numpedcl & "," & miRsAux!numlinea & "," & DBSet(miRsAux!codArtic, "T") & "," & DBSet(miRsAux!NomArtic, "T") & ","
        miSQL = miSQL & DBSet(miRsAux!cantidad, "N") & "," & DBSet(miRsAux!coddiren, "N", "S") & "," & miRsAux!codClien & ")"
        Aux = Aux & miSQL
    
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    If Aux <> "" Then
        miSQL = Mid(Aux, 2)
        miSQL = "INSERT INTO tmpsliped(codusu, numpedcl, numlinea, codartic, nomartic, cantidad,codzona,codclien) VALUES " & miSQL
    
    
    Else
        miSQL = "Select numpedcl from scaped where false"   'Para que no de error el SQL
    
    End If
    Set miRsAux = Nothing
    If ejecutar(miSQL, False) Then
        'Pondre a cero la codzona pq si no el rpt no enlaza bien
        miSQL = "UPDATE tmpsliped SET codzona = 0 where codusu = " & vUsu.codigo & " AND codzona is null"
        ejecutar miSQL, False
        CargarDatosImprimeAlbaranConTransporte = True
    End If
    
End Function






Private Sub ActualizarPreciosVentaCompra()
Dim RT As ADODB.Recordset

    cadFrom = ""
    If Me.optCopiaPrecio(1).Value Then
        cadParam = "slista"
        cadFrom = " AND codlista = " & vParamAplic.CodTarifa
        devuelve = "slispr"
    Else
        cadParam = "slispr"
        devuelve = "slista"
    End If
   

   
    campo = " WHERE l.codartic=sartic.codartic "
    
    'En cadselect tengo el where.  Ahora lo completo con las tablas y  joins
    Set RT = New ADODB.Recordset
    campo = campo & cadSelect & cadFrom
    campo = campo & " AND fechanue = " & DBSet(txtFecha(34).Text, "F")
    campo = "Select nomartic,sartic.codprove,l.* from sartic," & cadParam & " l" & campo & " ORDER BY codartic"
    RT.Open campo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RT.EOF
        Label3(107).Caption = RT!NomArtic
        Label3(107).Refresh
        'Vemos si existe ya la referencia
        cadSelect = "codartic=" & DBSet(RT!codArtic, "T")
        If Me.optCopiaPrecio(1).Value Then
            cadParam = "slispr"
            cadSelect = cadSelect & " AND codprove = " & RT!Codprove
       
        Else
            cadParam = "slista"
            cadSelect = cadSelect & " AND codlista = " & vParamAplic.CodTarifa
           
        End If
        ImpTot = RT!precionu
        If ImpTot > 0 Then
            If ExisteEnListaprecio Then
                'UPDATE
                campo = "UPDATE " & devuelve & " SET precionu = " & DBSet(ImpTot, "N") & ", fechanue=" & DBSet(Me.txtFecha(34).Text, "F")
                campo = campo & " WHERE " & cadSelect
                
            Else
                If Me.optCopiaPrecio(0).Value Then
                    campo = "INSERT INTO slista(codartic,codlista,precioac,dtopermi,fechanue,precionu) VALUES ("
                    campo = campo & DBSet(RT!codArtic, "T") & "," & vParamAplic.CodTarifa & "," & DBSet(RT!precioac, "N") & ",0,"
                    campo = campo & DBSet(txtFecha(34).Text, "F") & "," & DBSet(ImpTot, "N") & ")"
                Else
                    campo = "INSERT INTO slispr(codartic,codprove,precioac,dtopermi,dtoperm1,fechanue,precionu)  VALUES ("
                    campo = campo & DBSet(RT!codArtic, "T") & "," & RT!Codprove & "," & DBSet(RT!precioac, "N") & ",0,0,"
                    campo = campo & DBSet(txtFecha(34).Text, "F") & "," & DBSet(ImpTot, "N") & ")"
                End If
            End If
            ejecutar campo, True
            
            
            'Si actualizamos en slista (ventas), es decir, actualizamos desde compra, y tiene que actualizar precio especial
            If Me.optCopiaPrecio(0).Value Then
                If vParamAplic.ActualizaPrecioEspecial Then
                    campo = "UPDATE sprees SET precionu = " & DBSet(ImpTot, "N") & ", fechanue=" & DBSet(Me.txtFecha(34).Text, "F")
                    campo = campo & " WHERE codartic = " & DBSet(RT!codArtic, "T")
                     ejecutar campo, False 'Si no existe, NO lo creamos. Simplemente, elmupdate NO hara nada
                End If
            End If
        End If
        RT.MoveNext
    Wend
    RT.Close
    
    
    
    
    
    
    
    
    
    Set RT = Nothing
    
End Sub

    
Private Function ExisteEnListaprecio() As Boolean
    ExisteEnListaprecio = False
    cadParam = "select * from " & cadParam & " WHERE " & cadSelect
    miRsAux.Open cadParam, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        If Not IsNull(miRsAux!codArtic) Then ExisteEnListaprecio = True
    End If
    miRsAux.Close
End Function



'----------------------------------
Private Sub RecorrerRiesgo()
Dim TareaCompletada As Boolean
Dim fin As Boolean
Dim RI As ADODB.Recordset


    Label3(95).Caption = "Cargando clientes"
    Label3(95).Refresh
    
    miSQL = "Select * from tiposiva"
    Set RI = New ADODB.Recordset
    RI.Open miSQL, ConnConta, adOpenKeyset, adLockPessimistic, adCmdText
    
    miSQL = "Select * from tmpinformes where codusu = " & vUsu.codigo & " ORDER BY codigo1 " 'codclien
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    TareaCompletada = False
    While Not fin
    
        pb2.Value = pb2.Value + 1
        If (pb2.Value Mod 15) = 0 Then
            Me.Refresh
            DoEvents
        End If
        
        
        Label3(95).Caption = DBLet(miRsAux!nombre1, "T")
        Label3(95).Refresh
        
        
        RiesgoCliente miRsAux!Codigo1, miRsAux!campo2, Now, ImpTeo, ImpTot, RI, 60
        
            
        If ImpTeo <> 0 Or ImpTot <> 0 Then
            miSQL = "UPDATE tmpinformes set importe2=" & TransformaComasPuntos(CStr(ImpTeo))
            miSQL = miSQL & ", importe3=" & TransformaComasPuntos(CStr(ImpTot))
            miSQL = miSQL & ", porcen1 = 1"  'para luego buscar toos los que han cambiado
            miSQL = miSQL & " WHERE codusu = " & vUsu.codigo & " AND codigo1=" & miRsAux!Codigo1
            conn.Execute miSQL
        End If
        
        If Opcion = 0 Then
            miRsAux.MoveNext
            If miRsAux.EOF Then
                TareaCompletada = True
                fin = True
            End If
        Else
            fin = True
        End If
        
        
        
    Wend
    miRsAux.Close
    RI.Close
    Set RI = Nothing
    
    If TareaCompletada Then
        Label3(95).Caption = "Buscando cambios en situacion"
        pb2.Value = 0
        Me.Refresh
        Espera 0.5
            TareaCompletada = False
            fin = False
            Label3(95).Caption = "Actualizando datos en tabla clientes"
            Label3(95).Refresh
            Espera 0.2
            miSQL = "from tmpinformes where codusu = " & vUsu.codigo & " ORDER BY codigo1 " 'codclien
            miRsAux.Open "Select * " & miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not fin
            
                pb2.Value = pb2.Value + 1
                If (pb2.Value Mod 50) = 0 Then
                    Me.Refresh
                    DoEvents
                End If
            
                Label3(95).Caption = DBLet(miRsAux!nombre1, "T")
                Label3(95).Refresh
          
                'Sobrepasa el riesog si o no
                
                ImpTot = DBLet(miRsAux!Importe2, "N") + DBLet(miRsAux!Importe3, "N")
                ImpTeo = DBLet(miRsAux!Importe1, "N")
                miSQL = "UPDATE sclien SET UtFecrecal = " & DBSet(Now, "F")
                miSQL = miSQL & ", riesgoact = " & DBSet(ImpTot, "N")
        
                If ImpTeo >= ImpTot Then
        
                    'NO supera el limite
                    If miRsAux!campo1 > 0 Then
                        'Estaba bloqueado por riesgo. Le quito la marca
                        If CInt(miRsAux!campo1) = vParamAplic.SituacionBloqueoOpAseg Then miSQL = miSQL & " ,codsitua = 0"
                    End If
                Else
                    'SUPERA EL RIESGO
                    If miRsAux!campo1 = 0 Then miSQL = miSQL & " ,codsitua = " & vParamAplic.SituacionBloqueoOpAseg
                    
                End If
                miSQL = miSQL & " WHERE codclien = " & miRsAux!Codigo1
                conn.Execute miSQL
                
                
                 If Opcion = 0 Then
                    miRsAux.MoveNext
                    If miRsAux.EOF Then
                        TareaCompletada = True
                        fin = True
                    End If
                Else
                    fin = True
                End If
                
                
                
                
            Wend
            miRsAux.Close
            
            If TareaCompletada Then
                MsgBox "Proceso finalizado con exito", vbExclamation
                Unload Me
            End If
            
    End If
    Opcion = 31
End Sub






'----------------------------------------------------------------------------------
'----------------------------------------------------------------------------------
'----------------------------------------------------------------------------------
'
'INforme de pedido de proveedores. Despues podra generar un pedido desde aqui
'
'----------------------------------------------------------------------------------
'----------------------------------------------------------------------------------
'----------------------------------------------------------------------------------

Private Sub MontaSQL_InformePedidoProveedor()
    '---------------------------------------------------------------------------------------
    'proveedor
    If txtCodProve(17).Text <> "" Then miSQL = miSQL & " AND codProve = " & txtCodProve(17).Text
    'Situacion
    If Me.cboProPed(1).ListIndex > 0 Then miSQL = miSQL & " AND codstatu = " & Me.cboProPed(1).ListIndex - 1
    If Me.txtFamia(2).Text <> "" Then miSQL = miSQL & " AND codfamia >= " & txtFamia(2).Text
    If Me.txtFamia(3).Text <> "" Then miSQL = miSQL & " AND codfamia <= " & txtFamia(3).Text
    If Me.txtmarca(0).Text <> "" Then miSQL = miSQL & " AND codmarca >= " & txtmarca(0).Text
    If Me.txtmarca(1).Text <> "" Then miSQL = miSQL & " AND codmarca <= " & txtmarca(1).Text
    '---------------------------------------------------------------------------------------
End Sub


Private Function GeneraInformepedidoProv() As Boolean
Dim ColArt_ As Collection
Dim FI As Date
Dim RatioMensual As Currency
Dim AprovMesMin As Currency
Dim AprovMesMax As Currency
Dim Cantaux As Currency
Dim KK As Integer

    On Error GoTo Etmppedprov
    GeneraInformepedidoProv = False
    
    'Vacio temporales
    Label3(100).Caption = "Preparando datos"
    Label3(100).Refresh
    conn.Execute "DELETE FROM tmppedprov where codusu = " & vUsu.codigo
    
    
    'Monto el SQL
    miSQL = "Select " & vUsu.codigo & ",codprove,codfamia,codartic,1,artvario  from sartic WHERE ctrstock = 1 "
    
    
    MontaSQL_InformePedidoProveedor  'D/H familia y marca
    
    
    
    miSQL = "insert into `tmppedprov` (`codusu`,`codprove`,`codfamia`,`codartic`,TieneVtasEscandallo,deVarios) " & miSQL
    conn.Execute miSQL
    
    
    
    'Febrero 2013
    'Si marca rotacion añadiremos los articulos de varios (para los desde /hasta )
    If cboProPed(0).ListIndex = 1 Then '->ha marcado solo rotacion
        Label3(100).Caption = "Articulos varios"
        Label3(100).Refresh
        Espera 0.3
        miSQL = "Select " & vUsu.codigo & ",codprove,codfamia,codartic,0,1  from sartic WHERE artvario=1 "
        MontaSQL_InformePedidoProveedor  'D/H familia y marca
        
        miSQL = miSQL & " AND not codartic in (select codartic from tmppedprov WHERE codusu =" & vUsu.codigo & ")"
        miSQL = "insert  into `tmppedprov` (`codusu`,`codprove`,`codfamia`,`codartic`,TieneVtasEscandallo,deVarios) " & miSQL
        conn.Execute miSQL
    End If
    
    
    miSQL = DevuelveDesdeBD(conAri, "count(*)", "tmppedprov", "codusu", CStr(vUsu.codigo))
    If miSQL = "" Then miSQL = "0"
    If Val(miSQL) = 0 Then
        MsgBox "Ningun dato para procesar", vbExclamation
        Exit Function
    End If
    Label3(100).Tag = Val(miSQL)
    
    'AHora tengo cargada la tmp. La voy reccorriendo
    Label3(100).Caption = "leyendo tmp"
    Label3(100).Refresh
    Set ColArt_ = New Collection
    miSQL = "Select * from tmppedprov where codusu = " & vUsu.codigo
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    miSQL = ""
    NumRegElim = 0
    While Not miRsAux.EOF
        
        'Ya tengo el articulo
        NumRegElim = NumRegElim + 1
        miSQL = miSQL & ", " & DBSet(miRsAux!codArtic, "T")
    
    
    
        If NumRegElim = 30 Then
            ColArt_.Add Mid(miSQL, 2)
            miSQL = ""
            NumRegElim = 0
        End If
    
    
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    If NumRegElim > 0 Then
        ColArt_.Add Mid(miSQL, 2)
        miSQL = ""
        NumRegElim = 0
    End If
    
    'Tengo agrupado los articulos
    ' en ctos alb fras esta mes
    FI = DateAdd("m", -vParamAplic.Rot_ConsumMes1, Now)
    PedProv_SalidasPeriodo FI, ColArt_
    
    DoEvents
    
    
    'Cantidad del perido 1 y del 2
    FI = DateAdd("m", -vParamAplic.Rot_ConsumMes1, Now)
    PedProv_CantidadPeriodo FI, ColArt_, True
    DoEvents
    FI = DateAdd("m", -vParamAplic.Rot_ConsumMes2, Now)
    PedProv_CantidadPeriodo FI, ColArt_, False
    DoEvents
    PedProv_PedidosPendiente ColArt_, 1
    PedProv_PedidosPendiente ColArt_, 2
    PedProv_PedidosPendiente ColArt_, 3
    
    
    
    
    'Aqui ira lo del escandallo
    'Como por cada linea del ppal sale n del del escandallo,
    'ahora tendremos que ver
    HacerEscandalloPropuestaPedido ColArt_
    
    miSQL = "UPDATE tmppedprov SET pedpro=pedpro1+pedpro2+pedpro3 WHERE codusu = " & vUsu.codigo
    conn.Execute miSQL
    miSQL = "UPDATE tmppedprov SET pedcli=pedcli1+pedcli2+pedcli3 WHERE codusu = " & vUsu.codigo
    conn.Execute miSQL
    
    
    
    DoEvents
    PedProv_Stock ColArt_, 1
    PedProv_Stock ColArt_, 2
    PedProv_Stock ColArt_, 3
    miSQL = "UPDATE tmppedprov SET stock=stock1+stock2+stock3 WHERE codusu = " & vUsu.codigo
    conn.Execute miSQL
    
    
    If Me.cboProPed(0).ListIndex > 0 Then
        'Parametro de ROTACION
        'Solo quiere los de rotacion
        Label3(100).Caption = "Rotacion"
        Label3(100).Refresh
        miSQL = "DELETE tmppedprov.* from tmppedprov,sartic where codusu = " & vUsu.codigo & " AND tmppedprov.codartic=sartic.codartic"
        miSQL = miSQL & " and rotacion = 0  and pedcli =0 and devarios=0"
        conn.Execute miSQL
        
        
        Label3(100).Caption = "Ajuste stock para varios"
        Label3(100).Refresh
        miSQL = "DELETE  from tmppedprov where codusu = " & vUsu.codigo & " AND devarios=1 and pedcli =0"  'varios que no tengan pedidos
        conn.Execute miSQL
        
    End If
    
    
    
    'Para los que queden de varios, el stock almacen lo ponemos a cero
    Label3(100).Caption = "Ajuste stock para varios(II)"
    Label3(100).Refresh
    miSQL = "UPDATE tmppedprov set stock =0 where codusu = " & vUsu.codigo & " and devarios=1"
    conn.Execute miSQL
    
        
    
    
    'Si no ha indicado proveedor
    If txtCodProve(17).Text = "" Then
        Label3(100).Caption = "Minimo de salidas"
        Label3(100).Refresh
        'Vamos a eliminar de la tmp aquellas entradas que no superan el minimo de salidas
        miSQL = "DELETE  from tmppedprov  where codusu = " & vUsu.codigo & " AND nsal <" & txtAnyo(4).Text
        conn.Execute miSQL
        Espera 0.2
    End If
    
    
    'Recorremos el rs
    Label3(100).Caption = "Calculando datos"
    Label3(100).Refresh
    Espera 0.3
    NumRegElim = 0
    miSQL = "Select * from tmppedprov where codusu = " & vUsu.codigo & " ORDER BY codartic"
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Label3(100).Caption = miRsAux!codArtic
        Label3(100).Refresh
        NumRegElim = NumRegElim + 1
        
        'If miRsAux!codArtic = "48HD1/15" Then S top
        
        RatioMensual = miRsAux!ult_1 / vParamAplic.Rot_ConsumMes1
        If RatioMensual > 0 Then
                'Esto es fijo para el mes mas y min. Es la cantidad necesaria
                Cantaux = miRsAux!stock + miRsAux!pedpro - miRsAux!pedcli
                
                'Ya tengo lo que consumo por mes
                'Vamos a ver para el aprovisionamiento para el mes min
                AprovMesMin = vParamAplic.Rot_ConsumMesMin * RatioMensual 'NEcesito para n mese
                AprovMesMin = AprovMesMin - Cantaux
                
                'Vamos a ver para el aprovisionamiento para el mes max
                AprovMesMax = vParamAplic.Rot_ConsumMesMax * RatioMensual 'NEcesito para n mese
                AprovMesMax = AprovMesMax - Cantaux
                
                
                miSQL = "UPDATE tmppedprov SET permin = " & Val(Round2(AprovMesMin, 0))
                miSQL = miSQL & " , permax = " & Val(Round2(AprovMesMax, 0))
                miSQL = miSQL & " WHERE codusu = " & vUsu.codigo
                miSQL = miSQL & " and codartic  = " & DBSet(miRsAux!codArtic, "T")
                conn.Execute miSQL
                
                
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    'Para los almacenes cnsolidados voy a poner el punto de pedido
    
    'puntoped2
    If txtAlma(7).Text <> "" Or txtAlma(8).Text <> "" Then
        'CONSOLIDADO
        'CREIA que era punto pedido por eso los campos se llaman asi. Realmente es MINIMO
        Label3(100).Caption = "Stock minimo consolidado"
        Label3(100).Refresh
        miSQL = "Select codartic from tmppedprov WHERE codusu = " & vUsu.codigo
        miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Set ColArt_ = Nothing
        Set ColArt_ = New Collection
        miSQL = ""
        While Not miRsAux.EOF
            miSQL = miSQL & ", " & DBSet(miRsAux!codArtic, "T")
            miRsAux.MoveNext
    
            If Len(miSQL) > 100 Then
                ColArt_.Add Mid(miSQL, 2)
                miSQL = ""
            End If
        Wend
        miRsAux.Close
        If miSQL <> "" Then ColArt_.Add Mid(miSQL, 2)
    
        
        For numParam = 1 To ColArt_.Count
            Label3(100).Caption = "St minimo " & numParam & " / " & ColArt_.Count
            Label3(100).Refresh
            miSQL = ""
            If Me.txtAlma(7).Text <> "" Then miSQL = "," & txtAlma(7).Text
            If Me.txtAlma(8).Text <> "" Then miSQL = miSQL & "," & txtAlma(8).Text
            miSQL = Mid(miSQL, 2)
            
            
            miSQL = " codalmac IN (" & miSQL & ")"
            miSQL = "Select codartic,codalmac,stockmin,canstock FROM salmac where  " & miSQL
            miSQL = miSQL & " AND stockmin >0"
            miSQL = miSQL & " AND codartic IN (" & Mid(ColArt_(numParam), 2) & ")"
            
            miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            miSQL = ""
            While Not miRsAux.EOF
                If miSQL <> miRsAux!codArtic Then
                    If miSQL <> "" Then
                        codigo = ""
                        If ImpTot > 0 Or ImpTeo > 0 Then
                        
                            codigo = "UPDATE tmppedprov SET puntoped2 = " & Val(ImpTot) & ", puntoped3 = " & Val(ImpTeo)
                            
                            codigo = codigo & " WHERE codusu =" & vUsu.codigo & " AND codartic =" & DBSet(miSQL, "T")
                            conn.Execute codigo

                        End If
                    End If
                    miSQL = miRsAux!codArtic
                    ImpTot = 0: ImpTeo = 0
                End If
                
                If miRsAux!codAlmac = Val(txtAlma(8).Text) Then
                    ImpTeo = miRsAux!CanStock
                    If ImpTeo < 0 Then ImpTeo = 0
                    ImpTeo = miRsAux!stockmin - ImpTeo  'ALMACEN segundo consolidado
                    If ImpTeo < 0 Then ImpTeo = 0
                Else
                    ImpTot = miRsAux!CanStock
                    If ImpTot < 0 Then ImpTot = 0
                    ImpTot = miRsAux!stockmin - ImpTot 'ALMACEN primer consolidado
                    If ImpTot < 0 Then ImpTot = 0
                End If
                miRsAux.MoveNext
            Wend
            miRsAux.Close
            If miSQL <> "" Then
                codigo = "UPDATE tmppedprov SET puntoped2 = " & Val(ImpTot) & ", puntoped3 = " & Val(ImpTeo)
                codigo = codigo & " WHERE codusu =" & vUsu.codigo & " AND codartic =" & DBSet(miSQL, "T")
                conn.Execute codigo
            End If
            
        
        Next
    
    End If
    
    
    
    
    'Si ha puesto entradas minimias
    ' si el resultado es menor que cero, es decir, no neceista aprovisionarse, no lo muestro
    If Me.txtAnyo(4).Text <> "" Then
        Label3(100).Caption = "Eliminando datos "
        Label3(100).Refresh
        miSQL = "DELETE FROM  tmppedprov WHERE codusu = " & vUsu.codigo
        miSQL = miSQL & " AND permin<0 and permax <0"
        conn.Execute miSQL
        
        Espera 0.2
        miSQL = DevuelveDesdeBD(conAri, "count(*)", "tmppedprov", "codusu", CStr(vUsu.codigo))
        If miSQL = "" Then miSQL = "0"
        If Val(miSQL) = 0 Then NumRegElim = 0
    End If
    
    
    'Enero 2015
    'UNicajas minimo
    'Sale del la tabla precios proveedor
    Label3(100).Caption = "Ud minimo"
    Label3(100).Refresh
    DoEvents
    
    Set ColArt_ = Nothing
    Set ColArt_ = New Collection
    miSQL = "Select distinct(codprove) from tmppedprov where codusu = " & vUsu.codigo
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        ColArt_.Add CStr(miRsAux.Fields(0))
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    For KK = 1 To ColArt_.Count
        Label3(100).Caption = "Ud minimo. " & KK & " de " & ColArt_.Count
        Label3(100).Refresh
    
        miSQL = "select * from slispr  where codprove = " & ColArt_.item(KK)
        miSQL = miSQL & " AND cantmini>0 and codartic in "
        miSQL = miSQL & " (select codartic from tmppedprov where codprove = " & ColArt_.item(KK)
        miSQL = miSQL & " AND codusu =" & vUsu.codigo & ")   "
        miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            miSQL = "UPDATE tmppedprov SET caja=" & DBSet(miRsAux!cantmini, "N") & " WHERE codprove =" & ColArt_.item(KK)
            miSQL = miSQL & " AND codartic = " & DBSet(miRsAux!codArtic, "T") & " AND codusu =" & vUsu.codigo
            conn.Execute miSQL
            miRsAux.MoveNext
        Wend
        miRsAux.Close
    Next
    
    'Sobre lo que queda. Veremos los porcentajes de un mismo cliente
    If NumRegElim > 0 Then
        If Val(Me.txtAnyo(5).Text) > 0 Then
                DoEvents
                Label3(100).Caption = "(1) Porcentaje ventas"
                Label3(100).Refresh
                FI = DateAdd("m", -vParamAplic.Rot_ConsumMes1, Now)
                Set ColArt_ = Nothing
                Set ColArt_ = New Collection
                conn.Execute "DELETE from tmpcommand where codusu =" & vUsu.codigo
                
                'Vamos a ver todos los articulos que
                'esta en albaranes,fras nsal>1
                'han pedido mas de unos cuantos
                miSQL = "select codartic,ult_1 from tmppedprov where nsal>1 and ult_1>1 and codusu = " & vUsu.codigo
                miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                
                While Not miRsAux.EOF
                    'por posiciones  ###
                    miSQL = Mid(miRsAux!codArtic & Space(16), 1, 16)
                    ColArt_.Add miSQL & miRsAux!ult_1
                    'ColArt.Add CStr(miRsAux!codArtic)
                    miRsAux.MoveNext
                Wend
                miRsAux.Close
                
                
                For NumRegElim = 1 To ColArt_.Count
                 '   tmpcommand CodUsu, CodProve, nomprove, cantidad
                    Label3(100).Caption = "(2)% ventas: " & NumRegElim & " de " & ColArt_.Count
                    Label3(100).Refresh
                    
                    campo = Trim(Mid(ColArt_.item(NumRegElim), 1, 16)) 'codartic
                    devuelve = TransformaComasPuntos(Mid(ColArt_.item(NumRegElim), 17))    'total
                    
                    miSQL = "INSERT INTO tmpcommand (CodUsu, CodProve, nomprove, cantidad,importel)"
                    miSQL = miSQL & " select " & vUsu.codigo & ",codclien,codartic,sum(cantidad)," & devuelve & " from scaalb,slialb where scaalb.numalbar =slialb.numalbar and slialb.codtipom=slialb.codtipom"
                    
                    miSQL = miSQL & " AND fechaalb >=" & DBSet(FI, "F")
                    
                    'Noviembre 2014. No hacia el consolidado
                    'miSQL = miSQL & " AND codalmac = " & Me.txtAlma(0).Text
                    miSQL = miSQL & " AND codalmac IN (" & Me.txtAlma(0).Text
                    If Me.txtAlma(7).Text <> "" Then miSQL = miSQL & "," & Me.txtAlma(7).Text
                    If Me.txtAlma(8).Text <> "" Then miSQL = miSQL & "," & Me.txtAlma(8).Text
                    miSQL = miSQL & ") and codartic = " & DBSet(campo, "T") & " GROUP BY codclien"
                    conn.Execute miSQL
                    
                    miSQL = "INSERT INTO tmpcommand (CodUsu, CodProve, nomprove, cantidad,importel)"
                    miSQL = miSQL & " select " & vUsu.codigo & ",codclien,codartic,sum(cantidad)," & devuelve & " from scafac,slifac "
                    miSQL = miSQL & " Where scafac.NumFactu = slifac.NumFactu And scafac.codtipom = slifac.codtipom And scafac.FecFactu = slifac.FecFactu"
                    miSQL = miSQL & " AND scafac.fecfactu >=" & DBSet(FI, "F")
                    'Noviembre 2014. No hacia el consolidado
                    'miSQL = miSQL & " AND codalmac = " & Me.txtAlma(0).Text
                    miSQL = miSQL & " AND codalmac IN (" & Me.txtAlma(0).Text
                    If Me.txtAlma(7).Text <> "" Then miSQL = miSQL & "," & Me.txtAlma(7).Text
                    If Me.txtAlma(8).Text <> "" Then miSQL = miSQL & "," & Me.txtAlma(8).Text
                    miSQL = miSQL & ") and codartic = " & DBSet(campo, "T") & " group by 1,2"
                    conn.Execute miSQL
                Next
                 
                 
                'Ahora vere cual tienen un porcentaje mayor
                
                RatioMensual = Val(txtAnyo(5).Text)
                Label3(100).Caption = "(3) Actualizando reg: " & NumRegElim & " de " & ColArt_.Count
                Label3(100).Refresh
                miSQL = "select nomprove,codprove,importel,sum(cantidad) suma from tmpcommand where codusu= " & vUsu.codigo
                miSQL = miSQL & " group by nomprove,codprove order by nomprove,codprove"
                miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

                While Not miRsAux.EOF
                    If DBLet(miRsAux!ImporteL, "N") > 0 Then
                        Cantaux = (miRsAux!Suma / miRsAux!ImporteL) * 100
                        
                        If Cantaux > RatioMensual Then
                        
                            
                            'Enero 2016.
                            'Piden que no controle si es de varios el cliente
                            miSQL = DevuelveDesdeBD(conAri, "clivario", "sclien", "codclien", CStr(miRsAux!Codprove))
                            miSQL = "0"
                            If miSQL = "0" Then
                                'Debug.Print "->" & DBLet(miRsAux!coda, "T")
                                miSQL = "UPDATE tmppedprov set De1Cliente = 1 WHERE codusu =" & vUsu.codigo & " AND codartic = " & DBSet(miRsAux!nomprove, "T")
                                conn.Execute miSQL
                            End If
                        Else
                           ' St op
                        End If
                    End If
                    miRsAux.MoveNext
                Wend
                miRsAux.Close
                NumRegElim = 1 'para que luego vaya bien
        End If
    End If
    
    
    'Herbelca. Si el articulo es conjunto de otro (aunque no este en el select), que lo marque

    Label3(100).Caption = "Escandallo (II)"
    Label3(100).Refresh
    miSQL = "UPDATE tmppedprov set esEscandallo =2 where codusu =" & vUsu.codigo
    miSQL = miSQL & " and esescandallo=0 and codartic in (select codarti1 from sarti1)"
    conn.Execute miSQL
    
    If NumRegElim = 0 Then
        MsgBox "Ningún dato generado", vbExclamation
    Else
        GeneraInformepedidoProv = True
    End If
    
Etmppedprov:
    If Err.Number <> 0 Then MuestraError Err.Number
    
End Function

'E n cuantos albaranes/fras esta el articulo en el periodo pequeño (perido1) de parametros
Private Sub PedProv_SalidasPeriodo(FInicio As Date, ByRef CA As Collection)
Dim J As Integer

    'Vamos a ver en cuantos albaranes, facturas del periodo salen

    For J = 1 To CA.Count
        Label3(100).Caption = "Alb " & J & "/" & CA.Count
        Label3(100).Refresh
        
        
        'Docuemtnos en los que esta el articulo
        
        'En cuantos albaranes esta
        miSQL = "select codartic,count(distinct(concat(scaalb.codtipom,scaalb.numalbar))) from scaalb,slialb  WHERE"
        miSQL = miSQL & " scaalb.Codtipom = slialb.Codtipom And scaalb.NumAlbar = slialb.NumAlbar"
        miSQL = miSQL & " AND fechaalb >=" & DBSet(FInicio, "F")
        'Marzo 2011. falta el codalmac
        If Me.txtAlma(7).Text = "" Then
            miSQL = miSQL & " AND codalmac = " & Me.txtAlma(0).Text
        Else
            'MARZO 2014
            'AHORA con el segundo consolidado
            'miSQL = miSQL & " AND codalmac IN (" & Me.txtAlma(0).Text & "," & Me.txtAlma(7).Text & ")"
            miSQL = miSQL & " AND codalmac IN (" & Me.txtAlma(0).Text & "," & Me.txtAlma(7).Text
            If Me.txtAlma(8).Text <> "" Then miSQL = miSQL & "," & Me.txtAlma(8).Text
            miSQL = miSQL & ")"
        End If
        miSQL = miSQL & " AND codartic IN (" & CA.item(J) & ") GROUP BY 1"
        miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            
            miSQL = "UPDATE tmppedprov SET nsal = nsal + " & miRsAux.Fields(1)
            miSQL = miSQL & " WHERE codusu = " & vUsu.codigo
            miSQL = miSQL & " and codartic  = " & DBSet(miRsAux!codArtic, "T")
            conn.Execute miSQL
            miRsAux.MoveNext
        
        
        Wend
        miRsAux.Close
        Espera 0.1
        
        'En facturas
        miSQL = " select codartic,count(distinct(concat(scafac1.codtipom,scafac1.numfactu))) from scafac1,slifac  WHERE"
        miSQL = miSQL & " scafac1.Codtipom = slifac.Codtipom And scafac1.NumFactu = slifac.NumFactu And scafac1.FecFactu = slifac.FecFactu"
        miSQL = miSQL & " and scafac1.numalbar=slifac.numalbar and scafac1.codtipoa=slifac.codtipoa"
        miSQL = miSQL & " AND fechaalb >=" & DBSet(FInicio, "F")
        
        
        'ENERO 2013. NO estaba la linea de abajo
        If Me.txtAlma(7).Text = "" Then
            miSQL = miSQL & " AND codalmac = " & Me.txtAlma(0).Text
        Else
            'MARZO 2014
            'AHORA con el segundo consolidado
            'miSQL = miSQL & " AND codalmac IN (" & Me.txtAlma(0).Text & "," & Me.txtAlma(7).Text & ")"
            miSQL = miSQL & " AND codalmac IN (" & Me.txtAlma(0).Text & "," & Me.txtAlma(7).Text
            If Me.txtAlma(8).Text <> "" Then miSQL = miSQL & "," & Me.txtAlma(8).Text
            miSQL = miSQL & ")"
            
        End If
        
        miSQL = miSQL & " AND codartic IN (" & CA.item(J) & ") GROUP BY 1"
        miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
        
            miSQL = "UPDATE tmppedprov SET nsal = nsal + " & miRsAux.Fields(1)
            miSQL = miSQL & " WHERE codusu = " & vUsu.codigo
            miSQL = miSQL & " and codartic  = " & DBSet(miRsAux!codArtic, "T")
            conn.Execute miSQL
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        
        
        
        
    Next
End Sub


Private Sub PedProv_CantidadPeriodo(FInicio As Date, ByRef CA As Collection, Periodo1 As Boolean)
Dim J As Integer

    'Vamos a ver en cuantos albaranes, facturas del periodo salen

    For J = 1 To CA.Count
        If Periodo1 Then
            Label3(100).Caption = "Cantidad (I)  " & J & "/" & CA.Count
        Else
            Label3(100).Caption = "Cantidad (II)   " & J & "/" & CA.Count
        End If
        Label3(100).Refresh
        'Docuemtnos en los que esta el articulo
        
        'En cuantos albaranes esta
        miSQL = "select codartic,sum(cantidad) from scaalb,slialb  WHERE"
        miSQL = miSQL & " scaalb.Codtipom = slialb.Codtipom And scaalb.NumAlbar = slialb.NumAlbar"
        miSQL = miSQL & " AND fechaalb >=" & DBSet(FInicio, "F")
        
        'Si esta maarcado NO tiene en cuenta coddirec
        If chkPropPedido(1).Value = 0 Then miSQL = miSQL & " AND coddirec is null "
        
        'Marzo 2011. falta el codalmac
        If Me.txtAlma(7).Text = "" Then
            miSQL = miSQL & " AND codalmac = " & Me.txtAlma(0).Text
        Else
            'MARZO 2014
            'AHORA con el segundo consolidado
            'miSQL = miSQL & " AND codalmac IN (" & Me.txtAlma(0).Text & "," & Me.txtAlma(7).Text & ")"
            miSQL = miSQL & " AND codalmac IN (" & Me.txtAlma(0).Text & "," & Me.txtAlma(7).Text
            If Me.txtAlma(8).Text <> "" Then miSQL = miSQL & "," & Me.txtAlma(8).Text
            miSQL = miSQL & ")"
        End If

        
        miSQL = miSQL & " AND codartic IN (" & CA.item(J) & ") GROUP BY 1"
        miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            
            miSQL = "UPDATE tmppedprov SET "
            If Periodo1 Then
                miSQL = miSQL & "ult_1 = ult_1  + "
            Else
                miSQL = miSQL & "ult_2 = ult_2  + "
            End If
            miSQL = miSQL & TransformaComasPuntos(CStr(miRsAux.Fields(1))) & " WHERE codusu = " & vUsu.codigo
            miSQL = miSQL & " and codartic  = " & DBSet(miRsAux!codArtic, "T")
            conn.Execute miSQL
            
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        
        
        'En facturas
        miSQL = " select codartic,sum(cantidad) from scafac,scafac1,slifac  WHERE"
        miSQL = miSQL & " scafac1.Codtipom = scafac.Codtipom And scafac1.NumFactu = scafac.NumFactu And scafac1.FecFactu = scafac.FecFactu"
        miSQL = miSQL & " AND scafac1.Codtipom = slifac.Codtipom And scafac1.NumFactu = slifac.NumFactu And scafac1.FecFactu = slifac.FecFactu"
        miSQL = miSQL & " and scafac1.numalbar=slifac.numalbar and scafac1.codtipoa=slifac.codtipoa"
        miSQL = miSQL & " AND fechaalb >=" & DBSet(FInicio, "F")
        
        'Si esta maarcado NO tiene en cuenta coddirec
        If chkPropPedido(1).Value = 0 Then miSQL = miSQL & " AND coddirec is null "
        
        'Marzo 2011. falta el codalmac
        If Me.txtAlma(7).Text = "" Then
            miSQL = miSQL & " AND codalmac = " & Me.txtAlma(0).Text
        Else
            'MARZO 2014
            'AHORA con el segundo consolidado
            'miSQL = miSQL & " AND codalmac IN (" & Me.txtAlma(0).Text & "," & Me.txtAlma(7).Text & ")"
            miSQL = miSQL & " AND codalmac IN (" & Me.txtAlma(0).Text & "," & Me.txtAlma(7).Text
            If Me.txtAlma(8).Text <> "" Then miSQL = miSQL & "," & Me.txtAlma(8).Text
            miSQL = miSQL & ")"
        End If

        
        miSQL = miSQL & " AND codartic IN (" & CA.item(J) & ") GROUP BY 1"
        miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            
            miSQL = "UPDATE tmppedprov SET "
            If Periodo1 Then
                miSQL = miSQL & "ult_1 = ult_1  + "
            Else
                miSQL = miSQL & "ult_2 = ult_2  + "
            End If
            miSQL = miSQL & TransformaComasPuntos(CStr(miRsAux.Fields(1))) & " WHERE codusu = " & vUsu.codigo
            miSQL = miSQL & " and codartic  = " & DBSet(miRsAux!codArtic, "T")
            conn.Execute miSQL
            
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        
        
        
        
    Next
End Sub




        'En pedidos proveedor(cliente
' 1  Almacen solicitado
' 2  Almacen consolidado
' 3  Segundo almacen consolidado
Private Sub PedProv_PedidosPendiente(ByRef CA As Collection, Cual_ As Byte)
Dim J As Integer

    'Vamos a ver en cuantos albaranes, facturas del periodo salen
    If Cual_ = 2 And txtAlma(7).Text = "" Then Exit Sub 'Consolidado NO indicado
    If Cual_ = 3 And txtAlma(8).Text = "" Then Exit Sub 'Consolidado NO indicado
    
    For J = 1 To CA.Count
        Label3(100).Caption = "Pedidos pendiente " & J & "/" & CA.Count
        Label3(100).Refresh
        'Docuemtnos en los que esta el articulo
        
        'En pedidos proveedor
        miSQL = "select codartic,sum(cantidad) from slippr,scappr WHERE scappr.numpedpr = slippr.numpedpr "
        If Me.chkPropPedido(0).Value = 0 Then miSQL = miSQL & " AND scappr.obra=0"
        
        
        Select Case Cual_
        Case 1
            miSQL = miSQL & " AND codalmac = " & Me.txtAlma(0).Text
        Case 2
            miSQL = miSQL & " AND codalmac = " & Me.txtAlma(7).Text
            'miSQL = miSQL & " AND codalmac IN (" & Me.txtAlma(0).Text & "," & Me.txtAlma(7).Text & ")"
        Case 3
            miSQL = miSQL & " AND codalmac = " & Me.txtAlma(8).Text
        End Select

        
        miSQL = miSQL & " AND codartic IN (" & CA.item(J) & ") GROUP BY 1"
        miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            
            miSQL = "UPDATE tmppedprov SET pedpro"
            miSQL = miSQL & Cual_
            miSQL = miSQL & " = " & TransformaComasPuntos(CStr(miRsAux.Fields(1)))
            miSQL = miSQL & " WHERE codusu = " & vUsu.codigo
            miSQL = miSQL & " and codartic  = " & DBSet(miRsAux!codArtic, "T")
            conn.Execute miSQL
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        
        
        'En pedidos cliente
        
        miSQL = "select codartic,sum(cantidad) from sliped,scaped WHERE scaped.numpedcl=sliped.numpedcl "
        If Me.chkPropPedido(0).Value = 0 Then miSQL = miSQL & " AND scaped.coddirec is null "
        'Marzo 2011. falta el codalmac
        Select Case Cual_
        Case 1
            miSQL = miSQL & " AND codalmac = " & Me.txtAlma(0).Text
        Case 2
        '    miSQL = miSQL & " AND codalmac IN (" & Me.txtAlma(0).Text & "," & Me.txtAlma(7).Text & ")"
            miSQL = miSQL & " AND codalmac = " & Me.txtAlma(7).Text
        Case 3
            miSQL = miSQL & " AND codalmac = " & Me.txtAlma(8).Text
        End Select

        miSQL = miSQL & " AND codartic IN (" & CA.item(J) & ") GROUP BY 1"
        miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            miSQL = "UPDATE tmppedprov SET pedcli"
            miSQL = miSQL & Cual_ 'almacen consolidado
            miSQL = miSQL & " = " & TransformaComasPuntos(CStr(miRsAux.Fields(1)))
            miSQL = miSQL & " WHERE codusu = " & vUsu.codigo
            miSQL = miSQL & " and codartic  = " & DBSet(miRsAux!codArtic, "T")
            conn.Execute miSQL
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        
        
        
        
    Next
End Sub

Private Sub PedProv_Stock(ByRef CA As Collection, Cual_ As Byte)
Dim J As Integer


    If Cual_ = 2 And txtAlma(7).Text = "" Then Exit Sub 'Consolidado NO indicado
    If Cual_ = 3 And txtAlma(8).Text = "" Then Exit Sub 'Consolidado NO indicado

    For J = 1 To CA.Count
        Label3(100).Caption = Cual_ & "-> Stock " & J & "/" & CA.Count
        Label3(100).Refresh
        'Stock
        miSQL = "select codartic,canstock from salmac where codalmac = "
        Select Case Cual_
        Case 1
            miSQL = miSQL & txtAlma(0).Text
        Case 2
            miSQL = miSQL & txtAlma(7).Text
        Case 3
            miSQL = miSQL & txtAlma(8).Text
        End Select
        miSQL = miSQL & " AND codartic IN (" & CA.item(J) & ") "
        miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
           
            If Val(miRsAux!CanStock) <> 0 Then
                miSQL = "UPDATE tmppedprov SET stock"
                miSQL = miSQL & Cual_
                miSQL = miSQL & " = " & TransformaComasPuntos(CStr(miRsAux.Fields(1)))
                miSQL = miSQL & " WHERE codusu = " & vUsu.codigo
                miSQL = miSQL & " and codartic  = " & DBSet(miRsAux!codArtic, "T")
                conn.Execute miSQL
            End If
            miRsAux.MoveNext
         
        Wend
        miRsAux.Close
 
        
        
    Next
End Sub




Private Sub HacerEscandalloPropuestaPedido(ByRef CA As Collection)
Dim J As Integer
Dim RT As ADODB.Recordset

    Set RT = New ADODB.Recordset




    For J = 1 To CA.Count
        Label3(100).Caption = "Escandallo " & J & "/" & CA.Count
        Label3(100).Refresh
    
        'EsEscandallo,TieneVtasEscandallo
        'Stock
        miSQL = "select codarti1,sum(nsal) ns,sum(ult_1)*cantidad as u1,sum(ult_2)*cantidad as u2,"
        'Octubre 2018. No es pedcli, es pedcli por cada almacen (pedcli1,2 y 3)
        miSQL = miSQL & " sum(pedcli1)*cantidad as p1 ,sum(pedcli2)*cantidad as p2 ,sum(pedcli3)*cantidad as p3"
        miSQL = miSQL & " from tmppedprov,sarti1,sartic "
        miSQL = miSQL & " Where tmppedprov.codArtic = sarti1.codArtic AND tmppedprov.codArtic = sarti1.codArtic And sarti1.codarti1 = sartic.codArtic"
        
        'para que solo salgan los desde hastas marcados
        If txtCodProve(17).Text <> "" Then miSQL = miSQL & " AND sartic.codProve = " & txtCodProve(17).Text
        'Situacion
        If Me.cboProPed(1).ListIndex > 0 Then miSQL = miSQL & " AND sartic.codstatu = " & Me.cboProPed(1).ListIndex - 1
    
        If Me.txtFamia(2).Text <> "" Then miSQL = miSQL & " AND sartic.codfamia >= " & txtFamia(2).Text
        If Me.txtFamia(3).Text <> "" Then miSQL = miSQL & " AND sartic.codfamia <= " & txtFamia(3).Text
        If Me.txtmarca(0).Text <> "" Then miSQL = miSQL & " AND sartic.codmarca >= " & txtmarca(0).Text
        If Me.txtmarca(1).Text <> "" Then miSQL = miSQL & " AND sartic.codmarca <= " & txtmarca(1).Text
    '---------------------------------------------------------------------------------------
        
        
        miSQL = miSQL & " AND tmppedprov.codusu =" & vUsu.codigo  ' 2018/12/28 NOOOOOOOO estaba
        
        'ENero 2019. Miramos la marca de conjunto
        miSQL = miSQL & " AND tmppedprov.codArtic IN ("
        miSQL = miSQL & "       select codartic from sartic where conjunto=1 and codartic IN (" & CA.item(J) & ") )"
        miSQL = miSQL & " GROUP BY 1 "
        
        'If InStr(1, CA.item(j), "0010004426") > 0 Then Sto p
        
        
        miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            
         
            
            'Alguna tiene valor
            If DBLet(miRsAux!ns, "N") = 0 And DBLet(miRsAux!u1, "N") = 0 And DBLet(miRsAux!u2, "N") = 0 And DBLet(miRsAux!p1, "N") = 0 Then
                'TODAS SON CERO. NO hacemos nada
            Else
                miSQL = "Select tmppedprov.*,sartic.codprove,sartic.codfamia,canstock from tmppedprov,sartic,salmac WHERE tmppedprov.codartic=sartic.codartic and salmac.codartic=sartic.codartic"
                miSQL = miSQL & " and salmac.codalmac=" & txtAlma(0).Text & " and tmppedprov.codusu=" & vUsu.codigo
                miSQL = miSQL & " and tmppedprov.codartic= " & DBSet(miRsAux!codarti1, "T")
                
 
                RT.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If RT.EOF Then
                    'NO EXISTIA
                    'Hay que crearlo con las cantidades adecuadas
                    'insert into `tmppedprov` (`codusu`,`codprove`,`codfamia`,`codartic`,`nsal`,`ult_1`,`ult_2`,`pedcli`,`pedpro`,`stock`,`permin`,`permax`,`caja`,`De1Cliente`,`EsEscandallo`,`TieneVtasEscandallo`) values ( '1','22','1','1','0','0','0','0','0','1','0','0','0','0','1','0')
                    MsgBox "Escandallo no encontrado. Avise soporte tecnico. El programa continuará", vbExclamation
                    miSQL = "select * from tmpcrmmsg where codusu=-1"  'para que no de error y se salga
                Else
                    miSQL = "UPDATE tmppedprov SET esescandallo =1 ,nsal=nsal + " & DBLet(miRsAux!ns, "N")
                    miSQL = miSQL & " ,ult_1=ult_1 + " & DBLet(miRsAux!u1, "N")
                    miSQL = miSQL & " ,ult_2=ult_2 + " & DBLet(miRsAux!u2, "N")
                    miSQL = miSQL & " ,pedcli1=pedcli1 + " & DBLet(miRsAux!p1, "N")
                    miSQL = miSQL & " ,pedcli2=pedcli2 + " & DBLet(miRsAux!p2, "N")
                    miSQL = miSQL & " ,pedcli3=pedcli3 + " & DBLet(miRsAux!p3, "N")
                    miSQL = miSQL & " WHERE codusu = " & vUsu.codigo & " AND codartic= " & DBSet(miRsAux!codarti1, "T")
                    
                End If
                conn.Execute miSQL
                RT.Close
                
            End If
            miRsAux.MoveNext
        Wend
        miRsAux.Close
 
        
        
    Next
    Set RT = Nothing
    
End Sub















 


Private Sub MontaSQLVtasAgente()
    
    miSQL = miSQL & " FROM scafac, slifac,sartic,smarca,sclien WHERE scafac.codtipom=slifac.codtipom AND"
    miSQL = miSQL & " scafac.numfactu = slifac.numfactu AND scafac.fecfactu = slifac.fecfactu AND"
    miSQL = miSQL & " slifac.codArtic = sartic.codArtic AND  sartic.codmarca = smarca.codmarca "
    miSQL = miSQL & " AND scafac.codclien=sclien.codclien "
    
    
    'Lo vuelvo a poner
    'Enero 2016. Manolo Belarte. Lo quitamos
    'miSQL = miSQL & " AND slifac.codartic <> 'TASA RECICLAR'"
    
    'El D/H
    'Antes Nov 2013
    'If chkResVtaAgen(1).Value = 0 Then miSQL = miSQL & " AND scafac.codtipom <> 'FAZ'"
    miSQL = miSQL & " AND scafac.codtipom"
    If chkResVtaAgen(1).Value = 1 Then
        miSQL = miSQL & " = "
    Else
        miSQL = miSQL & " <> "
    End If
    miSQL = miSQL & " 'FAZ'"
    
    'portes
    If chkResVtaAgen(2).Value = 0 And vParamAplic.ArtPortesN <> "" Then miSQL = miSQL & " AND slifac.codartic <> " & DBSet(vParamAplic.ArtPortesN, "T")
    If chkResVtaAgen(3).Value = 0 Then miSQL = miSQL & " AND scafac.codtipom <> 'FRT'"
    
    If Me.txtFecha(39).Text <> "" Then miSQL = miSQL & " AND scafac.fecfactu >= " & DBSet(txtFecha(39).Text, "F")
    If Me.txtFecha(40).Text <> "" Then miSQL = miSQL & " AND scafac.fecfactu <= " & DBSet(txtFecha(40).Text, "F")
    
    If Me.txtmarca(4).Text <> "" Then miSQL = miSQL & " AND sartic.codmarca >= " & DBSet(txtmarca(4).Text, "N")
    If Me.txtmarca(5).Text <> "" Then miSQL = miSQL & " AND sartic.codmarca <= " & DBSet(txtmarca(5).Text, "N")
    
    codigo = "scafac.codagent"
    If Me.chkResVtaAgen(4).Value = 1 Then codigo = "sclien.visitador"
    If Me.txtAgente(4).Text <> "" Then miSQL = miSQL & " AND " & codigo & " >= " & DBSet(txtAgente(4).Text, "N")
    If Me.txtAgente(5).Text <> "" Then miSQL = miSQL & " AND " & codigo & " <= " & DBSet(txtAgente(5).Text, "N")
    codigo = ""
    
    
    
    
End Sub


Private Function CargaDatosResumenVtaAgente() As Boolean
Dim ColAgent As Collection
Dim J As Integer
Dim marca As Integer
Dim Aux As Currency
Dim Llevo As Currency
'Dim LlevoB As Currency  Ya no hay B, ahora es ECO
Dim LlevoEco_ As Currency
Dim LlevoSuperEco As Currency
Dim RS As ADODB.Recordset
Dim Visitador As Integer
Dim CodAgente As Integer
Dim Reestablecer As Boolean
    On Error GoTo ECargaDatosResumenVtaAgente
    CargaDatosResumenVtaAgente = False
    
       
    'Veremos los agentes
    Label3(122).Caption = "Obteniendo datos"
    Label3(122).Refresh
    If chkResVtaAgen(4).Value = 1 Then
        'Visitador
        miSQL = "SELECT  distinct(visitador)  "
    Else
        miSQL = "SELECT  distinct(scafac.codagent)  "
    End If
    MontaSQLVtasAgente  'añade los where ....
    miSQL = miSQL & " ORDER BY 1   "
       
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Set ColAgent = New Collection
    While Not miRsAux.EOF
        ColAgent.Add CStr(miRsAux.Fields(0))
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If ColAgent.Count = 0 Then
        MsgBox "No existen datos", vbExclamation
        Set ColAgent = Nothing
        Exit Function
    End If
    
    
    'FACTURAS
    '---------------------------------
    For J = 1 To ColAgent.Count
        Label3(122).Caption = "Agente " & J & " de " & ColAgent.Count
        Label3(122).Refresh
        
        
        miSQL = "SELECT  sartic.codmarca,scafac.codtipom,importel,scafac.dtoppago, scafac.dtognral,pvpInferior,scafac.codagent,visitador "
        MontaSQLVtasAgente  'añade los where ....
        If chkResVtaAgen(4).Value = 1 Then
            miSQL = miSQL & " AND visitador = " & ColAgent.item(J)
        Else
            miSQL = miSQL & " AND scafac.codagent = " & ColAgent.item(J)
        End If
        miSQL = miSQL & "  ORDER BY codmarca"
        If Me.chkResVtaAgen(4).Value = 1 Then miSQL = miSQL & ",visitador,codagent"
        
        miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        marca = -1
        miSQL = ""
        Visitador = -1
        While Not miRsAux.EOF
            Label3(122).Caption = ColAgent.item(J) & ".  Marca:  " & miRsAux!codmarca
            Label3(122).Refresh
            
            Reestablecer = False
            
            If miRsAux!codmarca <> marca Then
                Reestablecer = True
            Else
                If Me.chkResVtaAgen(4).Value = 1 Then
                    If miRsAux!Visitador <> Visitador Then
                         Reestablecer = True
                    Else
                        If miRsAux!CodAgent <> CodAgente Then Reestablecer = True
                    End If
                End If
            End If
            If Reestablecer Then
                'Otra marca
                If marca >= 0 Then
                    NumRegElim = NumRegElim + 1
                    'miSQL = miSQL & ", (" & vUsu.Codigo & "," & NumRegElim & "," & ColAgent.item(J) & "," & Marca & ","
                    miSQL = miSQL & ", (" & vUsu.codigo & "," & NumRegElim & "," & CodAgente & "," & marca & "," & Visitador & ","
                    miSQL = miSQL & DBSet(Llevo, "N") & "," & DBSet(LlevoEco_, "N") & "," & DBSet(LlevoSuperEco, "N") & ",0,0,0)"
                End If
                 
                'Reseteo
                Llevo = 0
                LlevoEco_ = 0
                LlevoSuperEco = 0
                marca = miRsAux!codmarca
                Visitador = miRsAux!Visitador
                CodAgente = miRsAux!CodAgent
            End If
            
            Aux = miRsAux!ImporteL
            If miRsAux!DtoPPago <> 0 Or miRsAux!DtoGnral <> 0 Then
                'Lleva algun descuento. De momento solo trato dtos aditivos
                Aux = Aux * ((100 - miRsAux!DtoPPago) / 100)
                Aux = Aux * ((100 - miRsAux!DtoGnral) / 100)
                Aux = Round(Aux, 2)
            End If
            
            
            
            
            
            If miRsAux!PVPInferior = 1 Then
                LlevoEco_ = LlevoEco_ + Aux
                
            ElseIf miRsAux!PVPInferior = 2 Then
                LlevoSuperEco = LlevoSuperEco + Aux
            Else
                Llevo = Llevo + Aux
            End If
            
'            If miRsAux!codtipom = "FAZ" Then
'                LlevoB = LlevoB + Aux
'            Else
'                Llevo = Llevo + Aux
'            End If
            
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        
    
        'El ultimo
        NumRegElim = NumRegElim + 1
        miSQL = miSQL & ", (" & vUsu.codigo & "," & NumRegElim & "," & CodAgente & "," & marca & "," & Visitador & ","
        miSQL = miSQL & DBSet(Llevo, "N") & "," & DBSet(LlevoEco_, "N") & "," & DBSet(LlevoSuperEco, "N") & ",0,0,0)"
        
        miSQL = Mid(miSQL, 2)
        miSQL = "INSERT INTO tmpinformes(codusu,codigo1,campo1,campo2,nombre3,importe1,importe2,importe5,importe3,importe4,importeb1) VALUES " & miSQL
        conn.Execute miSQL

    
    
    Next

    
    
    'Albaranes.
    'Si salen tb los albaranes
    If chkResVtaAgen(0).Value Then
            
            Label3(122).Caption = "Albaranes"
            Label3(122).Refresh
            miSQL = " SELECT  sartic.codmarca,scaalb.codagent,"
            'JULIO2013
            'SUM(IF(scaalb.codtipom='ALZ',0,importel)) ,SUM(IF(scaalb.codtipom='ALZ',importel,0) )"
            miSQL = miSQL & " SUM(IF(slialb.pvpinferior=0,importel,0))  ,SUM(IF(slialb.pvpinferior=1,importel,0) )"
            miSQL = miSQL & " ,SUM(IF(slialb.pvpinferior=2,importel,0)), visitador"
            miSQL = miSQL & " FROM scaalb, slialb,sartic,smarca,sclien WHERE scaalb.codtipom=slialb.codtipom AND"
            miSQL = miSQL & " scaalb.NumAlbar = slialb.NumAlbar And slialb.codArtic = sartic.codArtic And "
            miSQL = miSQL & " sartic.codmarca = smarca.codmarca AND scaalb.codclien=sclien.codclien"
            
            'antes NOV
            'If chkResVtaAgen(1).Value = 0 Then miSQL = miSQL & " AND scaalb.codtipom <> 'ALZ'"
            miSQL = miSQL & " AND scaalb.codtipom "
            If chkResVtaAgen(1).Value = 0 Then
                miSQL = miSQL & " <> "
            Else
                miSQL = miSQL & " = "
            End If
                
            miSQL = miSQL & "  'ALZ'"
            If Me.txtFecha(39).Text <> "" Then miSQL = miSQL & " AND scaalb.fechaalb >= " & DBSet(txtFecha(39).Text, "F")
            If Me.txtFecha(40).Text <> "" Then miSQL = miSQL & " AND scaalb.fechaalb <= " & DBSet(txtFecha(40).Text, "F")
            
            If Me.txtmarca(4).Text <> "" Then miSQL = miSQL & " AND sartic.codmarca >= " & DBSet(txtmarca(4).Text, "N")
            If Me.txtmarca(5).Text <> "" Then miSQL = miSQL & " AND sartic.codmarca <= " & DBSet(txtmarca(5).Text, "N")
            
            codigo = "scaalb.codagent"
            If Me.chkResVtaAgen(4).Value = 1 Then codigo = "sclien.visitador"
            If Me.txtAgente(4).Text <> "" Then miSQL = miSQL & " AND " & codigo & " >= " & DBSet(txtAgente(4).Text, "N")
            If Me.txtAgente(5).Text <> "" Then miSQL = miSQL & " AND " & codigo & " <= " & DBSet(txtAgente(5).Text, "N")
            
            
            
            'Vamos a quitar ciertos articulos
            'If vParamAplic.ArtReciclado <> "" Then miSQL = miSQL & " AND slialb.codartic <> " & DBSet(vParamAplic.ArtReciclado, "T")
            miSQL = miSQL & " GROUP BY 1,2"
            'If Me.chkResVtaAgen(4).Value = 1 Then  miSQL = miSQL & ",visitador"
            
            
            
            Set RS = New ADODB.Recordset
            miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not miRsAux.EOF
            
                Label3(122).Caption = miRsAux!CodAgent & " / " & miRsAux!codmarca
                Label3(122).Refresh
                
                miSQL = "Select codigo1 from tmpinformes where codusu = " & vUsu.codigo & " AND campo1 = " & miRsAux!CodAgent
                miSQL = miSQL & " AND campo2 = " & miRsAux!codmarca
                RS.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If RS.EOF Then
                   'NUEVO
                    NumRegElim = NumRegElim + 1
                    miSQL = " (" & vUsu.codigo & "," & NumRegElim & "," & miRsAux!CodAgent & "," & miRsAux!codmarca & ","
                    miSQL = miSQL & DBSet(DBLet(miRsAux.Fields(2), "N"), "N") & "," & DBSet(DBLet(miRsAux.Fields(3), "N"), "N") & ","
                    miSQL = miSQL & DBSet(DBLet(miRsAux.Fields(4), "N"), "N") & ",0,0,0," & miRsAux!Visitador & ")"
                    miSQL = "INSERT INTO tmpinformes(codusu,codigo1,campo1,campo2,importe3,importe4,importe5,importe1,importe2,importeb1,nombre3) VALUES " & miSQL
               
                 Else
                    miSQL = "UPDATE tmpinformes SET importe3= " & DBSet(DBLet(miRsAux.Fields(2), "N"), "N")
                    miSQL = miSQL & " , importe4= " & DBSet(DBLet(miRsAux.Fields(3), "N"), "N")
                    miSQL = miSQL & " , importeb1= " & DBSet(DBLet(miRsAux.Fields(4), "N"), "N")
                    miSQL = miSQL & " WHERE codusu = " & vUsu.codigo & " AND codigo1 = " & RS!Codigo1
                 
                 End If
                 RS.Close
                conn.Execute miSQL
                miRsAux.MoveNext
            Wend
            miRsAux.Close
            
            
            
            
            
            Set RS = Nothing
    End If

    If NumRegElim > 0 Then
            Label3(122).Caption = "Agente"
            Label3(122).Refresh
            miSQL = "Select distinct(campo1) from tmpinformes where codusu = " & vUsu.codigo
            miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not miRsAux.EOF
                campo = DevuelveDesdeBD(conAri, "nomagent", "sagent", "codagent", miRsAux.Fields(0), "N")
                miSQL = "UPDATE tmpinformes set nombre1=" & DBSet(campo, "T")
                miSQL = miSQL & " where codusu = " & vUsu.codigo & " AND campo1 = " & miRsAux.Fields(0)
                miRsAux.MoveNext
                conn.Execute miSQL
            Wend
            miRsAux.Close
            Label3(122).Caption = "Obt agente,marca"
            Label3(122).Refresh
            
            miSQL = "Select distinct(campo2) from tmpinformes where codusu = " & vUsu.codigo
            miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not miRsAux.EOF
                campo = DevuelveDesdeBD(conAri, "nommarca", "smarca", "codmarca", miRsAux.Fields(0), "N")
                miSQL = "UPDATE tmpinformes set nombre2=" & DBSet(campo, "T")
                miSQL = miSQL & " where codusu = " & vUsu.codigo & " AND campo2 = " & miRsAux.Fields(0)
                miRsAux.MoveNext
                conn.Execute miSQL
            Wend
            miRsAux.Close
            If Me.chkResVtaAgen(4).Value = 1 Then
                Label3(122).Caption = "Visitador"
                Label3(122).Refresh
            
                miSQL = "Select distinct(nombre3) from tmpinformes where codusu = " & vUsu.codigo
                miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                While Not miRsAux.EOF
                    
                    campo = DevuelveDesdeBD(conAri, "nomagent", "sagent", "codagent", miRsAux.Fields(0), "N")
                    campo = Format(miRsAux.Fields(0), "0000") & " - " & campo
                    miSQL = "UPDATE tmpinformes set nombre3=" & DBSet(campo, "T")
                    miSQL = miSQL & " where codusu = " & vUsu.codigo & " AND nombre3 = '" & miRsAux.Fields(0) & "'"
                    miRsAux.MoveNext
                    conn.Execute miSQL
                Wend
                miRsAux.Close
             End If
            CargaDatosResumenVtaAgente = True
    Else
        MsgBox "ningun datos entre esos parametros", vbExclamation
    End If

    Exit Function
ECargaDatosResumenVtaAgente:
    MuestraError Err.Number, Err.Description
End Function



'-------------Comparativo agentes
Private Sub ComparativoAgentes()
    
  
    conn.Execute "DELETE FROM tmpinformes where codusu = " & vUsu.codigo
    'codusu,codigo1,nombre1,nombre2,importe1,importe2,fecha1
    campo = "Select " & vUsu.codigo & "," & "codagent,codfamia,sum(cantidad),"
    'Febrero 2012
    campo = campo & " sum(importel-round((importel*(dtoppago+dtognral))/100,2))"
    
    campo = campo & " ,sartic.codprove,0,0 FROM scafac,slifac,sartic  WHERE scafac.codtipom=slifac.codtipom and scafac.numfactu=slifac.numfactu and scafac.fecfactu=slifac.fecfactu"
    campo = campo & " AND slifac.codartic=sartic.codartic AND scafac.codtipom"
    If chkBenAge(6).Value = 1 Then
        campo = campo & " ="
    Else
        campo = campo & " <>"
    End If
    campo = campo & " 'FAZ' "     'En el compartivo SI salen los articulos de varios. comento el trozo:    AND sartic.artvario =0"
    
    
    
    codigo = Replace(cadSelect, "{", "(")
    codigo = Replace(codigo, "}", ")")
    codigo = campo & " AND " & codigo & " GROUP BY 1,2,3"
    codigo = "INSERT INTO tmpinformes(codusu,codigo1,campo1,importe1,importe2,nombre1,importe3,importe4) " & codigo
    conn.Execute codigo
    
    'peridod ANTERIOR
    codigo = Replace(cadSelect, "{", "(")
    codigo = Replace(codigo, "}", ")")

    'replace de fecha
    codigo = Replace(codigo, "'" & txtAnyo(0).Text & "-", "'" & CStr(CInt(txtAnyo(0).Text) - 1) & "-")
    codigo = campo & " AND " & codigo & " GROUP BY 1,2,3"
    codigo = "INSERT INTO tmpinformes(codusu,codigo1,campo1,importe3,importe4,nombre1,importe1,importe2) " & codigo
    conn.Execute codigo
    
    
    'Cojo el proveedor y en nombre 2 pongo el nomprove
    Set miRsAux = New ADODB.Recordset
    codigo = "Select nombre1 from tmpinformes where codusu=" & vUsu.codigo & " GROUP BY 1"
    miRsAux.Open codigo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        codigo = miRsAux!nombre1
        codigo = DevuelveDesdeBD(conAri, "nomprove", "sprove", "codprove", codigo, "N")
        codigo = "UPDATE tmpinformes SET nombre2=" & DBSet(codigo, "T") & ",nombre1 = '" & Format(miRsAux!nombre1, "00000") & "'"
        codigo = codigo & " WHERE codusu = " & vUsu.codigo & " AND nombre1 = " & miRsAux!nombre1
        miRsAux.MoveNext
        conn.Execute codigo
    
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    
    
End Sub


Private Sub AgrupaVtasxProveedorxAgente()
Dim J As Integer

    Label3(142).Caption = "Obteniendo datos"  'indicador
    Label3(142).Refresh

    conn.Execute "DELETE FROM tmpinformes WHERE codusu = " & vUsu.codigo

        
    If cadSelect <> "" Then
        cadSelect = Replace(cadSelect, "{", "(")
        cadSelect = Replace(cadSelect, "}", ")")
    End If

    'JULIO 2013
    'Puede detallar cliente articulo
    'nombre1,nombre2,nombre3   codartic codclien nomclien
    
    
    'Campo2: campo1 los cambiamos para poder poner el codclien
    
    codigo = "Select " & vUsu.codigo & ",codprove,scafac.codagent,sum(importel),0"
    codigo = codigo & " ,sclien.codclien,sclien.nomclien,slifac.codartic,slifac.nomartic,sum(cantidad),0"
    
    codigo = codigo & " FROM scafac,slifac,sclien,sartic WHERE  scafac.fecfactu = slifac.fecfactu AND"
    codigo = codigo & " scafac.NumFactu = slifac.NumFactu And sclien.codclien = scafac.codclien And slifac.codArtic = sartic.codArtic"
    codigo = codigo & " and scafac.codtipom=slifac.codtipom  AND " & cadSelect
    codigo = codigo & " GROUP BY codprove,codagent,sclien.codclien,slifac.codartic "
    codigo = "INSERT INTO tmpinformes(codusu,codigo1,campo2,importe1,importe2,campo1,nombre1,nombre2,nombre3,importe3,importe4) " & codigo
    conn.Execute codigo
    
    'El comparativo
    Label3(142).Caption = "Comparativo"  'indicador
    Label3(142).Refresh
    devuelve = CStr(Year(Format(txtFecha(9).Text, FormatoFecha)))
    campo = "'" & Val(devuelve) - 1 & "-"
    devuelve = "'" & devuelve & "-"
    miSQL = Replace(cadSelect, devuelve, campo)
    codigo = "Select " & vUsu.codigo & ",codprove,scafac.codagent,sum(importel),0 "
    codigo = codigo & " ,sclien.codclien,sclien.nomclien,slifac.codartic,slifac.nomartic,0,sum(cantidad)"
    codigo = codigo & " FROM scafac,slifac,sclien,sartic WHERE  scafac.fecfactu = slifac.fecfactu AND"
    codigo = codigo & " scafac.NumFactu = slifac.NumFactu And sclien.codclien = scafac.codclien And slifac.codArtic = sartic.codArtic"
    codigo = codigo & " and scafac.codtipom=slifac.codtipom  AND " & miSQL
    codigo = codigo & " GROUP BY codprove,codagent,sclien.codclien,slifac.codartic "
    codigo = "INSERT INTO tmpinformes(codusu,codigo1,campo2,importe2,importe1,campo1,nombre1,nombre2,nombre3,importe3,importe4) " & codigo
    conn.Execute codigo
    
    
    
 
    
    If Me.txtImporte(1).Text <> "" Then
        Label3(142).Caption = "Aplicando importe minimo"  'indicador
        Label3(142).Refresh

        'Quiere un minimo
        ImpTot = ImporteFormateado(txtImporte(1).Text)
        codigo = "DELETE FROM tmpinformes WHERE importe1 < " & DBSet(ImpTot, "N") & " AND codusu = " & vUsu.codigo
        conn.Execute codigo
        
    End If
    
    
End Sub



Private Sub AgrupaVtasxProveedorxFamilia()
Dim J As Integer
    
    Label3(142).Caption = "Obteniendo datos"  'indicador
    Label3(142).Refresh
    'conn.Execute "DELETE FROM tmpinformes WHERE codusu = " & vUsu.codigo
    BorrarTempInformes
    
        
    If cadSelect <> "" Then
        cadSelect = Replace(cadSelect, "{", "(")
        cadSelect = Replace(cadSelect, "}", ")")
    End If


    codigo = "Select " & vUsu.codigo & ",codprove,sartic.codfamia,sum(importel),sum(cantidad),0,0 FROM scafac,slifac,sclien,sartic WHERE  scafac.fecfactu = slifac.fecfactu AND"
    codigo = codigo & " scafac.NumFactu = slifac.NumFactu And sclien.codclien = scafac.codclien And slifac.codArtic = sartic.codArtic"
    codigo = codigo & " and scafac.codtipom=slifac.codtipom  AND " & cadSelect
    codigo = codigo & " GROUP BY 1,2 "
    codigo = "INSERT INTO tmpinformes(codusu,codigo1,campo1,importe1,importe3,importe2,importe4) " & codigo
    conn.Execute codigo
    
    
    'Minimo para comparativo
    If Me.txtImporte(1).Text <> "" Then
        Label3(142).Caption = "Aplicando importe minimo"  'indicador
        Label3(142).Refresh

        'Quiere un minimo
        ImpTot = ImporteFormateado(txtImporte(1).Text)
        codigo = "DELETE FROM tmpinformes WHERE importe1 < " & DBSet(ImpTot, "N") & " AND codusu = " & vUsu.codigo
        conn.Execute codigo
        
    End If
    
    
    
    
    
    'El comparativo
    Label3(142).Caption = "Compartivo"  'indicador
    Label3(142).Refresh
    devuelve = CStr(Year(Format(txtFecha(9).Text, FormatoFecha)))
    campo = "'" & Val(devuelve) - 1 & "-"
    devuelve = "'" & devuelve & "-"
    miSQL = Replace(cadSelect, devuelve, campo)
    codigo = "Select " & vUsu.codigo & ",codprove,sartic.codfamia,sum(importel),sum(cantidad),0,0 FROM scafac,slifac,sclien,sartic WHERE  scafac.fecfactu = slifac.fecfactu AND"
    codigo = codigo & " scafac.NumFactu = slifac.NumFactu And sclien.codclien = scafac.codclien And slifac.codArtic = sartic.codArtic"
    codigo = codigo & " and scafac.codtipom=slifac.codtipom  AND " & miSQL
    codigo = codigo & " GROUP BY 1,2 "
    codigo = "INSERT INTO tmpinformes(codusu,codigo1,campo1,importe2,importe4,importe1,importe3) " & codigo
    conn.Execute codigo
    
    
    'Minimo para comparativo
    If Me.txtImporte(1).Text <> "" Then
        Label3(142).Caption = "Aplicando importe minimo"  'indicador
        Label3(142).Refresh

        'Quiere un minimo
        ImpTot = ImporteFormateado(txtImporte(1).Text)
        codigo = "DELETE FROM tmpinformes WHERE importe2<>0 and importe2 < " & DBSet(ImpTot, "N") & "  AND codusu = " & vUsu.codigo
        conn.Execute codigo
        
    End If
    
    
    
    
    
    
    Label3(142).Caption = "Leyendo familia"  'indicador
    Label3(142).Refresh
    Set miRsAux = New ADODB.Recordset
    codigo = "Select distinct(campo1) from tmpinformes where codusu = " & vUsu.codigo & " ORDER BY 1"
    miRsAux.Open codigo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Label3(142).Caption = "Fam:" & miRsAux!campo1  'indicador
        Label3(142).Refresh
        
        codigo = DevuelveDesdeBD(conAri, "nomfamia", "sfamia", "codfamia", miRsAux!campo1)
        codigo = "UPDATE tmpinformes set nombre1 = " & DBSet(codigo, "T")
        codigo = codigo & " WHERE campo1 = " & miRsAux!campo1 & " AND codusu = " & vUsu.codigo
        conn.Execute codigo
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
        

        

End Sub



Private Function DevuelvePrecioCosteListado(QueIndex As Integer, ParaSQL As Boolean) As String
    If ParaSQL Then
        If Me.cboCoste(QueIndex).ListIndex = 0 Then
            DevuelvePrecioCosteListado = "preciouc"
        ElseIf Me.cboCoste(QueIndex).ListIndex = 1 Then
            DevuelvePrecioCosteListado = "preciomp"
        Else
            DevuelvePrecioCosteListado = "preciost"
        End If
        
    Else
        If Me.cboCoste(QueIndex).ListIndex = 0 Then
            DevuelvePrecioCosteListado = "Ult. compra"
        ElseIf Me.cboCoste(QueIndex).ListIndex = 1 Then
            DevuelvePrecioCosteListado = "Precio Medio Pond."
        Else
            DevuelvePrecioCosteListado = "Precio St."
        End If
        
    End If
        
End Function


'El importe minimo afecta a la suma para un proveedor, con lo cual tenemos que grabar en tmp
Private Sub InsertarTmpBeneAgeProv()
Dim Col As Collection
Dim KK As Integer
Dim ExisteDto As Boolean



    On Error GoTo eInsertarTmpBeneAgeProv
    
    BorrarTempInformes
    
    'select codigo1 codprove,campo1 codagent, campo2 codfamia,nombre1 codartic,nombre2 nomartic,
    'importe1 cantidad,importe2 importel ,importe3 preciouc,nombre3 nomagent,
    'fecha1 FecFactu
    Label3(147).Caption = "Leyendo registros"
    Label3(147).Refresh
    
    
    
    
    
    miSQL = "Select " & vUsu.codigo & ",sartic.codprove,scafac.codagent,codfamia,slifac.codartic,slifac.nomartic,cantidad,"
    'Febrero2012. Tendra en cuenta el dtoppago,dtognral
    miSQL = miSQL & "ImporteL -round((importel * (scafac.dtoppago+scafac.dtognral)/100),2),"
    'miSQL = miSQL & "slifac.preciouc,nomagent,scafac.fecfactu FROM "
    miSQL = miSQL & "slifac." & DevuelvePrecioCosteListado(0, True) & "*cantidad,nomagent,scafac.fecfactu,sartic.codmarca FROM "
    miSQL = miSQL & "scafac,slifac,sartic,sagent"
    
    If InStr(1, cadSelect, "{sclien") > 0 Then miSQL = miSQL & ",sclien"
    
    miSQL = miSQL & " WHERE scafac.codtipom=slifac.codtipom and scafac.numfactu=slifac.numfactu and scafac.fecfactu=slifac.fecfactu "
    miSQL = miSQL & " AND slifac.codartic=sartic.codartic AND scafac.codagent=sagent.codagent "
    If InStr(1, cadSelect, "{sclien") > 0 Then miSQL = miSQL & " AND sclien.codclien=scafac.codclien "
    campo = " <> "
    If chkBenAge(6).Value = 1 Then campo = " = "
    campo = " AND scafac.codtipom " & campo & "'FAZ'"
    miSQL = miSQL & campo
    
    If cadSelect <> "" Then
        campo = QuitarCaracterACadena(cadSelect, "{")
        campo = QuitarCaracterACadena(campo, "}")
       
        miSQL = miSQL & " AND " & campo
   
   End If
    
   campo = "INSERT INTO tmpinformes(codusu,codigo1 ,campo1 , campo2 ,nombre1 ,nombre2 ,importe1 ,importe2 ,importe3 ,nombre3 ,fecha1,importeb1)  "
   campo = campo & miSQL
   conn.Execute campo
   
   
   
    If Me.chkBenAge(9).Value = 1 Then
        'Aplica DTO
        'Aplicando descuentos al coste
        Label3(147).Caption = "Leyendo descuentos"
        Label3(147).Refresh
        Set Col = New Collection
        
        'Para ello agruparemos por proveedores,codfamia
        codigo = "select codigo1,campo2 from tmpinformes  where codusu =" & vUsu.codigo & " group by 1,2 order by 1,2"
        Set miRsAux = New ADODB.Recordset
        miRsAux.Open codigo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        codigo = ""
        KK = 0
        While Not miRsAux.EOF
            KK = KK + 1
            codigo = codigo & ", (" & miRsAux!Codigo1 & "," & miRsAux!campo2 & ")"
            If KK > 30 Then
                codigo = Mid(codigo, 2)
                Col.Add codigo
                codigo = ""
            End If
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        
        If codigo <> "" Then
            codigo = Mid(codigo, 2)
            Col.Add codigo
        End If
        
        ExisteDto = False 'Si hay que ejecutar el update
        If Col.Count > 0 Then
            For KK = 1 To Col.Count
                'Montamos el SQL
                Label3(147).Caption = KK & " / " & Col.Count
                Label3(147).Refresh
                codigo = "select * from sdtomp where dtosincargo>0 and (codprove,codfamia) in ( " & Col.item(KK) & ") ORDER BY 1,2,3"
                miRsAux.Open codigo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If Not miRsAux.EOF Then
                    ExisteDto = True
                    While Not miRsAux.EOF
                        codigo = "UPDATE tmpinformes set importe4=" & DBSet(miRsAux!dtosincargo, "N") & " WHERE codusu =" & vUsu.codigo
                        codigo = codigo & " AND codigo1= " & miRsAux!Codprove & " AND campo2 = " & miRsAux!Codfamia
                        If Not IsNull(miRsAux!codmarca) Then codigo = codigo & " AND importeb1 =" & miRsAux!codmarca
                        conn.Execute codigo
                        miRsAux.MoveNext
                    Wend
                End If
                miRsAux.Close
            Next KK
               
            If ExisteDto Then
                'Es que ha habiado alguna actualizacion del coste por la columna dtosincargo
                codigo = " update tmpinformes set importe3=(importe3*(100-importe4))/100"
                codigo = codigo & " Where CodUsu = " & vUsu.codigo & " And importe4 > 0"
                conn.Execute codigo
            End If
        End If
 
   
   
   
   
   End If
eInsertarTmpBeneAgeProv:
    If Err.Number <> 0 Then MuestraError Err.Number
    Set miRsAux = Nothing
End Sub

Private Sub QuitarProveedoresImporteMenor()
Dim EsDelComparativo As Boolean
    ImpTot = ImporteFormateado(txtImporte(2).Text)
    
    '    codigo = "{slifac.importel} >= " & TransformaComasPuntos(CStr(ImpTot))
    Label3(147).Caption = "Aplicando importe prov <" & txtImporte(2).Text
    Label3(147).Refresh
    
    EsDelComparativo = False
    If Opcion = 37 Then
        If chkBenAge(2).Value = 1 Then
            'Del comparativo quitaremos aquellos que en el perido actual NO haya superado las ventas minimas
            EsDelComparativo = True
        End If
    End If
    
    Set miRsAux = New ADODB.Recordset
    If EsDelComparativo Then
        '----------------------------------------------------------------
        'Comparativo AGENTES
        miSQL = "nombre1 ,sum(importe2) " 'importe del actual
        
    Else
        miSQL = "codigo1 ,sum(importe2) "
    End If
    miSQL = "select " & miSQL & " from tmpinformes where codusu=" & vUsu.codigo
    miSQL = miSQL & " GROUP BY 1 HAVING sum(importe2)<" & TransformaComasPuntos(CStr(ImpTot))
    
     miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Label3(147).Caption = "Prov: " & miRsAux.Fields(0)
        Label3(147).Refresh
        If EsDelComparativo Then
            miSQL = "DELETE FROM tmpinformes where codusu = " & vUsu.codigo & " AND nombre1='" & miRsAux!nombre1 & "'"
        Else
            miSQL = "DELETE FROM tmpinformes where codusu = " & vUsu.codigo & " AND codigo1=" & miRsAux!Codigo1
        End If
        miRsAux.MoveNext
        conn.Execute miSQL
    Wend
    miRsAux.Close

    Set miRsAux = Nothing
End Sub



Private Sub benexClien()
Dim Col As Collection
Dim KK As Integer
    
    Label3(156).Caption = "Prepara datos"
    Label3(156).Refresh
    conn.Execute "DELETE FROM tmpinformes where codusu = " & vUsu.codigo
    Espera 0.2
    'codusu,codigo1,nombre1,nombre2,importe1,importe2,fecha1
    Label3(156).Caption = "Leyendo BD"
    Label3(156).Refresh
    codigo = "Select  " & vUsu.codigo & ",codclien,sartic.codmarca,slifac.codartic,sum(cantidad),"
    If True Then
        codigo = codigo & " sum(importel)"
    Else
        codigo = codigo & " sum(ImporteL -round((importel * (dtoppago+dtognral)/100),2))"
        cadSelect = cadSelect & " AND scafac.codtipom  <> 'FAZ' AND sartic.artvario =0  "
    End If
    codigo = codigo & " ,slifac.nomartic,nommarca,slifac." & DevuelvePrecioCosteListado(1, True)
    
    'Para aplicar dto necesito familia y proveedor
    codigo = codigo & " ,codfamia,codprove"
    'Para hace lo mismo que BenexProv y BenexAgen habria que multiplicarlo por cantidad y en el rpt el campo coste que sea directamente tmpimport3
    codigo = codigo & " FROM scafac,slifac,sartic ,smarca WHERE scafac.codtipom = slifac.codtipom And scafac.NumFactu = "
    codigo = codigo & " slifac.NumFactu And scafac.FecFactu = slifac.FecFactu AND slifac.codartic=sartic.codartic  and sartic.codmarca=smarca.codmarca"
    If cadSelect <> "" Then
        campo = Replace(cadSelect, "{", "(")
        campo = Replace(campo, "}", ")")
        codigo = codigo & " AND " & campo
    End If
    codigo = codigo & " group by 2,3,4 "
    codigo = "INSERT INTO tmpinformes(codusu,codigo1,campo1,nombre1,importe1,importe2,nombre2,nombre3,importe3,campo2,importeb1) " & codigo
    
    conn.Execute codigo
    

    'Quio los ceros
    Label3(156).Caption = "Ceros"
    Label3(156).Refresh
    codigo = "DELETE FROM tmpinformes WHERE codusu = " & vUsu.codigo & " AND importe1 = 0 and importe2=0 "
    conn.Execute codigo
    
    
    Label3(156).Caption = "Valores nulos"
    Label3(156).Refresh
    codigo = "UPDATE tmpinformes set importe3 =0 WHERE codusu = " & vUsu.codigo & " AND importe3 is null "
    conn.Execute codigo
    
    
    
    
    
    If chkBenAge(10).Value = 1 Then
        'Aplicando descuentos al coste
        Label3(156).Caption = "Leyendo descuentos"
        Label3(156).Refresh
        Set Col = New Collection
        Set miRsAux = New ADODB.Recordset
        'Para ello agruparemos por proveedores,codfamia
        codigo = "select importeb1,campo2 from tmpinformes  where codusu =" & vUsu.codigo & " group by 1,2 order by 1,2"
        miRsAux.Open codigo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        codigo = ""
        KK = 0
        While Not miRsAux.EOF
            KK = KK + 1
            codigo = codigo & ", (" & miRsAux!importeb1 & "," & miRsAux!campo2 & ")"
            If KK > 30 Then
                codigo = Mid(codigo, 2)
                Col.Add codigo
                codigo = ""
            End If
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        
        If codigo <> "" Then
            codigo = Mid(codigo, 2)
            Col.Add codigo
        End If
        
        conSubRPT = False 'Si hay que ejecutar el update
        If Col.Count > 0 Then
            For KK = 1 To Col.Count
                'Montamos el SQL
                Label3(156).Caption = KK & " / " & Col.Count
                Label3(156).Refresh
                codigo = "select * from sdtomp where dtosincargo>0 and (codprove,codfamia) in ( " & Col.item(KK) & ") ORDER BY 1,2,3"
                miRsAux.Open codigo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If Not miRsAux.EOF Then
                    conSubRPT = True
                    While Not miRsAux.EOF
                        codigo = "UPDATE tmpinformes set importe4=" & DBSet(miRsAux!dtosincargo, "N") & " WHERE codusu =" & vUsu.codigo
                        codigo = codigo & " AND importeb1= '" & miRsAux!Codprove & "' AND campo2 = " & miRsAux!Codfamia
                        If Not IsNull(miRsAux!codmarca) Then codigo = codigo & " AND campo1 =" & miRsAux!codmarca
                        conn.Execute codigo
                        miRsAux.MoveNext
                    Wend
                End If
                miRsAux.Close
            Next KK
               
            If conSubRPT Then
                'Es que ha habiado alguna actualizacion del coste por la columna dtosincargo
                codigo = " update tmpinformes set importe3=(importe3*(100-importe4))/100"
                codigo = codigo & " Where CodUsu = " & vUsu.codigo & " And importe4 > 0"
                conn.Execute codigo
            End If
        End If
 
        
        Set miRsAux = Nothing
        
    End If
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    Label3(156).Caption = "Registros"
    Label3(156).Refresh
End Sub











Private Sub CargaArbolTablas()
Dim N As Node
Dim SQL As String
Dim i As Integer

    Set miRsAux = New ADODB.Recordset
    miRsAux.Open "show tables", conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        SQL = miRsAux.Fields(0)
        If LCase(Mid(SQL, 1, 3)) = "tmp" Then SQL = ""
        
        If SQL <> "" Then
            Set N = TreeView1.Nodes.Add(, , miRsAux.Fields(0), miRsAux.Fields(0))
            N.Checked = True
            N.Expanded = True
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    For i = 1 To TreeView1.Nodes.Count
        lblMultibase.Caption = Space(20) & TreeView1.Nodes(i).Text
        lblMultibase.Refresh
        miRsAux.Open "show columns from " & TreeView1.Nodes(i), conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        While Not miRsAux.EOF
            
            SQL = miRsAux!Field
            If DBLet(miRsAux!Key, "T") <> "" Then
                If DBLet(miRsAux!Key, "T") = "PRI" Then SQL = ""
 
             
                
            End If
            If SQL <> "" Then
                'Solo los textos
                If UCase(Mid(miRsAux!Type, 1, 5)) <> "VARCH" Then SQL = ""
            End If
            miRsAux.MoveNext
            
            If SQL <> "" Then
                Set N = TreeView1.Nodes.Add(TreeView1.Nodes(i).Key, tvwChild, , SQL)
                N.Checked = True
                
            End If
                
        Wend
        miRsAux.Close
   Next

    'Quito los que no voy a procesar
   Set N = TreeView1.Nodes(1)
   Set N = N.LastSibling
   While Not (N Is Nothing)
        i = 0
        If N.Children = 0 Then i = N.Index
        If N.Previous Is Nothing Then
            Set N = Nothing
        Else
            Set N = N.Previous
        End If
        If i > 0 Then TreeView1.Nodes.Remove i
    Wend
End Sub








'------------------------------------------------------------------

Private Function crearCliente() As Boolean
Dim OK As Boolean
    On Error GoTo ecrearCliente
    
    crearCliente = False
    
    
'    campo = DevuelveDesdeBD(conAri, "max(codclien)", "sclien", "1", "1")
    'NumRegElim = Val(campo) + 1
    NumRegElim = Val(txtNumero(2).Text)
    
    
    'codmacta
    'numParam = vEmpresa.DigitosUltimoNivel - 2 'Menos el 43 del principio de la codmacta
    'campo = "43" & Right(String(10, "0") & NumRegElim, numParam)
    campo = txtNumero(3).Text
    
    numParam = InStr(1, Me.txtTextoNoEditable(0).Text, "-")
    devuelve = Trim(Mid(txtTextoNoEditable(0).Text, 1, numParam - 1))
    
    'ORDEN INSERT
    ' codclien , Nomclien, nomcomer, domclien, codpobla, pobclien, proclien, nifClien, wwwclien, fechaalt,
    ' codactiv, CodEnvio, codzonas, codrutas, CodAgent, codforpa, Clivario, TipoIVA, tipofact, albarcon,tipclien,
    ' periodof, codTarif, DtoPPago, DtoGnral, codsitua, referobl, cliabono, pasclien, ManipuladortipoCarnet ,codmacta

    cadSelect = NumRegElim & ", Nomclien, nomcomer, domclien, codpobla, pobclien, proclien, nifClien, wwwclien,now() fechaalt, codactiv,"
    cadSelect = cadSelect & " CodEnvio, codzonas, codrutas," & Me.txtAgente(10).Text & " agente," & Me.txtForpa(3).Text & " forpa,"
    cadSelect = cadSelect & " 0 Clivario, 0 TipoIVA, 0 tipofact, 0 albarcon,0 tipclien, 0 periodof," & vParamAplic.PorDefecto_Tarifa & " tarifa, "
    cadSelect = cadSelect & " 0.00,0.00,0 codsitua,0 referob,0 cliabo, nifclien pasweb, 0 manicarn, '" & campo & "' codmacta"
    cadSelect = cadSelect & " , perclie1 , telclie1, faxclie1, maiclie1, perclie2, telclie2, faxclie2, maiclie2, observac"
    cadSelect = cadSelect & ",  " & Me.txtAgente(10).Text & " visitador, 9 credipriv "
    
    cadSelect = cadSelect & " from sclipot where codclien=" & devuelve
    
    codigo = "INSERT INTO sclien( codclien , Nomclien, nomcomer, domclien, codpobla, pobclien, proclien, nifClien, wwwclien, fechaalt,"
    codigo = codigo & " codactiv, CodEnvio, codzonas, codrutas, CodAgent, codforpa, Clivario, TipoIVA, tipofact, albarcon,tipclien,"
    codigo = codigo & " periodof, codTarif, DtoPPago, DtoGnral, codsitua, referobl, cliabono, pasclien, ManipuladortipoCarnet ,codmacta"
    codigo = codigo & " , perclie1 , telclie1, faxclie1, maiclie1, perclie2, telclie2, faxclie2, maiclie2, observac,visitador,credipriv)"
    codigo = codigo & " SELECT " & cadSelect
    conn.Execute codigo
    
    
    
    'Si llega aqui es que va bien
    Espera 0.75
   
    'Si ya existe, NO la creo. NO hago nada con la cuenta
    
    codigo = DevuelveDesdeBD(conConta, "nommacta", "cuentas", "codmacta", campo)
    If codigo <> "" Then
        OK = True
    Else
        OK = False
        If InsertarCuentaCble(campo, CStr(NumRegElim)) Then OK = True
    End If
    If OK Then
        crearCliente = True 'OK, todo perfecto
        CadenaDesdeOtroForm = NumRegElim
        
        'Si huberia o hubiesen metido mas contactos
        codigo = ",id,nombre,dpto,cargo,telefono,ext,maidirec,movil,observa "
        cadSelect = "select " & NumRegElim & codigo & " from sclipotdp "
        cadSelect = "INSERT INTO scliendp(codclien" & codigo & ") " & cadSelect
        ejecutar cadSelect, True
        
    End If
    
    
    
        
    Exit Function
ecrearCliente:
    MuestraError Err.Number, Err.Description
End Function




'-------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------
'
'       Beneficio por Marca Agente Proveedor  (29/08/2016)
'
'-------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------
Private Sub BenexMarcaAgenProv()
Dim Col As Collection
Dim KK As Integer
    Label3(183).Caption = "Prepara datos"
    Label3(183).Refresh
    conn.Execute "DELETE FROM tmpinformes where codusu = " & vUsu.codigo
    Espera 0.2
    'codusu,codigo1,nombre1,nombre2,importe1,importe2,fecha1
    Label3(183).Caption = "Leyendo BD"
    Label3(183).Refresh
    codigo = "Select  " & vUsu.codigo & ",sartic.codfamia,codmarca,scafac.codagent,slifac.codartic,sum(cantidad),"
    codigo = codigo & " sum(ImporteL -round((importel * (dtoppago+dtognral)/100),2)),slifac.nomartic,sartic.codprove,sum(slifac." & DevuelvePrecioCosteListado(2, True)
    codigo = codigo & " *cantidad) FROM scafac,slifac,sartic ,sfamia WHERE scafac.codtipom = slifac.codtipom And scafac.NumFactu = "
    codigo = codigo & " slifac.NumFactu And scafac.FecFactu = slifac.FecFactu AND slifac.codartic=sartic.codartic  and sartic.codfamia=sfamia.codfamia"
    If cadSelect <> "" Then
        campo = Replace(cadSelect, "{", "(")
        campo = Replace(campo, "}", ")")
        codigo = codigo & " AND " & campo
    End If
    codigo = codigo & " group by codartic,codagent "
    codigo = "INSERT INTO tmpinformes(codusu,codigo1,campo1,campo2,nombre1,importe1,importe2,nombre2,nombre3,importe3) " & codigo
    
    conn.Execute codigo
    
    
    
    Set miRsAux = New ADODB.Recordset
    
    

    'Quio los ceros
    Label3(183).Caption = "Ceros"
    Label3(183).Refresh
    codigo = "DELETE FROM tmpinformes WHERE codusu = " & vUsu.codigo & " AND importe1 = 0 and importe2=0 "
    conn.Execute codigo
    
    
    
    Label3(183).Caption = "Valores nulos"
    Label3(183).Refresh
    codigo = "UPDATE tmpinformes set importe3 =0 WHERE codusu = " & vUsu.codigo & " AND importe3 is null "
    conn.Execute codigo
    
    
    
    
    
    If chkBeneMarcaAgen(1).Value Then
        'Aplicando descuentos al coste
        Label3(183).Caption = "Leyendo descuentos"
        Label3(183).Refresh
        Set Col = New Collection
        
        'Para ello agruparemos por proveedores,codfamia
        codigo = "select nombre3,codigo1 from tmpinformes  where codusu =" & vUsu.codigo & " group by 1,2 order by 1,2"
        miRsAux.Open codigo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        codigo = ""
        KK = 0
        While Not miRsAux.EOF
            KK = KK + 1
            codigo = codigo & ", (" & miRsAux!nombre3 & "," & miRsAux!Codigo1 & ")"
            If KK > 30 Then
                codigo = Mid(codigo, 2)
                Col.Add codigo
                codigo = ""
            End If
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        
        If codigo <> "" Then
            codigo = Mid(codigo, 2)
            Col.Add codigo
        End If
        
        conSubRPT = False 'Si hay que ejecutar el update
        If Col.Count > 0 Then
            For KK = 1 To Col.Count
                'Montamos el SQL
                Label3(183).Caption = KK & " / " & Col.Count
                Label3(183).Refresh
                codigo = "select * from sdtomp where dtosincargo>0 and (codprove,codfamia) in ( " & Col.item(KK) & ") ORDER BY 1,2,3"
                miRsAux.Open codigo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If Not miRsAux.EOF Then
                    conSubRPT = True
                    While Not miRsAux.EOF
                        codigo = "UPDATE tmpinformes set importe4=" & DBSet(miRsAux!dtosincargo, "N") & " WHERE codusu =" & vUsu.codigo
                        codigo = codigo & " AND nombre3= '" & miRsAux!Codprove & "' AND codigo1 = " & miRsAux!Codfamia
                        If Not IsNull(miRsAux!codmarca) Then codigo = codigo & " AND campo1 =" & miRsAux!codmarca
                        conn.Execute codigo
                        miRsAux.MoveNext
                    Wend
                End If
                miRsAux.Close
            Next KK
               
            If conSubRPT Then
                'Es que ha habiado alguna actualizacion del coste por la columna dtosincargo
                codigo = " update tmpinformes set importe3=(importe3*(100-importe4))/100"
                codigo = codigo & " Where CodUsu = " & vUsu.codigo & " And importe4 > 0"
                conn.Execute codigo
            End If
        End If
 
        
        
        
    End If
    
    
    Label3(183).Caption = "Proveedor"
    Label3(183).Refresh
    codigo = "Select nombre3 from tmpinformes where codusu =" & vUsu.codigo & " GROUP BY 1"
    miRsAux.Open codigo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        codigo = DevuelveDesdeBD(conAri, "nomprove", "sprove", "codprove", miRsAux!nombre3)
        If codigo = "" Then codigo = "N/A"
        codigo = miRsAux!nombre3 & "-" & codigo
        Label3(183).Caption = "Prov: " & codigo
        Label3(183).Refresh
        
        codigo = "UPDATE tmpinformes set nombre3=" & DBSet(codigo, "T") & " WHERE codusu=" & vUsu.codigo & " AND nombre3=" & DBSet(miRsAux!nombre3, "T")
        miRsAux.MoveNext
        conn.Execute codigo
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    
    Label3(183).Caption = ""
    
End Sub





'-------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------
'
'       Ventas marca-familia  (29/08/2016)
'
'-------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------
Private Sub VentasMarcaFamilia()
    
    Label3(188).Caption = "Prepara datos"
    Label3(188).Refresh
    conn.Execute "DELETE FROM tmpinformes where codusu = " & vUsu.codigo
    Espera 0.2
    'codusu,codigo1,nombre1,nombre2,importe1,importe2,fecha1
    Label3(188).Caption = "Leyendo BD"
    Label3(188).Refresh
    
    codigo = "Select  " & vUsu.codigo & ",sartic.codfamia,codmarca,"
    
    If Opcion = 49 Then
        'VENTAS
        'codusu,campo1,campo2,codigo1,nombre1,importe1,importe2,nombre2,nombre3
        codigo = codigo & "scafac.codagent,slifac.codartic,sum(cantidad),"
        codigo = codigo & "sum(importel) "   ' sum(ImporteL -round((importel * (dtoppago+dtognral)/100),2))"
        codigo = codigo & " ,slifac.nomartic,sartic.codprove"
        'codigo = codigo & " ,sum(slifac." & DevuelvePrecioCosteListado(2, True) *cantidad)        No es de costes ni de beneficios
        codigo = codigo & "  FROM scafac,slifac,sartic  WHERE scafac.codtipom = slifac.codtipom And scafac.NumFactu = "
        codigo = codigo & " slifac.NumFactu And scafac.FecFactu = slifac.FecFactu AND slifac.codartic=sartic.codartic  "
    
    
    Else
        'COMPRAS
        codigo = codigo & "scafpc.Codprove , slifpc.codArtic, Sum(cantidad), Sum(ImporteL)"
        codigo = codigo & ",slifpc.nomartic,sartic.codprove"
        codigo = codigo & " FROM scafpc,slifpc,sartic  WHERE scafpc.codprove = slifpc.codprove And scafpc.NumFactu =  slifpc.NumFactu And"
        codigo = codigo & " scafpc.FecFactu = slifpc.FecFactu  AND slifpc.codartic=sartic.codartic "
    
    End If
    
    If cadSelect <> "" Then
        campo = Replace(cadSelect, "{", "(")
        campo = Replace(campo, "}", ")")
        codigo = codigo & " AND " & campo
    End If
    codigo = codigo & " group by codartic,"
    codigo = codigo & IIf(Opcion = 49, "codagent", "slifpc.codprove")
        

    codigo = "INSERT INTO tmpinformes(codusu,campo1,campo2,codigo1,nombre1,importe1,importe2,nombre2,nombre3) " & codigo
    conn.Execute codigo
    
    
    
  
    
    
    

    'Quio los ceros
    Label3(188).Caption = "Ceros"
    Label3(188).Refresh
    codigo = "DELETE FROM tmpinformes WHERE codusu = " & vUsu.codigo & " AND importe1 = 0 and importe2=0 "
    conn.Execute codigo
    

    
End Sub






Private Sub GenerearFicheroTxtAlbaranRuta()
Dim NF2 As Integer

    On Error GoTo eGenerearFicheroTxtAlbaranRuta
    
    codigo = App.Path & "\Rutas" & Format(Now, "yymmddhhnnss") & ".dat"
    NF2 = FreeFile
    Open codigo For Output As #NF2
    Print #NF2, cadSelect
    Print #NF2, ""
    Print #NF2, ""
    Set miRsAux = New ADODB.Recordset
    codigo = "Select codtipom,numalbar FROM scaalb WHERE " & cadSelect
    
    miRsAux.Open codigo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    codigo = ""
    While Not miRsAux.EOF
        codigo = codigo & ", ('" & miRsAux!codtipom & "'," & miRsAux!NUmAlbar & ")"
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Print #NF2, codigo
    
    Print #NF2, ""
    Print #NF2, "FIN"
eGenerearFicheroTxtAlbaranRuta:
    If Err.Number <> 0 Then Err.Clear
    CierraF NF2
    Set miRsAux = Nothing
    
End Sub
Private Sub CierraF(ByRef N As Integer)
    On Error Resume Next
    Close #N
    Err.Clear
End Sub

