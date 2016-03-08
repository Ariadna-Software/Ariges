VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmListadoNomi 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7815
   Icon            =   "frmListadoNomi.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   7815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame FrameGenerar 
      Height          =   3375
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   4695
      Begin VB.CheckBox Check1 
         Caption         =   "Copiar gastos"
         Height          =   255
         Left            =   360
         TabIndex        =   74
         Top             =   2280
         Width           =   2055
      End
      Begin VB.CommandButton cmdCancelarGen 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   2640
         TabIndex        =   5
         Top             =   2760
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptarGen 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   1440
         TabIndex        =   4
         Top             =   2760
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   2835
         MaxLength       =   4
         TabIndex        =   3
         Top             =   1740
         Width           =   765
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   2835
         MaxLength       =   4
         TabIndex        =   2
         Top             =   1380
         Width           =   765
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   915
         MaxLength       =   4
         TabIndex        =   1
         Top             =   1740
         Width           =   765
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   915
         MaxLength       =   4
         TabIndex        =   0
         Top             =   1380
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Introduzca el período del que desea duplicar las nóminas y el que desea generar:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   6
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   4140
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Mes"
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
         Left            =   2400
         TabIndex        =   12
         Top             =   1740
         Width           =   285
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
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
         Height          =   315
         Index           =   3
         Left            =   2400
         TabIndex        =   11
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Año"
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
         Left            =   2400
         TabIndex        =   10
         Top             =   1365
         Width           =   285
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Mes"
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
         Left            =   480
         TabIndex        =   9
         Top             =   1740
         Width           =   285
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
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
         Height          =   315
         Index           =   4
         Left            =   480
         TabIndex        =   8
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Año"
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
         Left            =   480
         TabIndex        =   7
         Top             =   1365
         Width           =   285
      End
   End
   Begin VB.Frame FrameRecibo 
      Height          =   4935
      Left            =   120
      TabIndex        =   50
      Top             =   0
      Visible         =   0   'False
      Width           =   6135
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   18
         Left            =   3720
         MaxLength       =   10
         TabIndex        =   58
         Top             =   3720
         Width           =   1215
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   17
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   57
         Top             =   3720
         Width           =   1215
      End
      Begin VB.OptionButton optList 
         Caption         =   "Anticipo"
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   60
         Top             =   4440
         Width           =   975
      End
      Begin VB.OptionButton optList 
         Caption         =   "Banco"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   59
         Top             =   4440
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.CommandButton cmdListaRecibos 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3480
         TabIndex        =   61
         Top             =   4320
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   4680
         TabIndex        =   62
         Top             =   4320
         Width           =   1095
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   16
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   70
         Text            =   "Text5"
         Top             =   2880
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   16
         Left            =   1260
         MaxLength       =   4
         TabIndex        =   56
         Top             =   2880
         Width           =   615
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   15
         Left            =   1935
         Locked          =   -1  'True
         TabIndex        =   67
         Text            =   "Text5"
         Top             =   2400
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   15
         Left            =   1260
         MaxLength       =   4
         TabIndex        =   55
         Top             =   2400
         Width           =   615
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   14
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   54
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   13
         Left            =   2115
         MaxLength       =   4
         TabIndex        =   53
         Top             =   840
         Width           =   765
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   12
         Left            =   600
         MaxLength       =   4
         TabIndex        =   52
         Top             =   840
         Width           =   765
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Importe2"
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
         Left            =   2760
         TabIndex        =   73
         Top             =   3480
         Width           =   810
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Importe1"
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
         TabIndex        =   72
         Top             =   3480
         Width           =   810
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   16
         Left            =   960
         Picture         =   "frmListadoNomi.frx":000C
         ToolTipText     =   "Buscar Trabajador"
         Top             =   2880
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
         Index           =   6
         Left            =   360
         TabIndex        =   71
         Top             =   2880
         Width           =   420
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   15
         Left            =   960
         Picture         =   "frmListadoNomi.frx":010E
         ToolTipText     =   "Buscar Trabajador"
         Top             =   2400
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
         Left            =   360
         TabIndex        =   69
         Top             =   2400
         Width           =   450
      End
      Begin VB.Label Label1 
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
         Index           =   15
         Left            =   240
         TabIndex        =   68
         Top             =   2160
         Width           =   945
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   14
         Left            =   1440
         Picture         =   "frmListadoNomi.frx":0210
         Top             =   1560
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha recibo"
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
         TabIndex        =   66
         Top             =   1560
         Width           =   1065
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nombre Mes"
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
         Left            =   3000
         TabIndex        =   65
         Top             =   885
         Width           =   1245
      End
      Begin VB.Label Label1 
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
         Index           =   14
         Left            =   1680
         TabIndex        =   64
         Top             =   840
         Width           =   345
      End
      Begin VB.Label Label1 
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
         Index           =   13
         Left            =   240
         TabIndex        =   63
         Top             =   840
         Width           =   330
      End
      Begin VB.Label Label14 
         Caption         =   "Recibos desplazamiento"
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
         Left            =   240
         TabIndex        =   51
         Top             =   240
         Width           =   3975
      End
   End
   Begin VB.Frame FrameListado 
      Height          =   3855
      Left            =   120
      TabIndex        =   14
      Top             =   0
      Width           =   6495
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   7
         Left            =   1500
         MaxLength       =   4
         TabIndex        =   18
         Top             =   2520
         Width           =   615
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   7
         Left            =   2175
         Locked          =   -1  'True
         TabIndex        =   27
         Text            =   "Text5"
         Top             =   2520
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   6
         Left            =   1500
         MaxLength       =   4
         TabIndex        =   17
         Top             =   2160
         Width           =   615
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   6
         Left            =   2175
         Locked          =   -1  'True
         TabIndex        =   24
         Text            =   "Text5"
         Top             =   2160
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   4
         Left            =   915
         MaxLength       =   4
         TabIndex        =   15
         Top             =   1020
         Width           =   765
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   915
         MaxLength       =   4
         TabIndex        =   16
         Top             =   1380
         Width           =   765
      End
      Begin VB.CommandButton cmdAceptarLis 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3600
         TabIndex        =   19
         Top             =   3120
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelarLis 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   4800
         TabIndex        =   20
         Top             =   3120
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nombre Mes"
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
         Left            =   1800
         TabIndex        =   29
         Top             =   1420
         Width           =   1245
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
         Left            =   600
         TabIndex        =   28
         Top             =   2520
         Width           =   420
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   7
         Left            =   1200
         Picture         =   "frmListadoNomi.frx":029B
         ToolTipText     =   "Buscar Trabajador"
         Top             =   2520
         Width           =   240
      End
      Begin VB.Label Label1 
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
         Index           =   9
         Left            =   480
         TabIndex        =   26
         Top             =   1875
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
         Index           =   15
         Left            =   600
         TabIndex        =   25
         Top             =   2160
         Width           =   450
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   6
         Left            =   1200
         Picture         =   "frmListadoNomi.frx":039D
         ToolTipText     =   "Buscar Trabajador"
         Top             =   2160
         Width           =   240
      End
      Begin VB.Label Label14 
         Caption         =   "Listado Nóminas y Gastos"
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
         Left            =   240
         TabIndex        =   23
         Top             =   360
         Width           =   4815
      End
      Begin VB.Label Label1 
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
         Index           =   7
         Left            =   480
         TabIndex        =   22
         Top             =   1005
         Width           =   330
      End
      Begin VB.Label Label1 
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
         Index           =   8
         Left            =   480
         TabIndex        =   21
         Top             =   1380
         Width           =   345
      End
   End
   Begin VB.Frame FrameNorma34 
      Height          =   5175
      Left            =   120
      TabIndex        =   38
      Top             =   0
      Width           =   6375
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   480
         TabIndex        =   48
         Top             =   4560
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   240
         Top             =   3360
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   0
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   47
         Text            =   "Text5"
         Top             =   2540
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   9
         Left            =   2580
         MaxLength       =   10
         TabIndex        =   35
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Frame Frame1 
         Caption         =   "Importe"
         ForeColor       =   &H00000080&
         Height          =   975
         Left            =   3720
         TabIndex        =   45
         Top             =   840
         Width           =   1575
         Begin VB.OptionButton OptImporte 
            Caption         =   "Gastos"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   33
            Top             =   540
            Width           =   1215
         End
         Begin VB.OptionButton OptImporte 
            Caption         =   "Nómina"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   32
            Top             =   260
            Value           =   -1  'True
            Width           =   1215
         End
      End
      Begin VB.CommandButton cmdCancelarNorma34 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   4320
         TabIndex        =   37
         Top             =   3520
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptarNorma34 
         Caption         =   "&Generar Fichero"
         Height          =   375
         Left            =   2640
         TabIndex        =   36
         Top             =   3520
         Width           =   1455
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   11
         Left            =   915
         MaxLength       =   4
         TabIndex        =   31
         Top             =   1380
         Width           =   765
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   10
         Left            =   915
         MaxLength       =   4
         TabIndex        =   30
         Top             =   1020
         Width           =   765
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   8
         Left            =   1695
         Locked          =   -1  'True
         TabIndex        =   39
         Text            =   "Text5"
         Top             =   2240
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   8
         Left            =   900
         MaxLength       =   4
         TabIndex        =   34
         Top             =   2240
         Width           =   735
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   480
         X2              =   5760
         Y1              =   4215
         Y2              =   4215
      End
      Begin VB.Label lblProgreso 
         AutoSize        =   -1  'True
         Caption         =   "Generando Norma 34 ..."
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
         Left            =   480
         TabIndex        =   49
         Top             =   4335
         Width           =   5295
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Transferencia"
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
         Left            =   480
         TabIndex        =   46
         Top             =   3000
         Width           =   1710
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   9
         Left            =   2280
         Picture         =   "frmListadoNomi.frx":049F
         Top             =   3000
         Width           =   240
      End
      Begin VB.Label Label1 
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
         Index           =   12
         Left            =   480
         TabIndex        =   44
         Top             =   1380
         Width           =   345
      End
      Begin VB.Label Label1 
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
         Index           =   11
         Left            =   480
         TabIndex        =   43
         Top             =   1005
         Width           =   330
      End
      Begin VB.Label Label14 
         Caption         =   "Generar Norma 34"
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
         Left            =   240
         TabIndex        =   42
         Top             =   360
         Width           =   4815
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   8
         Left            =   600
         Picture         =   "frmListadoNomi.frx":052A
         ToolTipText     =   "Buscar banco propio"
         Top             =   2240
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Banco Propio"
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
         Left            =   480
         TabIndex        =   41
         Top             =   1995
         Width           =   1095
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nombre Mes"
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
         Left            =   1800
         TabIndex        =   40
         Top             =   1420
         Width           =   1245
      End
   End
End
Attribute VB_Name = "frmListadoNomi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public parOpcion As Integer
'opcion del frame a abrir
'   1: listado de nominas
'   2: Generar mes automaticamente
'   3: Norma 34
'   4: Recibo gastos

Private WithEvents frmF As frmCal 'Calendario de Fechas
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmMtoTraba As frmAdmTrabajadores
Attribute frmMtoTraba.VB_VarHelpID = -1
Private WithEvents frmMtoBancosPro As frmFacBancosPropios
Attribute frmMtoBancosPro.VB_VarHelpID = -1


Dim PrimeraVez As Boolean
Dim indCodigo As Integer

Dim cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Dim cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Dim cadParam As String 'cadena con los parametros q se pasan a Crystal R.
Dim numParam As Byte 'numero de parametros


Private Sub InicializarVbles()
    cadFormula = ""
    cadSelect = ""
    cadParam = ""
    numParam = 0
End Sub



Private Sub cmdAceptarGen_Click()
'Generar automaticamente las nominas
Dim cad As String
Dim BorrarPeriodo As Boolean
    On Error GoTo ErrGen

    '- comprobar que los campos origen y destino mes/año tienen valor
    If Me.txtCodigo(0).Text = "" Or Me.txtCodigo(1).Text = "" Then
        MsgBox "El año y mes origen deben tener valor.", vbInformation
        Exit Sub
    End If
    
    If Me.txtCodigo(2).Text = "" Or Me.txtCodigo(3).Text = "" Then
        MsgBox "El año y mes destino deben tener valor.", vbInformation
        Exit Sub
    End If
    
    
    '- comprobar que los campos origen y destino mes/año NO son mismo valor
    If Me.txtCodigo(0).Text = Me.txtCodigo(2).Text Then
      If Me.txtCodigo(1).Text = Me.txtCodigo(3).Text Then
            MsgBox "Mismo año y mes.", vbInformation
            Exit Sub
        End If
    End If
    
    
    
    
    InicializarVbles
    
    '- montar la cadena de seleccion y comprobar q NO hay datos para el
    '  destino q queremos generar
    cad = "{snomin.anynomi}=" & Val(txtCodigo(2).Text)
    If Not AnyadirAFormula(cadFormula, cad) Then Exit Sub
    
    cad = "{snomin.mesnomi}=" & Val(txtCodigo(3).Text)
    If Not AnyadirAFormula(cadFormula, cad) Then Exit Sub
    
    '- comprabar si ya hay lineas de nominas generadas para el destino
    '  q queremos generar en ese caso aviso y salir
    cadSelect = cadFormula
    BorrarPeriodo = False
    If HayRegParaInforme("snomin", cadSelect, True) Then
       
        cad = "Ya existen lineas de nómina para el período destino seleccionado. ¿Desea borrarlas y volverlas a crear?"
        If MsgBox(cad, vbQuestion + vbYesNo) = vbNo Then Exit Sub
        BorrarPeriodo = True
    End If
    
    InicializarVbles
    
    '- montar la cadena de seleccion y comprobar q hay datos para el
    '  origen y se puede duplicar
    cad = "{snomin.anynomi}=" & Val(txtCodigo(0).Text)
    If Not AnyadirAFormula(cadFormula, cad) Then Exit Sub
    
    cad = "{snomin.mesnomi}=" & Val(txtCodigo(1).Text)
    If Not AnyadirAFormula(cadFormula, cad) Then Exit Sub
    
    
    '==== comprobar q hay registros para mostrar en el informe ====
    cadSelect = cadFormula
    If Not HayRegParaInforme("snomin", cadSelect) Then Exit Sub
    Screen.MousePointer = vbHourglass
    
    
    '- duplicar las nominas para el nuevo periodo año/mes
    If BorrarPeriodo Then
        Set miRsAux = New ADODB.Recordset
        cad = "select codtraba from snomin WHERE " & cadSelect & " ORDER BY codtraba"
        miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        cad = ""
        While Not miRsAux.EOF
            cad = cad & ", " & miRsAux.Fields(0)
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        Set miRsAux = Nothing
        If cad <> "" Then
            cad = Mid(cad, 2)
            cad = " and codtraba IN (" & cad & ")"
            cad = "snomin.anynomi=" & txtCodigo(2).Text & " AND  snomin.mesnomi=" & txtCodigo(3).Text & cad
            cad = "DELETE FROM snomin WHERE " & cad
            
            conn.Execute cad
        End If
    End If
    cad = "INSERT INTO snomin (codtraba, anynomi, mesnomi, impnomi, impgasto,impanti) "
    cad = cad & "SELECT codtraba," & Val(txtCodigo(2).Text) & " as anynomi, "
    cad = cad & Val(txtCodigo(3).Text) & " as mesnomi, impnomi, "
    If Me.Check1.Value = 0 Then cad = cad & "0 as "  'para que meta un cero
    cad = cad & "impgasto,0  "
    cad = cad & "FROM snomin "
    cad = cad & "WHERE " & cadSelect
    
    conn.Execute cad
    Screen.MousePointer = vbDefault
    MsgBox "Se han generado correctamente las lineas de nóminas para el mes destino.", vbInformation
    
    cmdCancelarGen_Click
    Exit Sub
    
ErrGen:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Error al generar lineas de nóminas", Err.Description
    Set miRsAux = Nothing
End Sub

Private Sub cmdAceptarLis_Click()
'abrir el listado de nominas
Dim cad As String

    If txtCodigo(4).Text = "" And txtCodigo(5).Text = "" And txtCodigo(6).Text = "" And txtCodigo(7).Text = "" Then
        MsgBox "Debe seleccionar algún criterio para el informe.", vbInformation
        Exit Sub
    End If


    '==== montar la cadena de seleccion de registros ====
    InicializarVbles
    
    '============ PARAMETROS ===========================
    'Nombre de la Empresa
    cadParam = "|pEmpresa=""" & vParam.NombreEmpresa & """|"
    numParam = numParam + 1
    
    
    'Cadena para seleccion del año
    '-------------------------------
    If txtCodigo(4).Text <> "" Then
        cad = "{snomin.anynomi}=" & Val(txtCodigo(4).Text)
        If Not AnyadirAFormula(cadFormula, cad) Then Exit Sub
        'Parametro año
        cadParam = cadParam & "pAnyo=""Año: " & txtCodigo(4).Text & """|"
        numParam = numParam + 1
    End If
    
    'Cadena para seleccion del mes
    '-------------------------------
    If txtCodigo(5).Text <> "" Then
        cad = "{snomin.mesnomi}=" & Val(txtCodigo(5).Text)
        If Not AnyadirAFormula(cadFormula, cad) Then Exit Sub
        'Parametro mes
        cadParam = cadParam & "pMes=""Mes: " & txtCodigo(5).Text & """|"
        numParam = numParam + 1
    End If
    
    'Cadena para seleccion del trabajador
    '------------------------------------
    If txtCodigo(6).Text <> "" Or txtCodigo(7).Text <> "" Then
        cad = CadenaDesdeHasta(txtCodigo(6).Text, txtCodigo(7).Text, "{snomin.codtraba}", "N", "Trabajador")
        If Not AnyadirAFormula(cadFormula, cad) Then Exit Sub
        'Parametro trabajador
        cad = "pDHTrabajador=""Trabajador: "
        cadParam = cadParam & AnyadirParametroDH(cad, 6, 7) & """|"
        numParam = numParam + 1
        
    End If
    
    cadSelect = cadFormula
    
    '==== comprobar q hay registros para mostrar en el informe ====
    If Not HayRegParaInforme("snomin", cadSelect) Then Exit Sub
    
    
    '==== abrir el informe ====
    With frmImprimir
        .FormulaSeleccion = cadFormula
        .OtrosParametros = cadParam
        .NumeroParametros = numParam
        .SoloImprimir = False
        .EnvioEMail = False
        .Opcion = 501
        .Titulo = "Inf. Nóminas y Gastos"
        .NombreRPT = "rAdmNominas.rpt"
        .ConSubInforme = False
        .Show vbModal
    End With
End Sub



Private Function AnyadirParametroDH(cad As String, indD As Byte, indH As Byte) As String
On Error Resume Next
    If txtCodigo(indD).Text <> "" Then
        cad = cad & "desde " & txtCodigo(indD).Text
        If txtNombre(indD).Text <> "" Then cad = cad & " - " & txtNombre(indD).Text
    End If
    If txtCodigo(indH).Text <> "" Then
        cad = cad & "  hasta " & txtCodigo(indH).Text
        If txtNombre(indH).Text <> "" Then cad = cad & " - " & txtNombre(indH).Text
    End If
    AnyadirParametroDH = cad
    If Err.Number <> 0 Then Err.Clear
End Function



Private Function DatosOk_Norma34() As Boolean
Dim b As Boolean
Dim SQL As String

    On Error GoTo EDatosOK

    b = True
    
    '- comprobar q el año y mes tiene valor
    If Me.txtCodigo(10).Text = "" Or Me.txtCodigo(11).Text = "" Then
        b = False
        MsgBox "Debe introducir el año y mes a procesar.", vbExclamation
    ElseIf Not EsMesOK(txtCodigo(11).Text) Then
        b = False
        MsgBox "El mes introducido no es correcto.", vbExclamation
    End If
    


    '- comprobar q la fecha transferencia tiene valor
    If b Then
        If txtCodigo(9).Text = "" Then
            MsgBox "Debe introducir la Fecha de Transferencia.", vbExclamation
            txtCodigo(9).Text = ""
            PonerFoco txtCodigo(9)
            b = False
        ElseIf Not IsDate(Me.txtCodigo(9).Text) Then
            MsgBox "La Fecha de transferencia no es válida.", vbExclamation
            b = False
        End If
    End If
    
    '- comprobar q el banco propio tiene valor
    If b Then
        If txtCodigo(8).Text = "" Then
            MsgBox "Debe introducir el Banco propio.", vbExclamation
            txtCodigo(8).Text = ""
            PonerFoco txtCodigo(8)
            b = False
        Else
            'comprobar q el banco propio es correcto
            SQL = DevuelveDesdeBDNew(conAri, "sbanpr", "codbanpr", "codbanpr", txtCodigo(8).Text, "N")
            If Trim(SQL) = "" Then
                b = False
                MsgBox "No existe el Banco propio seleccionado.", vbExclamation
            Else
                ObtenerCtasBancoPropio2 txtCodigo(8).Text, SQL, ""
                SQL = Replace(SQL, "-", "")
                If SQL = "" Then
                    b = False
                    MsgBox "El banco propio seleccionado no tiene cuenta bancaria.", vbExclamation
                ElseIf Not Comprueba_CuentaBan2(SQL) Then
'                    b = False
                End If
            End If
        End If
    End If
    
    DatosOk_Norma34 = b
    Exit Function
    
EDatosOK:
   MuestraError Err.Number, "Datos OK", Err.Description
End Function






Private Sub cmdAceptarNorma34_Click()
'generar el fichero de Norma 34
Dim cad As String
Dim cadAux As String

    '- comprobar q el valor introducido en los campos son correctos
    If Not DatosOk_Norma34 Then Exit Sub

    InicializarVbles
    
    '==== montar la cadena de seleccion de registros ====
    '- seleccionar mes/año
    cad = "snomin.anynomi=" & Val(txtCodigo(10).Text)
    If Not AnyadirAFormula(cadSelect, cad) Then Exit Sub
    
    cad = "snomin.mesnomi=" & Val(txtCodigo(11).Text)
    If Not AnyadirAFormula(cadSelect, cad) Then Exit Sub
    
    
    '- seleccionar los q el importe no sea 0
    If Me.OptImporte(0).Value = True Then 'Nomina
        cad = "snomin.impnomi<>0"
    Else 'Gastos
        cad = "snomin.impgasto<>0"
    End If
    If Not AnyadirAFormula(cadSelect, cad) Then Exit Sub
    
    
    '- seleccionar los no pasados ya a norma 34
    If Me.OptImporte(0).Value = True Then 'Nomina
        cad = "snomin.n34nomi=0"
        cadAux = "snomin.n34nomi=1"
    Else 'Gastos
        cad = "snomin.n34gast=0"
        cadAux = "snomin.n34nomi=1"
    End If
    If Not AnyadirAFormula(cadAux, cadSelect) Then Exit Sub
    If Not AnyadirAFormula(cadSelect, cad) Then Exit Sub
    
    
    '==== comprobar q hay registros para mostrar en el informe ====
    If HayRegParaInforme("snomin", cadSelect, True) Then
        Screen.MousePointer = vbHourglass
        MostrarProgreso True
        
        GenerarNorma34 (cadSelect)
        
        MostrarProgreso False
        Screen.MousePointer = vbDefault
        cmdCancelarNorma34_Click
        
    ElseIf HayRegParaInforme("snomin", cadAux, True) Then
        MsgBox "Ya se han traspasado las Nóminas del mes " & txtCodigo(11).Text & " a Norma 34.", vbExclamation
    Else
        MsgBox "No hay datos con esos criterios para pasar a Norma 34.", vbExclamation
    End If

End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdCancelarGen_Click()
'cancelar generar automaticamente las nominas
    Unload Me
End Sub

Private Sub cmdCancelarLis_Click()
'cancelar listado
    Unload Me
End Sub

Private Sub cmdCancelarNorma34_Click()
    Unload Me
End Sub

Private Sub cmdListaRecibos_Click()
Dim b As Boolean
Dim cad As String

    cadParam = ""
    For indCodigo = 12 To 18
        If indCodigo < 15 Or indCodigo > 16 Then
            If txtCodigo(indCodigo).Text = "" Then
                cadParam = "m"
                Exit For
            End If
        End If
    Next
    If cadParam <> "" Then
        MsgBox "Todos los campos, execepto trabajador, son obligatorios", vbExclamation
        Exit Sub
    End If
        
    If ImporteFormateado(txtCodigo(17).Text) = 0 Or ImporteFormateado(txtCodigo(18).Text) = 0 Then
        MsgBox "Importes no pueden ser CERO", vbExclamation
        Exit Sub
    End If
        

    Screen.MousePointer = vbHourglass
    b = GeneraDatosRecibos
    Screen.MousePointer = vbDefault
    If b Then
        InicializarVbles
        If Me.optList(0).Value Then
            indCodigo = 44
        Else
            indCodigo = 43
        End If
        
        If Not PonerParamRPT2(CByte(indCodigo), cadParam, numParam, cad, pImprimeDirecto, pPdfRpt, pRptvMultiInforme) Then Exit Sub
        'NO imprimo, de momento, via mail
        pPdfRpt = cad  'por nio meter mas variables
        
    
        cad = "{snomin.anynomi}=" & Val(txtCodigo(12).Text)
        If Not AnyadirAFormula(cadFormula, cad) Then Exit Sub
    
        cad = "{snomin.mesnomi}=" & Val(txtCodigo(13).Text)
        If Not AnyadirAFormula(cadFormula, cad) Then Exit Sub
    
        If txtCodigo(15).Text <> "" Then
                cad = "{snomin.codtraba}>=" & Val(txtCodigo(15).Text)
                If Not AnyadirAFormula(cadFormula, cad) Then Exit Sub
        End If
    
        If txtCodigo(16).Text <> "" Then
                cad = "{snomin.codtraba}<=" & Val(txtCodigo(16).Text)
                If Not AnyadirAFormula(cadFormula, cad) Then Exit Sub
        End If
        
        
        If Me.optList(0).Value Then
            cad = "{snomin.impgasto}>0"
        Else
            'Segun ISABEL TENISA
            cad = "({snomin.impgasto}-{snomin.impanti})>0"
            cad = "({snomin.impanti})>0"
        End If
        If Not AnyadirAFormula(cadFormula, cad) Then Exit Sub
         
        'Montamos el select para los registros
        cadParam = cadParam & "pcodusu= " & vUsu.codigo & "|"
        numParam = numParam + 1
        
        'Fecha recibo
        cadParam = cadParam & "FechaRec=""" & vParam.Poblacion & " " & txtCodigo(14).Text & """|"
        numParam = numParam + 1
        
        
        With frmImprimir
            .FormulaSeleccion = cadFormula
            .OtrosParametros = cadParam
            .NumeroParametros = numParam
            .SoloImprimir = False
            .EnvioEMail = False
            .Opcion = 501
            .Titulo = "Inf. recibo gastos"
            .NombreRPT = pPdfRpt
            .ConSubInforme = False
            .SeleccionaRPTCodigo = pRptvMultiInforme
            .Show vbModal
        End With

    
    
    End If
    
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    If PrimeraVez Then
        Select Case Me.parOpcion
            Case 1: PonerFoco Me.txtCodigo(4)
            Case 2: PonerFoco Me.txtCodigo(0)
        End Select
        
        PrimeraVez = False
    End If
End Sub

Private Sub Form_Load()
Dim H As Integer, W As Integer 'Alto, Ancho

    '- Icono del formulario
    Me.Icon = frmPpal.Icon
    
    '- Iniciar formularios
    PrimeraVez = True
    limpiar Me 'limpiar los campos Text
    Me.Label4(1).Caption = ""
    Me.Label4(2).Caption = ""
    FrameRecibo.visible = False
    '- ocultar/mostrar los frames correspondientes a la opcion
    'Ajustar Tamaño del Frame para ajustar tamaño de Formulario al del Frame
    Select Case Me.parOpcion
        Case 1 'Listado
            H = 3855
            W = 6495
            'valores por defecto
            Me.txtCodigo(4).Text = Year(Now)
            Me.txtCodigo(5).Text = Month(Now)
        Case 2 'Generar aut.
            H = 3375
            W = 4695
            Me.Caption = "Generar nóminas mes automát."
            Me.txtCodigo(0).Text = Year(Now)
            Me.txtCodigo(2).Text = Year(Now)
        Case 3 'Norma 34
            H = 4215
            W = 6375
            Me.Caption = "Generar Norma 34"
            Me.txtCodigo(10).Text = Year(Now)
            Me.txtCodigo(11).Text = Month(Now)
            Me.txtCodigo(9).Text = Format(Now, "dd/mm/yyyy")
            
        Case 4
            Me.txtCodigo(12).Text = Year(Now)
            Me.txtCodigo(13).Text = Month(Now)
            Me.txtCodigo(14).Text = Format(Now, "dd/mm/yyyy")
            H = FrameRecibo.Height
            W = FrameRecibo.Width
            PonerFrameVisible FrameRecibo, True, H, W
            LeerGuardarImportes True
            Label4(3).Caption = ""
    End Select
    
    PonerFrameVisible Me.frameListado, (Me.parOpcion = 1), H, W
    PonerFrameVisible Me.FrameGenerar, (Me.parOpcion = 2), H, W
    PonerFrameVisible Me.FrameNorma34, (Me.parOpcion = 3), H, W
    
    '- ajustar tamaño del form
    Me.Width = W + 70
    Me.Height = H + 350
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If parOpcion = 4 Then LeerGuardarImportes False
End Sub

Private Sub frmF_Selec(vFecha As Date)
'Calendario de Fecha
    txtCodigo(9).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmMtoBancosPro_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Bancos Propios
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoTraba_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Trabajadores
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub imgBuscar_Click(Index As Integer)
    indCodigo = Index
    
    Select Case Index
        Case 6, 7, 15, 16 'TRABAJADOR
            Set frmMtoTraba = New frmAdmTrabajadores
            frmMtoTraba.DatosADevolverBusqueda = "0|1|"
            frmMtoTraba.Show vbModal
            Set frmMtoTraba = Nothing
            
        Case 8 'BANCO PROPIO
            Set frmMtoBancosPro = New frmFacBancosPropios
            frmMtoBancosPro.DatosADevolverBusqueda = "0|1|"
            frmMtoBancosPro.Show vbModal
            Set frmMtoBancosPro = Nothing
    End Select
    
    PonerFoco Me.txtCodigo(indCodigo)
End Sub


Private Sub imgFecha_Click(Index As Integer)
    Screen.MousePointer = vbHourglass
    Set frmF = New frmCal
    frmF.Fecha = Now
   
    PonerFormatoFecha txtCodigo(Index)
    If txtCodigo(Index).Text <> "" Then frmF.Fecha = CDate(txtCodigo(Index).Text)

    Screen.MousePointer = vbDefault
    frmF.Show vbModal
    Set frmF = Nothing
    PonerFoco txtCodigo(Index)
End Sub




Private Sub OptImporte_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub optList_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtCodigo_GotFocus(Index As Integer)
    ConseguirFoco txtCodigo(Index), 3
End Sub

Private Sub txtCodigo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtCodigo_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub txtCodigo_LostFocus(Index As Integer)
    'quitar espacios en blanco por los lados
    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    Select Case Index
        Case 0, 2, 10: PonerFormatoEntero txtCodigo(Index) 'AÑO
        Case 1, 3, 5, 11, 13 'MES
            Me.Label4(1).Caption = ""
            Me.Label4(2).Caption = ""
            If PonerFormatoEntero(txtCodigo(Index)) Then
                If EsMesOK(txtCodigo(Index)) Then
                    'Es un poco gore, pero nio lo estaba haciendo yo
                    Me.Label4(1).Caption = UCase(MonthName(txtCodigo(Index).Text))
                    Me.Label4(2).Caption = UCase(MonthName(txtCodigo(Index).Text))
                     Me.Label4(3).Caption = UCase(MonthName(txtCodigo(Index).Text))
                Else
                    MsgBox "El mes introducido no es correcto.", vbInformation
                    Me.txtCodigo(Index).Text = ""
                    PonerFoco Me.txtCodigo(Index)
                End If
            End If
            
        Case 6, 7, 15, 16 'Trabajador
            If PonerFormatoEntero(txtCodigo(Index)) Then
                txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, "straba", "nomtraba", "codtraba", "Trabajador", "N")
                If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "0000")
            Else
                txtNombre(Index).Text = ""
            End If
            
        Case 8 'BANCO PROPIO
            txtNombre(Index).Text = ""
            txtNombre(0).Text = ""
            
            If PonerFormatoEntero(txtCodigo(Index)) Then
                txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, "sbanpr", "nombanpr", "codbanpr", "Bancos Propios", "N")
                If txtCodigo(Index).Text <> "" And txtNombre(Index).Text <> "" Then
                    PonerCamposBanco Index, 0
                Else
                    PonerFoco txtCodigo(Index)
                End If
            End If
        
        Case 9, 14 'FECHA
            PonerFormatoFecha txtCodigo(Index)
            
        Case 17, 18  'importes
            If Not PonerFormatoDecimal(txtCodigo(Index), 3) Then
                If txtCodigo(Index).Text <> "" Then PonerFoco txtCodigo(Index)
                txtCodigo(Index).Text = ""
                
            End If
    End Select
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 2, cerrar
    If cerrar Then Unload Me
End Sub



Private Sub PonerCamposBanco(indCod As Integer, indNom As String)
Dim ctaB As String
Dim ctaC As String
    
    txtCodigo(indCod).Text = Format(txtCodigo(indCod).Text, "0000")
    
    ObtenerCtasBancoPropio2 txtCodigo(indCod).Text, ctaB, ctaC
    Me.txtNombre(indNom).Text = ctaC
End Sub



Private Sub MostrarProgreso(mostrar As Boolean)
    Me.Line1.visible = mostrar
    Me.lblProgreso.visible = mostrar
    Me.ProgressBar1.visible = mostrar
    
    If mostrar Then
        Me.FrameNorma34.Height = 5175
        Me.FrameNorma34.Width = 6375
    Else
        Me.FrameNorma34.Height = 4215
        Me.FrameNorma34.Width = 6375
    End If
    
    '- ajustar tamaño del form
    Me.Width = Me.FrameNorma34.Width + 70
    Me.Height = Me.FrameNorma34.Height + 350
    Me.Refresh
End Sub


Private Sub GenerarNorma34(cadWHERE As String)
Dim totReg As Integer
Dim SQL As String
Dim cad As String
Dim b As Boolean
Dim RS As ADODB.Recordset

Dim CuentaPropia As String
Dim IdOrdenante As String
Dim CadenaSQL As String
Dim ConceptoTr As String  'concepto de la orden
Dim DescripTr As String 'descripcion de la orden

    On Error GoTo ErrGenNorma34
    conn.BeginTrans
    
        
    '-- total registros a processar para ProgressBar
    SQL = "SELECT count(*) FROM snomin WHERE " & cadWHERE
    totReg = TotalRegistros(SQL)
    CargarProgresNew Me.ProgressBar1, totReg

    
    '-- seleccionar registros a procesar: datos y cuenta bancaria trabajador
    If Me.OptImporte(0).Value = True Then 'Nominas
        SQL = "SELECT snomin.codtraba,sum(impnomi) as Importe"
        SQL = SQL & ",straba.codbanco as entidad,straba.codsucur as oficina"
        SQL = SQL & ",straba.digcontr as CC,straba.cuentaba as cuentaba ,straba.iban iban"
        ConceptoTr = "1"
        DescripTr = "Pago Nomina " & UCase(Format("20/" & txtCodigo(11).Text & "/2000", "mmmm")) & " - " & txtCodigo(10).Text
    Else 'Gastos
        'Marzo 2010. Gastos - anticipado
        SQL = "SELECT snomin.codtraba,sum(impgasto-impanti) as Importe"
        SQL = SQL & ",straba.codbanc1 as entidad,straba.codsucu1 as oficina"
        SQL = SQL & ",straba.digcont1 as CC,straba.cuentab1 as cuentaba,straba.iban1 iban "
        ConceptoTr = "9"
        DescripTr = "Transferencia"
    End If
    SQL = SQL & ",straba.nomtraba as nommacta, straba.domtraba as dirdatos, straba.codpobla as codposta,straba.pobtraba as despobla,straba.niftraba as refbenef,iban"
    SQL = SQL & " FROM snomin, straba "
    SQL = SQL & " WHERE snomin.codtraba = straba.codtraba and " & cadWHERE
    SQL = SQL & " GROUP BY codtraba"
    CadenaSQL = SQL


    '-- obtener la cuenta bancaria del banco propio (Ordenante)
    SQL = "select codbanco, codsucur, digcontr, cuentaba,idnorma34,iban from sbanpr where codbanpr = " & DBSet(txtCodigo(8).Text, "N")
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If RS.EOF Then
        cad = ""
        IdOrdenante = ""
    Else
        IdOrdenante = DBLet(RS!idnorma34, "T")
        If IsNull(RS!codbanco) Then
            cad = ""
        Else
            cad = Format(RS!codbanco, "0000") & "|" & Format(DBLet(RS!codsucur, "T"), "0000") & "|" & DBLet(RS!digcontr, "T") & "|" & Format(DBLet(RS!cuentaba, "T"), "0000000000") & "|" & DBLet(RS!IBAN, "T") & "|"
        End If
    End If
    Set RS = Nothing
    CuentaPropia = cad
    
    '-- generar el fichero
    If Trim(IdOrdenante) = "" Then IdOrdenante = vParam.CifEmpresa
    b = GeneraFicheroNorma34_ARIGES(IdOrdenante, CDate(txtCodigo(9).Text), CuentaPropia, ConceptoTr, 0, DescripTr, False, CadenaSQL)
    
    
    '-- marcar los registros procesados de snomin
    If b Then
        If Me.OptImporte(0).Value = True Then 'nomina procesada
            SQL = "UPDATE snomin SET n34nomi=1"
        Else 'gastos procesados
            SQL = "UPDATE snomin SET n34gast=1"
        End If
        SQL = SQL & " WHERE " & cadWHERE
        conn.Execute SQL
    End If
    
    
    If b Then
        b = CopiarFichero
'        If CopiarFichero Then
'            If MsgBox("¿Proceso realizado correctamente para actualizar?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
'                SQL = "update horas, straba, forpago set horas.intconta = 1 where horas.codtraba = straba.codtraba and straba.codforpa = forpago.codforpa and " & cadWhere
'                conn.Execute SQL
'            End If
'        End If
    End If


    If b Then
        conn.CommitTrans
        MsgBox "Proceso realizado correctamente.", vbInformation
    Else
        conn.RollbackTrans
        MsgBox "Error. Proceso NO realizado.", vbExclamation
    End If
    
    Exit Sub

ErrGenNorma34:
    conn.RollbackTrans
    MuestraError Err.Number, "Generar Norma 34.", Err.Description
End Sub






Private Sub ProcesarCambios(cadWHERE As String)
'Dim totReg As Long
'
'Dim SQL As String
'Dim SQL2 As String
'Dim Sql3 As String
'Dim cad As String
'Dim I As Integer
'Dim HayReg As Integer
'Dim b As Boolean
'Dim Rs As ADODB.Recordset
'Dim Rs2 As ADODB.Recordset
'Dim mens As String
'
'Dim ImpHoras As Currency
'Dim ImpHorasE As Currency
'Dim ImpBruto As Currency
'Dim IRPF As Currency
'Dim SegSoc As Currency
'Dim Neto As Currency
'Dim Bruto As Currency
'Dim CuentaPropia As String
'
'    On Error GoTo eProcesarCambios
'
'    conn.BeginTrans
'
'    If cadWhere <> "" Then
'        cadWhere = QuitarCaracterACadena(cadWhere, "{")
'        cadWhere = QuitarCaracterACadena(cadWhere, "}")
'        cadWhere = QuitarCaracterACadena(cadWhere, "_1")
'    End If
'
'
'    SQL = "select count(distinct horas.codtraba) from (horas inner join straba on horas.codtraba = straba.codtraba) inner join forpago on straba.codforpa = forpago.codforpa where " & cadWhere
'
'    Set Rs = New ADODB.Recordset
'    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'    Me.ProgressBar1.visible = True
'    CargarProgres Me.ProgressBar1, Rs.Fields(0).Value
'
'    Rs.Close
'
'
'
'    SQL = "select horas.codtraba, sum(horasdia), sum(compleme), sum(horasext) from (horas inner join straba on horas.codtraba = straba.codtraba) inner join forpago on straba.codforpa = forpago.codforpa where " & cadWhere
'    SQL = SQL & " group by horas.codtraba "
'
'    BorrarTMP
'    CrearTMP
'
'    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'    While Not Rs.EOF
'        IncrementarProgres Me.ProgressBar1, 1
'        mens = "Calculando Importes" & vbCrLf & vbCrLf & "Trabajador: " & Rs!CodTraba & vbCrLf
'
'        SQL2 = "select salarios.* from salarios, straba where straba.codtraba = " & DBSet(Rs!CodTraba, "N")
'        SQL2 = SQL2 & " and salarios.codcateg = straba.codcateg "
'
'        Set Rs2 = New ADODB.Recordset
'        Rs2.Open SQL2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'        ImpHoras = Round2(DBLet(Rs.Fields(1).Value, "N") * DBLet(Rs2!impsalar, "N"), 2)
'        ImpHorasE = Round2(DBLet(Rs.Fields(3).Value, "N") * DBLet(Rs2!imphorae, "N"), 2)
'        ImpBruto = Round2(ImpHoras + ImpHorasE + DBLet(Rs.Fields(2).Value, "N"), 2)
'
'        IRPF = Round2(ImpBruto * DBLet(Rs2!dtosirpf, "N") / 100, 2)
'        SegSoc = Round2(ImpBruto * DBLet(Rs2!dtosegso, "N") / 100, 2)
'
'        Neto = Round2(ImpBruto - IRPF - SegSoc, 2)
'
'        Sql3 = "insert into tmpImpor (codtraba, importe) values ("
'        Sql3 = Sql3 & DBSet(Rs.Fields(0).Value, "N") & "," & DBSet(ImporteSinFormato(CStr(Neto)), "N") & ")"
'
'        conn.Execute Sql3
'
'        Set Rs2 = Nothing
'
'        Rs.MoveNext
'    Wend
'    Set Rs = Nothing
'
'    SQL = "select codbanco, codsucur, digcontr, cuentaba from banpropi where codbanpr = " & DBSet(txtCodigo(18).Text, "N")
'    Set Rs = New ADODB.Recordset
'    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'    If Rs.EOF Then
'        cad = ""
'    Else
'        If IsNull(Rs!codbanco) Then
'            cad = ""
'        Else
'            cad = Format(Rs!codbanco, "0000") & "|" & Format(DBLet(Rs!codsucur, "T"), "0000") & "|" & DBLet(Rs!digcontr, "T") & "|" & Format(DBLet(Rs!CuentaBa, "T"), "0000000000") & "|"
'        End If
'    End If
'
'    Set Rs = Nothing
'
'    CuentaPropia = cad
'    b = GeneraFicheroNorma34New(vParam.CifEmpresa, CDate(txtCodigo(20).Text), CuentaPropia, 9, 0, "Pago Nómina", False)
'    If b Then
'        If CopiarFichero Then
'            If MsgBox("¿Proceso realizado correctamente para actualizar?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
'                SQL = "update horas, straba, forpago set horas.intconta = 1 where horas.codtraba = straba.codtraba and straba.codforpa = forpago.codforpa and " & cadWhere
'                conn.Execute SQL
'            End If
'        End If
'    End If
'
'eProcesarCambios:
'    If Err.Number <> 0 Then
'        mens = Err.Description
'        b = False
'    End If
'    If b Then
'        conn.CommitTrans
'        MsgBox "Proceso realizado correctamente.", vbExclamation
'        cmdCancelarNorma34_Click
'    Else
'        conn.RollbackTrans
'        MsgBox "Error " & mens, vbExclamation
'    End If
End Sub



Public Function CopiarFichero() As Boolean
Dim nomFich As String
Dim FichOrig As String
Dim FichDest As String



    CopiarFichero = False
    
    'nombre del fichero
    nomFich = "norma34.txt"
    
    ' abrimos el commondialog para indicar donde guardarlo
'    Me.CommonDialog1.InitDir = App.path

    '- abrir el dialog "Guardar como" con valores por defecto
    CommonDialog1.CancelError = True
    CommonDialog1.DefaultExt = ".txt" 'extension por defecto
    CommonDialog1.Filter = "Archivos txt|*.txt|" 'extensiones a mostrar
    CommonDialog1.FilterIndex = 1
    CommonDialog1.FileName = nomFich 'nombre fichero por defecto
    On Error Resume Next
    Me.CommonDialog1.ShowSave

    If Err.Number <> 0 Then
        Err.Clear
        FichDest = ""
    Else
        FichDest = CommonDialog1.FileName
    End If
    
On Error GoTo ecopiarfichero
    
    '- copiar fichero origen en destino seleccionado en el Dialog
    FichOrig = App.Path & "\" & nomFich
    
    
    If FichDest <> "" Then
        If Dir(FichDest) <> "" Then
            If MsgBox("Ya existe el fichero " & FichDest & vbCrLf & "¿Desea remplazarlo?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                FileCopy FichOrig, FichDest
                CopiarFichero = True
            Else
                FileCopy FichOrig, Replace(FichDest, ".txt", "-copia.txt")
                CopiarFichero = True
            End If
        Else
            FileCopy FichOrig, FichDest
            CopiarFichero = True
        End If
    End If
    
    
    Exit Function

ecopiarfichero:
    MuestraError Err.Number, Err.Description
End Function


Private Sub LeerGuardarImportes(Leer As Boolean)
    On Error GoTo ELeerGuardarImportes
    
    cadParam = App.Path & "\nomin.xdf"
    indCodigo = FreeFile
    If Leer Then
        cadSelect = "||"
        If Dir(cadParam, vbArchive) <> "" Then
            Open cadParam For Input As #indCodigo
            Line Input #indCodigo, cadSelect
            Close #indCodigo
        End If
        txtCodigo(17).Text = RecuperaValor(cadSelect, 1)
        txtCodigo(18).Text = RecuperaValor(cadSelect, 2)
        
    Else
            cadSelect = txtCodigo(17).Text & "|" & txtCodigo(18).Text & "|"
            Open cadParam For Output As #indCodigo
            Print #indCodigo, cadSelect
            Close #indCodigo
            
    End If
    
ELeerGuardarImportes:
    If Err.Number <> 0 Then Err.Clear
    
End Sub



Private Function GeneraDatosRecibos() As Boolean
Dim ListaClientes As Collection
Dim Cuantos As Integer
Dim Aux As Long
Dim Importe As Currency
Dim Base As Currency

    On Error GoTo EGeneraDatosRecibos
    GeneraDatosRecibos = False
    Set miRsAux = New ADODB.Recordset
    conn.Execute "DELETE FROM tmpinformes WHERE codusu = " & vUsu.codigo
    
    'Creo la lista clientes
    cadSelect = "Select codclien,nomclien,pobclien from sclien WHERE not codpobla like '460%' AND clivario=0 "
    cadSelect = cadSelect & " and mid(nifclien,1,1) >=""A"" ORDER BY codclien"
    miRsAux.Open cadSelect, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Set ListaClientes = New Collection
    While Not miRsAux.EOF
        ListaClientes.Add CStr(miRsAux!codClien) & "|" & miRsAux!Nomclien & "|" & DBLet(miRsAux!pobclien, "T") & "|"
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    cadSelect = "select impgasto,impanti ,codtraba  from snomin where snomin.anynomi=" & txtCodigo(12).Text & " AND snomin.mesnomi=" & txtCodigo(13).Text
    If txtCodigo(15).Text <> "" Then cadSelect = cadSelect & " AND codtraba >=" & txtCodigo(15).Text
    If txtCodigo(16).Text <> "" Then cadSelect = cadSelect & " AND codtraba <=" & txtCodigo(16).Text
    
    miRsAux.Open cadSelect, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If miRsAux.EOF Then
        MsgBox "No ha datos", vbExclamation
        miRsAux.Close
        Exit Function
    End If
    
    If Me.optList(0).Value Then
        Base = ImporteFormateado(txtCodigo(17).Text)
    Else
        Base = ImporteFormateado(txtCodigo(18).Text)
    End If
    
    indCodigo = 0
    NumRegElim = 0
    While Not miRsAux.EOF
    
        
        If Me.optList(0).Value Then
            Importe = miRsAux!impgasto
        Else
            'Segun ISABEL es el anticipo
            'Importe = (miRsAux!impgasto - DBLet(miRsAux!impanti, "N"))
            Importe = DBLet(miRsAux!impanti, "N")
        End If
        
        
        
        If Importe = 0 Then
            'Stop
        Else
            cadSelect = miRsAux!CodTraba & txtCodigo(12).Text & Format(txtCodigo(13).Text, "00")
            Aux = CLng(cadSelect)
            
            Randomize Aux
                    
            
            Cuantos = Importe / Base
            
            If Cuantos = 0 Then Cuantos = 1
            
            'Inserto los clientes que necesite
            cadSelect = ""
            For indCodigo = 1 To Cuantos
                NumRegElim = NumRegElim + 1
                Aux = Int((ListaClientes.Count * Rnd) + 1)
                cadSelect = "insert into `tmpinformes` (`codusu`,`codigo1`,`campo1`,`campo2`,nombre1,nombre2) VALUES ("
                cadSelect = cadSelect & vUsu.codigo & "," & NumRegElim & "," & miRsAux!CodTraba & ","
                cadSelect = cadSelect & RecuperaValor(ListaClientes(Aux), 1) & ",'"
                cadSelect = cadSelect & DevNombreSQL(RecuperaValor(ListaClientes(Aux), 2)) & "','"
                cadSelect = cadSelect & DevNombreSQL(RecuperaValor(ListaClientes(Aux), 3)) & "')"
                conn.Execute cadSelect
            Next
            
       
            
            
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    GeneraDatosRecibos = True
    
EGeneraDatosRecibos:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set miRsAux = Nothing
    Set ListaClientes = Nothing
End Function



