VERSION 5.00
Begin VB.Form frmCRMVarios 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CRM"
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9330
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6990
   ScaleWidth      =   9330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameClientesxAccion 
      Height          =   6135
      Left            =   1530
      TabIndex        =   61
      Top             =   90
      Visible         =   0   'False
      Width           =   6615
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   2
         Left            =   3600
         TabIndex        =   83
         Text            =   "Text1"
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox txtNumero 
         Height          =   285
         Index           =   9
         Left            =   1200
         TabIndex        =   79
         Text            =   "Text1"
         Top             =   3120
         Width           =   975
      End
      Begin VB.TextBox txtDescNumero 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   9
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   78
         Text            =   "Text1"
         Top             =   3120
         Width           =   3495
      End
      Begin VB.TextBox txtNumero 
         Height          =   285
         Index           =   8
         Left            =   1200
         TabIndex        =   75
         Text            =   "Text1"
         Top             =   2760
         Width           =   975
      End
      Begin VB.TextBox txtDescNumero 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   8
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   74
         Text            =   "Text1"
         Top             =   2760
         Width           =   3495
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   2
         Left            =   4920
         TabIndex        =   73
         Top             =   5520
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Generar"
         Height          =   375
         Left            =   1680
         TabIndex        =   72
         Top             =   5520
         Width           =   1095
      End
      Begin VB.TextBox txtDescClie 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   5
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   70
         Text            =   "Text1"
         Top             =   2040
         Width           =   3495
      End
      Begin VB.TextBox txtCliente 
         Height          =   285
         Index           =   5
         Left            =   1200
         TabIndex        =   69
         Text            =   "Text1"
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox txtDescClie 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   4
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   66
         Text            =   "Text1"
         Top             =   1680
         Width           =   3495
      End
      Begin VB.TextBox txtCliente 
         Height          =   285
         Index           =   4
         Left            =   1200
         TabIndex        =   65
         Text            =   "Text1"
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   1
         Left            =   1200
         TabIndex        =   63
         Text            =   "Text1"
         Top             =   960
         Width           =   1335
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   2
         Left            =   3360
         Picture         =   "frmCRMVarios.frx":0000
         Top             =   960
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   15
         Left            =   2880
         TabIndex        =   82
         Top             =   960
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   14
         Left            =   360
         TabIndex        =   81
         Top             =   960
         Width           =   465
      End
      Begin VB.Image imgVarios 
         Height          =   240
         Index           =   9
         Left            =   960
         Picture         =   "frmCRMVarios.frx":008B
         Top             =   3120
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   13
         Left            =   360
         TabIndex        =   80
         Top             =   3120
         Width           =   465
      End
      Begin VB.Image imgVarios 
         Height          =   240
         Index           =   8
         Left            =   960
         Picture         =   "frmCRMVarios.frx":018D
         Top             =   2760
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
         Index           =   10
         Left            =   240
         TabIndex        =   77
         Top             =   2520
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   12
         Left            =   360
         TabIndex        =   76
         Top             =   2760
         Width           =   465
      End
      Begin VB.Image imgCliente 
         Height          =   240
         Index           =   5
         Left            =   960
         Picture         =   "frmCRMVarios.frx":028F
         Top             =   2040
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   11
         Left            =   360
         TabIndex        =   71
         Top             =   2040
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
         Index           =   9
         Left            =   240
         TabIndex        =   68
         Top             =   1440
         Width           =   585
      End
      Begin VB.Image imgCliente 
         Height          =   240
         Index           =   4
         Left            =   960
         Picture         =   "frmCRMVarios.frx":0391
         Top             =   1680
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   10
         Left            =   360
         TabIndex        =   67
         Top             =   1680
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
         Index           =   7
         Left            =   240
         TabIndex        =   64
         Top             =   720
         Width           =   540
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   1
         Left            =   960
         Picture         =   "frmCRMVarios.frx":0493
         Top             =   960
         Width           =   240
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Listado Clientes x Acción comercial"
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
         Left            =   720
         TabIndex        =   62
         Top             =   240
         Width           =   5295
      End
   End
   Begin VB.Frame FrameGenerar 
      Height          =   6735
      Left            =   120
      TabIndex        =   13
      Top             =   0
      Width           =   6855
      Begin VB.TextBox txtDescAccion 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   0
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   39
         Text            =   "Text1"
         Top             =   1560
         Width           =   3735
      End
      Begin VB.TextBox txtAccion 
         Height          =   285
         Index           =   0
         Left            =   1200
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   1560
         Width           =   855
      End
      Begin VB.CommandButton cmdGeneAcciones 
         Caption         =   "Generar"
         Height          =   375
         Left            =   5280
         TabIndex        =   11
         Top             =   5640
         Width           =   1095
      End
      Begin VB.TextBox txtNumero 
         Height          =   285
         Index           =   5
         Left            =   1200
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   6120
         Width           =   975
      End
      Begin VB.TextBox txtDescNumero 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   5
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   28
         Text            =   "Text1"
         Top             =   6120
         Width           =   2295
      End
      Begin VB.TextBox txtNumero 
         Height          =   285
         Index           =   4
         Left            =   1200
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   5760
         Width           =   975
      End
      Begin VB.TextBox txtDescNumero 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   4
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   26
         Text            =   "Text1"
         Top             =   5760
         Width           =   2295
      End
      Begin VB.TextBox txtNumero 
         Height          =   285
         Index           =   3
         Left            =   1200
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   5280
         Width           =   975
      End
      Begin VB.TextBox txtDescNumero 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   3
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   25
         Text            =   "Text1"
         Top             =   5280
         Width           =   2295
      End
      Begin VB.TextBox txtNumero 
         Height          =   285
         Index           =   2
         Left            =   1200
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   4920
         Width           =   975
      End
      Begin VB.TextBox txtDescNumero 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   2
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   23
         Text            =   "Text1"
         Top             =   4920
         Width           =   2295
      End
      Begin VB.TextBox txtNumero 
         Height          =   285
         Index           =   1
         Left            =   1200
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   4320
         Width           =   975
      End
      Begin VB.TextBox txtDescNumero 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   1
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   21
         Text            =   "Text1"
         Top             =   4320
         Width           =   3495
      End
      Begin VB.TextBox txtDescNumero 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   0
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   20
         Text            =   "Text1"
         Top             =   3960
         Width           =   3495
      End
      Begin VB.TextBox txtNumero 
         Height          =   285
         Index           =   0
         Left            =   1200
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   3960
         Width           =   975
      End
      Begin VB.TextBox txtDescClie 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   1
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   19
         Text            =   "Text1"
         Top             =   3360
         Width           =   3495
      End
      Begin VB.TextBox txtCliente 
         Height          =   285
         Index           =   1
         Left            =   1200
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   3360
         Width           =   975
      End
      Begin VB.TextBox txtCliente 
         Height          =   285
         Index           =   0
         Left            =   1200
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   3000
         Width           =   975
      End
      Begin VB.TextBox txtDescClie 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   0
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   3000
         Width           =   3495
      End
      Begin VB.TextBox txtDescTra 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   0
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   1080
         Width           =   3735
      End
      Begin VB.TextBox txtTrab 
         Height          =   285
         Index           =   0
         Left            =   1200
         TabIndex        =   1
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   0
         Left            =   1200
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   0
         Left            =   5280
         TabIndex        =   12
         Top             =   6120
         Width           =   1095
      End
      Begin VB.Image imgAccion 
         Height          =   240
         Index           =   0
         Left            =   960
         Picture         =   "frmCRMVarios.frx":051E
         Top             =   1560
         Width           =   240
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Accion"
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
         TabIndex        =   38
         Top             =   1560
         Width           =   555
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   6000
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Generar entrada de acciones comerciales"
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
         TabIndex        =   37
         Top             =   240
         Width           =   6495
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   5
         Left            =   360
         TabIndex        =   36
         Top             =   5760
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   4
         Left            =   360
         TabIndex        =   35
         Top             =   6120
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   3
         Left            =   360
         TabIndex        =   34
         Top             =   4920
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   2
         Left            =   360
         TabIndex        =   33
         Top             =   5280
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   32
         Top             =   3960
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   31
         Top             =   4320
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   58
         Left            =   360
         TabIndex        =   30
         Top             =   3000
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   59
         Left            =   360
         TabIndex        =   29
         Top             =   3360
         Width           =   465
      End
      Begin VB.Image imgVarios 
         Height          =   240
         Index           =   5
         Left            =   960
         Picture         =   "frmCRMVarios.frx":0620
         Top             =   6120
         Width           =   240
      End
      Begin VB.Label lblDpto 
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
         Index           =   3
         Left            =   240
         TabIndex        =   27
         Top             =   5520
         Width           =   405
      End
      Begin VB.Image imgVarios 
         Height          =   240
         Index           =   4
         Left            =   960
         Picture         =   "frmCRMVarios.frx":0722
         Top             =   5760
         Width           =   240
      End
      Begin VB.Image imgVarios 
         Height          =   240
         Index           =   3
         Left            =   960
         Picture         =   "frmCRMVarios.frx":0824
         Top             =   4320
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
         Index           =   2
         Left            =   240
         TabIndex        =   24
         Top             =   4680
         Width           =   420
      End
      Begin VB.Image imgVarios 
         Height          =   240
         Index           =   2
         Left            =   960
         Picture         =   "frmCRMVarios.frx":0926
         Top             =   4920
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
         Index           =   1
         Left            =   240
         TabIndex        =   22
         Top             =   3720
         Width           =   615
      End
      Begin VB.Image imgVarios 
         Height          =   240
         Index           =   1
         Left            =   960
         Picture         =   "frmCRMVarios.frx":0A28
         Top             =   4320
         Width           =   240
      End
      Begin VB.Image imgVarios 
         Height          =   240
         Index           =   0
         Left            =   960
         Picture         =   "frmCRMVarios.frx":0B2A
         Top             =   3960
         Width           =   240
      End
      Begin VB.Image imgCliente 
         Height          =   240
         Index           =   1
         Left            =   960
         Picture         =   "frmCRMVarios.frx":0C2C
         Top             =   3360
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
         Index           =   0
         Left            =   240
         TabIndex        =   18
         Top             =   1080
         Width           =   645
      End
      Begin VB.Image imgCliente 
         Height          =   240
         Index           =   0
         Left            =   960
         Picture         =   "frmCRMVarios.frx":0D2E
         Top             =   3000
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
         Index           =   8
         Left            =   240
         TabIndex        =   17
         Top             =   2760
         Width           =   585
      End
      Begin VB.Image imgTecnico 
         Height          =   240
         Index           =   0
         Left            =   960
         Picture         =   "frmCRMVarios.frx":0E30
         Top             =   1080
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   960
         Picture         =   "frmCRMVarios.frx":0F32
         Top             =   720
         Width           =   240
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
         Index           =   33
         Left            =   240
         TabIndex        =   14
         Top             =   720
         Width           =   540
      End
   End
   Begin VB.Frame FrameCRMresumen 
      Height          =   4695
      Left            =   840
      TabIndex        =   40
      Top             =   240
      Visible         =   0   'False
      Width           =   6015
      Begin VB.TextBox txtDescNumero 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   7
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   59
         Text            =   "Text1"
         Top             =   2760
         Width           =   3495
      End
      Begin VB.TextBox txtNumero 
         Height          =   285
         Index           =   7
         Left            =   1200
         TabIndex        =   45
         Text            =   "Text1"
         Top             =   2760
         Width           =   975
      End
      Begin VB.TextBox txtNumero 
         Height          =   285
         Index           =   6
         Left            =   1200
         TabIndex        =   44
         Text            =   "Text1"
         Top             =   2400
         Width           =   975
      End
      Begin VB.TextBox txtDescNumero 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   6
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   56
         Text            =   "Text1"
         Top             =   2400
         Width           =   3495
      End
      Begin VB.CommandButton cmdResumen 
         Caption         =   "Generar"
         Height          =   375
         Left            =   3360
         TabIndex        =   48
         Top             =   4200
         Width           =   1095
      End
      Begin VB.CheckBox chkVarios 
         Caption         =   "Check1"
         Height          =   255
         Index           =   1
         Left            =   1920
         TabIndex        =   47
         Top             =   3360
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CheckBox chkVarios 
         Caption         =   "Salta pagina"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   46
         Top             =   3360
         Width           =   1215
      End
      Begin VB.TextBox txtCliente 
         Height          =   285
         Index           =   3
         Left            =   1200
         TabIndex        =   43
         Text            =   "Text1"
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox txtDescClie 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   3
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   53
         Text            =   "Text1"
         Top             =   1680
         Width           =   3495
      End
      Begin VB.TextBox txtDescClie 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   2
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   50
         Text            =   "Text1"
         Top             =   1320
         Width           =   3495
      End
      Begin VB.TextBox txtCliente 
         Height          =   285
         Index           =   2
         Left            =   1200
         TabIndex        =   42
         Text            =   "Text1"
         Top             =   1320
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   1
         Left            =   4680
         TabIndex        =   49
         Top             =   4200
         Width           =   1095
      End
      Begin VB.Image imgVarios 
         Height          =   240
         Index           =   7
         Left            =   960
         Picture         =   "frmCRMVarios.frx":0FBD
         Top             =   2760
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   9
         Left            =   360
         TabIndex        =   60
         Top             =   2760
         Width           =   465
      End
      Begin VB.Image imgVarios 
         Height          =   240
         Index           =   6
         Left            =   960
         Picture         =   "frmCRMVarios.frx":10BF
         Top             =   2400
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
         Index           =   6
         Left            =   240
         TabIndex        =   58
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   315
         Index           =   8
         Left            =   360
         TabIndex        =   57
         Top             =   2400
         Width           =   465
      End
      Begin VB.Label lblIndicador 
         Caption         =   "Label1"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   55
         Top             =   3840
         Width           =   5295
      End
      Begin VB.Image imgCliente 
         Height          =   240
         Index           =   3
         Left            =   960
         Picture         =   "frmCRMVarios.frx":11C1
         Top             =   1680
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   7
         Left            =   360
         TabIndex        =   54
         Top             =   1680
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
         Index           =   5
         Left            =   240
         TabIndex        =   52
         Top             =   960
         Width           =   585
      End
      Begin VB.Image imgCliente 
         Height          =   240
         Index           =   2
         Left            =   960
         Picture         =   "frmCRMVarios.frx":12C3
         Top             =   1320
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   6
         Left            =   360
         TabIndex        =   51
         Top             =   1320
         Width           =   465
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Impresión CRM resumida"
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
         Left            =   1320
         TabIndex        =   41
         Top             =   360
         Width           =   3735
      End
   End
End
Attribute VB_Name = "frmCRMVarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public Opcion As Byte
    '   0.- Generacion
    '   1.- Resumen
    '   2.- Acciones x cliente

Private WithEvents frmCli As frmBasico2 'frmFacClientesGr
Attribute frmCli.VB_VarHelpID = -1
Private WithEvents frmT As frmBasico2 'frmAdmTrabajadores
Attribute frmT.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1

Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmAcc As frmBasico2 'frmCRMtipos
Attribute frmAcc.VB_VarHelpID = -1

Dim IndiceImg As Integer
Dim miSQL As String
Dim Codigo As String


Private Sub chkVarios_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 3, False
End Sub

Private Sub cmdCancelar_Click(Index As Integer)
    Unload Me
End Sub

Private Sub cmdGeneAcciones_Click()
    miSQL = ""
    If txtFecha(0).Text = "" Then miSQL = miSQL & "- Fecha debe tener valor" & vbCrLf
    If txtTrab(0).Text = "" Then miSQL = miSQL & "- Trabajador debe tener valor" & vbCrLf
    If txtAccion(0).Text = "" Then miSQL = miSQL & "- Indique la accion comercial" & vbCrLf
    
    If miSQL <> "" Then
        MsgBox miSQL, vbExclamation
        Exit Sub
    End If
    
    CadenaDesdeOtroForm = ""
    miSQL = ""
    
    'Desde hasta
    If txtCliente(0).Text <> "" Then miSQL = miSQL & " AND sclien.codclien >= " & txtCliente(0).Text
    If txtCliente(1).Text <> "" Then miSQL = miSQL & " AND sclien.codclien <= " & txtCliente(1).Text
    'Agente
    If txtNumero(0).Text <> "" Then miSQL = miSQL & " AND sclien.codagent >= " & txtNumero(0).Text
    If txtNumero(1).Text <> "" Then miSQL = miSQL & " AND sclien.codagent <= " & txtNumero(1).Text
    'ZOna
    If txtNumero(2).Text <> "" Then miSQL = miSQL & " AND sclien.codzonas >= " & txtNumero(2).Text
    If txtNumero(3).Text <> "" Then miSQL = miSQL & " AND sclien.codzonas <= " & txtNumero(3).Text
    'RUTA
    If txtNumero(4).Text <> "" Then miSQL = miSQL & " AND sclien.codrutas >= " & txtNumero(4).Text
    If txtNumero(5).Text <> "" Then miSQL = miSQL & " AND sclien.codrutas <= " & txtNumero(5).Text
    
    If miSQL <> "" Then miSQL = Mid(miSQL, 5) 'quito el primer AND
    
    If Not HayRegParaInforme("sclien", miSQL, True) Then
        MsgBox "No hay clientes con estos valores", vbExclamation
        Exit Sub
    End If
    
    CadenaDesdeOtroForm = miSQL
    If CadenaDesdeOtroForm <> "" Then CadenaDesdeOtroForm = " WHERE " & CadenaDesdeOtroForm
    
    frmVarios.Opcion = 4
    frmVarios.Show vbModal
    
    Screen.MousePointer = vbHourglass
    If CadenaDesdeOtroForm <> "" Then
        DoEvents
        GenerarEntradaMasivaAccionesComerciales
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdResumen_Click()
    Screen.MousePointer = vbHourglass
    Set miRsAux = New ADODB.Recordset
    If GenerarResumenCRM Then
        lblIndicador(0).Caption = "Mostrando rpt"
        lblIndicador(0).Refresh
        
        miSQL = "" 'cadparam
        NumRegElim = 3 'numparam
        If Me.txtCliente(2).Text <> "" Then miSQL = " desde  " & Me.txtCliente(2).Text & " " & txtDescClie(2).Text
        If Me.txtCliente(3).Text <> "" Then miSQL = miSQL & " hasta " & Me.txtCliente(3).Text & " " & txtDescClie(3).Text
        If miSQL <> "" Then miSQL = "Cliente: " & miSQL
        CadenaDesdeOtroForm = miSQL
        miSQL = ""
        If txtNumero(6).Text <> "" Then miSQL = miSQL & " desde " & txtNumero(6).Text & " " & Me.txtDescNumero(6).Text
        If txtNumero(7).Text <> "" Then miSQL = miSQL & " hasta " & txtNumero(7).Text & " " & Me.txtDescNumero(7).Text
        If miSQL <> "" Then miSQL = "      Agente: " & miSQL
        miSQL = Trim(CadenaDesdeOtroForm & miSQL)
        
        
        miSQL = "pdh="" " & miSQL & """|SaltaPagina=" & Abs(Me.chkVarios(0).Value) & "|"
   
        miSQL = "|" & vEmpresa.nomempre & "|" & miSQL
        
        
        Codigo = "{tmpcommand.codusu} = " & vUsu.Codigo 'cadformula
        
        LlamarImprimir "rCRMres.rpt", "CRM resumen"
        
    
    
    
    End If
    CadenaDesdeOtroForm = ""
    lblIndicador(0).Caption = ""
    Set miRsAux = Nothing
    Screen.MousePointer = vbDefault
End Sub


Private Sub LlamarImprimir(ByRef cadNomRPT As String, Titulo As String)
    
    With frmImprimir
        .FormulaSeleccion = Codigo
        .OtrosParametros = miSQL
        .NumeroParametros = NumRegElim
        
        .SoloImprimir = False
        .EnvioEMail = False
        .Opcion = 2000 'generico
        .NombrePDF = ""
        .NombreRPT = cadNomRPT
        .Titulo = Titulo
        .ConSubInforme = True
        .Show vbModal
    End With
End Sub


Private Sub Form_Load()
Dim H As Integer
Dim W As Integer

    Me.Icon = frmPpal.Icon
    FrameGenerar.visible = False
    Me.FrameCRMresumen.visible = False
    FrameClientesxAccion.visible = False
    limpiar Me
    
    Select Case Opcion
    Case 0
        PonerFrameVisible FrameGenerar, H, W
        txtFecha(0).Text = Format(Now, "dd/mm/yyyy")
        
        
        txtTrab(0).Text = PonerTrabajadorConectado(miSQL)
        Me.txtDescTra(0).Text = miSQL
        miSQL = ""
        
    Case 1
        'Opcion=1
        lblIndicador(0).Caption = ""
        PonerFrameVisible FrameCRMresumen, H, W
        
    Case 2
        PonerFrameVisible FrameClientesxAccion, H, W
    
    End Select
    
    Me.cmdCancelar(Opcion).Cancel = True
    Me.Height = H
    Me.Width = W
    
End Sub

Private Sub PonerFrameVisible(ByRef F As Frame, ByRef CH As Integer, CW As Integer)
    F.Top = 0
    F.Left = 120
    F.visible = True
    CH = F.Height + 450
    CW = F.Width + 240
End Sub


Private Sub frmAcc_DatoSeleccionado(CadenaSeleccion As String)
    miSQL = CadenaSeleccion
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
    miSQL = CadenaDevuelta
End Sub

Private Sub frmC_Selec(vFecha As Date)
    Me.txtFecha(IndiceImg).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmCli_DatoSeleccionado(CadenaSeleccion As String)
    Me.txtCliente(IndiceImg).Text = RecuperaValor(CadenaSeleccion, 1)
    Me.txtDescClie(IndiceImg).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmT_DatoSeleccionado(CadenaSeleccion As String)
    Me.txtTrab(IndiceImg).Text = RecuperaValor(CadenaSeleccion, 1)
    txtDescTra(IndiceImg).Text = RecuperaValor(CadenaSeleccion, 2)
 
End Sub

Private Sub imgAccion_Click(Index As Integer)
        miSQL = ""
'        Set frmAcc = New frmCRMtipos
'        frmAcc.DatosADevolverBusqueda = "0|1|"
'        frmAcc.Show vbModal
        Set frmAcc = New frmBasico2
        AyudaCRMTipos frmAcc, txtAccion(Index)
        Set frmAcc = Nothing
        If miSQL <> "" Then
            'Por defecto
            'NO dejo que la accon sea del 1 al 20 ya que las reservamos para otros menesteres
            If Val(RecuperaValor(miSQL, 1)) <= 20 Then
                MsgBox "Codigos reservados para la aplicacion", vbExclamation
                
            Else
                txtAccion(Index).Text = RecuperaValor(miSQL, 1)  'Pongo EL ID
                txtDescAccion(Index).Text = RecuperaValor(miSQL, 2)
            End If
        End If
End Sub

Private Sub imgCliente_Click(Index As Integer)
    
    Screen.MousePointer = vbHourglass
    IndiceImg = Index
'    Set frmCli = New frmFacClientesGr
'    frmCli.DatosADevolverBusqueda = "0|1|"
'    frmCli.Show vbModal
    Set frmCli = New frmBasico2
    AyudaClientes frmCli, txtCliente(Index).Text
    Set frmCli = Nothing

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



Private Sub imgTecnico_Click(Index As Integer)
   
'        Set frmT = New frmAdmTrabajadores
'        frmT.DatosADevolverBusqueda = "0|1|"
'        frmT.Show vbModal
        Set frmT = New frmBasico2
        AyudaTrabajadores frmT, txtTrab(Index)
        Set frmT = Nothing
End Sub

Private Sub imgVarios_Click(Index As Integer)
Dim campo As String

    miSQL = ""
    
    'EN codigo:
    'titulo|tabla|sql|
    
    Set frmB = New frmBuscaGrid
    Select Case Index
    Case 0, 1
        'AGENTE
        campo = "Cod.|sagent|codagent|N||20·"
        campo = campo & "Nombre|sagent|nomagent|T||40·"
        
        Codigo = "Agente|sagent||"
    Case 2, 3
        'ZONA
        campo = "Cod.|szonas|codzonas|N||20·"
        campo = campo & "Desc.|szonas|nomzonas|T||40·"
        
        Codigo = "ZONAS|szonas||"
    
    Case 4, 5
        'RUTA
        campo = "Cod.|srutas|codrutas|N||20·"
        campo = campo & "Desc.|srutas|nomrutas|T||40·"
        
        Codigo = "Rutas|srutas||"
    End Select
    frmB.vCampos = campo
    frmB.vTitulo = RecuperaValor(Codigo, 1)
    frmB.vTabla = RecuperaValor(Codigo, 2)
    frmB.vSQL = RecuperaValor(Codigo, 3)
    frmB.vCargaFrame = False
    frmB.vDevuelve = "0|1|"
    frmB.vselElem = 1
    frmB.vConexionGrid = 1  'ODBC Ariges
    frmB.Show vbModal
    Set frmB = Nothing
    Screen.MousePointer = vbDefault
            
    If miSQL <> "" Then
        txtNumero(Index).Text = RecuperaValor(miSQL, 1)
        txtDescNumero(Index) = RecuperaValor(miSQL, 2)
            
            
        miSQL = ""
    End If

End Sub

Private Sub txtAccion_GotFocus(Index As Integer)
    ConseguirFoco txtAccion(Index), 3
End Sub

Private Sub txtAccion_KeyPress(Index As Integer, KeyAscii As Integer)
     KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtAccion_LostFocus(Index As Integer)
     miSQL = ""
    txtAccion(Index).Text = Trim(txtAccion(Index).Text)
    If txtAccion(Index).Text <> "" Then
        If Not IsNumeric(txtAccion(Index).Text) Then
            MsgBox "Campo accion debe ser numérico", vbExclamation
            txtAccion(Index).Text = ""
            PonerFoco txtAccion(Index)
        Else
            If Val(txtAccion(Index).Text) < 21 Then
                MsgBox "Las 20 primeras se las reserva la aplicacion", vbExclamation
                miSQL = ""
            Else
                miSQL = DevuelveDesdeBD(conAri, "denominacion", "scrmtipo", "codigo", txtAccion(Index).Text, "N")
                If miSQL = "" Then MsgBox "No existe la accion comercial : " & txtAccion(Index).Text, vbExclamation
            End If
            If miSQL = "" Then
                txtAccion(Index).Text = ""
                PonerFoco txtAccion(Index)
            End If
        End If
    End If
    Me.txtDescAccion(Index).Text = miSQL
End Sub

Private Sub txtCliente_GotFocus(Index As Integer)
    ConseguirFoco txtCliente(Index), 3
End Sub

Private Sub txtCliente_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtCliente_LostFocus(Index As Integer)

    
    miSQL = ""
    txtCliente(Index).Text = Trim(txtCliente(Index).Text)
    If txtCliente(Index).Text <> "" Then
        If Not IsNumeric(txtCliente(Index).Text) Then
            MsgBox "Campo codigo cliente debe ser numérico", vbExclamation
            txtCliente(Index).Text = ""
            PonerFoco txtCliente(Index)
        Else
            miSQL = DevuelveDesdeBD(conAri, "nomclien", "sclien", "codclien", txtCliente(Index).Text, "N")
            If miSQL = "" Then
                MsgBox "No existe el cliente : " & txtCliente(Index).Text, vbExclamation
            End If
        End If
    End If
    Me.txtDescClie(Index).Text = miSQL
    
    
    
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
            PonerFoco txtFecha(Index)
        End If
    End If
    
End Sub





Private Sub txtNumero_GotFocus(Index As Integer)
     ConseguirFoco txtNumero(Index), 3
End Sub

Private Sub txtNumero_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtNumero_LostFocus(Index As Integer)

    txtNumero(Index).Text = Trim(txtNumero(Index).Text)
    miSQL = ""
    If txtNumero(Index).Text <> "" Then
        If Not IsNumeric(txtNumero(Index).Text) Then
            MsgBox "Campo debe ser numérico: " & txtNumero(Index).Text, vbExclamation
            txtNumero(Index).Text = ""
            PonerFoco txtNumero(Index)
        Else
            'Segun sea
            Select Case Index
            Case 0, 1, 6, 7
                'AGENTE
                miSQL = DevuelveDesdeBD(conAri, "nomagent", "sagent", "codagent", txtNumero(Index).Text, "N")
            Case 2, 3
                'ZONA
                miSQL = DevuelveDesdeBD(conAri, "nomzonas", "szonas", "codzonas", txtNumero(Index).Text, "N")
            Case 4, 5
                'RUTA
                miSQL = DevuelveDesdeBD(conAri, "nomrutas", "srutas", "codrutas", txtNumero(Index).Text, "N")
            End Select

            If miSQL = "" Then
                MsgBox "No existe el codigo: " & txtNumero(Index).Text, vbExclamation
                
                'Si obligaramos a que existiera el codig
                
            End If
        End If
    End If
    txtDescNumero(Index).Text = miSQL
    
End Sub

Private Sub txtTrab_GotFocus(Index As Integer)
    ConseguirFoco txtTrab(Index), 3
End Sub

Private Sub txtTrab_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtTrab_LostFocus(Index As Integer)


    txtTrab(Index).Text = Trim(txtTrab(Index).Text)
    Codigo = ""
    miSQL = ""

    If txtTrab(Index).Text <> "" Then
        If IsNumeric(txtTrab(Index).Text) Then
            Codigo = DevuelveDesdeBD(conAri, "nomtraba", "straba", "codtraba", txtTrab(Index).Text, "N")
            If Codigo = "" Then miSQL = "El código no pertence a ningun trabajador"
        Else
            miSQL = "Campo numerico"
        End If
    End If
    
    Me.txtDescTra(Index).Text = Codigo
    If miSQL <> "" Then
        MsgBox miSQL, vbExclamation
        If Index < 2 Then
            txtTrab(Index).Text = ""
            PonerFoco txtTrab(Index)
        End If
    End If
End Sub

Private Sub GenerarEntradaMasivaAccionesComerciales()

    
    Set miRsAux = New ADODB.Recordset
    miSQL = "Select * from scrmtipo where codigo = " & txtAccion(0).Text
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If miRsAux.EOF Then
        MsgBox "No se encuentra la accion: " & txtAccion(0).Text & " " & txtDescAccion(0).Text, vbExclamation
        
    Else
    
        miSQL = "insert into `scrmacciones` (`usuario`,`fechora`,`codclien`,`agente`,`codtraba`,`estado`,"
        miSQL = miSQL & "`tipo`,`medio`,`observaciones`) select '" & DevNombreSQL(vUsu.Login) & "','"
        miSQL = miSQL & Format(txtFecha(0).Text, FormatoFecha) & " " & Format(Now, "hh:mm:ss") & "',"
        miSQL = miSQL & "codclien,codagent," & txtTrab(0).Text & ",0,"
        'tipo, medio observaciones
        miSQL = miSQL & txtAccion(0).Text & "," & DBSet(miRsAux!medio, "T") & "," & DBSet(miRsAux!Observaciones, "T")
        miSQL = miSQL & " FROM sclien where codclien in (" & CadenaDesdeOtroForm & ")"
        ejecutar miSQL, False
    
    End If
    miRsAux.Close
    Set miRsAux = Nothing

End Sub

        
        


'**************************************************
'**************************************************
'
'       Resmune con D/H cliente para
'
Private Function GenerarResumenCRM() As Boolean
Dim linea() As String
Dim J As Integer
Dim K As Integer
Dim R As ADODB.Recordset
Dim Aux2 As String
Dim Importe As Currency
Dim Cad As String
Dim Vec As Byte

    On Error GoTo eGenerarResumenCRM

    lblIndicador(0).Caption = "Preparando datos"
    conn.Execute "DELETE FROM tmpcommand WHERE codusu =" & vUsu.Codigo
    conn.Execute "DELETE FROM tmpinformes WHERE codusu =" & vUsu.Codigo


    'Los datos del cliente, de momento, estamos guardandolos en la tmp
    'Puede llegar un punto en el que necesitemos mas campos por lo tanto
    ' podriamos linkar en el rpt tmpinformes con cliente y sacar de ahi datos


    'Metemos los clientes con sus datos basicos
    '                   cliente  nombre    tfno     Fpago    situacin  cta
    'tmpcommand(codusu,codprove,nomprove,nomfamia,nomartic,codartic,importel)
    miSQL = "select " & vUsu.Codigo & ",codclien,nomclien,telclie1,"
    'miSQL = miSQL & " concat(right(concat(""0000"",sclien.codforpa),4),' ',nomforpa),"
    miSQL = miSQL & " nomforpa,"
    'miSQL = miSQL & " concat(right(concat(""000"",sclien.codsitua),3),' ',nomsitua),codmacta" ''
    miSQL = miSQL & " nomsitua,codmacta,codagent"
    miSQL = miSQL & " From sclien, sforpa, ssitua"
    miSQL = miSQL & " Where sclien.codforpa = sforpa.codforpa And sclien.codsitua = ssitua.codsitua"
    'Desde hasta
    If txtCliente(2).Text <> "" Then miSQL = miSQL & " AND sclien.codclien >= " & txtCliente(2).Text
    If txtCliente(3).Text <> "" Then miSQL = miSQL & " AND sclien.codclien <= " & txtCliente(3).Text
    
    'El agente
   'Desde hasta
    If txtNumero(6).Text <> "" Then miSQL = miSQL & " AND sclien.codagent >= " & txtNumero(6).Text
    If txtNumero(7).Text <> "" Then miSQL = miSQL & " AND sclien.codagent <= " & txtNumero(7).Text
  

    If vParamAplic.ContabilidadNueva Then
        Codigo = "Select codmacta from cobros group by 1"
    Else
        Codigo = "Select codmacta from scobro group by 1"
    End If
    miRsAux.Open Codigo, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
    Codigo = ""
    While Not miRsAux.EOF
        Codigo = Codigo & ", '" & miRsAux!Codmacta & "'"
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If Codigo = "" Then
        MsgBox "Ningun dato en tesoreria!!!!", vbExclamation
        Exit Function
    End If
    
    Codigo = Mid(Codigo, 2)
    Codigo = "(" & Codigo & ")"
    
    miSQL = miSQL & " AND codmacta in " & Codigo
    
    
    miSQL = "insert into tmpcommand(codusu,codprove,nomprove,nomfamia,nomartic,codartic,importel,codfamia) " & miSQL
    conn.Execute miSQL


    'Ahora recorremos
    'Para buscar los pendientes...
    
    miSQL = "Select * from tmpcommand where codusu = " & vUsu.Codigo
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0

    Vec = 0
    While Not miRsAux.EOF
        Me.lblIndicador(0).Caption = miRsAux!nomprove
        Me.lblIndicador(0).Refresh
        If Vec = 20 Then
            Vec = 0
            DoEvents
        End If
        Vec = Vec + 1
        
        Set R = New ADODB.Recordset
        Codigo = CStr(Int(DBLet(miRsAux!ImporteL, "N")))
        
        
        'Para cada vencimiento indicaremos si ha sido reclamado (veces) y si esta en situacion juridica
        'El vto sera Serie00Fra DBLet(R!numSerie, "T") & Format(R!Codfaccl, "000000")
        
        'Empezamos
        
        If vParamAplic.ContabilidadNueva Then
            miSQL = "SELECT cobros.*,numfactu codfaccl, fecfactu fecfaccl FROM cobros "
        Else
            miSQL = "SELECT scobro.* FROM scobro "
        End If
        miSQL = miSQL & " WHERE recedocu=0 AND codmacta = '" & Codigo & "'"
        miSQL = miSQL & " ORDER BY fecvenci desc"
        R.Open miSQL, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
        miSQL = ""
        Aux2 = "|"  'para buscar las reclamaciones efectuadas de ese vto
        J = 0
        While Not R.EOF
            
            Importe = R!ImpVenci + DBLet(R!gastos, "N") - DBLet(R!impcobro, "N")
            If Importe <> 0 Then
                J = J + 1
                Codigo = DBLet(R!numSerie, "T") & Format(R!Codfaccl, "0000000")
                miSQL = miSQL & DBLet(R!numSerie, "T") & Format(R!Codfaccl, "0000000") & "·" & R!FecVenci & "·" & CStr(Importe) & "·" & R!situacionjuri & "·|"
                
                Aux2 = Aux2 & Codigo & "|"
                
                
            End If
            
            R.MoveNext
        Wend
        R.Close
        
        
        'ya tenemos el maximo de reclamaciones y/o cobro
        If J = 0 Then
            'No tiene ningun cobro pendiente
            Aux2 = "DELETE FROM tmpcommand WHERE codusu = " & vUsu.Codigo & " AND codprove= " & miRsAux!Codprove
            conn.Execute Aux2
        Else
            

            
            'Hay mas cobros que en situacion juridica (LO NORMAL)
            ReDim linea(J)
            
            'Las dos conjuntas
            Codigo = ""
            For K = 1 To J
                'miSQL cobros
                
                Codigo = RecuperaValor(miSQL, K) & "|"
                linea(K) = Codigo
            Next K
            
            
            
            'Volvemos a popner copdmacta
            Codigo = CStr(Int(DBLet(miRsAux!ImporteL, "N")))
            
            'Hco reclamacioones
            If vParamAplic.ContabilidadNueva Then
                miSQL = "select reclama.codigo,numserie,numfactu codfaccl,fecfactu fecfaccl,fecreclama,impvenci,codmacta,observaciones,importes "
                miSQL = miSQL & " from reclama  INNER join reclama_facturas  on reclama.codigo=reclama_facturas.codigo"
            Else
        
                miSQL = "SELECT numserie,codfaccl,count(*) from shcocob "
            End If
            miSQL = miSQL & " WHERE codmacta = '" & Codigo & "'  group by 1,2 order by 3"
            R.Open miSQL, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
            Codigo = ""
            While Not R.EOF
                    
                    miSQL = "|" & DBLet(R!numSerie, "T") & Format(R!Codfaccl, "0000000") & "|"
                    If InStr(1, Aux2, miSQL) > 0 Then
                        'Este vto esta
                        Codigo = Codigo & DBLet(R!numSerie, "T") & Format(DBLet(R!Codfaccl, "N"), "0000000") & "(" & R.Fields(2) & ")" & "|"
                        
                    End If
                    
            
                    
                
                R.MoveNext
            Wend
            R.Close
            
            For K = 1 To UBound(linea)
                miSQL = linea(K)
                If Codigo = "" Then
                    'Ninugna reclamacion
                    Aux2 = "0"
                Else
                    'Veamos la reclamacion
                    
                    Aux2 = Replace(miSQL, "·", "|")
                    Aux2 = RecuperaValor(Aux2, 1)
                    
                    J = InStr(1, Codigo, Aux2)
                    If J > 0 Then
                        'La ha encontrado
                        Aux2 = Mid(Codigo, J + Len(Aux2))
                        Aux2 = RecuperaValor(Aux2, 1)
                        Aux2 = Replace(Aux2, "(", "")
                        Aux2 = Replace(Aux2, ")", "")
                        miSQL = miSQL & Aux2
                    Else
                        Aux2 = "0"
                    End If
                    
                End If
                
                linea(K) = miSQL & "|"
            Next K
            'Acciones comerciales
            miSQL = "SELECT usuario,fechora,observaciones from scrmacciones "
            miSQL = miSQL & " WHERE codclien = " & miRsAux!Codprove & " ORDER BY fechora desc"
            R.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            J = 0
            While Not R.EOF
                J = J + 1
                
                If J > UBound(linea) Then
                   ' If J > 2 Then
                        'No vamos a insertar mas de dos acciones comerciales
                        MueveHastaElFinal R
                        J = 0 'Para que se salga sin insertar
                   ' Else
                   '     miSQL = "||||"
                   '     Lineas(J) = miSQL
                   ' End If
                Else
                    miSQL = linea(J)
                End If
                If J > 0 Then
                    Aux2 = Format(R!fechora, "dd/mm/yyyy") & " " & R!Usuario & " .-" & DBLet(R!Observaciones, "T")
                    Aux2 = Mid(Aux2, 1, 255)
                    miSQL = miSQL & Aux2 & "|"
                        
                    linea(J) = miSQL
                    
                    R.MoveNext
                End If
            Wend
            R.Close
            
            
            
            
            
            '                   cliente  secue  pdte de cobro     Reclamacin                 situjur histo
            'tmpinformes(codusu,codigo1,campo1,nombre1,importe1,fecha1,nombre2,importe2,fecha2,nombre3,obser)
            miSQL = ""
            NumRegElim = 0
            For J = 1 To UBound(linea)
                NumRegElim = NumRegElim + 1
                'Para cada linea
                miSQL = miSQL & ", (" & vUsu.Codigo & "," & miRsAux!Codprove & "," & NumRegElim
                'Cobro pdte
                Aux2 = RecuperaValor(CStr(linea(J)), 1)
                If Aux2 <> "" Then
                    Aux2 = Replace(Aux2, "·", "|")
                    Codigo = RecuperaValor(Aux2, 3)
                    Importe = CCur(Codigo)
                    Codigo = RecuperaValor(Aux2, 1) 'vto
                    miSQL = miSQL & ",'" & Codigo & "'," & TransformaComasPuntos(CStr(Importe))
                    Codigo = RecuperaValor(Aux2, 4) 'juridica
                    Aux2 = RecuperaValor(Aux2, 2) 'vto
                                        'fecha                 juridica
                    miSQL = miSQL & "," & DBSet(Aux2, "F") & "," & CStr(Val(Codigo))
                    
                Else
                    miSQL = miSQL & ",NULL,NULL,NULL,NULL"
                End If
                
                'Recalma
                Aux2 = RecuperaValor(CStr(linea(J)), 2)
                If Aux2 <> "" Then
                    
                    miSQL = miSQL & "," & DBSet(Aux2, "T")
                Else
                    miSQL = miSQL & ",NULL"
                End If
                
                
                'HCO
                Aux2 = RecuperaValor(CStr(linea(J)), 3)
                miSQL = miSQL & "," & DBSet(Aux2, "T") & ")"
            Next
            If miSQL <> "" Then
                miSQL = Mid(miSQL, 2)                                                       'juridica
                Aux2 = "INSERT INTO tmpinformes(codusu,codigo1,campo1,nombre1,importe1,fecha1,importeb1,importe2,obser) VALUES"
                Aux2 = Aux2 & miSQL
                conn.Execute Aux2
            End If
            Set R = Nothing
        
        End If 'de si tiene datos
        miRsAux.MoveNext
    Wend
    
    miRsAux.Close
    
    Me.lblIndicador(0).Caption = "Ver agentes"
    Me.lblIndicador(0).Refresh
    
    
    Codigo = "Select codprove,codfamia from tmpcommand WHERE codusu = " & vUsu.Codigo & " ORDER BY 1,2"
    miRsAux.Open Codigo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    J = -1
    miSQL = ""
    While Not miRsAux.EOF
        If J <> miRsAux!Codfamia Then
            
            miSQL = DevuelveDesdeBD(conAri, "nomagent", "sagent", "codagent", CStr(miRsAux!Codfamia))
            J = miRsAux!Codfamia
            Me.lblIndicador(0).Caption = miSQL
            Me.lblIndicador(0).Refresh
        End If
        Codigo = "UPDATE tmpinformes SET nombre2=" & DBSet(miSQL, "T")
        Codigo = Codigo & " WHERE codusu = " & vUsu.Codigo
        Codigo = Codigo & " AND codigo1 = " & miRsAux!Codprove
        conn.Execute Codigo
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    'Cuando tiene mas de un cobro, hare una sum
    Me.lblIndicador(0).Caption = "Totales"
    Me.lblIndicador(0).Refresh
    
    K = 0 'para saber si hay datos en tablas
    Codigo = "select codigo1,max(campo1) cuantos,sum(importe1) suma from tmpinformes where codusu =" & vUsu.Codigo & " group by 1"
    miRsAux.Open Codigo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        K = K + 1
        If miRsAux!Cuantos > 1 Then
            Codigo = "UPDATE tmpinformes set nombre3='Total pdte: " & Format(miRsAux!Suma, FormatoImporte) & "'"
            Codigo = Codigo & " WHERE codusu = " & vUsu.Codigo & " AND codigo1= " & miRsAux!Codigo1
            Codigo = Codigo & " ANd campo1 = " & miRsAux!Cuantos
            conn.Execute Codigo
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    GenerarResumenCRM = K > 0
    
    
    
    
eGenerarResumenCRM:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set R = Nothing
    
End Function


Private Sub MueveHastaElFinal(ByRef RS As ADODB.Recordset)
    While Not RS.EOF
        RS.MoveNext
    Wend
End Sub
