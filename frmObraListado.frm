VERSION 5.00
Begin VB.Form frmObraListado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado"
   ClientHeight    =   7575
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   10320
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameActuaciones 
      Height          =   4815
      Left            =   2760
      TabIndex        =   66
      Top             =   0
      Visible         =   0   'False
      Width           =   6015
      Begin VB.CommandButton cmdFraActuacion 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3240
         TabIndex        =   71
         Top             =   4200
         Width           =   1215
      End
      Begin VB.TextBox txtActua 
         Height          =   285
         Index           =   3
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   70
         Top             =   3480
         Width           =   2055
      End
      Begin VB.TextBox txtActua 
         Height          =   285
         Index           =   2
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   69
         Top             =   3000
         Width           =   2055
      End
      Begin VB.TextBox txtDpto 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   4
         Left            =   1200
         MaxLength       =   6
         TabIndex        =   68
         Top             =   2040
         Width           =   735
      End
      Begin VB.TextBox txtNomDpto 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   4
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   77
         Text            =   "Text5"
         Top             =   2040
         Width           =   3735
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   3
         Left            =   4560
         TabIndex        =   72
         Top             =   4200
         Width           =   1215
      End
      Begin VB.TextBox txtNomCli 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   6
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   73
         Text            =   "Text5"
         Top             =   1080
         Width           =   3735
      End
      Begin VB.TextBox txtCli 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   6
         Left            =   1200
         MaxLength       =   6
         TabIndex        =   67
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblInd 
         Caption         =   "Label1"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   83
         Top             =   4320
         Width           =   2055
      End
      Begin VB.Image imgActuacion 
         Height          =   240
         Index           =   3
         Left            =   840
         Picture         =   "frmObraListado.frx":0000
         Top             =   3480
         Width           =   240
      End
      Begin VB.Image imgActuacion 
         Height          =   240
         Index           =   2
         Left            =   840
         Picture         =   "frmObraListado.frx":0102
         Top             =   3000
         Width           =   240
      End
      Begin VB.Label Label14 
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
         Left            =   360
         TabIndex        =   82
         Top             =   3480
         Width           =   420
      End
      Begin VB.Label Label14 
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
         Index           =   19
         Left            =   360
         TabIndex        =   81
         Top             =   3000
         Width           =   450
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Actuación"
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
         TabIndex        =   80
         Top             =   2640
         Width           =   840
      End
      Begin VB.Label Label14 
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
         Index           =   17
         Left            =   360
         TabIndex        =   79
         Top             =   2040
         Width           =   450
      End
      Begin VB.Image imgDpto 
         Height          =   240
         Index           =   4
         Left            =   840
         Picture         =   "frmObraListado.frx":0204
         Top             =   2040
         Width           =   240
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Obra"
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
         TabIndex        =   78
         Top             =   1560
         Width           =   405
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   "Facturas x Actuacion"
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
         TabIndex        =   76
         Top             =   240
         Width           =   3195
      End
      Begin VB.Label Label10 
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
         Left            =   240
         TabIndex        =   75
         Top             =   720
         Width           =   585
      End
      Begin VB.Image imgCli 
         Height          =   240
         Index           =   6
         Left            =   840
         Picture         =   "frmObraListado.frx":0306
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label Label14 
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
         Index           =   16
         Left            =   360
         TabIndex        =   74
         Top             =   1080
         Width           =   450
      End
   End
   Begin VB.Frame framePartesTrabajo 
      Height          =   7095
      Left            =   240
      TabIndex        =   14
      Top             =   0
      Visible         =   0   'False
      Width           =   6015
      Begin VB.OptionButton optParte 
         Caption         =   "Cliente"
         Height          =   255
         Index           =   1
         Left            =   3000
         TabIndex        =   11
         Top             =   5640
         Width           =   975
      End
      Begin VB.OptionButton optParte 
         Caption         =   "Trabajador "
         Height          =   255
         Index           =   0
         Left            =   1200
         TabIndex        =   10
         Top             =   5640
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.CommandButton cmdPartesTra 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3120
         TabIndex        =   12
         Top             =   6480
         Width           =   1215
      End
      Begin VB.TextBox txtFecha 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   3720
         MaxLength       =   10
         TabIndex        =   9
         Top             =   5040
         Width           =   1215
      End
      Begin VB.TextBox txtNomTra 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   1
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   32
         Text            =   "Text5"
         Top             =   4320
         Width           =   3735
      End
      Begin VB.TextBox txtTra 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   1200
         MaxLength       =   6
         TabIndex        =   7
         Top             =   4320
         Width           =   735
      End
      Begin VB.TextBox txtNomTra 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   0
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   29
         Text            =   "Text5"
         Top             =   3960
         Width           =   3735
      End
      Begin VB.TextBox txtTra 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1200
         MaxLength       =   6
         TabIndex        =   6
         Top             =   3960
         Width           =   735
      End
      Begin VB.TextBox txtActua 
         Height          =   285
         Index           =   1
         Left            =   3840
         MaxLength       =   20
         TabIndex        =   5
         Top             =   3195
         Width           =   1935
      End
      Begin VB.TextBox txtActua 
         Height          =   285
         Index           =   0
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   4
         Top             =   3195
         Width           =   1815
      End
      Begin VB.TextBox txtNomDpto 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   1
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   24
         Text            =   "Text5"
         Top             =   2520
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.TextBox txtDpto 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   1200
         MaxLength       =   6
         TabIndex        =   3
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox txtNomDpto 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   0
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   21
         Text            =   "Text5"
         Top             =   2160
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.TextBox txtDpto 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1200
         MaxLength       =   6
         TabIndex        =   2
         Top             =   2160
         Width           =   735
      End
      Begin VB.TextBox txtNomCli 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   1
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   19
         Text            =   "Text5"
         Top             =   1440
         Width           =   3735
      End
      Begin VB.TextBox txtCli 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   1200
         MaxLength       =   6
         TabIndex        =   1
         Top             =   1440
         Width           =   735
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   0
         Left            =   4440
         TabIndex        =   13
         Top             =   6480
         Width           =   1215
      End
      Begin VB.TextBox txtCli 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1200
         MaxLength       =   6
         TabIndex        =   0
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox txtNomCli 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   0
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   17
         Text            =   "Text5"
         Top             =   1080
         Width           =   3735
      End
      Begin VB.TextBox txtFecha 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   8
         Top             =   5040
         Width           =   1215
      End
      Begin VB.Image imgActuacion 
         Height          =   240
         Index           =   1
         Left            =   3600
         Picture         =   "frmObraListado.frx":0408
         Top             =   3240
         Width           =   240
      End
      Begin VB.Image imgActuacion 
         Height          =   240
         Index           =   0
         Left            =   960
         Picture         =   "frmObraListado.frx":050A
         Top             =   3240
         Width           =   240
      End
      Begin VB.Label Label14 
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
         Index           =   9
         Left            =   3000
         TabIndex        =   36
         Top             =   5040
         Width           =   450
      End
      Begin VB.Label Label14 
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
         Index           =   8
         Left            =   360
         TabIndex        =   35
         Top             =   5040
         Width           =   450
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   1
         Left            =   3480
         Picture         =   "frmObraListado.frx":060C
         Top             =   5040
         Width           =   240
      End
      Begin VB.Label Label10 
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
         Index           =   4
         Left            =   240
         TabIndex        =   34
         Top             =   4800
         Width           =   495
      End
      Begin VB.Image imgT 
         Height          =   240
         Index           =   1
         Left            =   840
         Picture         =   "frmObraListado.frx":0697
         Top             =   4320
         Width           =   240
      End
      Begin VB.Label Label14 
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
         TabIndex        =   33
         Top             =   4320
         Width           =   420
      End
      Begin VB.Label Label10 
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
         Index           =   3
         Left            =   240
         TabIndex        =   31
         Top             =   3720
         Width           =   945
      End
      Begin VB.Image imgT 
         Height          =   240
         Index           =   0
         Left            =   840
         Picture         =   "frmObraListado.frx":0799
         Top             =   3960
         Width           =   240
      End
      Begin VB.Label Label14 
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
         TabIndex        =   30
         Top             =   3960
         Width           =   450
      End
      Begin VB.Label Label14 
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
         Index           =   4
         Left            =   3120
         TabIndex        =   28
         Top             =   3240
         Width           =   420
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Actuación"
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
         TabIndex        =   27
         Top             =   3000
         Width           =   840
      End
      Begin VB.Label Label14 
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
         Left            =   360
         TabIndex        =   26
         Top             =   3240
         Width           =   450
      End
      Begin VB.Image imgDpto 
         Height          =   240
         Index           =   1
         Left            =   840
         Picture         =   "frmObraListado.frx":089B
         Top             =   2520
         Width           =   240
      End
      Begin VB.Label Label14 
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
         Left            =   360
         TabIndex        =   25
         Top             =   2520
         Width           =   420
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Obra"
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
         TabIndex        =   23
         Top             =   1920
         Width           =   405
      End
      Begin VB.Image imgDpto 
         Height          =   240
         Index           =   0
         Left            =   840
         Picture         =   "frmObraListado.frx":099D
         Top             =   2160
         Width           =   240
      End
      Begin VB.Label Label14 
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
         Index           =   1
         Left            =   360
         TabIndex        =   22
         Top             =   2160
         Width           =   450
      End
      Begin VB.Image imgCli 
         Height          =   240
         Index           =   1
         Left            =   840
         Picture         =   "frmObraListado.frx":0A9F
         Top             =   1440
         Width           =   240
      End
      Begin VB.Label Label14 
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
         Left            =   360
         TabIndex        =   20
         Top             =   1440
         Width           =   420
      End
      Begin VB.Label Label14 
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
         Left            =   360
         TabIndex        =   18
         Top             =   1080
         Width           =   450
      End
      Begin VB.Image imgCli 
         Height          =   240
         Index           =   0
         Left            =   840
         Picture         =   "frmObraListado.frx":0BA1
         Top             =   1080
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   840
         Picture         =   "frmObraListado.frx":0CA3
         Top             =   5040
         Width           =   240
      End
      Begin VB.Label Label10 
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
         TabIndex        =   16
         Top             =   720
         Width           =   585
      End
      Begin VB.Label Label10 
         Caption         =   "Partes de trabajo"
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
         Left            =   1560
         TabIndex        =   15
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.Frame FrameObra 
      Height          =   2655
      Left            =   240
      TabIndex        =   55
      Top             =   0
      Visible         =   0   'False
      Width           =   6015
      Begin VB.CommandButton cmdObras 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3120
         TabIndex        =   58
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox txtCli 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   1200
         MaxLength       =   6
         TabIndex        =   57
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox txtNomCli 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   5
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   64
         Text            =   "Text5"
         Top             =   1440
         Width           =   3735
      End
      Begin VB.TextBox txtCli 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   4
         Left            =   1200
         MaxLength       =   6
         TabIndex        =   56
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox txtNomCli 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   4
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   61
         Text            =   "Text5"
         Top             =   1080
         Width           =   3735
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   2
         Left            =   4560
         TabIndex        =   59
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label14 
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
         Index           =   15
         Left            =   360
         TabIndex        =   65
         Top             =   1440
         Width           =   420
      End
      Begin VB.Image imgCli 
         Height          =   240
         Index           =   5
         Left            =   840
         Picture         =   "frmObraListado.frx":0D2E
         Top             =   1440
         Width           =   240
      End
      Begin VB.Label Label14 
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
         Index           =   14
         Left            =   360
         TabIndex        =   63
         Top             =   1080
         Width           =   450
      End
      Begin VB.Image imgCli 
         Height          =   240
         Index           =   4
         Left            =   840
         Picture         =   "frmObraListado.frx":0E30
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label Label10 
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
         Index           =   10
         Left            =   120
         TabIndex        =   62
         Top             =   720
         Width           =   585
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   "a"
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
         TabIndex        =   60
         Top             =   360
         Width           =   5655
      End
   End
   Begin VB.Frame FrameActua 
      Height          =   3735
      Left            =   240
      TabIndex        =   37
      Top             =   0
      Visible         =   0   'False
      Width           =   6135
      Begin VB.CommandButton cmdActua 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3120
         TabIndex        =   42
         Top             =   3120
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   1
         Left            =   4560
         TabIndex        =   43
         Top             =   3120
         Width           =   1215
      End
      Begin VB.TextBox txtDpto 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   1200
         MaxLength       =   6
         TabIndex        =   40
         Top             =   2160
         Width           =   735
      End
      Begin VB.TextBox txtNomDpto 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   3
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   51
         Text            =   "Text5"
         Top             =   2160
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.TextBox txtDpto 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   1200
         MaxLength       =   6
         TabIndex        =   41
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox txtNomDpto 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   2
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   50
         Text            =   "Text5"
         Top             =   2520
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.TextBox txtCli 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   1200
         MaxLength       =   6
         TabIndex        =   39
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox txtNomCli 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   3
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   48
         Text            =   "Text5"
         Top             =   1560
         Width           =   3735
      End
      Begin VB.TextBox txtNomCli 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   2
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   45
         Text            =   "Text5"
         Top             =   1200
         Width           =   3735
      End
      Begin VB.TextBox txtCli 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   1200
         MaxLength       =   6
         TabIndex        =   38
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label14 
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
         Left            =   360
         TabIndex        =   54
         Top             =   2160
         Width           =   450
      End
      Begin VB.Image imgDpto 
         Height          =   240
         Index           =   3
         Left            =   840
         Picture         =   "frmObraListado.frx":0F32
         Top             =   2160
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Obra"
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
         TabIndex        =   53
         Top             =   1920
         Width           =   405
      End
      Begin VB.Label Label14 
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
         Left            =   360
         TabIndex        =   52
         Top             =   2520
         Width           =   420
      End
      Begin VB.Image imgDpto 
         Height          =   240
         Index           =   2
         Left            =   840
         Picture         =   "frmObraListado.frx":1034
         Top             =   2520
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label14 
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
         Index           =   11
         Left            =   360
         TabIndex        =   49
         Top             =   1560
         Width           =   420
      End
      Begin VB.Image imgCli 
         Height          =   240
         Index           =   3
         Left            =   840
         Picture         =   "frmObraListado.frx":1136
         Top             =   1560
         Width           =   240
      End
      Begin VB.Label Label10 
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
         Index           =   7
         Left            =   120
         TabIndex        =   47
         Top             =   840
         Width           =   585
      End
      Begin VB.Image imgCli 
         Height          =   240
         Index           =   2
         Left            =   840
         Picture         =   "frmObraListado.frx":1238
         Top             =   1200
         Width           =   240
      End
      Begin VB.Label Label14 
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
         Index           =   10
         Left            =   360
         TabIndex        =   46
         Top             =   1200
         Width           =   450
      End
      Begin VB.Label Label10 
         Caption         =   "Actuaciones"
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
         Left            =   2040
         TabIndex        =   44
         Top             =   240
         Width           =   2655
      End
   End
End
Attribute VB_Name = "frmObraListado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Opcion As Byte
    '0 .-   Listado partes trabajo
    '1 .-   Actuaciones
    '2 .-   Obras/departamentos  x cliente
    '3.-    Facturas x Actuacion
    
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmCli As frmBasico2 'frmFacClientesGr
Attribute frmCli.VB_VarHelpID = -1
Private WithEvents frmT As frmBasico2 'frmAdmTrabajadores
Attribute frmT.VB_VarHelpID = -1
Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1


Dim PrimeraVez As Boolean

'Uso gral
Dim CadenaDesdeForms As String
Dim devuelve As String
Dim campo As String
Dim Aux As String  '
'---- Variables para el INFORME ----
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private cadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para el frmImprimir
Private cadNomRPT As String 'Nombre del informe a Imprimir
Private conSubRPT As Boolean 'Si el informe tiene subreports


Private Sub cmdActua_Click()
  InicializarVbles
    'Añadir el nombre de la Empresa como parametro
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    
    Aux = ""
    If txtCli(2).Text <> "" Or txtCli(3).Text <> "" Then
        campo = "{sactuaobra.codclien}"
        devuelve = "|"
        If Not PonerDesdeHasta(campo, "CLI", 2, 3, devuelve) Then Exit Sub
        Aux = "Cliente: " & Mid(devuelve, 2)
    End If
    
    
    If txtDpto(2).Text <> "" Or txtDpto(3).Text <> "" Then
        campo = "{sactuaobra.coddirec}"
        devuelve = "|"
        If Not PonerDesdeHasta(campo, "DPT", 2, 3, devuelve) Then Exit Sub
        Aux = Trim(Aux & "      " & "Dpto: " & Mid(devuelve, 2))
    End If
    cadParam = cadParam & "|pDesde=""" & Aux & """|"
    numParam = numParam + 1

    If cadSelect <> "" Then campo = campo & " AND " & cadSelect
    If Not HayRegParaInforme("sactuaobra", campo) Then Exit Sub
    
    
    cadNomRPT = "rfacactua.rpt"
    cadTitulo = "Listado actuaciones "
    LlamarImprimir
    
End Sub

Private Sub cmdCancelar_Click(Index As Integer)
    Unload Me
End Sub

Private Sub cmdFraActuacion_Click()
    If txtCli(6).Text = "" Then
        MsgBox "Indique el cliente", vbExclamation
        Exit Sub
    End If
    
    
    Screen.MousePointer = vbHourglass
    Set miRsAux = New ADODB.Recordset
    If GeneraFraActuacion Then
        cadParam = "|pEmpresa=""" & vEmpresa.nomempre & """|"
        cadParam = cadParam & "pDHCliente=""" & Format(Me.txtCli(6).Text, "0000") & " " & Me.txtNomCli(6).Text & """|"
        
        
        Aux = ""
        'Si ha puesto OBRA
        If Me.txtDpto(4).Text <> "" Then Aux = Aux & "Departamento: " & txtDpto(4).Text & " " & Me.txtNomDpto(4).Text
        'Si ha puesto desde hasta actuacion
        devuelve = ""
        If Me.txtActua(2).Text <> "" Then devuelve = devuelve & " desde " & txtActua(2).Text
        If Me.txtActua(3).Text <> "" Then devuelve = devuelve & " hasta " & txtActua(3).Text
        If devuelve <> "" Then
            devuelve = "    Actuación: " & Trim(devuelve)
            Aux = Trim(Aux & devuelve)
        End If
        CadenaDesdeOtroForm = "=""" & Aux & """|"
        cadParam = cadParam & "pObraDir=""" & Aux & """|"
        
        numParam = 3
        
    
        With frmImprimir
            .FormulaSeleccion = "{tmpinformes.codusu}=" & vUsu.Codigo
            
            'D/H
            .NumeroParametros = 1
            .OtrosParametros = cadParam
            .SoloImprimir = False
            .EnvioEMail = False
            .Opcion = 5
            .Titulo = "Facturas por obra-actuacion"
            .NombreRPT = "saiObrasFras.rpt"
            
            .MostrarTreeDesdeFuera = True
            .Show vbModal
         End With
    
    
    
    End If
    lblInd(0).Caption = ""
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdObras_Click()
     InicializarVbles
    'Añadir el nombre de la Empresa como parametro
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    Aux = ""
    If txtCli(4).Text <> "" Or txtCli(5).Text <> "" Then
        campo = "{sclien.codclien}"
        devuelve = "|"
        If Not PonerDesdeHasta(campo, "CLI", 4, 5, devuelve) Then Exit Sub
        Aux = "Cliente: " & Mid(devuelve, 2)
        
        
    End If
    cadParam = cadParam & "|pDesde=""" & Aux & """|"
    numParam = numParam + 1
    
    If vParamAplic.HayDeparNuevo = 2 Then
        cadTitulo = "Obras x Cliente"
    Else
        cadTitulo = "Departamento-direcciones"
    End If
    cadNomRPT = "rfacobras.rpt"
    
    LlamarImprimir
End Sub



Private Sub cmdPartesTra_Click()

    'En AUX meteremos los desde hasta
    'para meter en pdh1 y 2 del rpt
    Aux = ""
    InicializarVbles
    
    'Añadir el nombre de la Empresa como parametro
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    

     If txtCli(0).Text <> "" Or txtCli(1).Text <> "" Then
        campo = "{sliparte.codclien}"
        devuelve = "|"
        If Not PonerDesdeHasta(campo, "CLI", 0, 1, devuelve) Then Exit Sub
        Aux = "Cliente: " & Mid(devuelve, 2)
    End If
    
    
    If txtDpto(0).Text <> "" Or txtDpto(1).Text <> "" Then
        campo = "{sliparte.coddirec}"
        devuelve = "|"
        If Not PonerDesdeHasta(campo, "DPT", 0, 1, devuelve) Then Exit Sub
        Aux = Trim(Aux & "      " & "Dpto: " & Mid(devuelve, 2))
    End If
    
    If txtActua(0).Text <> "" Or txtActua(1).Text <> "" Then
        campo = "{sliparte.actuacion}"
        devuelve = "|"
        If Not PonerDesdeHasta(campo, "ACT", 0, 1, devuelve) Then Exit Sub
        Aux = Trim(Aux & "      " & "Actuacion: " & Mid(devuelve, 2))
    End If
    
    'Añadir el nombre de la Empresa como parametro
    cadParam = cadParam & "|pdh1=""" & Aux & """|"
    numParam = numParam + 1
    
    
    
    
    
    'Al paremetro pdh2
    Aux = ""
     If txtFecha(0).Text <> "" Or txtFecha(1).Text <> "" Then
        campo = "{scaparte.fecha}"
        devuelve = "|"
        'devuelve = "pDHFamilia=""Fecha: "
        If Not PonerDesdeHasta(campo, "F", 0, 1, devuelve) Then Exit Sub
        Aux = "Fecha parte: " & Mid(devuelve, 2)
    End If
    
     
    If txtTra(0).Text <> "" Or txtTra(1).Text <> "" Then
        campo = "{scaparte.codtraba}"
        devuelve = "|"
        'devuelve = "pDHFamilia=""Fecha: "
        If Not PonerDesdeHasta(campo, "TRA", 0, 1, devuelve) Then Exit Sub
        Aux = Trim(Aux & "      " & "Trabajador: " & Mid(devuelve, 2))
    End If
    'Añadir el nombre de la Empresa como parametro
    cadParam = cadParam & "|pdh2=""" & Aux & """|"
    numParam = numParam + 1
    
    
    
    

    
    
    
    campo = "scaparte.numparte=sliparte.numparte"
    If cadSelect <> "" Then campo = campo & " AND " & cadSelect
    If Not HayRegParaInforme("scaparte,sliparte", campo) Then Exit Sub
    
    If Me.optParte(0).Value Then
        cadNomRPT = "saiparteop.rpt"
        cadTitulo = "Trabajador"
    Else
        cadNomRPT = "saipartecl.rpt"
        cadTitulo = "Cliente"
    End If
    cadTitulo = "Partes trabajo (" & cadTitulo & ")"
    LlamarImprimir
End Sub



Private Sub Form_Activate()

    If PrimeraVez Then
        PrimeraVez = False
        Select Case Opcion
        Case 0
            PonerFoco txtCli(0)
        Case 1
            PonerFoco txtCli(2)
        End Select
    
    End If
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Form_Load()
Dim W As Integer
Dim H As Integer

        Screen.MousePointer = vbHourglass
        Me.Icon = frmPpal.Icon
        PrimeraVez = True
        framePartesTrabajo.visible = False
        FrameObra.visible = False
        FrameActua.visible = False
        FrameActuaciones.visible = False
        limpiar Me
        Select Case Opcion
        Case 0
            H = Me.framePartesTrabajo.Height
            W = Me.framePartesTrabajo.Width
            PonerFrameVisible framePartesTrabajo, True, H, W
        
        Case 1
            H = Me.FrameActua.Height
            W = Me.FrameActua.Width
            PonerFrameVisible FrameActua, True, H, W
        
        Case 2
            H = Me.FrameObra.Height
            W = Me.FrameObra.Width
            PonerFrameVisible FrameObra, True, H, W
            Label10(9).Caption = "Listado " & DevuelveTextoDepto(False)
            
        Case 3
            H = Me.FrameActuaciones.Height
            W = Me.FrameActuaciones.Width
            PonerFrameVisible FrameActuaciones, True, H, W
            lblInd(0).Caption = "" 'indicador
        End Select
        
        Me.cmdCancelar(Opcion).Cancel = True
        Me.Width = W + 70
        Me.Height = H + 350
        
End Sub



Private Sub frmB_Selecionado(CadenaDevuelta As String)
    Aux = CadenaDevuelta
End Sub



Private Sub frmC_Selec(vFecha As Date)
    CadenaDesdeForms = Format(vFecha, "dd/mm/yyyy")
End Sub



Private Sub frmCli_DatoSeleccionado(CadenaSeleccion As String)
    CadenaDesdeForms = CadenaSeleccion
End Sub

Private Sub frmT_DatoSeleccionado(CadenaSeleccion As String)
     CadenaDesdeForms = CadenaSeleccion
End Sub

Private Sub imgActuacion_Click(Index As Integer)
    Aux = ""
    NumRegElim = -1
    If Index >= 2 Then
        'Fra por client obra
        NumRegElim = 6
        If Me.txtCli(6).Text = "" Then Aux = "Indique el cliente"
            
            
        'La obra
        If Me.txtDpto(4).Text = "" Then Aux = Aux & vbCrLf & "Indique la obra"
            
            
    Else
        'D/H obra actuacion
        NumRegElim = 0
        If txtCli(0).Text = "" Or txtCli(1).Text = "" Then
            Aux = "Indique el cliente"
        Else
            If txtCli(0).Text <> txtCli(1).Text Then Aux = "El cliente debe ser el mismo"
        End If
        
        
        If txtDpto(0).Text = "" Or txtDpto(1).Text = "" Then
            Aux = Aux & vbCrLf & "Indique la obra"
        Else
            If txtDpto(0).Text <> txtDpto(1).Text Then Aux = "La obra debe ser la misma"
        End If
    End If
    

    
    If Aux <> "" Then
        MsgBox Aux, vbExclamation
        PonerFoco txtCli(NumRegElim)
        Exit Sub
    End If
    
    
    
     'Llamamos a al form
    Aux = ""
    'Estamos en Modo de Cabeceras
    'Registro de la tabla de cabeceras: slista
    Aux = Aux & "Cliente|sactuaobra|codclien|N|0000|9·"
    'Aux = Aux & "Nombre|sclien|nomclien|T||39·"
    Aux = Aux & "Obra|sactuaobra|coddirec|T|000|10·"
    Aux = Aux & "Desc. obra|sdirec|nomdirec|T||24·"
    Aux = Aux & "Actuacion|sactuaobra|actuacion|T||18·"
    Aux = Aux & "Actuacion|sactuaobra|observa|T||38·"
    
               
 
    Screen.MousePointer = vbHourglass
    Set frmB = New frmBuscaGrid
    frmB.vCampos = Aux
    frmB.vTabla = "sactuaobra,sclien,sdirec"
    
    Aux = "sactuaobra.codclien=sclien.codclien and sdirec.codclien=sactuaobra.codclien and sdirec.coddirec=sactuaobra.coddirec"
    If Index >= 2 Then
        frmB.vTitulo = "Actuaciones: " & txtCli(6).Text & " " & Me.txtNomCli(6).Text
        Aux = Aux & " AND sactuaobra.codclien = " & Me.txtCli(6).Text
        Aux = Aux & " AND sactuaobra.coddirec = " & Me.txtDpto(4).Text
    Else
        frmB.vTitulo = "Actuaciones: " & txtCli(0).Text & " " & Me.txtNomCli(0).Text
        Aux = Aux & " AND sactuaobra.codclien = " & Me.txtCli(0).Text
        Aux = Aux & " AND sactuaobra.coddirec = " & Me.txtDpto(0).Text
    End If
    frmB.vSQL = Aux

    frmB.vDevuelve = "0|1|3|"
    
    frmB.vselElem = 0
    frmB.vConexionGrid = conAri 'Conexion a BD Ariges
    Aux = ""
    frmB.Show vbModal
    Set frmB = Nothing
    If Aux <> "" Then
        txtActua(Index).Text = RecuperaValor(Aux, 3)
       
        Aux = ""
    End If
    
    
End Sub

Private Sub imgCli_Click(Index As Integer)
    CadenaDesdeForms = ""
'    Set frmCli = New frmFacClientesGr
'    frmCli.DatosADevolverBusqueda = "0|1|"
'    frmCli.Show vbModal
    Set frmCli = New frmBasico2
    AyudaClientes frmCli, txtCli(Index).Text
    Set frmCli = Nothing
    If CadenaDesdeForms <> "" Then
        txtCli(Index).Text = RecuperaValor(CadenaDesdeForms, 1)
        txtNomCli(Index).Text = RecuperaValor(CadenaDesdeForms, 2)
    End If
End Sub

Private Sub imgDpto_Click(Index As Integer)

    If Index = 4 Or Index <= 1 Then
        Aux = ""
        If Index = 4 Then
            If Me.txtCli(6).Text = "" Then Aux = "Ponga el cliente"
        Else
            If txtCli(0).Text <> txtCli(1).Text Then
                Aux = "Ponga el mismo cliente"
            Else
                If txtCli(0).Text = "" Then Aux = "Ponga el cliente"
            End If
            
            If Aux <> "" Then
                MsgBox Aux, vbExclamation
                Exit Sub
            End If
        End If
        
        Set frmB = New frmBuscaGrid
        Aux = "Obra|sdirec|coddirec|N|000|15·"
        Aux = Aux & "Desc. obra|sdirec|nomdirec|T||55·"
        frmB.vCampos = Aux
        frmB.vTabla = "sdirec"
        
        frmB.vDevuelve = "0|1|"
        If Index = 4 Then
            Aux = "codclien = " & Me.txtCli(6).Text
            frmB.vTitulo = "Obras: " & txtCli(6).Text & " - " & Me.txtNomCli(6).Text
        Else
            Aux = "codclien = " & Me.txtCli(0).Text
            frmB.vTitulo = "Obras: " & txtCli(0).Text & " - " & Me.txtNomCli(0).Text
        End If
        frmB.vSQL = Aux
        frmB.vselElem = 1
        frmB.vConexionGrid = conAri  'Conexión a BD: Ariges
        frmB.Label1.FontSize = 11
        Aux = ""
        frmB.Show vbModal
        Set frmB = Nothing
            

        If Aux <> "" Then
            Me.txtDpto(Index).Text = RecuperaValor(Aux, 1)
            Me.txtNomDpto(Index).Text = RecuperaValor(Aux, 2)
        End If
    End If
End Sub

Private Sub imgFecha_Click(Index As Integer)

    CadenaDesdeForms = ""
    Set frmC = New frmCal
    frmC.Fecha = Now
    If txtFecha(Index).Text <> "" Then
         If IsDate(txtFecha(Index).Text) Then frmC.Fecha = CDate(txtFecha(Index).Text)
     End If
    frmC.Show vbModal
    Set frmC = Nothing
    If CadenaDesdeForms <> "" Then txtFecha(Index).Text = CadenaDesdeForms

End Sub



Private Sub imgT_Click(Index As Integer)
    CadenaDesdeForms = ""
'        Set frmT = New frmAdmTrabajadores
'        frmT.DatosADevolverBusqueda = "0|1|"
'        frmT.Show vbModal
        Set frmT = New frmBasico2
        AyudaTrabajadores frmT, txtTra(Index)
        Set frmT = Nothing
    If CadenaDesdeForms <> "" Then
        txtTra(Index).Text = RecuperaValor(CadenaDesdeForms, 1)
        txtNomTra(Index).Text = RecuperaValor(CadenaDesdeForms, 2)
    End If
End Sub



Private Sub optParte_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 3, False
End Sub

Private Sub txtActua_KeyPress(Index As Integer, KeyAscii As Integer)
     KEYpressGnral KeyAscii, 3, False
End Sub

Private Sub txtCli_GotFocus(Index As Integer)
    ConseguirFoco txtCli(Index), 3
End Sub

Private Sub txtCli_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 3, False
End Sub

Private Sub txtCli_LostFocus(Index As Integer)
    txtCli(Index).Text = Trim(txtCli(Index).Text)
    CadenaDesdeForms = ""
    If txtCli(Index).Text <> "" Then
        If PonerFormatoEntero(txtCli(Index)) Then
            CadenaDesdeForms = DevuelveDesdeBD(conAri, "nomclien", "sclien", "codclien", txtCli(Index).Text, "N")
            If CadenaDesdeForms = "" Then
                '** NO EXISTE
                If Index = 6 Then
                    'ES OBLIGADO QUE EXISTA
                    MsgBox "No existe cliente", vbExclamation
                    txtCli(Index).Text = ""
                    PonerFoco txtCli(Index)
                End If
            End If
        End If
    End If
    txtNomCli(Index).Text = CadenaDesdeForms
End Sub

Private Sub txtDpto_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 3, False
End Sub

Private Sub txtDpto_LostFocus(Index As Integer)
    txtDpto(Index).Text = Trim(txtDpto(Index).Text)
    CadenaDesdeForms = ""
    If txtDpto(Index).Text <> "" Then
        If PonerFormatoEntero(txtDpto(Index)) Then
            
            If Index = 4 Then
                If Me.txtCli(6).Text = "" Then
                    MsgBox "Indique el cliente", vbExclamation
                    Me.txtDpto(Index).Text = ""
                    PonerFoco txtCli(6)
                Else
                    CadenaDesdeForms = "codclien = " & Me.txtCli(6).Text & " AND coddirec"
                    CadenaDesdeForms = DevuelveDesdeBD(conAri, "nomdirec", "sdirec", CadenaDesdeForms, txtDpto(Index).Text, "N")
                    If CadenaDesdeForms = "" Then
                        MsgBox "No existe obra", vbExclamation
                        Me.txtDpto(Index).Text = ""
                        PonerFoco txtDpto(Index)
                    End If
                End If
            End If
        Else
            txtDpto(Index).Text = ""
        End If
    End If
    txtNomDpto(Index).Text = CadenaDesdeForms
End Sub

Private Sub txtFecha_GotFocus(Index As Integer)
    ConseguirFoco txtFecha(Index), 3
End Sub

Private Sub txtFecha_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 3, False
End Sub

Private Sub txtFecha_LostFocus(Index As Integer)
    txtFecha(Index).Text = Trim(txtFecha(Index).Text)
    PonerFormatoFecha txtFecha(Index)
End Sub



Private Sub txtTra_GotFocus(Index As Integer)
    ConseguirFoco txtTra(Index), 3
End Sub

Private Sub txtTra_KeyPress(Index As Integer, KeyAscii As Integer)
     KEYpressGnral KeyAscii, 3, False
End Sub

Private Sub txtTra_LostFocus(Index As Integer)
    txtTra(Index).Text = Trim(txtTra(Index).Text)
    CadenaDesdeForms = ""
    If txtTra(Index).Text <> "" Then
        If PonerFormatoEntero(txtTra(Index)) Then
            
            CadenaDesdeForms = DevuelveDesdeBD(conAri, "nomtraba", "straba", "codtraba", txtTra(Index).Text, "N")
            If CadenaDesdeForms = "" Then
                '** NO EXISTE
                
            End If
        End If
    End If
    txtNomTra(Index).Text = CadenaDesdeForms
End Sub


Private Sub InicializarVbles()
    cadFormula = ""
    cadSelect = ""
    cadParam = ""
    numParam = 0
    cadTitulo = ""
    cadNomRPT = ""
    conSubRPT = False
End Sub


Private Sub LlamarImprimir()
    With frmImprimir
        .FormulaSeleccion = cadFormula
        .OtrosParametros = cadParam
        .NumeroParametros = numParam
        .Titulo = cadTitulo
        .SoloImprimir = False
        .EnvioEMail = False
        .Opcion = 2000 '2000 generico
        .NombrePDF = ""
        ' PongoNombrePDF Then .NombrePDF = cadPDFrpt
        .NombreRPT = cadNomRPT
        .ConSubInforme = conSubRPT
        .Show vbModal
    End With
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
        Set TDes = txtCli(indD)
        Set THas = txtCli(indH)
        Set DesD = txtNomCli(indD)
        Set DesH = txtNomCli(indH)
        Subtipo = "N"
    Case "DPT"
        'DEpartamento  /OBRA
        Set TDes = txtDpto(indD)
        Set THas = txtDpto(indH)
        Set DesD = txtNomDpto(indD)
        Set DesH = txtNomDpto(indH)
        Subtipo = "N"
        
    Case "ACT"
        'Actuacion
        Set TDes = Me.txtActua(indD)
        Set THas = txtActua(indH)
        Subtipo = "T"
        
    Case "TRA"
        'TRABAJADOR
         
        Set TDes = Me.txtTra(indD)
        Set THas = txtTra(indH)
        Subtipo = "N"

        Set DesD = txtNomTra(indD)
        Set DesH = txtNomTra(indH)

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

Private Sub Indicador(Indice As Integer, ByRef texto)
    lblInd(Indice).Caption = texto
    lblInd(Indice).Refresh
End Sub

Private Function GeneraFraActuacion() As Boolean

On Error GoTo eGeneraFraActuacion
    
    GeneraFraActuacion = False
    Indicador 0, "Preparando datos"
    
    conn.Execute "DELETE FROM tmpinformes WHERE codusu =" & vUsu.Codigo
    
    
    'INSERT PARA TODOS
    devuelve = "insert into `tmpinformes` (`codusu`,`codigo1`,`campo1`,`campo2`,"
    devuelve = devuelve & "`nombre1`,`nombre2`,`importe1`,importe2,`fecha1`,`obser`,nombre3)"

    
    'Empezamos metiendo en tmp los albaranes venta
    Indicador 0, "Albaran venta"
    Aux = "select " & vUsu.Codigo & ",codclien,coddirec,0,actuacion,"
    Aux = Aux & "concat(slialb.codtipom,right(concat(""000000"",slialb.numalbar),6)),"
    Aux = Aux & " sum(importel),0,fechaalb,null,null from "
    Aux = Aux & "slialb,scaalb where slialb.numalbar=scaalb.numalbar and  slialb.codtipom=slialb.codtipom "
    
    PonDesdehastaObraActuacion ""
    Aux = Aux & " group by 2,3,5,6" 'codclien,codobra,actua, fra
    Aux = devuelve & Aux
    conn.Execute Aux
    
    
    
    'Factura venta
    Indicador 0, "Factura venta"
    Aux = "select " & vUsu.Codigo & ",codclien,coddirec,1,actuacion,"
    Aux = Aux & "concat(scafac.codtipom,right(concat(""000000"",scafac.numfactu),6)),"
    Aux = Aux & "brutofac,0,scafac.fecfactu,null,null from scafac,scafac1"
    Aux = Aux & " Where scafac.NumFactu = scafac1.NumFactu And scafac.codtipom = scafac1.codtipom"
    Aux = Aux & " and scafac.fecfactu=scafac1.fecfactu"
    PonDesdehastaObraActuacion ""
    Aux = Aux & " group by 2,3,5,6" 'codclien,codobra,actua, fra
    Aux = devuelve & Aux
    conn.Execute Aux
    

    'Albaran compra
    Indicador 0, "albaran compra"
    Aux = "select " & vUsu.Codigo & ",codclien,coddirec,2,actuacion,"
    Aux = Aux & " concat(numalbar,""("",codprove,"")""),0,sum(importel),fechaalb,null,codprove"
    Aux = Aux & " From slialp WHERE 1=1"
    PonDesdehastaObraActuacion ""
    Aux = Aux & " group by 2,3,5,6" 'codclien,codobra,actua, fra
    Aux = devuelve & Aux
    conn.Execute Aux
    
    
    'Factura compra
    Indicador 0, "Factura compra"
    Aux = "select " & vUsu.Codigo & ",codclien,coddirec,3,actuacion,"
    Aux = Aux & " concat(numfactu,""("",codprove,"")""),0,sum(importel),fecfactu,null,codprove"
    Aux = Aux & " From slifpc WHERE 1=1"
    PonDesdehastaObraActuacion ""
    Aux = Aux & " group by 2,3,5,6" 'codclien,codobra,actua, fra
    Aux = devuelve & Aux
    conn.Execute Aux
    
    
    
    'Pedido compra
    Indicador 0, "Ped compra"
      
    Aux = "select " & vUsu.Codigo & ",slippr.codclien,coddirec,4,actuacion,"
    Aux = Aux & " concat(slippr.numpedpr,""("",codprove,"")""),0,sum(importel),fecpedpr,null,codprove"
    Aux = Aux & " From slippr,scappr WHERE slippr.numpedpr=scappr.numpedpr"
    PonDesdehastaObraActuacion "slippr."
    Aux = Aux & " group by 2,3,5,6" 'codclien,codobra,actua, fra
    Aux = devuelve & Aux
    conn.Execute Aux
    
    'ACtualizamos proveedores
    Indicador 0, "Ajustar proveedor"
    Aux = "Select nombre3 from tmpinformes where codusu = " & vUsu.Codigo & " GROUP BY 1"
    miRsAux.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    devuelve = ""
    While Not miRsAux.EOF
        devuelve = "1"
        If Not IsNull(miRsAux!nombre3) Then
            Aux = DevuelveDesdeBD(conAri, "nomprove", "sprove", "codprove", miRsAux!nombre3)
            If Aux <> "" Then

                Aux = "UPDATE tmpinformes SET nombre3=" & DBSet(Aux, "T") & " WHERE tmpinformes.codusu = " & vUsu.Codigo
                Aux = Aux & " AND nombre3 = '" & miRsAux!nombre3 & "'"
                conn.Execute Aux
            End If
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    

    
    If devuelve <> "" Then
        GeneraFraActuacion = True
    Else
        MsgBox "No existen datos para estos valores", vbExclamation
    End If
    
    Exit Function
eGeneraFraActuacion:
    MuestraError Err.Number, Err.Description

End Function


Private Sub PonDesdehastaObraActuacion(NombreTablas As String)
    Aux = Aux & " and  " & NombreTablas & "codclien= " & txtCli(6).Text
'Si ha puesto OBRA
    If Me.txtDpto(4).Text <> "" Then Aux = Aux & " and  coddirec= " & txtDpto(4).Text
    'Si ha puesto desde hasta actuacion
    If Me.txtActua(2).Text <> "" Then Aux = Aux & " and  actuacion >= " & DBSet(txtActua(2).Text, "T")
    If Me.txtActua(3).Text <> "" Then Aux = Aux & " and  actuacion <= " & DBSet(txtActua(3).Text, "T")
End Sub
