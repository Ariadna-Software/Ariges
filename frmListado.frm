VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmListado 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   9075
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   10890
   Icon            =   "frmListado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9075
   ScaleWidth      =   10890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameDtosFM 
      Height          =   5775
      Left            =   480
      TabIndex        =   336
      Top             =   600
      Width           =   6915
      Begin VB.CheckBox chkVarios 
         Caption         =   "CABEL"
         Height          =   195
         Index           =   4
         Left            =   3840
         TabIndex        =   763
         Top             =   840
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CheckBox chkVarios 
         Caption         =   "Solo rotaci�n"
         Height          =   195
         Index           =   2
         Left            =   1920
         TabIndex        =   732
         Top             =   840
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.OptionButton optFrDto 
         Caption         =   "CLIENTE/ACT"
         Height          =   255
         Index           =   5
         Left            =   5280
         TabIndex        =   722
         Top             =   4920
         Width           =   1575
      End
      Begin VB.CheckBox chkVarios 
         Caption         =   "Mostrar precio neto"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   720
         Top             =   840
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.ComboBox cboVarios 
         Height          =   315
         Index           =   0
         ItemData        =   "frmListado.frx":000C
         Left            =   1320
         List            =   "frmListado.frx":0019
         Style           =   2  'Dropdown List
         TabIndex        =   327
         Top             =   5280
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.OptionButton optFrDto 
         Caption         =   "Proveedor"
         Height          =   255
         Index           =   4
         Left            =   4080
         TabIndex        =   326
         Top             =   4920
         Width           =   1215
      End
      Begin VB.OptionButton optFrDto 
         Caption         =   "Marca"
         Height          =   255
         Index           =   3
         Left            =   3120
         TabIndex        =   325
         Top             =   4920
         Width           =   975
      End
      Begin VB.OptionButton optFrDto 
         Caption         =   "Familia"
         Height          =   255
         Index           =   2
         Left            =   2160
         TabIndex        =   324
         Top             =   4920
         Width           =   855
      End
      Begin VB.OptionButton optFrDto 
         Caption         =   "Actividad"
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   323
         Top             =   4920
         Width           =   1335
      End
      Begin VB.OptionButton optFrDto 
         Caption         =   "Cliente"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   322
         Top             =   4920
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.Frame Frame6 
         BorderStyle     =   0  'None
         Height          =   1095
         Left            =   360
         TabIndex        =   373
         Top             =   840
         Width           =   6135
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   74
            Left            =   2340
            Locked          =   -1  'True
            TabIndex        =   375
            Text            =   "Text5"
            Top             =   720
            Width           =   3495
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   74
            Left            =   1560
            MaxLength       =   6
            TabIndex        =   315
            Top             =   720
            Width           =   735
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   73
            Left            =   2340
            Locked          =   -1  'True
            TabIndex        =   374
            Text            =   "Text5"
            Top             =   360
            Width           =   3495
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   73
            Left            =   1560
            MaxLength       =   6
            TabIndex        =   314
            Top             =   360
            Width           =   735
         End
         Begin VB.Image imgBuscarG 
            Height          =   240
            Index           =   1
            Left            =   1275
            Picture         =   "frmListado.frx":004B
            ToolTipText     =   "Buscar cliente"
            Top             =   720
            Width           =   240
         End
         Begin VB.Label Label3 
            Caption         =   "Desde"
            Height          =   195
            Index           =   61
            Left            =   720
            TabIndex        =   378
            Top             =   360
            Width           =   465
         End
         Begin VB.Image imgBuscarG 
            Height          =   240
            Index           =   0
            Left            =   1275
            Picture         =   "frmListado.frx":014D
            ToolTipText     =   "Buscar cliente"
            Top             =   360
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
            Index           =   44
            Left            =   240
            TabIndex        =   377
            Top             =   120
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
            Index           =   45
            Left            =   720
            TabIndex        =   376
            Top             =   720
            Width           =   420
         End
      End
      Begin VB.Frame Frame5 
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   360
         TabIndex        =   367
         Top             =   2880
         Width           =   6135
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   77
            Left            =   1560
            MaxLength       =   4
            TabIndex        =   318
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   78
            Left            =   1560
            MaxLength       =   4
            TabIndex        =   319
            Top             =   720
            Width           =   735
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   77
            Left            =   2340
            Locked          =   -1  'True
            TabIndex        =   369
            Text            =   "Text5"
            Top             =   360
            Width           =   3135
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   78
            Left            =   2340
            Locked          =   -1  'True
            TabIndex        =   368
            Text            =   "Text5"
            Top             =   720
            Width           =   3135
         End
         Begin VB.Label Label3 
            Caption         =   "Desde"
            Height          =   195
            Index           =   66
            Left            =   720
            TabIndex        =   372
            Top             =   360
            Width           =   465
         End
         Begin VB.Label Label3 
            Caption         =   "Hasta"
            Height          =   195
            Index           =   67
            Left            =   720
            TabIndex        =   371
            Top             =   720
            Width           =   420
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
            Index           =   42
            Left            =   240
            TabIndex        =   370
            Top             =   120
            Width           =   525
         End
         Begin VB.Image imgBuscarG 
            Height          =   240
            Index           =   4
            Left            =   1275
            Picture         =   "frmListado.frx":024F
            ToolTipText     =   "Buscar marca"
            Top             =   360
            Width           =   240
         End
         Begin VB.Image imgBuscarG 
            Height          =   240
            Index           =   5
            Left            =   1275
            Picture         =   "frmListado.frx":0351
            ToolTipText     =   "Buscar marca"
            Top             =   720
            Width           =   240
         End
      End
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         Height          =   1095
         Left            =   360
         TabIndex        =   361
         Top             =   3720
         Width           =   6255
         Begin VB.TextBox txtCodigo 
            Height          =   285
            Index           =   79
            Left            =   1560
            TabIndex        =   320
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox txtCodigo 
            Height          =   285
            Index           =   80
            Left            =   1560
            TabIndex        =   321
            Top             =   720
            Width           =   735
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   79
            Left            =   2355
            Locked          =   -1  'True
            TabIndex        =   363
            Text            =   "Text5"
            Top             =   360
            Width           =   3615
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   80
            Left            =   2355
            Locked          =   -1  'True
            TabIndex        =   362
            Text            =   "Text5"
            Top             =   720
            Width           =   3615
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
            Index           =   46
            Left            =   240
            TabIndex        =   364
            Top             =   120
            Width           =   885
         End
         Begin VB.Label Label3 
            Caption         =   "Desde"
            Height          =   195
            Index           =   65
            Left            =   720
            TabIndex        =   366
            Top             =   360
            Width           =   465
         End
         Begin VB.Label Label3 
            Caption         =   "Hasta"
            Height          =   195
            Index           =   64
            Left            =   720
            TabIndex        =   365
            Top             =   720
            Width           =   420
         End
         Begin VB.Image imgBuscarG 
            Height          =   240
            Index           =   63
            Left            =   1275
            Picture         =   "frmListado.frx":0453
            ToolTipText     =   "Buscar proveedor"
            Top             =   360
            Width           =   240
         End
         Begin VB.Image imgBuscarG 
            Height          =   240
            Index           =   64
            Left            =   1275
            Picture         =   "frmListado.frx":0555
            ToolTipText     =   "Buscar proveedor"
            Top             =   720
            Width           =   240
         End
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   12
         Left            =   5760
         TabIndex        =   329
         Top             =   5280
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptarDtosFM 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4680
         TabIndex        =   328
         Top             =   5280
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   75
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   316
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   76
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   317
         Top             =   2640
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   75
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   338
         Text            =   "Text5"
         Top             =   2280
         Width           =   2895
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   76
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   337
         Text            =   "Text5"
         Top             =   2640
         Width           =   2895
      End
      Begin VB.Image imgayuda 
         Height          =   240
         Index           =   2
         Left            =   4200
         ToolTipText     =   "Buscar cliente"
         Top             =   5280
         Width           =   240
      End
      Begin VB.Label Label4 
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
         Index           =   103
         Left            =   3360
         TabIndex        =   721
         Top             =   840
         Visible         =   0   'False
         Width           =   1845
      End
      Begin VB.Label Label3 
         Caption         =   "Dto. especial"
         Height          =   195
         Index           =   110
         Left            =   120
         TabIndex        =   677
         Top             =   5310
         Width           =   1050
      End
      Begin VB.Label Label10 
         Caption         =   "Listado Descuentos Familia/Marca"
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
         TabIndex        =   342
         Top             =   360
         Width           =   5655
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   63
         Left            =   1080
         TabIndex        =   341
         Top             =   2280
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   62
         Left            =   1080
         TabIndex        =   340
         Top             =   2640
         Width           =   420
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
         Left            =   600
         TabIndex        =   339
         Top             =   2040
         Width           =   600
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   2
         Left            =   1635
         Picture         =   "frmListado.frx":0657
         Top             =   2280
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   3
         Left            =   1635
         Picture         =   "frmListado.frx":0759
         ToolTipText     =   "Buscar familia"
         Top             =   2640
         Width           =   240
      End
   End
   Begin VB.Frame FrameTarifas 
      Height          =   7335
      Left            =   1800
      TabIndex        =   102
      Top             =   240
      Width           =   7635
      Begin VB.CheckBox chkVarios 
         Caption         =   "CABEL"
         Height          =   255
         Index           =   3
         Left            =   5880
         TabIndex        =   762
         Top             =   2520
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CheckBox chkSoloRotacion 
         Caption         =   "S�lo rotaci�n"
         Height          =   255
         Left            =   840
         TabIndex        =   676
         Top             =   6600
         Width           =   2055
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   135
         Left            =   1920
         TabIndex        =   112
         Top             =   5640
         Width           =   1455
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   135
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   670
         Text            =   "Text5"
         Top             =   5640
         Width           =   3975
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   134
         Left            =   1920
         TabIndex        =   111
         Top             =   5280
         Width           =   1455
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   134
         Left            =   3420
         Locked          =   -1  'True
         TabIndex        =   667
         Text            =   "Text5"
         Top             =   5280
         Width           =   3975
      End
      Begin VB.ComboBox cboDecimales 
         Height          =   315
         ItemData        =   "frmListado.frx":085B
         Left            =   3240
         List            =   "frmListado.frx":086E
         Style           =   2  'Dropdown List
         TabIndex        =   601
         Top             =   6480
         Width           =   1335
      End
      Begin VB.CheckBox chkMostrarErrores 
         Caption         =   "Mostrar solo tarifas con error"
         Height          =   255
         Left            =   840
         TabIndex        =   421
         Top             =   6600
         Width           =   2415
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   24
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   156
         Text            =   "Text5"
         Top             =   1560
         Width           =   3495
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   24
         Left            =   1920
         TabIndex        =   104
         Top             =   1560
         Width           =   735
      End
      Begin VB.CheckBox chkSaltaPagTarif 
         Caption         =   "Salta p�g. en Familia"
         Height          =   255
         Left            =   840
         TabIndex        =   122
         Top             =   6240
         Width           =   2055
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   26
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   121
         Text            =   "Text5"
         Top             =   2520
         Width           =   2895
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   25
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   120
         Text            =   "Text5"
         Top             =   2160
         Width           =   2895
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   30
         Left            =   3420
         Locked          =   -1  'True
         TabIndex        =   119
         Text            =   "Text5"
         Top             =   4680
         Width           =   3975
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   29
         Left            =   3420
         Locked          =   -1  'True
         TabIndex        =   118
         Text            =   "Text5"
         Top             =   4320
         Width           =   3975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   26
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   106
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   25
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   105
         Top             =   2160
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   30
         Left            =   1920
         TabIndex        =   110
         Top             =   4680
         Width           =   1455
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   29
         Left            =   1920
         TabIndex        =   109
         Top             =   4320
         Width           =   1455
      End
      Begin VB.CommandButton cmdAceptarTarif 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   5400
         TabIndex        =   113
         Top             =   6600
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   5
         Left            =   6480
         TabIndex        =   114
         Top             =   6600
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   27
         Left            =   1920
         TabIndex        =   107
         Top             =   3240
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   28
         Left            =   1920
         TabIndex        =   108
         Top             =   3600
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   27
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   117
         Text            =   "Text5"
         Top             =   3240
         Width           =   2895
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   28
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   116
         Text            =   "Text5"
         Top             =   3600
         Width           =   2895
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   23
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   115
         Text            =   "Text5"
         Top             =   1200
         Width           =   3495
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   23
         Left            =   1920
         TabIndex        =   103
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   106
         Left            =   1080
         TabIndex        =   671
         Top             =   5640
         Width           =   465
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   104
         Left            =   1635
         Picture         =   "frmListado.frx":088F
         ToolTipText     =   "Buscar art�culo"
         Top             =   5640
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   105
         Left            =   1080
         TabIndex        =   669
         Top             =   5280
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
         Index           =   97
         Left            =   600
         TabIndex        =   668
         Top             =   5040
         Width           =   885
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   103
         Left            =   1635
         Picture         =   "frmListado.frx":0991
         ToolTipText     =   "Buscar art�culo"
         Top             =   5280
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Decimales"
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
         Index           =   88
         Left            =   3240
         TabIndex        =   600
         Top             =   6120
         Width           =   870
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   56
         Left            =   1635
         Picture         =   "frmListado.frx":0A93
         ToolTipText     =   "Buscar tarifa"
         Top             =   1560
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   21
         Left            =   1080
         TabIndex        =   155
         Top             =   1200
         Width           =   465
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   58
         Left            =   1635
         Picture         =   "frmListado.frx":0B95
         ToolTipText     =   "Buscar familia"
         Top             =   2520
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   57
         Left            =   1635
         Picture         =   "frmListado.frx":0C97
         ToolTipText     =   "Buscar familia"
         Top             =   2160
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
         Left            =   600
         TabIndex        =   146
         Top             =   1920
         Width           =   600
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   18
         Left            =   1080
         TabIndex        =   145
         Top             =   2520
         Width           =   420
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   17
         Left            =   1080
         TabIndex        =   144
         Top             =   2160
         Width           =   465
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   62
         Left            =   1635
         Picture         =   "frmListado.frx":0D99
         ToolTipText     =   "Buscar art�culo"
         Top             =   4680
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   61
         Left            =   1635
         Picture         =   "frmListado.frx":0E9B
         ToolTipText     =   "Buscar art�culo"
         Top             =   4320
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
         Index           =   14
         Left            =   600
         TabIndex        =   143
         Top             =   4080
         Width           =   660
      End
      Begin VB.Label lblTituloTarif 
         Caption         =   "Informe Precios y Descuentos"
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
         TabIndex        =   142
         Top             =   360
         Width           =   5655
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   16
         Left            =   1080
         TabIndex        =   141
         Top             =   4680
         Width           =   420
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   15
         Left            =   1080
         TabIndex        =   140
         Top             =   4320
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   14
         Left            =   1080
         TabIndex        =   139
         Top             =   3240
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   13
         Left            =   1080
         TabIndex        =   138
         Top             =   3600
         Width           =   420
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
         Index           =   13
         Left            =   600
         TabIndex        =   137
         Top             =   3000
         Width           =   525
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   59
         Left            =   1635
         Picture         =   "frmListado.frx":0F9D
         ToolTipText     =   "Buscar marca"
         Top             =   3240
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   60
         Left            =   1635
         Picture         =   "frmListado.frx":109F
         ToolTipText     =   "Buscar marca"
         Top             =   3600
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Opciones"
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
         TabIndex        =   125
         Top             =   6000
         Width           =   765
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   55
         Left            =   1635
         Picture         =   "frmListado.frx":11A1
         ToolTipText     =   "Buscar tarifa"
         Top             =   1200
         Width           =   240
      End
      Begin VB.Label Label4 
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
         Index           =   11
         Left            =   600
         TabIndex        =   124
         Top             =   960
         Width           =   495
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
         Left            =   1080
         TabIndex        =   123
         Top             =   1560
         Width           =   420
      End
   End
   Begin VB.Frame FrameEstMargenes 
      Height          =   5295
      Left            =   240
      TabIndex        =   422
      Top             =   120
      Width           =   7815
      Begin VB.CheckBox chkMargen 
         Caption         =   "Margen sobre venta"
         Height          =   255
         Index           =   4
         Left            =   5280
         TabIndex        =   761
         Top             =   4320
         Width           =   2295
      End
      Begin VB.CheckBox chkMargen 
         Caption         =   "Agrupa proveedor"
         Height          =   255
         Index           =   3
         Left            =   3120
         TabIndex        =   733
         Top             =   4680
         Width           =   1695
      End
      Begin VB.CheckBox chkMargen 
         Caption         =   "Incluir articulos de varios"
         Height          =   255
         Index           =   2
         Left            =   3120
         TabIndex        =   717
         Top             =   4320
         Width           =   2295
      End
      Begin VB.CheckBox chkMargen 
         Caption         =   "Detalla art�culo"
         Height          =   255
         Index           =   1
         Left            =   3120
         TabIndex        =   716
         Top             =   3960
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CheckBox chkMargen 
         Caption         =   "Detalla serie factura"
         Height          =   255
         Index           =   0
         Left            =   5280
         TabIndex        =   679
         Top             =   3960
         Width           =   2055
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   131
         Left            =   5040
         TabIndex        =   427
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   130
         Left            =   1800
         TabIndex        =   426
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Frame FrameValorar2 
         Caption         =   "Valorar Con:"
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
         Height          =   1215
         Left            =   360
         TabIndex        =   442
         Top             =   3840
         Width           =   2535
         Begin VB.OptionButton optPrecioMP2 
            Caption         =   "Precio Medio Ponderado"
            Height          =   255
            Left            =   240
            TabIndex        =   445
            Top             =   240
            Width           =   2055
         End
         Begin VB.OptionButton optPrecioUC2 
            Caption         =   "Precio Ultima Compra"
            Height          =   255
            Left            =   240
            TabIndex        =   444
            Top             =   525
            Value           =   -1  'True
            Width           =   2055
         End
         Begin VB.OptionButton optPrecioStd2 
            Caption         =   "Precio Standard"
            Height          =   255
            Left            =   240
            TabIndex        =   443
            Top             =   840
            Width           =   2055
         End
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   15
         Left            =   6600
         TabIndex        =   433
         Top             =   4680
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptarEst 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   5400
         TabIndex        =   432
         Top             =   4680
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   90
         Left            =   1800
         TabIndex        =   430
         Top             =   2880
         Width           =   1455
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   91
         Left            =   1800
         TabIndex        =   431
         Top             =   3240
         Width           =   1455
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   90
         Left            =   3300
         Locked          =   -1  'True
         TabIndex        =   438
         Text            =   "Text5"
         Top             =   2880
         Width           =   3975
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   91
         Left            =   3300
         Locked          =   -1  'True
         TabIndex        =   437
         Text            =   "Text5"
         Top             =   3240
         Width           =   3975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   88
         Left            =   1800
         MaxLength       =   4
         TabIndex        =   428
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   89
         Left            =   1800
         MaxLength       =   4
         TabIndex        =   429
         Top             =   2160
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   88
         Left            =   2600
         Locked          =   -1  'True
         TabIndex        =   425
         Text            =   "Text5"
         Top             =   1800
         Width           =   3615
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   89
         Left            =   2600
         Locked          =   -1  'True
         TabIndex        =   424
         Text            =   "Text5"
         Top             =   2160
         Width           =   3615
      End
      Begin VB.Image imgayuda 
         Height          =   240
         Index           =   3
         Left            =   4920
         ToolTipText     =   "Listado m�rgenes"
         Top             =   4680
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   100
         Left            =   4200
         TabIndex        =   651
         Top             =   1200
         Width           =   420
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   20
         Left            =   4680
         Picture         =   "frmListado.frx":12A3
         Top             =   1200
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   99
         Left            =   960
         TabIndex        =   650
         Top             =   1200
         Width           =   465
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   19
         Left            =   1560
         Picture         =   "frmListado.frx":132E
         Top             =   1200
         Width           =   240
      End
      Begin VB.Label Label4 
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
         Index           =   95
         Left            =   480
         TabIndex        =   649
         Top             =   840
         Width           =   1155
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   74
         Left            =   960
         TabIndex        =   441
         Top             =   2880
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   73
         Left            =   960
         TabIndex        =   440
         Top             =   3240
         Width           =   420
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
         Index           =   54
         Left            =   480
         TabIndex        =   439
         Top             =   2640
         Width           =   660
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   69
         Left            =   1515
         Picture         =   "frmListado.frx":13B9
         ToolTipText     =   "Buscar art�culo"
         Top             =   2880
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   70
         Left            =   1515
         Picture         =   "frmListado.frx":14BB
         ToolTipText     =   "Buscar art�culo"
         Top             =   3240
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   72
         Left            =   960
         TabIndex        =   436
         Top             =   1800
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   71
         Left            =   960
         TabIndex        =   435
         Top             =   2160
         Width           =   420
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
         Index           =   53
         Left            =   480
         TabIndex        =   434
         Top             =   1560
         Width           =   600
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   67
         Left            =   1515
         Picture         =   "frmListado.frx":15BD
         ToolTipText     =   "Buscar familia"
         Top             =   1800
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   68
         Left            =   1515
         Picture         =   "frmListado.frx":16BF
         ToolTipText     =   "buscar familia"
         Top             =   2160
         Width           =   240
      End
      Begin VB.Label lblTituloEst 
         Caption         =   "Informe Margenes de Venta por Art�culo"
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
         Left            =   720
         TabIndex        =   423
         Top             =   240
         Width           =   5895
      End
   End
   Begin VB.Frame FrameFraProveedor 
      Height          =   4455
      Left            =   2640
      TabIndex        =   734
      Top             =   240
      Visible         =   0   'False
      Width           =   4815
      Begin VB.TextBox txtImporte 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   6
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   759
         Text            =   "Text5"
         Top             =   3840
         Width           =   975
      End
      Begin VB.CommandButton cmdActVtosFraPro 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3480
         TabIndex        =   740
         Top             =   3360
         Width           =   975
      End
      Begin VB.TextBox txtImporte 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   5
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   754
         Text            =   "Text5"
         Top             =   3840
         Width           =   1335
      End
      Begin VB.TextBox txtImporte 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   4
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   752
         Text            =   "Text5"
         Top             =   2520
         Width           =   1335
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   153
         Left            =   1560
         TabIndex        =   739
         Top             =   2520
         Width           =   1215
      End
      Begin VB.TextBox txtImporte 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   3
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   750
         Text            =   "Text5"
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   152
         Left            =   1560
         TabIndex        =   738
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox txtImporte 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   2
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   748
         Text            =   "Text5"
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   151
         Left            =   1560
         TabIndex        =   737
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox txtImporte 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   1
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   746
         Text            =   "Text5"
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   150
         Left            =   1560
         TabIndex        =   736
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtImporte 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   0
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   744
         Text            =   "Text5"
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   149
         Left            =   1560
         TabIndex        =   735
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Salir"
         Height          =   375
         Index           =   16
         Left            =   3480
         TabIndex        =   741
         Top             =   3840
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "N� Asiento"
         Height          =   195
         Index           =   126
         Left            =   1680
         TabIndex        =   760
         Top             =   3600
         Width           =   1425
      End
      Begin VB.Label Label3 
         Caption         =   "Registro factura"
         Height          =   195
         Index           =   125
         Left            =   240
         TabIndex        =   758
         Top             =   3600
         Width           =   1425
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Contabilidad"
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
         Index           =   107
         Left            =   240
         TabIndex        =   755
         Top             =   3240
         Width           =   1050
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "... mas de cinco vencimientos"
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
         Index           =   106
         Left            =   1920
         TabIndex        =   753
         Top             =   2880
         Width           =   2460
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   25
         Left            =   1320
         Picture         =   "frmListado.frx":17C1
         Top             =   2520
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Quinto"
         Height          =   195
         Index           =   124
         Left            =   480
         TabIndex        =   751
         Top             =   2520
         Width           =   465
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   24
         Left            =   1320
         Picture         =   "frmListado.frx":184C
         Top             =   2040
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Cuarto"
         Height          =   195
         Index           =   123
         Left            =   480
         TabIndex        =   749
         Top             =   2040
         Width           =   465
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   23
         Left            =   1320
         Picture         =   "frmListado.frx":18D7
         Top             =   1560
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Tercero"
         Height          =   195
         Index           =   122
         Left            =   480
         TabIndex        =   747
         Top             =   1560
         Width           =   465
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   22
         Left            =   1320
         Picture         =   "frmListado.frx":1962
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Segundo"
         Height          =   195
         Index           =   121
         Left            =   480
         TabIndex        =   745
         Top             =   1080
         Width           =   705
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   21
         Left            =   1320
         Picture         =   "frmListado.frx":19ED
         Top             =   600
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Vencimientos"
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
         Index           =   105
         Left            =   240
         TabIndex        =   743
         Top             =   240
         Width           =   1140
      End
      Begin VB.Label Label3 
         Caption         =   "Primero"
         Height          =   195
         Index           =   120
         Left            =   480
         TabIndex        =   742
         Top             =   600
         Width           =   585
      End
   End
   Begin VB.Frame FrameInfArticulos 
      Height          =   6855
      Left            =   1440
      TabIndex        =   272
      Top             =   0
      Width           =   8715
      Begin VB.CheckBox chkImpEtiq 
         Caption         =   "Rotaci�n"
         Height          =   195
         Index           =   2
         Left            =   7560
         TabIndex        =   719
         Top             =   5760
         Width           =   1335
      End
      Begin VB.CheckBox chkImpEtiq 
         Caption         =   "P.V.P."
         Height          =   195
         Index           =   1
         Left            =   6240
         TabIndex        =   688
         Top             =   5760
         Width           =   1335
      End
      Begin VB.CheckBox chkImpEtiq 
         Caption         =   "Etiquetas"
         Height          =   195
         Index           =   0
         Left            =   5160
         TabIndex        =   672
         Top             =   5760
         Width           =   1095
      End
      Begin VB.Frame FrameSituacionArticulo 
         Caption         =   "Situaci�n art�culo"
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
         Height          =   735
         Left            =   120
         TabIndex        =   612
         Top             =   5880
         Width           =   4695
         Begin VB.CheckBox chkSitaucionArticulo2 
            Caption         =   "Obsoleto"
            Height          =   255
            Index           =   1
            Left            =   1200
            TabIndex        =   614
            Top             =   360
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.CheckBox chkSitaucionArticulo2 
            Caption         =   "Caducado"
            Height          =   255
            Index           =   3
            Left            =   3480
            TabIndex        =   616
            Top             =   360
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.CheckBox chkSitaucionArticulo2 
            Caption         =   "Bloqueado"
            Height          =   255
            Index           =   2
            Left            =   2280
            TabIndex        =   615
            Top             =   360
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.CheckBox chkSitaucionArticulo2 
            Caption         =   "Normal"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   613
            Top             =   360
            Value           =   1  'Checked
            Width           =   1095
         End
      End
      Begin VB.CheckBox chkMinimoCorreg 
         Caption         =   "No mostrar tarifas por encima de margen"
         Height          =   195
         Left            =   600
         TabIndex        =   558
         Top             =   5280
         Width           =   6015
      End
      Begin VB.Frame FrameStockMaxMin 
         Caption         =   "Imprimir Stocks"
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
         Height          =   615
         Left            =   360
         TabIndex        =   335
         Top             =   5880
         Width           =   4455
         Begin VB.OptionButton optPuntoPedido 
            Caption         =   "Punto de pedido"
            Height          =   255
            Left            =   2520
            TabIndex        =   289
            Top             =   280
            Width           =   1575
         End
         Begin VB.OptionButton optStockMin 
            Caption         =   "M�nimos"
            Height          =   255
            Left            =   1320
            TabIndex        =   288
            Top             =   280
            Width           =   975
         End
         Begin VB.OptionButton optStockMax 
            Caption         =   "M�ximos"
            Height          =   255
            Left            =   120
            TabIndex        =   287
            Top             =   280
            Width           =   975
         End
      End
      Begin VB.ComboBox cmbDecimales 
         Height          =   315
         ItemData        =   "frmListado.frx":1A78
         Left            =   600
         List            =   "frmListado.frx":1A85
         Style           =   2  'Dropdown List
         TabIndex        =   291
         Top             =   6120
         Width           =   1335
      End
      Begin VB.Frame FrameTapaINCORRECTO 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   1440
         TabIndex        =   544
         Top             =   840
         Width           =   4215
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   107
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   545
            Text            =   "Text5"
            Top             =   45
            Width           =   3015
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   107
            Left            =   360
            MaxLength       =   4
            TabIndex        =   275
            Top             =   45
            Width           =   615
         End
         Begin VB.Image imgBuscarG 
            Height          =   240
            Index           =   87
            Left            =   80
            Picture         =   "frmListado.frx":1AA4
            ToolTipText     =   "Buscar almacen"
            Top             =   60
            Width           =   240
         End
      End
      Begin VB.Frame FrameOrden 
         BorderStyle     =   0  'None
         Height          =   1935
         Left            =   5760
         TabIndex        =   379
         Top             =   840
         Width           =   2655
         Begin VB.CommandButton cmdBajar 
            Height          =   510
            Left            =   1920
            Picture         =   "frmListado.frx":1BA6
            Style           =   1  'Graphical
            TabIndex        =   381
            Top             =   1305
            Width           =   510
         End
         Begin VB.CommandButton cmdSubir 
            Caption         =   " "
            Height          =   510
            Left            =   1920
            Picture         =   "frmListado.frx":1EB0
            Style           =   1  'Graphical
            TabIndex        =   380
            Top             =   600
            Width           =   510
         End
         Begin MSComctlLib.ListView ListView2 
            Height          =   1335
            Left            =   120
            TabIndex        =   382
            Top             =   480
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   2355
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
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
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Orden del Informe"
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
            TabIndex        =   383
            Top             =   240
            Width           =   1545
         End
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   72
         Left            =   1800
         MaxLength       =   4
         TabIndex        =   276
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   72
         Left            =   2445
         Locked          =   -1  'True
         TabIndex        =   333
         Text            =   "Text5"
         Top             =   840
         Width           =   3015
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   69
         Left            =   2445
         Locked          =   -1  'True
         TabIndex        =   313
         Text            =   "Text5"
         Top             =   4470
         Width           =   2895
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   68
         Left            =   2445
         Locked          =   -1  'True
         TabIndex        =   312
         Text            =   "Text5"
         Top             =   4150
         Width           =   2895
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   69
         Left            =   1800
         MaxLength       =   2
         TabIndex        =   284
         Top             =   4470
         Width           =   615
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   68
         Left            =   1800
         MaxLength       =   2
         TabIndex        =   283
         Top             =   4155
         Width           =   615
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   65
         Left            =   2445
         Locked          =   -1  'True
         TabIndex        =   308
         Text            =   "Text5"
         Top             =   2590
         Width           =   3015
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   64
         Left            =   2445
         Locked          =   -1  'True
         TabIndex        =   307
         Text            =   "Text5"
         Top             =   2280
         Width           =   3015
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   65
         Left            =   1800
         MaxLength       =   4
         TabIndex        =   280
         Top             =   2590
         Width           =   615
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   64
         Left            =   1800
         MaxLength       =   4
         TabIndex        =   279
         Top             =   2280
         Width           =   615
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   63
         Left            =   2445
         Locked          =   -1  'True
         TabIndex        =   296
         Text            =   "Text5"
         Top             =   1750
         Width           =   3015
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   62
         Left            =   2445
         Locked          =   -1  'True
         TabIndex        =   295
         Text            =   "Text5"
         Top             =   1440
         Width           =   3015
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   71
         Left            =   3285
         Locked          =   -1  'True
         TabIndex        =   294
         Text            =   "Text5"
         Top             =   5400
         Width           =   3855
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   70
         Left            =   3285
         Locked          =   -1  'True
         TabIndex        =   293
         Text            =   "Text5"
         Top             =   5080
         Width           =   3855
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   63
         Left            =   1800
         MaxLength       =   4
         TabIndex        =   278
         Top             =   1750
         Width           =   615
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   62
         Left            =   1800
         MaxLength       =   4
         TabIndex        =   277
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   71
         Left            =   1800
         MaxLength       =   16
         TabIndex        =   286
         Top             =   5400
         Width           =   1455
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   70
         Left            =   1800
         MaxLength       =   16
         TabIndex        =   285
         Top             =   5080
         Width           =   1455
      End
      Begin VB.CommandButton cmdAceptarArtic 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   5280
         TabIndex        =   290
         Top             =   6120
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   11
         Left            =   6360
         TabIndex        =   292
         Top             =   6120
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   66
         Left            =   1800
         MaxLength       =   6
         TabIndex        =   281
         Top             =   3190
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   67
         Left            =   1800
         MaxLength       =   6
         TabIndex        =   282
         Top             =   3510
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   66
         Left            =   2565
         Locked          =   -1  'True
         TabIndex        =   274
         Text            =   "Text5"
         Top             =   3190
         Width           =   4575
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   67
         Left            =   2565
         Locked          =   -1  'True
         TabIndex        =   273
         Text            =   "Text5"
         Top             =   3510
         Width           =   4575
      End
      Begin VB.ComboBox cmbProduccion 
         Height          =   315
         ItemData        =   "frmListado.frx":21BA
         Left            =   2280
         List            =   "frmListado.frx":21C4
         Style           =   2  'Dropdown List
         TabIndex        =   610
         Top             =   6120
         Width           =   2415
      End
      Begin VB.Image imgayuda 
         Height          =   240
         Index           =   0
         Left            =   4920
         ToolTipText     =   "Facturacion Renting y servicios"
         Top             =   6120
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Verificar sobre"
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
         Index           =   90
         Left            =   2280
         TabIndex        =   611
         Top             =   5880
         Width           =   1230
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Decimales"
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
         Index           =   75
         Left            =   600
         TabIndex        =   546
         Top             =   5880
         Width           =   870
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
         Index           =   39
         Left            =   600
         TabIndex        =   305
         Top             =   1200
         Width           =   600
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
         Index           =   36
         Left            =   600
         TabIndex        =   334
         Top             =   890
         Width           =   735
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   18
         Left            =   1515
         Picture         =   "frmListado.frx":21F5
         Top             =   840
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   26
         Left            =   1515
         Picture         =   "frmListado.frx":22F7
         ToolTipText     =   "Buscar tipo articulo"
         Top             =   4485
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   25
         Left            =   1515
         Picture         =   "frmListado.frx":23F9
         ToolTipText     =   "Buscar tipo articulo"
         Top             =   4155
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Articulo"
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
         Left            =   600
         TabIndex        =   332
         Top             =   3900
         Width           =   1320
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   60
         Left            =   960
         TabIndex        =   331
         Top             =   4470
         Width           =   420
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   59
         Left            =   960
         TabIndex        =   330
         Top             =   4155
         Width           =   465
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   22
         Left            =   1515
         Picture         =   "frmListado.frx":24FB
         ToolTipText     =   "Buscar marca"
         Top             =   2625
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   21
         Left            =   1515
         Picture         =   "frmListado.frx":25FD
         ToolTipText     =   "Buscar marca"
         Top             =   2280
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
         Index           =   35
         Left            =   600
         TabIndex        =   311
         Top             =   2040
         Width           =   525
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   58
         Left            =   960
         TabIndex        =   310
         Top             =   2595
         Width           =   420
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   57
         Left            =   960
         TabIndex        =   309
         Top             =   2280
         Width           =   465
      End
      Begin VB.Label Label9 
         Caption         =   "Informe de Articulos"
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
         TabIndex        =   306
         Top             =   360
         Width           =   6735
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   20
         Left            =   1515
         Picture         =   "frmListado.frx":26FF
         ToolTipText     =   "Buscar familia"
         Top             =   1770
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   19
         Left            =   1515
         Picture         =   "frmListado.frx":2801
         ToolTipText     =   "Buscar familia"
         Top             =   1440
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   56
         Left            =   960
         TabIndex        =   304
         Top             =   1755
         Width           =   420
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   55
         Left            =   960
         TabIndex        =   303
         Top             =   1440
         Width           =   465
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   28
         Left            =   1515
         Picture         =   "frmListado.frx":2903
         ToolTipText     =   "Buscar art�culo"
         Top             =   5400
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   27
         Left            =   1515
         Picture         =   "frmListado.frx":2A05
         ToolTipText     =   "Buscar art�culo"
         Top             =   5085
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
         Index           =   38
         Left            =   600
         TabIndex        =   302
         Top             =   4820
         Width           =   660
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   54
         Left            =   960
         TabIndex        =   301
         Top             =   5400
         Width           =   420
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   51
         Left            =   960
         TabIndex        =   300
         Top             =   5085
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   50
         Left            =   960
         TabIndex        =   299
         Top             =   3195
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   48
         Left            =   960
         TabIndex        =   298
         Top             =   3510
         Width           =   420
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
         Left            =   600
         TabIndex        =   297
         Top             =   2950
         Width           =   885
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   23
         Left            =   1515
         Picture         =   "frmListado.frx":2B07
         ToolTipText     =   "Buscar proveedor"
         Top             =   3195
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   24
         Left            =   1515
         Picture         =   "frmListado.frx":2C09
         ToolTipText     =   "Buscar proveedor"
         Top             =   3540
         Width           =   240
      End
   End
   Begin VB.Frame FrameBultos 
      Height          =   6975
      Left            =   0
      TabIndex        =   475
      Top             =   0
      Width           =   6735
      Begin VB.OptionButton optBultos 
         Caption         =   "Direcci�n env�o"
         Height          =   255
         Index           =   1
         Left            =   3840
         TabIndex        =   757
         Top             =   1320
         Width           =   2415
      End
      Begin VB.OptionButton optBultos 
         Caption         =   "Option1"
         Height          =   255
         Index           =   0
         Left            =   1320
         TabIndex        =   756
         Top             =   1320
         Value           =   -1  'True
         Width           =   2295
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   148
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   730
         Text            =   "Text5"
         Top             =   1080
         Width           =   4215
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   148
         Left            =   1440
         MaxLength       =   6
         TabIndex        =   478
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox txtBultos 
         Height          =   285
         Index           =   7
         Left            =   3000
         TabIndex        =   487
         Text            =   "Text1"
         Top             =   6480
         Width           =   615
      End
      Begin VB.TextBox txtBultos 
         Height          =   320
         Index           =   6
         Left            =   1320
         TabIndex        =   484
         Text            =   "Text1"
         Top             =   3600
         Visible         =   0   'False
         Width           =   5175
      End
      Begin VB.TextBox txtBultos 
         Height          =   320
         Index           =   5
         Left            =   2280
         TabIndex        =   483
         Text            =   "Text1"
         Top             =   3120
         Width           =   4215
      End
      Begin VB.TextBox txtBultos 
         Height          =   320
         Index           =   4
         Left            =   1320
         TabIndex        =   482
         Text            =   "Text1"
         Top             =   3120
         Width           =   855
      End
      Begin VB.TextBox txtBultos 
         Height          =   320
         Index           =   3
         Left            =   1320
         TabIndex        =   481
         Text            =   "Text1"
         Top             =   2640
         Width           =   5175
      End
      Begin VB.TextBox txtBultos 
         Height          =   320
         Index           =   2
         Left            =   1320
         TabIndex        =   480
         Text            =   "Text1"
         Top             =   2160
         Width           =   5175
      End
      Begin VB.ComboBox cmbBulto 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   479
         Top             =   1620
         Width           =   5175
      End
      Begin VB.TextBox txtBultos 
         Height          =   285
         Index           =   1
         Left            =   1320
         TabIndex        =   486
         Text            =   "Text1"
         Top             =   6480
         Width           =   615
      End
      Begin VB.CommandButton cmdEtiqBulto 
         Caption         =   "Imprimir"
         Height          =   375
         Left            =   4440
         TabIndex        =   488
         Top             =   6480
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   95
         Left            =   5520
         TabIndex        =   489
         Top             =   6480
         Width           =   975
      End
      Begin VB.TextBox txtBultos 
         Height          =   1695
         Index           =   0
         Left            =   1320
         MultiLine       =   -1  'True
         TabIndex        =   485
         Text            =   "frmListado.frx":2D0B
         Top             =   4200
         Width           =   5175
      End
      Begin VB.TextBox txtClie 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         TabIndex        =   477
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   10
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   490
         Text            =   "Text5"
         Top             =   840
         Width           =   4335
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
         Index           =   104
         Left            =   240
         TabIndex        =   731
         Top             =   1080
         Width           =   885
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   148
         Left            =   1200
         Picture         =   "frmListado.frx":2D11
         ToolTipText     =   "Buscar art�culo"
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "En blanco"
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
         Index           =   91
         Left            =   2040
         TabIndex        =   617
         Top             =   6480
         Width           =   795
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Pais"
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
         Index           =   71
         Left            =   240
         TabIndex        =   516
         Top             =   3663
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Poblaci�n"
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
         Index           =   70
         Left            =   240
         TabIndex        =   515
         Top             =   2703
         Width           =   675
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Provincia"
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
         Index           =   69
         Left            =   240
         TabIndex        =   514
         Top             =   3183
         Width           =   645
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Direcci�n"
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
         Index           =   68
         Left            =   240
         TabIndex        =   513
         Top             =   2223
         Width           =   645
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "N� Copias"
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
         TabIndex        =   494
         Top             =   6480
         Width           =   795
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Texto"
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
         TabIndex        =   493
         Top             =   4200
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Direcci�n"
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
         TabIndex        =   492
         Top             =   1680
         Width           =   780
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
         Index           =   61
         Left            =   240
         TabIndex        =   491
         Top             =   840
         Width           =   705
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   75
         Left            =   1080
         Picture         =   "frmListado.frx":2E13
         ToolTipText     =   "Buscar cliente"
         Top             =   840
         Width           =   240
      End
      Begin VB.Label Label6 
         Caption         =   "Etiquetas de bultos"
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
         Height          =   465
         Left            =   1680
         TabIndex        =   476
         Top             =   240
         Width           =   3375
      End
   End
   Begin VB.Frame FrameInvArtComp 
      Height          =   4455
      Left            =   2160
      TabIndex        =   618
      Top             =   1200
      Width           =   7335
      Begin VB.Frame FrameAlmacenesListadoComponentes 
         Height          =   1455
         Left            =   960
         TabIndex        =   723
         Top             =   2160
         Visible         =   0   'False
         Width           =   5175
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   147
            Left            =   1080
            MaxLength       =   4
            TabIndex        =   630
            Top             =   960
            Width           =   615
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   147
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   726
            Text            =   "Text5"
            Top             =   960
            Width           =   3015
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   146
            Left            =   1080
            MaxLength       =   4
            TabIndex        =   629
            Top             =   600
            Width           =   615
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   146
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   725
            Text            =   "Text5"
            Top             =   600
            Width           =   3015
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   145
            Left            =   1080
            MaxLength       =   4
            TabIndex        =   628
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   145
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   724
            Text            =   "Text5"
            Top             =   240
            Width           =   3015
         End
         Begin VB.Label Label3 
            Caption         =   "Alm 3"
            Height          =   195
            Index           =   119
            Left            =   120
            TabIndex        =   729
            Top             =   960
            Width           =   465
         End
         Begin VB.Image imgBuscarG 
            Height          =   240
            Index           =   109
            Left            =   675
            Picture         =   "frmListado.frx":2F15
            ToolTipText     =   "Buscar art�culo"
            Top             =   960
            Width           =   240
         End
         Begin VB.Label Label3 
            Caption         =   "Alm 2"
            Height          =   195
            Index           =   118
            Left            =   120
            TabIndex        =   728
            Top             =   600
            Width           =   465
         End
         Begin VB.Image imgBuscarG 
            Height          =   240
            Index           =   108
            Left            =   675
            Picture         =   "frmListado.frx":3017
            ToolTipText     =   "Buscar art�culo"
            Top             =   600
            Width           =   240
         End
         Begin VB.Label Label3 
            Caption         =   "Alm 1"
            Height          =   195
            Index           =   117
            Left            =   120
            TabIndex        =   727
            Top             =   240
            Width           =   465
         End
         Begin VB.Image imgBuscarG 
            Height          =   240
            Index           =   107
            Left            =   675
            Picture         =   "frmListado.frx":3119
            ToolTipText     =   "Buscar art�culo"
            Top             =   240
            Width           =   240
         End
      End
      Begin VB.CheckBox chkCompo 
         Caption         =   "Listado informativo componentes x articulo"
         Height          =   255
         Left            =   240
         TabIndex        =   623
         Top             =   1800
         Width           =   3735
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   0
         Left            =   4080
         TabIndex        =   632
         Top             =   3840
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   0
         Left            =   5280
         TabIndex        =   637
         Top             =   3840
         Width           =   975
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   126
         Left            =   2955
         Locked          =   -1  'True
         TabIndex        =   635
         Text            =   "Text5"
         Top             =   1320
         Width           =   3855
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   126
         Left            =   1440
         MaxLength       =   16
         TabIndex        =   622
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   125
         Left            =   2955
         Locked          =   -1  'True
         TabIndex        =   631
         Text            =   "Text5"
         Top             =   960
         Width           =   3855
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   125
         Left            =   1440
         MaxLength       =   16
         TabIndex        =   621
         Top             =   960
         Width           =   1455
      End
      Begin VB.Frame FrameValorar3 
         Caption         =   "Valorar Con:"
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
         Height          =   1575
         Left            =   4560
         TabIndex        =   620
         Top             =   1920
         Width           =   2535
         Begin VB.OptionButton optPrecioMP3 
            Caption         =   "Precio Medio Ponderado"
            Height          =   255
            Left            =   240
            TabIndex        =   624
            Top             =   240
            Value           =   -1  'True
            Width           =   2055
         End
         Begin VB.OptionButton optPrecioMA3 
            Caption         =   "Precio Medio Acumulado"
            Height          =   255
            Left            =   240
            TabIndex        =   625
            Top             =   560
            Width           =   2175
         End
         Begin VB.OptionButton optPrecioUC3 
            Caption         =   "Precio Ultima Compra"
            Height          =   255
            Left            =   240
            TabIndex        =   626
            Top             =   880
            Width           =   2055
         End
         Begin VB.OptionButton optPrecioStd3 
            Caption         =   "Precio Standard"
            Height          =   255
            Left            =   240
            TabIndex        =   627
            Top             =   1200
            Width           =   2055
         End
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   98
         Left            =   1155
         Picture         =   "frmListado.frx":321B
         ToolTipText     =   "Buscar art�culo"
         Top             =   1320
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   98
         Left            =   600
         TabIndex        =   636
         Top             =   1320
         Width           =   420
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   95
         Left            =   1155
         Picture         =   "frmListado.frx":331D
         ToolTipText     =   "Buscar art�culo"
         Top             =   960
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
         Index           =   92
         Left            =   120
         TabIndex        =   634
         Top             =   720
         Width           =   2460
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   97
         Left            =   600
         TabIndex        =   633
         Top             =   960
         Width           =   465
      End
      Begin VB.Label Label12 
         Caption         =   "Listado art�culos - componentes"
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
         Left            =   240
         TabIndex        =   619
         Top             =   240
         Width           =   6855
      End
   End
   Begin VB.Frame FrameInventario 
      Height          =   6735
      Left            =   240
      TabIndex        =   73
      Top             =   0
      Width           =   7995
      Begin VB.Frame Frame8 
         Height          =   1455
         Left            =   5400
         TabIndex        =   673
         Top             =   4320
         Width           =   2295
         Begin VB.CheckBox chkProv2 
            Caption         =   "Varios"
            Height          =   255
            Index           =   2
            Left            =   1200
            TabIndex        =   678
            Top             =   480
            Width           =   975
         End
         Begin VB.CheckBox chkProv2 
            Caption         =   "Detalla"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   55
            Top             =   480
            Value           =   1  'Checked
            Width           =   975
         End
         Begin VB.ComboBox cboStokFecha 
            Height          =   315
            ItemData        =   "frmListado.frx":341F
            Left            =   120
            List            =   "frmListado.frx":342C
            Style           =   2  'Dropdown List
            TabIndex        =   57
            Top             =   1080
            Width           =   975
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   136
            Left            =   1200
            MaxLength       =   5
            TabIndex        =   58
            Text            =   "0.00"
            Top             =   1080
            Width           =   855
         End
         Begin VB.CheckBox chkProv2 
            Caption         =   "Agrupa proveedor"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   54
            Top             =   120
            Width           =   1815
         End
         Begin VB.Label Label3 
            Caption         =   "Valores"
            Height          =   195
            Index           =   108
            Left            =   120
            TabIndex        =   674
            Top             =   840
            Width           =   825
         End
         Begin VB.Label Label3 
            Caption         =   "Corrector(%)"
            Height          =   195
            Index           =   107
            Left            =   1200
            TabIndex        =   56
            Top             =   840
            Width           =   945
         End
      End
      Begin VB.Frame FrameOpciones2 
         Height          =   1575
         Left            =   2880
         TabIndex        =   384
         Top             =   4320
         Width           =   2415
         Begin VB.CheckBox chkValorDesdeArticulo 
            Caption         =   "Desde art."
            Height          =   255
            Left            =   1200
            TabIndex        =   718
            Top             =   1200
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.CheckBox chkValorado 
            Caption         =   "Valorado"
            Height          =   255
            Left            =   120
            TabIndex        =   388
            Top             =   1200
            Value           =   1  'Checked
            Width           =   975
         End
         Begin VB.CheckBox chkImprimeStock 
            Caption         =   "Imprimir Stock"
            Height          =   255
            Left            =   120
            TabIndex        =   387
            Top             =   840
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin VB.CheckBox chkSinStock 
            Caption         =   "Imprimir Art. sin Stock"
            Height          =   255
            Left            =   120
            TabIndex        =   386
            Top             =   480
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.CheckBox chkSaltaPag 
            Caption         =   "Salta p�g. en Familia"
            Height          =   255
            Left            =   120
            TabIndex        =   385
            Top             =   120
            Width           =   2055
         End
      End
      Begin VB.Frame FrameValorar 
         Caption         =   "Valorar Con:"
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
         Height          =   1575
         Left            =   240
         TabIndex        =   95
         Top             =   4320
         Visible         =   0   'False
         Width           =   2535
         Begin VB.OptionButton optPrecioStd 
            Caption         =   "Precio Standard"
            Height          =   255
            Left            =   240
            TabIndex        =   99
            Top             =   1200
            Width           =   2055
         End
         Begin VB.OptionButton optPrecioUC 
            Caption         =   "Precio Ultima Compra"
            Height          =   255
            Left            =   240
            TabIndex        =   98
            Top             =   880
            Width           =   2055
         End
         Begin VB.OptionButton optPrecioMA 
            Caption         =   "Precio Medio Acumulado"
            Height          =   255
            Left            =   240
            TabIndex        =   97
            Top             =   560
            Width           =   2175
         End
         Begin VB.OptionButton optPrecioMP 
            Caption         =   "Precio Medio Ponderado"
            Height          =   255
            Left            =   240
            TabIndex        =   96
            Top             =   240
            Value           =   -1  'True
            Width           =   2055
         End
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   22
         Left            =   4920
         TabIndex        =   52
         Top             =   4440
         Width           =   1215
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   21
         Left            =   1920
         TabIndex        =   53
         Top             =   4680
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   21
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   93
         Text            =   "Text5"
         Top             =   4680
         Width           =   4215
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   19
         Left            =   2720
         Locked          =   -1  'True
         TabIndex        =   89
         Text            =   "Text5"
         Top             =   3960
         Width           =   4575
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   18
         Left            =   2720
         Locked          =   -1  'True
         TabIndex        =   88
         Text            =   "Text5"
         Top             =   3600
         Width           =   4575
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   19
         Left            =   1920
         TabIndex        =   50
         Top             =   3960
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   18
         Left            =   1920
         TabIndex        =   49
         Top             =   3600
         Width           =   735
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   4
         Left            =   6600
         TabIndex        =   60
         Top             =   5760
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   4
         Left            =   5520
         TabIndex        =   59
         Top             =   5760
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   14
         Left            =   1920
         MaxLength       =   16
         TabIndex        =   45
         Top             =   1680
         Width           =   1455
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   15
         Left            =   1920
         MaxLength       =   16
         TabIndex        =   46
         Top             =   2040
         Width           =   1455
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   16
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   47
         Top             =   2640
         Width           =   615
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   17
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   48
         Top             =   3000
         Width           =   615
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   20
         Left            =   2440
         TabIndex        =   51
         Top             =   4440
         Width           =   1215
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   13
         Left            =   1920
         MaxLength       =   3
         TabIndex        =   44
         Top             =   1080
         Width           =   495
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   14
         Left            =   3440
         Locked          =   -1  'True
         TabIndex        =   78
         Text            =   "Text5"
         Top             =   1680
         Width           =   3855
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   15
         Left            =   3440
         Locked          =   -1  'True
         TabIndex        =   77
         Text            =   "Text5"
         Top             =   2040
         Width           =   3855
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   16
         Left            =   2600
         Locked          =   -1  'True
         TabIndex        =   76
         Text            =   "Text5"
         Top             =   2640
         Width           =   3135
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   17
         Left            =   2600
         Locked          =   -1  'True
         TabIndex        =   75
         Text            =   "Text5"
         Top             =   3000
         Width           =   3135
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   13
         Left            =   2480
         Locked          =   -1  'True
         TabIndex        =   74
         Text            =   "Text5"
         Top             =   1080
         Width           =   3255
      End
      Begin VB.Label Label3 
         Caption         =   "Indicador"
         Height          =   195
         Index           =   109
         Left            =   120
         TabIndex        =   675
         Top             =   6600
         Width           =   4305
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   3
         Left            =   4670
         Picture         =   "frmListado.frx":344D
         Top             =   4440
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
         Left            =   4200
         TabIndex        =   101
         Top             =   4440
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
         Height          =   195
         Index           =   8
         Left            =   3720
         TabIndex        =   100
         Top             =   4440
         Width           =   450
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
         Index           =   7
         Left            =   600
         TabIndex        =   94
         Top             =   4680
         Width           =   945
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   17
         Left            =   1635
         Picture         =   "frmListado.frx":34D8
         Top             =   4680
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   16
         Left            =   1635
         Picture         =   "frmListado.frx":35DA
         ToolTipText     =   "Buscar provedor"
         Top             =   3960
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   15
         Left            =   1635
         Picture         =   "frmListado.frx":36DC
         ToolTipText     =   "Buscar proveedor"
         Top             =   3600
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
         Left            =   600
         TabIndex        =   92
         Top             =   3360
         Width           =   885
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   25
         Left            =   1080
         TabIndex        =   91
         Top             =   3960
         Width           =   420
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   24
         Left            =   1080
         TabIndex        =   90
         Top             =   3600
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   23
         Left            =   1080
         TabIndex        =   87
         Top             =   1680
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   22
         Left            =   1080
         TabIndex        =   86
         Top             =   2040
         Width           =   420
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
         Left            =   600
         TabIndex        =   84
         Top             =   1440
         Width           =   660
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   11
         Left            =   1635
         Picture         =   "frmListado.frx":37DE
         ToolTipText     =   "Buscar art�culo"
         Top             =   1680
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   12
         Left            =   1635
         Picture         =   "frmListado.frx":38E0
         ToolTipText     =   "Buscar art�culo"
         Top             =   2040
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   20
         Left            =   1080
         TabIndex        =   83
         Top             =   2640
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   19
         Left            =   1080
         TabIndex        =   82
         Top             =   3000
         Width           =   420
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
         Index           =   3
         Left            =   600
         TabIndex        =   81
         Top             =   2400
         Width           =   600
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   13
         Left            =   1635
         Picture         =   "frmListado.frx":39E2
         ToolTipText     =   "Buscar familia"
         Top             =   2640
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   14
         Left            =   1635
         Picture         =   "frmListado.frx":3AE4
         ToolTipText     =   "Buscar familia"
         Top             =   3000
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Inventario"
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
         Left            =   600
         TabIndex        =   80
         Top             =   4440
         Width           =   1440
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
         Index           =   1
         Left            =   600
         TabIndex        =   79
         Top             =   1080
         Width           =   735
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   10
         Left            =   1635
         Picture         =   "frmListado.frx":3BE6
         ToolTipText     =   "Buscar almacen"
         Top             =   1080
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   2
         Left            =   2140
         Picture         =   "frmListado.frx":3CE8
         Top             =   4440
         Width           =   240
      End
      Begin VB.Label lbltituloInven 
         Caption         =   "Informe Toma de Inventario Articulos"
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
         Left            =   240
         TabIndex        =   85
         Top             =   360
         Width           =   7575
      End
   End
   Begin VB.Frame FrameAlmacenStkMin 
      Height          =   5655
      Left            =   240
      TabIndex        =   689
      Top             =   0
      Visible         =   0   'False
      Width           =   5775
      Begin VB.CheckBox chkVarios 
         Caption         =   "Articulos sin stock m�nimo"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   714
         Top             =   4800
         Width           =   2535
      End
      Begin VB.CommandButton cmdStockMin 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3360
         TabIndex        =   696
         Top             =   4920
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   100
         Left            =   4560
         TabIndex        =   697
         Top             =   4920
         Width           =   975
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   144
         Left            =   2115
         Locked          =   -1  'True
         TabIndex        =   712
         Text            =   "Text5"
         Top             =   4200
         Width           =   3375
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   144
         Left            =   1245
         MaxLength       =   6
         TabIndex        =   695
         Top             =   4200
         Width           =   830
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   143
         Left            =   2115
         Locked          =   -1  'True
         TabIndex        =   709
         Text            =   "Text5"
         Top             =   3840
         Width           =   3375
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   143
         Left            =   1245
         MaxLength       =   6
         TabIndex        =   694
         Top             =   3840
         Width           =   830
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   142
         Left            =   1245
         MaxLength       =   4
         TabIndex        =   693
         Top             =   2880
         Width           =   830
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   142
         Left            =   2115
         Locked          =   -1  'True
         TabIndex        =   707
         Text            =   "Text5"
         Top             =   2880
         Width           =   3375
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   141
         Left            =   1245
         MaxLength       =   4
         TabIndex        =   692
         Top             =   2520
         Width           =   830
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   141
         Left            =   2115
         Locked          =   -1  'True
         TabIndex        =   704
         Text            =   "Text5"
         Top             =   2520
         Width           =   3375
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   140
         Left            =   1245
         TabIndex        =   691
         Top             =   1680
         Width           =   830
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   139
         Left            =   2115
         Locked          =   -1  'True
         TabIndex        =   699
         Text            =   "Text5"
         Top             =   1320
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   140
         Left            =   2115
         Locked          =   -1  'True
         TabIndex        =   698
         Text            =   "Text5"
         Top             =   1680
         Width           =   3375
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   139
         Left            =   1245
         TabIndex        =   690
         Top             =   1320
         Width           =   830
      End
      Begin VB.Label Label3 
         Height          =   195
         Index           =   116
         Left            =   240
         TabIndex        =   715
         Top             =   5160
         Width           =   2505
      End
      Begin VB.Image imgayuda 
         Height          =   240
         Index           =   1
         Left            =   2880
         ToolTipText     =   "Facturacion Renting y servicios"
         Top             =   5040
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   140
         Left            =   960
         Picture         =   "frmListado.frx":3D73
         ToolTipText     =   "Buscar familia"
         Top             =   1680
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   139
         Left            =   960
         Picture         =   "frmListado.frx":3E75
         ToolTipText     =   "Buscar familia"
         Top             =   1320
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   144
         Left            =   960
         Picture         =   "frmListado.frx":3F77
         ToolTipText     =   "Buscar proveedor"
         Top             =   4230
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   115
         Left            =   360
         TabIndex        =   713
         Top             =   4200
         Width           =   420
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   143
         Left            =   960
         Picture         =   "frmListado.frx":4079
         ToolTipText     =   "Buscar proveedor"
         Top             =   3840
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
         Index           =   102
         Left            =   360
         TabIndex        =   711
         Top             =   3480
         Width           =   885
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   114
         Left            =   360
         TabIndex        =   710
         Top             =   3840
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   113
         Left            =   360
         TabIndex        =   708
         Top             =   2880
         Width           =   420
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   142
         Left            =   960
         Picture         =   "frmListado.frx":417B
         ToolTipText     =   "Buscar familia"
         Top             =   2895
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   112
         Left            =   360
         TabIndex        =   706
         Top             =   2520
         Width           =   465
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   141
         Left            =   960
         Picture         =   "frmListado.frx":427D
         ToolTipText     =   "Buscar familia"
         Top             =   2520
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
         Index           =   101
         Left            =   360
         TabIndex        =   705
         Top             =   2280
         Width           =   600
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Listado almacen con stock m�nimo"
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
         Left            =   480
         TabIndex        =   703
         Top             =   360
         Width           =   5055
      End
      Begin VB.Label Label1 
         Caption         =   "Desde"
         Height          =   195
         Index           =   9
         Left            =   360
         TabIndex        =   702
         Top             =   1365
         Width           =   465
      End
      Begin VB.Label Label1 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   8
         Left            =   360
         TabIndex        =   701
         Top             =   1680
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Almac�n"
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
         Left            =   360
         TabIndex        =   700
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame FrameEtiqEstanteria 
      Height          =   5175
      Left            =   0
      TabIndex        =   446
      Top             =   0
      Width           =   7815
      Begin VB.Frame FrameTapaEtiq 
         Height          =   3615
         Left            =   120
         TabIndex        =   680
         Top             =   240
         Width           =   7575
         Begin VB.Label Label3 
            Caption         =   "Desd"
            Height          =   1095
            Index           =   111
            Left            =   360
            TabIndex        =   681
            Top             =   720
            Width           =   6975
         End
      End
      Begin VB.CheckBox chkDtoFM 
         Caption         =   "Mostrar descuento fam/marca"
         Height          =   255
         Left            =   360
         TabIndex        =   455
         Top             =   4680
         Width           =   3015
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   124
         Left            =   4140
         TabIndex        =   452
         Top             =   3480
         Width           =   1095
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   123
         Left            =   1800
         TabIndex        =   451
         Top             =   3480
         Width           =   1095
      End
      Begin VB.ComboBox cboDecimal 
         Height          =   315
         ItemData        =   "frmListado.frx":437F
         Left            =   1800
         List            =   "frmListado.frx":4392
         Style           =   2  'Dropdown List
         TabIndex        =   453
         Top             =   4080
         Width           =   735
      End
      Begin VB.CheckBox chkImprimeCodigoBarras 
         Caption         =   "Impime codigo barras"
         Height          =   255
         Left            =   2760
         TabIndex        =   454
         Top             =   4080
         Width           =   2055
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   95
         Left            =   2600
         Locked          =   -1  'True
         TabIndex        =   461
         Text            =   "Text5"
         Top             =   1680
         Width           =   3615
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   94
         Left            =   2600
         Locked          =   -1  'True
         TabIndex        =   460
         Text            =   "Text5"
         Top             =   1320
         Width           =   3615
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   95
         Left            =   1800
         MaxLength       =   4
         TabIndex        =   448
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   94
         Left            =   1800
         MaxLength       =   4
         TabIndex        =   447
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   93
         Left            =   3300
         Locked          =   -1  'True
         TabIndex        =   459
         Text            =   "Text5"
         Top             =   2760
         Width           =   3975
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   92
         Left            =   3300
         Locked          =   -1  'True
         TabIndex        =   457
         Text            =   "Text5"
         Top             =   2400
         Width           =   3975
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   93
         Left            =   1800
         TabIndex        =   450
         Top             =   2760
         Width           =   1455
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   92
         Left            =   1800
         TabIndex        =   449
         Top             =   2400
         Width           =   1455
      End
      Begin VB.CommandButton cmdEtiqEstanteria 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   5640
         TabIndex        =   456
         Top             =   4560
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   94
         Left            =   6720
         TabIndex        =   458
         Top             =   4560
         Width           =   975
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   18
         Left            =   3840
         Picture         =   "frmListado.frx":43A5
         Top             =   3480
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   17
         Left            =   1515
         Picture         =   "frmListado.frx":4430
         Top             =   3480
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   96
         Left            =   3315
         TabIndex        =   609
         Top             =   3480
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   7
         Left            =   960
         TabIndex        =   608
         Top             =   3480
         Width           =   465
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha ult. cambio precio P.V.P."
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
         Index           =   89
         Left            =   480
         TabIndex        =   607
         Top             =   3240
         Width           =   2655
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Decimales"
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
         TabIndex        =   469
         Top             =   4080
         Width           =   870
      End
      Begin VB.Label lblTituloEst 
         Caption         =   "Etiquetas estanterias"
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
         TabIndex        =   468
         Top             =   360
         Width           =   5895
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   74
         Left            =   1515
         Picture         =   "frmListado.frx":44BB
         ToolTipText     =   "Buscar familia"
         Top             =   1680
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   73
         Left            =   1515
         Picture         =   "frmListado.frx":45BD
         ToolTipText     =   "Buscar familia"
         Top             =   1320
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
         Index           =   56
         Left            =   480
         TabIndex        =   467
         Top             =   1080
         Width           =   600
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   78
         Left            =   960
         TabIndex        =   466
         Top             =   1680
         Width           =   420
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   77
         Left            =   960
         TabIndex        =   465
         Top             =   1320
         Width           =   465
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   72
         Left            =   1515
         Picture         =   "frmListado.frx":46BF
         ToolTipText     =   "Buscar art�culo"
         Top             =   2760
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   71
         Left            =   1515
         Picture         =   "frmListado.frx":47C1
         ToolTipText     =   "Buscar art�culo"
         Top             =   2400
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
         Index           =   55
         Left            =   480
         TabIndex        =   464
         Top             =   2160
         Width           =   660
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   76
         Left            =   960
         TabIndex        =   463
         Top             =   2760
         Width           =   420
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   75
         Left            =   960
         TabIndex        =   462
         Top             =   2400
         Width           =   465
      End
   End
   Begin VB.Frame FrameFichasMan2 
      Height          =   5295
      Left            =   0
      TabIndex        =   256
      Top             =   0
      Width           =   7395
      Begin VB.CheckBox chkMante 
         Caption         =   "Informe completo"
         Height          =   255
         Index           =   4
         Left            =   2280
         TabIndex        =   648
         Top             =   4800
         Width           =   1815
      End
      Begin VB.CheckBox chkMante 
         Caption         =   "Imprimir art�culos"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   552
         Top             =   4800
         Width           =   1815
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   108
         Left            =   2520
         MaxLength       =   4
         TabIndex        =   131
         Top             =   3240
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   108
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   550
         Text            =   "Text5"
         Top             =   3240
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   106
         Left            =   2520
         MaxLength       =   4
         TabIndex        =   130
         Top             =   2880
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   106
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   547
         Text            =   "Text5"
         Top             =   2880
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   59
         Left            =   2520
         TabIndex        =   132
         Top             =   3840
         Width           =   1335
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   60
         Left            =   2520
         TabIndex        =   133
         Top             =   4200
         Width           =   1335
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   55
         Left            =   2520
         TabIndex        =   126
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   55
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   260
         Text            =   "Text5"
         Top             =   960
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   58
         Left            =   2520
         TabIndex        =   129
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   57
         Left            =   2520
         TabIndex        =   128
         Top             =   1920
         Width           =   735
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   10
         Left            =   6120
         TabIndex        =   136
         Top             =   4680
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptarFichas 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4800
         TabIndex        =   135
         Top             =   4680
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   56
         Left            =   2520
         TabIndex        =   127
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   56
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   259
         Text            =   "Text5"
         Top             =   1320
         Width           =   3735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   58
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   258
         Text            =   "Text5"
         Top             =   2280
         Width           =   2895
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   57
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   257
         Text            =   "Text5"
         Top             =   1920
         Width           =   2895
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   61
         Left            =   5880
         TabIndex        =   134
         Top             =   3840
         Width           =   855
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
         Index           =   78
         Left            =   1680
         TabIndex        =   551
         Top             =   3240
         Width           =   420
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   88
         Left            =   2280
         Picture         =   "frmListado.frx":48C3
         ToolTipText     =   "Buscar ruta"
         Top             =   3240
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
         Index           =   76
         Left            =   1680
         TabIndex        =   549
         Top             =   2925
         Width           =   450
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
         Index           =   74
         Left            =   360
         TabIndex        =   548
         Top             =   2880
         Width           =   405
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   86
         Left            =   2280
         Picture         =   "frmListado.frx":49C5
         ToolTipText     =   "Buscar ruta"
         Top             =   2902
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   39
         Left            =   2280
         Picture         =   "frmListado.frx":4AC7
         ToolTipText     =   "Buscar contrato"
         Top             =   3840
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   47
         Left            =   1680
         TabIndex        =   271
         Top             =   3840
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   46
         Left            =   1680
         TabIndex        =   270
         Top             =   4200
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "N� Contrato"
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
         TabIndex        =   269
         Top             =   3720
         Width           =   990
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   40
         Left            =   2280
         Picture         =   "frmListado.frx":4BC9
         ToolTipText     =   "Buscar contrato"
         Top             =   4200
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
         Index           =   34
         Left            =   1680
         TabIndex        =   268
         Top             =   1320
         Width           =   420
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
         Index           =   33
         Left            =   240
         TabIndex        =   267
         Top             =   960
         Width           =   585
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   35
         Left            =   2160
         Picture         =   "frmListado.frx":4CCB
         ToolTipText     =   "Buscar cliente"
         Top             =   960
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   38
         Left            =   2235
         Picture         =   "frmListado.frx":4DCD
         ToolTipText     =   "Buscar tipo contrato"
         Top             =   2280
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   37
         Left            =   2235
         Picture         =   "frmListado.frx":4ECF
         ToolTipText     =   "Buscar tipo contrato"
         Top             =   1920
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Contrato"
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
         Left            =   240
         TabIndex        =   266
         Top             =   1920
         Width           =   1155
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   53
         Left            =   1680
         TabIndex        =   265
         Top             =   2280
         Width           =   420
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   52
         Left            =   1680
         TabIndex        =   264
         Top             =   1920
         Width           =   465
      End
      Begin VB.Label Label8 
         Caption         =   "Informe Fichas de Mantenimientos"
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
         TabIndex        =   263
         Top             =   360
         Width           =   5775
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   49
         Left            =   1680
         TabIndex        =   262
         Top             =   960
         Width           =   465
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   36
         Left            =   2160
         Picture         =   "frmListado.frx":4FD1
         ToolTipText     =   "Buscar cliente"
         Top             =   1320
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Ejercicio"
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
         Height          =   255
         Index           =   0
         Left            =   4800
         TabIndex        =   261
         Top             =   3840
         Width           =   735
      End
   End
   Begin VB.Frame FrameRepxDia 
      Height          =   5415
      Left            =   120
      TabIndex        =   180
      Top             =   0
      Width           =   6075
      Begin VB.Frame FrameCliRepDia 
         BorderStyle     =   0  'None
         Caption         =   "Frame8"
         Height          =   1215
         Left            =   240
         TabIndex        =   659
         Top             =   600
         Width           =   5775
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   133
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   664
            Text            =   "Text5"
            Top             =   720
            Width           =   3735
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   133
            Left            =   1080
            TabIndex        =   666
            Top             =   720
            Width           =   735
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   132
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   662
            Text            =   "Text5"
            Top             =   360
            Width           =   3735
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   132
            Left            =   1080
            TabIndex        =   665
            Top             =   360
            Width           =   735
         End
         Begin VB.Image imgBuscarG 
            Height          =   240
            Index           =   102
            Left            =   840
            Picture         =   "frmListado.frx":50D3
            ToolTipText     =   "Buscar cliente"
            Top             =   720
            Width           =   240
         End
         Begin VB.Label Label3 
            Caption         =   "Hasta"
            Height          =   195
            Index           =   104
            Left            =   240
            TabIndex        =   663
            Top             =   720
            Width           =   465
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
            Index           =   96
            Left            =   240
            TabIndex        =   661
            Top             =   0
            Width           =   585
         End
         Begin VB.Image imgBuscarG 
            Height          =   240
            Index           =   101
            Left            =   840
            Picture         =   "frmListado.frx":51D5
            ToolTipText     =   "Buscar cliente"
            Top             =   360
            Width           =   240
         End
         Begin VB.Label Label3 
            Caption         =   "Desde"
            Height          =   195
            Index           =   103
            Left            =   240
            TabIndex        =   660
            Top             =   360
            Width           =   465
         End
      End
      Begin VB.Frame FrameProgress 
         Height          =   1200
         Left            =   360
         TabIndex        =   351
         Top             =   4080
         Visible         =   0   'False
         Width           =   4935
         Begin MSComctlLib.ProgressBar ProgressBarContab 
            Height          =   400
            Left            =   120
            TabIndex        =   353
            Top             =   640
            Width           =   4575
            _ExtentX        =   8070
            _ExtentY        =   714
            _Version        =   393216
            Appearance      =   1
         End
         Begin VB.Label lblProgess2 
            Caption         =   "Comprobaciones:"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   354
            Top             =   135
            Width           =   4455
         End
         Begin VB.Label lblProgess2 
            Caption         =   "Iniciando el proceso ..."
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   352
            Top             =   375
            Width           =   4575
         End
      End
      Begin VB.Frame FrameTipMov 
         BorderStyle     =   0  'None
         Caption         =   "N� Factura"
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
         Height          =   990
         Left            =   360
         TabIndex        =   602
         Top             =   2560
         Width           =   4815
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   122
            Left            =   3555
            TabIndex        =   179
            Top             =   440
            Width           =   1040
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   121
            Left            =   2360
            TabIndex        =   178
            Top             =   440
            Width           =   1040
         End
         Begin VB.ComboBox cboTipMov 
            Height          =   315
            ItemData        =   "frmListado.frx":52D7
            Left            =   110
            List            =   "frmListado.frx":52D9
            Style           =   2  'Dropdown List
            TabIndex        =   177
            Top             =   440
            Width           =   2060
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "N� Factura: "
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
            TabIndex        =   606
            Top             =   0
            Width           =   1095
         End
         Begin VB.Label Label3 
            Caption         =   "Tip. Mov."
            Height          =   195
            Index           =   95
            Left            =   110
            TabIndex        =   605
            Top             =   240
            Width           =   780
         End
         Begin VB.Label Label3 
            Caption         =   "Hasta"
            Height          =   195
            Index           =   94
            Left            =   3555
            TabIndex        =   604
            Top             =   240
            Width           =   420
         End
         Begin VB.Label Label2 
            Caption         =   "Desde"
            Height          =   195
            Index           =   5
            Left            =   2360
            TabIndex        =   603
            Top             =   240
            Width           =   555
         End
      End
      Begin VB.CommandButton cmdAceptarRepxDia 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   2640
         TabIndex        =   181
         Top             =   3600
         Width           =   975
      End
      Begin VB.Frame FrameContab 
         Caption         =   " Facturas "
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
         Height          =   620
         Left            =   480
         TabIndex        =   350
         Top             =   960
         Width           =   4455
         Begin VB.OptionButton OptProve 
            Caption         =   "Proveedores"
            Height          =   255
            Left            =   2280
            TabIndex        =   172
            Top             =   250
            Width           =   1695
         End
         Begin VB.OptionButton OptClientes 
            Caption         =   "Clientes"
            Height          =   255
            Left            =   600
            TabIndex        =   170
            Top             =   250
            Value           =   -1  'True
            Width           =   1215
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   120
         TabIndex        =   346
         Top             =   1680
         Width           =   5415
         Begin VB.TextBox txtCodigo 
            Height          =   285
            Index           =   31
            Left            =   1200
            TabIndex        =   174
            Top             =   480
            Width           =   1095
         End
         Begin VB.TextBox txtCodigo 
            Height          =   285
            Index           =   32
            Left            =   3660
            TabIndex        =   176
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label2 
            Caption         =   "Desde"
            Height          =   195
            Index           =   0
            Left            =   360
            TabIndex        =   349
            Top             =   480
            Width           =   465
         End
         Begin VB.Label Label3 
            Caption         =   "Hasta"
            Height          =   195
            Index           =   29
            Left            =   2840
            TabIndex        =   348
            Top             =   480
            Width           =   420
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Reparaci�n:"
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
            Left            =   360
            TabIndex        =   347
            Top             =   200
            Width           =   1665
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   4
            Left            =   915
            Picture         =   "frmListado.frx":52DB
            Top             =   480
            Width           =   240
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   5
            Left            =   3360
            Picture         =   "frmListado.frx":5366
            Top             =   480
            Width           =   240
         End
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   7
         Left            =   3840
         TabIndex        =   182
         Top             =   3600
         Width           =   975
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Reparaciones por D�a"
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
         Left            =   480
         TabIndex        =   183
         Top             =   240
         Width           =   5055
      End
   End
   Begin VB.Frame FrameConta1FRAPRO 
      Height          =   3135
      Left            =   3600
      TabIndex        =   652
      Top             =   0
      Visible         =   0   'False
      Width           =   6135
      Begin MSComctlLib.ProgressBar pg1 
         Height          =   405
         Left            =   600
         TabIndex        =   654
         Top             =   2520
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   714
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lblProvCon 
         Caption         =   "Comprobaciones:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   360
         TabIndex        =   658
         Top             =   1320
         Width           =   5415
      End
      Begin VB.Label lblProvCon 
         Caption         =   "Comprobaciones:"
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   657
         Top             =   2160
         Width           =   4455
      End
      Begin VB.Label lblProvCon 
         Caption         =   "Comprobaciones:"
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   656
         Top             =   1800
         Width           =   4455
      End
      Begin VB.Label lblProvCon 
         Caption         =   "Comprobaciones:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   360
         TabIndex        =   655
         Top             =   960
         Width           =   5295
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Contabilizar factura proveedor"
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
         Left            =   600
         TabIndex        =   653
         Top             =   360
         Width           =   5055
      End
   End
   Begin VB.Frame FrameEnvioMail 
      Height          =   1215
      Left            =   120
      TabIndex        =   559
      Top             =   0
      Visible         =   0   'False
      Width           =   6615
      Begin MSComctlLib.ProgressBar PBMail 
         Height          =   375
         Left            =   360
         TabIndex        =   560
         Top             =   360
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Preparando datos envio"
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
         Index           =   22
         Left            =   360
         TabIndex        =   561
         Top             =   840
         Width           =   5805
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   10680
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameListMant2 
      Height          =   4215
      Left            =   1080
      TabIndex        =   523
      Top             =   0
      Width           =   6495
      Begin VB.CheckBox chkMante 
         Caption         =   "Imprimir art�culos"
         Height          =   255
         Index           =   0
         Left            =   840
         TabIndex        =   541
         Top             =   3120
         Width           =   1815
      End
      Begin VB.CommandButton cmdManteTeorico 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3840
         TabIndex        =   540
         Top             =   3600
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   77
         Left            =   5040
         TabIndex        =   539
         Top             =   3600
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   105
         Left            =   1680
         TabIndex        =   536
         Top             =   2640
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   105
         Left            =   2460
         Locked          =   -1  'True
         TabIndex        =   535
         Text            =   "Text5"
         Top             =   2640
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   104
         Left            =   1680
         TabIndex        =   533
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   104
         Left            =   2460
         Locked          =   -1  'True
         TabIndex        =   532
         Text            =   "Text5"
         Top             =   2280
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   103
         Left            =   1680
         TabIndex        =   529
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   103
         Left            =   2460
         Locked          =   -1  'True
         TabIndex        =   528
         Text            =   "Text5"
         Top             =   1560
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   102
         Left            =   1680
         TabIndex        =   526
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   102
         Left            =   2460
         Locked          =   -1  'True
         TabIndex        =   525
         Text            =   "Text5"
         Top             =   1200
         Width           =   3735
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Contrato"
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
         Index           =   73
         Left            =   240
         TabIndex        =   538
         Top             =   2040
         Width           =   1155
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   83
         Left            =   1395
         Picture         =   "frmListado.frx":53F1
         ToolTipText     =   "Buscar tipo contrato"
         Top             =   2640
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   87
         Left            =   840
         TabIndex        =   537
         Top             =   2640
         Width           =   465
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   82
         Left            =   1395
         Picture         =   "frmListado.frx":54F3
         ToolTipText     =   "Buscar tipo contrato"
         Top             =   2280
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   86
         Left            =   840
         TabIndex        =   534
         Top             =   2280
         Width           =   465
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
         Index           =   72
         Left            =   240
         TabIndex        =   531
         Top             =   960
         Width           =   585
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   81
         Left            =   1395
         Picture         =   "frmListado.frx":55F5
         ToolTipText     =   "Buscar cliente"
         Top             =   1560
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   85
         Left            =   840
         TabIndex        =   530
         Top             =   1560
         Width           =   465
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   80
         Left            =   1395
         Picture         =   "frmListado.frx":56F7
         ToolTipText     =   "Buscar cliente"
         Top             =   1200
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   84
         Left            =   840
         TabIndex        =   527
         Top             =   1200
         Width           =   465
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Informe te�rico de mantenimientos"
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
         TabIndex        =   524
         Top             =   480
         Width           =   5100
      End
   End
   Begin VB.Frame FrameRepSustNSerie 
      Height          =   3735
      Left            =   240
      TabIndex        =   389
      Top             =   0
      Width           =   5715
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   81
         Left            =   1080
         MaxLength       =   15
         TabIndex        =   390
         Top             =   2160
         Width           =   1935
      End
      Begin VB.CommandButton cmdAceptarSustNSerie 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   0
         Left            =   2040
         TabIndex        =   391
         Top             =   3000
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   13
         Left            =   3120
         TabIndex        =   392
         Top             =   3000
         Width           =   975
      End
      Begin VB.Label lblNumSerie 
         Caption         =   "num serie"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   406
         Top             =   1560
         Width           =   4215
      End
      Begin VB.Label lblNumSerie 
         Caption         =   "num serie"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   1080
         TabIndex        =   396
         Top             =   1320
         Width           =   4215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Introduce el nuevo N� de Serie que va a sustituir al: "
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
         Left            =   360
         TabIndex        =   395
         Top             =   1000
         Width           =   3780
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Sustituci�n N� de Serie"
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
         TabIndex        =   394
         Top             =   480
         Width           =   5055
      End
      Begin VB.Label Label2 
         Caption         =   "N� Serie"
         Height          =   195
         Index           =   3
         Left            =   360
         TabIndex        =   393
         Top             =   2160
         Width           =   705
      End
   End
   Begin VB.Frame FrameInfAlmacen 
      Height          =   3495
      Left            =   1560
      TabIndex        =   30
      Top             =   1080
      Width           =   5835
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   2
         Left            =   3720
         TabIndex        =   9
         Top             =   2880
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   2
         Left            =   2640
         TabIndex        =   8
         Top             =   2880
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   3
         Left            =   1200
         TabIndex        =   6
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   4
         Left            =   3480
         TabIndex        =   7
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   6
         Left            =   360
         TabIndex        =   34
         Top             =   1800
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   3
         Left            =   2640
         TabIndex        =   33
         Top             =   1800
         Width           =   420
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Informes Almacenes"
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
         TabIndex        =   32
         Top             =   600
         Width           =   5055
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "N� Traspaso"
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
         Left            =   360
         TabIndex        =   31
         Top             =   1320
         Width           =   1020
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   920
         Picture         =   "frmListado.frx":57F9
         ToolTipText     =   "Buscar almac�n"
         Top             =   1800
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   4
         Left            =   3200
         Picture         =   "frmListado.frx":58FB
         ToolTipText     =   "Buscar almac�n"
         Top             =   1800
         Width           =   240
      End
   End
   Begin VB.Frame FrEliminarFacturas 
      Height          =   4215
      Left            =   120
      TabIndex        =   517
      Top             =   0
      Width           =   6375
      Begin VB.CommandButton cmdElimiaFacturas 
         Caption         =   "Eliminar"
         Height          =   375
         Left            =   3840
         TabIndex        =   521
         Top             =   3600
         Width           =   975
      End
      Begin VB.ComboBox cmbEliFac 
         Height          =   315
         Left            =   3360
         Style           =   2  'Dropdown List
         TabIndex        =   520
         Top             =   3000
         Width           =   2655
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   97
         Left            =   5040
         TabIndex        =   518
         Top             =   3600
         Width           =   975
      End
      Begin VB.Label Label11 
         Caption         =   "lore ipsum lorem ipsum lorem ipsum"
         Height          =   615
         Index           =   1
         Left            =   360
         TabIndex        =   543
         Top             =   2160
         Width           =   5775
      End
      Begin VB.Label Label11 
         Caption         =   "lore"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Index           =   0
         Left            =   360
         TabIndex        =   542
         Top             =   360
         Width           =   5775
      End
      Begin VB.Label Label3 
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
         Height          =   315
         Index           =   83
         Left            =   120
         TabIndex        =   522
         Top             =   3600
         Width           =   3525
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Eliminar facturas hasta: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   82
         Left            =   360
         TabIndex        =   519
         Top             =   3000
         Width           =   2370
      End
   End
   Begin VB.Frame FrameRepNSerie 
      Height          =   5415
      Left            =   360
      TabIndex        =   158
      Top             =   0
      Width           =   6795
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   37
         Left            =   1920
         TabIndex        =   147
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   37
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   162
         Text            =   "Text5"
         Top             =   1320
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   42
         Left            =   1920
         TabIndex        =   152
         Top             =   3720
         Width           =   1335
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   41
         Left            =   1920
         TabIndex        =   151
         Top             =   3360
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   6
         Left            =   4560
         TabIndex        =   154
         Top             =   4680
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptarNSerie 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3480
         TabIndex        =   153
         Top             =   4680
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   39
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   149
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   40
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   150
         Top             =   2640
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   39
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   161
         Text            =   "Text5"
         Top             =   2280
         Width           =   3735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   40
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   160
         Text            =   "Text5"
         Top             =   2640
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   38
         Left            =   1920
         TabIndex        =   148
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   38
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   159
         Text            =   "Text5"
         Top             =   1680
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
         Index           =   20
         Left            =   1080
         TabIndex        =   175
         Top             =   1680
         Width           =   420
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
         Index           =   19
         Left            =   600
         TabIndex        =   173
         Top             =   1080
         Width           =   585
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   49
         Left            =   1635
         Picture         =   "frmListado.frx":59FD
         ToolTipText     =   "Buscar cliente"
         Top             =   1320
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   54
         Left            =   1635
         Picture         =   "frmListado.frx":5AFF
         ToolTipText     =   "Buscar contrato"
         Top             =   3720
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   53
         Left            =   1635
         Picture         =   "frmListado.frx":5C01
         ToolTipText     =   "Buscar  contrato"
         Top             =   3360
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "N� Contrato"
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
         Left            =   600
         TabIndex        =   171
         Top             =   3120
         Width           =   990
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   32
         Left            =   1080
         TabIndex        =   169
         Top             =   3720
         Width           =   420
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   31
         Left            =   1080
         TabIndex        =   168
         Top             =   3360
         Width           =   465
      End
      Begin VB.Label Label5 
         Caption         =   "Informe N� Serie"
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
         TabIndex        =   167
         Top             =   360
         Width           =   4815
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   28
         Left            =   1080
         TabIndex        =   166
         Top             =   2280
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   27
         Left            =   1080
         TabIndex        =   165
         Top             =   2640
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Direc/Dpto"
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
         Left            =   600
         TabIndex        =   164
         Top             =   2040
         Width           =   930
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   51
         Left            =   1635
         Picture         =   "frmListado.frx":5D03
         ToolTipText     =   "Buscar dir/dpto"
         Top             =   2280
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   52
         Left            =   1635
         Picture         =   "frmListado.frx":5E05
         ToolTipText     =   "Buscar dir/dpto"
         Top             =   2640
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   26
         Left            =   1080
         TabIndex        =   163
         Top             =   1320
         Width           =   465
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   50
         Left            =   1635
         Picture         =   "frmListado.frx":5F07
         ToolTipText     =   "Buscar cliente"
         Top             =   1680
         Width           =   240
      End
   End
   Begin VB.Frame FrameHcoMante 
      Height          =   3495
      Left            =   0
      TabIndex        =   562
      Top             =   -120
      Width           =   6495
      Begin VB.CommandButton cmdHcoMante 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3840
         TabIndex        =   567
         Top             =   2880
         Width           =   1095
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   112
         Left            =   1680
         MaxLength       =   4
         TabIndex        =   566
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   112
         Left            =   2460
         Locked          =   -1  'True
         TabIndex        =   572
         Text            =   "Text5"
         Top             =   2280
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   111
         Left            =   1680
         MaxLength       =   4
         TabIndex        =   565
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   111
         Left            =   2460
         Locked          =   -1  'True
         TabIndex        =   570
         Text            =   "Text5"
         Top             =   1560
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   110
         Left            =   1680
         TabIndex        =   564
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   99
         Left            =   5160
         TabIndex        =   569
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Motivo baja"
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
         Index           =   81
         Left            =   240
         TabIndex        =   573
         Top             =   2280
         Width           =   1005
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   90
         Left            =   1395
         Picture         =   "frmListado.frx":6009
         ToolTipText     =   "Buscar motivo baja"
         Top             =   2280
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
         Index           =   80
         Left            =   240
         TabIndex        =   571
         Top             =   1560
         Width           =   945
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   89
         Left            =   1395
         Picture         =   "frmListado.frx":610B
         ToolTipText     =   "Buscar trabajador"
         Top             =   1560
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   12
         Left            =   1395
         Picture         =   "frmListado.frx":620D
         Top             =   960
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha baja"
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
         Index           =   79
         Left            =   240
         TabIndex        =   568
         Top             =   960
         Width           =   915
      End
      Begin VB.Label Label7 
         Caption         =   "Paso a mantenimientos anulados"
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
         TabIndex        =   563
         Top             =   240
         Width           =   5775
      End
   End
   Begin VB.Frame FrameAlbaranesMarcaFacturar 
      Height          =   3735
      Left            =   0
      TabIndex        =   583
      Top             =   0
      Width           =   6495
      Begin VB.Frame Frame7 
         BorderStyle     =   0  'None
         Height          =   1215
         Left            =   120
         TabIndex        =   599
         Top             =   1680
         Width           =   6135
      End
      Begin VB.CommandButton cmdFactAlbaranes 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3840
         TabIndex        =   589
         Top             =   3120
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   82
         Left            =   5160
         TabIndex        =   590
         Top             =   3120
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   120
         Left            =   3960
         TabIndex        =   586
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   119
         Left            =   1680
         TabIndex        =   585
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   118
         Left            =   1680
         TabIndex        =   588
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   118
         Left            =   2460
         Locked          =   -1  'True
         TabIndex        =   592
         Text            =   "Text5"
         Top             =   2520
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   117
         Left            =   1680
         TabIndex        =   587
         Top             =   2160
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   117
         Left            =   2460
         Locked          =   -1  'True
         TabIndex        =   591
         Text            =   "Text5"
         Top             =   2160
         Width           =   3735
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   16
         Left            =   3600
         Picture         =   "frmListado.frx":6298
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
         Index           =   87
         Left            =   3000
         TabIndex        =   598
         Top             =   1440
         Width           =   420
      End
      Begin VB.Label Label4 
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
         Index           =   86
         Left            =   240
         TabIndex        =   597
         Top             =   960
         Width           =   1185
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   93
         Left            =   720
         TabIndex        =   596
         Top             =   1440
         Width           =   465
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   15
         Left            =   1320
         Picture         =   "frmListado.frx":6323
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
         Index           =   85
         Left            =   840
         TabIndex        =   595
         Top             =   2520
         Width           =   420
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
         Index           =   84
         Left            =   360
         TabIndex        =   594
         Top             =   1920
         Width           =   585
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   94
         Left            =   1395
         Picture         =   "frmListado.frx":63AE
         Top             =   2520
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   92
         Left            =   840
         TabIndex        =   593
         Top             =   2160
         Width           =   465
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   93
         Left            =   1395
         Picture         =   "frmListado.frx":64B0
         Top             =   2160
         Width           =   240
      End
      Begin VB.Label Label7 
         Caption         =   "Marcar facturar albaranes"
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
         Left            =   360
         TabIndex        =   584
         Top             =   360
         Width           =   5775
      End
   End
   Begin VB.Frame FrameRepxClien 
      Height          =   5415
      Left            =   240
      TabIndex        =   186
      Top             =   240
      Width           =   6795
      Begin VB.Frame Frame1 
         Height          =   855
         Left            =   3720
         TabIndex        =   343
         Top             =   3240
         Width           =   2415
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   120
            MaxLength       =   4
            TabIndex        =   193
            Text            =   "1"
            Top             =   420
            Width           =   855
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "reparaciones"
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
            Index           =   43
            Left            =   1200
            TabIndex        =   345
            Top             =   420
            Width           =   930
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Mostrar equipos con m�s de:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   195
            Index           =   41
            Left            =   120
            TabIndex        =   344
            Top             =   120
            Width           =   2070
         End
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   34
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   199
         Text            =   "Text5"
         Top             =   1680
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   34
         Left            =   1920
         TabIndex        =   188
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   36
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   198
         Text            =   "Text5"
         Top             =   2640
         Width           =   3735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   35
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   197
         Text            =   "Text5"
         Top             =   2280
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   36
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   190
         Top             =   2640
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   35
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   189
         Top             =   2280
         Width           =   735
      End
      Begin VB.CommandButton cmdAceptarRepxClien 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3960
         TabIndex        =   194
         Top             =   4680
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   8
         Left            =   5040
         TabIndex        =   195
         Top             =   4680
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   43
         Left            =   1920
         TabIndex        =   191
         Top             =   3360
         Width           =   1215
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   44
         Left            =   1920
         TabIndex        =   192
         Top             =   3720
         Width           =   1215
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   33
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   196
         Text            =   "Text5"
         Top             =   1320
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   33
         Left            =   1920
         TabIndex        =   187
         Top             =   1320
         Width           =   735
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   7
         Left            =   1640
         Picture         =   "frmListado.frx":65B2
         Top             =   3720
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   6
         Left            =   1640
         Picture         =   "frmListado.frx":663D
         Top             =   3360
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   7
         Left            =   1635
         Picture         =   "frmListado.frx":66C8
         ToolTipText     =   "Buscar cliente"
         Top             =   1680
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   37
         Left            =   1080
         TabIndex        =   209
         Top             =   1320
         Width           =   465
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   9
         Left            =   1635
         Picture         =   "frmListado.frx":67CA
         ToolTipText     =   "Buscar dir/dpto"
         Top             =   2640
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   8
         Left            =   1635
         Picture         =   "frmListado.frx":68CC
         ToolTipText     =   "buscar dir/dpto"
         Top             =   2280
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Direc/Dpto"
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
         Left            =   600
         TabIndex        =   208
         Top             =   2040
         Width           =   930
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   36
         Left            =   1080
         TabIndex        =   207
         Top             =   2640
         Width           =   420
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   35
         Left            =   1080
         TabIndex        =   206
         Top             =   2280
         Width           =   465
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Reparaciones  por Cliente"
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
         TabIndex        =   205
         Top             =   360
         Width           =   4815
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   34
         Left            =   1080
         TabIndex        =   204
         Top             =   3360
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   30
         Left            =   1080
         TabIndex        =   203
         Top             =   3720
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha albar�n"
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
         Left            =   600
         TabIndex        =   202
         Top             =   3120
         Width           =   1185
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   6
         Left            =   1635
         Picture         =   "frmListado.frx":69CE
         ToolTipText     =   "Buscar cliente"
         Top             =   1320
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
         Index           =   18
         Left            =   600
         TabIndex        =   201
         Top             =   1080
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
         Index           =   16
         Left            =   1080
         TabIndex        =   200
         Top             =   1680
         Width           =   420
      End
   End
   Begin VB.Frame FrameFrecuencia 
      Height          =   3855
      Left            =   120
      TabIndex        =   495
      Top             =   0
      Width           =   6015
      Begin VB.Frame Frame13 
         Height          =   615
         Left            =   360
         TabIndex        =   645
         Top             =   2880
         Width           =   2655
         Begin VB.OptionButton OptFrecFicha 
            Caption         =   "Ficha"
            Height          =   255
            Left            =   120
            TabIndex        =   647
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton OptFrecResumen 
            Caption         =   "Resumen"
            Height          =   255
            Left            =   1320
            TabIndex        =   646
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   99
         Left            =   2100
         Locked          =   -1  'True
         TabIndex        =   506
         Text            =   "Text5"
         Top             =   1440
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   99
         Left            =   1320
         TabIndex        =   498
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   101
         Left            =   2100
         Locked          =   -1  'True
         TabIndex        =   505
         Text            =   "Text5"
         Top             =   2400
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   100
         Left            =   2100
         Locked          =   -1  'True
         TabIndex        =   504
         Text            =   "Text5"
         Top             =   2040
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   101
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   500
         Top             =   2400
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   100
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   499
         Top             =   2040
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   98
         Left            =   2100
         Locked          =   -1  'True
         TabIndex        =   503
         Text            =   "Text5"
         Top             =   1080
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   98
         Left            =   1320
         TabIndex        =   497
         Top             =   1080
         Width           =   735
      End
      Begin VB.CommandButton cmdFrecuencias 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3600
         TabIndex        =   501
         Top             =   3120
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   96
         Left            =   4800
         TabIndex        =   502
         Top             =   3120
         Width           =   975
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   77
         Left            =   1035
         Picture         =   "frmListado.frx":6AD0
         ToolTipText     =   "Buscar cliente"
         Top             =   1440
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   81
         Left            =   480
         TabIndex        =   512
         Top             =   1080
         Width           =   465
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   79
         Left            =   1035
         Picture         =   "frmListado.frx":6BD2
         ToolTipText     =   "Buscar dir/dpto"
         Top             =   2400
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   78
         Left            =   1035
         Picture         =   "frmListado.frx":6CD4
         ToolTipText     =   "Buscar dir/dpto"
         Top             =   2040
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Direc/Dpto"
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
         Left            =   120
         TabIndex        =   511
         Top             =   1800
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   80
         Left            =   480
         TabIndex        =   510
         Top             =   2400
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   79
         Left            =   480
         TabIndex        =   509
         Top             =   2040
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   76
         Left            =   1035
         Picture         =   "frmListado.frx":6DD6
         ToolTipText     =   "Buscar cliente"
         Top             =   1080
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
         Index           =   66
         Left            =   120
         TabIndex        =   508
         Top             =   720
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
         Index           =   65
         Left            =   480
         TabIndex        =   507
         Top             =   1440
         Width           =   420
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Datos de frecuencias  clientes"
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
         Left            =   480
         TabIndex        =   496
         Top             =   240
         Width           =   5055
      End
   End
   Begin VB.Frame FrameListAvisosPtes 
      Height          =   4815
      Left            =   0
      TabIndex        =   407
      Top             =   0
      Width           =   6315
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   97
         Left            =   1740
         MaxLength       =   4
         TabIndex        =   402
         Top             =   3480
         Width           =   615
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   97
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   473
         Text            =   "Text5"
         Top             =   3480
         Width           =   3375
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   96
         Left            =   1740
         MaxLength       =   4
         TabIndex        =   401
         Top             =   3120
         Width           =   615
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   96
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   470
         Text            =   "Text5"
         Top             =   3120
         Width           =   3375
      End
      Begin VB.ComboBox cboSituaAviso 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   403
         Top             =   4200
         Width           =   1695
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   82
         Left            =   1740
         MaxLength       =   10
         TabIndex        =   397
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton cmdAceptarAviPtes 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3720
         TabIndex        =   404
         Top             =   4200
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   14
         Left            =   4800
         TabIndex        =   405
         Top             =   4200
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   83
         Left            =   4260
         MaxLength       =   10
         TabIndex        =   398
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   84
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   409
         Text            =   "Text5"
         Top             =   1920
         Width           =   3375
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   84
         Left            =   1740
         MaxLength       =   4
         TabIndex        =   399
         Top             =   1920
         Width           =   615
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   85
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   408
         Text            =   "Text5"
         Top             =   2280
         Width           =   3375
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   85
         Left            =   1740
         MaxLength       =   4
         TabIndex        =   400
         Top             =   2280
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
         Index           =   60
         Left            =   960
         TabIndex        =   474
         Top             =   3480
         Width           =   420
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   97
         Left            =   1440
         Picture         =   "frmListado.frx":6ED8
         ToolTipText     =   "Buscar tecnico"
         Top             =   3480
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "T�cnico"
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
         Left            =   600
         TabIndex        =   472
         Top             =   2880
         Width           =   645
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
         Index           =   58
         Left            =   960
         TabIndex        =   471
         Top             =   3120
         Width           =   450
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   96
         Left            =   1440
         Picture         =   "frmListado.frx":6FDA
         ToolTipText     =   "Buscar tecnico"
         Top             =   3120
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Situaci�n"
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
         Left            =   600
         TabIndex        =   417
         Top             =   4200
         Width           =   780
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
         Index           =   51
         Left            =   3480
         TabIndex        =   416
         Top             =   1080
         Width           =   420
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   10
         Left            =   1440
         Picture         =   "frmListado.frx":70DC
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Avisos de aver�a pendientes"
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
         Left            =   1080
         TabIndex        =   415
         Top             =   240
         Width           =   4455
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha aviso"
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
         Left            =   600
         TabIndex        =   414
         Top             =   840
         Width           =   990
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
         Index           =   49
         Left            =   960
         TabIndex        =   413
         Top             =   1080
         Width           =   450
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   11
         Left            =   3960
         Picture         =   "frmListado.frx":7167
         Top             =   1080
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   84
         Left            =   1440
         Picture         =   "frmListado.frx":71F2
         ToolTipText     =   "Buscar ruta"
         Top             =   1920
         Width           =   240
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
         Index           =   48
         Left            =   600
         TabIndex        =   412
         Top             =   1680
         Width           =   405
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
         Index           =   47
         Left            =   960
         TabIndex        =   411
         Top             =   1920
         Width           =   450
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   85
         Left            =   1440
         Picture         =   "frmListado.frx":72F4
         ToolTipText     =   "Buscar ruta"
         Top             =   2280
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
         Left            =   960
         TabIndex        =   410
         Top             =   2280
         Width           =   420
      End
   End
   Begin VB.Frame FrameMovArtic 
      Height          =   5535
      Left            =   0
      TabIndex        =   35
      Top             =   0
      Width           =   10635
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   86
         Left            =   1485
         TabIndex        =   25
         Top             =   4560
         Width           =   855
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   87
         Left            =   1485
         TabIndex        =   26
         Top             =   4920
         Width           =   855
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   5
         Left            =   1440
         MaxLength       =   16
         TabIndex        =   17
         Top             =   1080
         Width           =   1455
      End
      Begin VB.CommandButton cmdDeselTodos 
         Height          =   435
         Left            =   9000
         Picture         =   "frmListado.frx":73F6
         Style           =   1  'Graphical
         TabIndex        =   185
         Top             =   740
         Width           =   585
      End
      Begin VB.CommandButton cmdSelTodos 
         Height          =   435
         Left            =   9720
         Picture         =   "frmListado.frx":7AE0
         Style           =   1  'Graphical
         TabIndex        =   184
         Top             =   740
         Width           =   585
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3375
         Left            =   6960
         TabIndex        =   27
         Top             =   1200
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   5953
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
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
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   12
         Left            =   2120
         Locked          =   -1  'True
         TabIndex        =   72
         Text            =   "Text5"
         Top             =   3960
         Width           =   2775
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   11
         Left            =   2120
         Locked          =   -1  'True
         TabIndex        =   71
         Text            =   "Text5"
         Top             =   3600
         Width           =   2775
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   8
         Left            =   2120
         Locked          =   -1  'True
         TabIndex        =   70
         Text            =   "Text5"
         Top             =   2400
         Width           =   2775
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   7
         Left            =   2120
         Locked          =   -1  'True
         TabIndex        =   69
         Text            =   "Text5"
         Top             =   2040
         Width           =   2775
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   6
         Left            =   2940
         Locked          =   -1  'True
         TabIndex        =   68
         Text            =   "Text5"
         Top             =   1440
         Width           =   3615
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   5
         Left            =   2940
         Locked          =   -1  'True
         TabIndex        =   67
         Text            =   "Text5"
         Top             =   1080
         Width           =   3615
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   12
         Left            =   1440
         TabIndex        =   24
         Top             =   3960
         Width           =   615
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   11
         Left            =   1440
         TabIndex        =   23
         Top             =   3600
         Width           =   615
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   10
         Left            =   3600
         TabIndex        =   22
         Top             =   3000
         Width           =   1215
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   9
         Left            =   1440
         TabIndex        =   21
         Top             =   3000
         Width           =   1215
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   8
         Left            =   1440
         MaxLength       =   4
         TabIndex        =   20
         Top             =   2400
         Width           =   615
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   7
         Left            =   1440
         MaxLength       =   4
         TabIndex        =   19
         Top             =   2040
         Width           =   615
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   6
         Left            =   1440
         MaxLength       =   16
         TabIndex        =   18
         Top             =   1440
         Width           =   1455
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   3
         Left            =   4440
         TabIndex        =   28
         Top             =   4800
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   3
         Left            =   5520
         TabIndex        =   29
         Top             =   4800
         Width           =   975
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   65
         Left            =   1200
         Picture         =   "frmListado.frx":81CA
         ToolTipText     =   "Cliente"
         Top             =   4560
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   66
         Left            =   1200
         Picture         =   "frmListado.frx":82CC
         Top             =   4920
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   70
         Left            =   600
         TabIndex        =   420
         Top             =   4560
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   69
         Left            =   600
         TabIndex        =   419
         Top             =   4920
         Width           =   420
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Cliente/Proveedor"
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
         Left            =   360
         TabIndex        =   418
         Top             =   4320
         Width           =   1560
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tipos de Movimiento"
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
         Left            =   6960
         TabIndex        =   66
         Top             =   960
         Width           =   1755
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   1
         Left            =   3315
         Picture         =   "frmListado.frx":83CE
         Top             =   3000
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   1155
         Picture         =   "frmListado.frx":8459
         Top             =   3000
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   34
         Left            =   1155
         Picture         =   "frmListado.frx":84E4
         ToolTipText     =   "Buscar almacen"
         Top             =   3960
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   33
         Left            =   1155
         Picture         =   "frmListado.frx":85E6
         ToolTipText     =   "Buscar almacen"
         Top             =   3600
         Width           =   240
      End
      Begin VB.Label Label3 
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
         Index           =   11
         Left            =   360
         TabIndex        =   65
         Top             =   3360
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   8
         Left            =   600
         TabIndex        =   64
         Top             =   3960
         Width           =   420
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   7
         Left            =   600
         TabIndex        =   63
         Top             =   3600
         Width           =   465
      End
      Begin VB.Label Label3 
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
         TabIndex        =   62
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   6
         Left            =   2760
         TabIndex        =   61
         Top             =   3000
         Width           =   420
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   5
         Left            =   600
         TabIndex        =   43
         Top             =   3000
         Width           =   465
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   32
         Left            =   1155
         Picture         =   "frmListado.frx":86E8
         ToolTipText     =   "Buscar familia"
         Top             =   2400
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   31
         Left            =   1155
         Picture         =   "frmListado.frx":87EA
         ToolTipText     =   "Buscar familia"
         Top             =   2040
         Width           =   240
      End
      Begin VB.Label Label3 
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
         Index           =   9
         Left            =   360
         TabIndex        =   42
         Top             =   1800
         Width           =   600
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   4
         Left            =   600
         TabIndex        =   41
         Top             =   2400
         Width           =   420
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   0
         Left            =   600
         TabIndex        =   40
         Top             =   2040
         Width           =   465
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   30
         Left            =   1155
         Picture         =   "frmListado.frx":88EC
         ToolTipText     =   "Buscar art�culo"
         Top             =   1440
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   29
         Left            =   1155
         Picture         =   "frmListado.frx":89EE
         ToolTipText     =   "Buscar art�culo"
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label Label3 
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
         Index           =   1
         Left            =   360
         TabIndex        =   39
         Top             =   840
         Width           =   660
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Informes Movimiento Art�culos"
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
         Left            =   240
         TabIndex        =   38
         Top             =   360
         Width           =   5055
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   33
         Left            =   600
         TabIndex        =   37
         Top             =   1440
         Width           =   420
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   2
         Left            =   600
         TabIndex        =   36
         Top             =   1080
         Width           =   465
      End
   End
   Begin VB.Frame frameListado 
      Height          =   4695
      Left            =   360
      TabIndex        =   10
      Top             =   120
      Width           =   6555
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   1605
         TabIndex        =   0
         Top             =   1560
         Width           =   830
      End
      Begin VB.Frame frameOrdenar 
         Caption         =   "Ordenar por"
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
         Height          =   735
         Left            =   720
         TabIndex        =   157
         Top             =   2640
         Width           =   3375
         Begin VB.OptionButton OptNombre 
            Caption         =   "Descripci�n"
            Height          =   375
            Left            =   1680
            TabIndex        =   3
            Top             =   240
            Width           =   1455
         End
         Begin VB.OptionButton Optcodigo 
            Caption         =   "C�digo"
            Height          =   375
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   2
         Left            =   2480
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "Text5"
         Top             =   2040
         Width           =   3255
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   1
         Left            =   2480
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "Text5"
         Top             =   1560
         Width           =   3255
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   1605
         TabIndex        =   1
         Top             =   2040
         Width           =   830
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   1
         Left            =   3600
         TabIndex        =   4
         Top             =   3960
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   1
         Left            =   4680
         TabIndex        =   5
         Top             =   3960
         Width           =   975
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   1320
         Picture         =   "frmListado.frx":8AF0
         ToolTipText     =   "Buscar marca"
         Top             =   2040
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1320
         Picture         =   "frmListado.frx":8BF2
         ToolTipText     =   "Buscar marca"
         Top             =   1560
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "C�digo"
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
         Left            =   720
         TabIndex        =   16
         Top             =   1200
         Width           =   570
      End
      Begin VB.Label Label1 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   3
         Left            =   720
         TabIndex        =   14
         Top             =   2040
         Width           =   420
      End
      Begin VB.Label Label1 
         Caption         =   "Desde"
         Height          =   195
         Index           =   2
         Left            =   720
         TabIndex        =   13
         Top             =   1605
         Width           =   465
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Listado Marcas"
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
         Left            =   600
         TabIndex        =   15
         Top             =   480
         Width           =   5055
      End
   End
   Begin VB.Frame FrameMantenimientos 
      Height          =   7695
      Left            =   3480
      TabIndex        =   210
      Top             =   0
      Width           =   6735
      Begin VB.Frame FrameRuta 
         Height          =   1095
         Left            =   600
         TabIndex        =   682
         Top             =   4800
         Visible         =   0   'False
         Width           =   6015
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   138
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   686
            Text            =   "Text5"
            Top             =   720
            Width           =   3735
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   138
            Left            =   1320
            MaxLength       =   4
            TabIndex        =   223
            Top             =   720
            Width           =   735
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   137
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   683
            Text            =   "Text5"
            Top             =   240
            Width           =   3735
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   137
            Left            =   1320
            MaxLength       =   4
            TabIndex        =   222
            Top             =   240
            Width           =   735
         End
         Begin VB.Image imgBuscarG 
            Height          =   240
            Index           =   106
            Left            =   1080
            Picture         =   "frmListado.frx":8CF4
            ToolTipText     =   "Buscar ruta"
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
            Index           =   100
            Left            =   480
            TabIndex        =   687
            Top             =   720
            Width           =   420
         End
         Begin VB.Image imgBuscarG 
            Height          =   240
            Index           =   105
            Left            =   1080
            Picture         =   "frmListado.frx":8DF6
            ToolTipText     =   "Buscar ruta"
            Top             =   255
            Width           =   240
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
            Index           =   99
            Left            =   0
            TabIndex        =   685
            Top             =   0
            Width           =   405
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
            Index           =   98
            Left            =   480
            TabIndex        =   684
            Top             =   285
            Width           =   450
         End
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   2
         Left            =   360
         TabIndex        =   556
         Top             =   5880
         Width           =   6255
         Begin VB.CheckBox chkMante 
            Caption         =   "Copia remitente"
            Height          =   255
            Index           =   3
            Left            =   360
            TabIndex        =   229
            Top             =   360
            Width           =   1935
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   109
            Left            =   1440
            TabIndex        =   230
            Top             =   720
            Width           =   1215
         End
         Begin VB.OptionButton optMante 
            Caption         =   "Comercial"
            Height          =   195
            Index           =   1
            Left            =   4200
            TabIndex        =   228
            Top             =   0
            Width           =   1455
         End
         Begin VB.OptionButton optMante 
            Caption         =   "Administracion"
            Height          =   195
            Index           =   0
            Left            =   2520
            TabIndex        =   227
            Top             =   0
            Value           =   -1  'True
            Width           =   1455
         End
         Begin VB.CheckBox chkMante 
            Caption         =   "Enviar e-mail"
            Height          =   255
            Index           =   2
            Left            =   360
            TabIndex        =   226
            Top             =   0
            Width           =   1455
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Fecha carta"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   195
            Index           =   77
            Left            =   120
            TabIndex        =   557
            Top             =   720
            Width           =   990
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   109
            Left            =   1155
            Picture         =   "frmListado.frx":8EF8
            Top             =   720
            Width           =   240
         End
      End
      Begin VB.Frame FrameManteActi 
         BorderStyle     =   0  'None
         Height          =   1455
         Left            =   120
         TabIndex        =   638
         Top             =   4800
         Width           =   6495
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   129
            Left            =   1800
            TabIndex        =   234
            Top             =   1080
            Width           =   1095
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   127
            Left            =   2580
            Locked          =   -1  'True
            TabIndex        =   640
            Text            =   "Text5"
            Top             =   240
            Width           =   3735
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   127
            Left            =   1800
            TabIndex        =   220
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   128
            Left            =   2580
            Locked          =   -1  'True
            TabIndex        =   639
            Text            =   "Text5"
            Top             =   600
            Width           =   3735
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   128
            Left            =   1800
            TabIndex        =   221
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "C�d. Postal"
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
            Index           =   93
            Left            =   480
            TabIndex        =   644
            Top             =   1080
            Width           =   930
         End
         Begin VB.Label Label3 
            Caption         =   "Desde"
            Height          =   195
            Index           =   102
            Left            =   960
            TabIndex        =   643
            Top             =   240
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
            Index           =   94
            Left            =   525
            TabIndex        =   642
            Top             =   0
            Width           =   795
         End
         Begin VB.Image imgBuscarG 
            Height          =   240
            Index           =   99
            Left            =   1515
            Picture         =   "frmListado.frx":8F83
            ToolTipText     =   "Buscar actividad"
            Top             =   240
            Width           =   240
         End
         Begin VB.Label Label3 
            Caption         =   "Hasta"
            Height          =   195
            Index           =   101
            Left            =   960
            TabIndex        =   641
            Top             =   600
            Width           =   465
         End
         Begin VB.Image imgBuscarG 
            Height          =   240
            Index           =   100
            Left            =   1515
            Picture         =   "frmListado.frx":9085
            ToolTipText     =   "Buscar actividad"
            Top             =   600
            Width           =   240
         End
      End
      Begin VB.Frame FrameManteAnu 
         BorderStyle     =   0  'None
         Height          =   1455
         Left            =   120
         TabIndex        =   574
         Top             =   4800
         Width           =   6495
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   116
            Left            =   5040
            TabIndex        =   233
            Top             =   1080
            Width           =   1215
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   115
            Left            =   2400
            TabIndex        =   232
            Top             =   1080
            Width           =   1215
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   114
            Left            =   1800
            TabIndex        =   225
            Top             =   600
            Width           =   735
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   114
            Left            =   2580
            Locked          =   -1  'True
            TabIndex        =   578
            Text            =   "Text5"
            Top             =   600
            Width           =   3735
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   113
            Left            =   1800
            TabIndex        =   224
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   113
            Left            =   2580
            Locked          =   -1  'True
            TabIndex        =   575
            Text            =   "Text5"
            Top             =   240
            Width           =   3735
         End
         Begin VB.Label Label3 
            Caption         =   "Hasta"
            Height          =   195
            Index           =   91
            Left            =   4200
            TabIndex        =   582
            Top             =   1080
            Width           =   465
         End
         Begin VB.Label Label3 
            Caption         =   "Desde"
            Height          =   195
            Index           =   90
            Left            =   1560
            TabIndex        =   581
            Top             =   1080
            Width           =   465
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   14
            Left            =   4680
            Picture         =   "frmListado.frx":9187
            Top             =   1080
            Width           =   240
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   13
            Left            =   2160
            Picture         =   "frmListado.frx":9212
            Top             =   1080
            Width           =   240
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Fecha baja"
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
            Index           =   83
            Left            =   480
            TabIndex        =   580
            Top             =   1080
            Width           =   915
         End
         Begin VB.Image imgBuscarG 
            Height          =   240
            Index           =   92
            Left            =   1515
            Picture         =   "frmListado.frx":929D
            ToolTipText     =   "Buscar motivo baja"
            Top             =   600
            Width           =   240
         End
         Begin VB.Label Label3 
            Caption         =   "Hasta"
            Height          =   195
            Index           =   89
            Left            =   960
            TabIndex        =   579
            Top             =   600
            Width           =   465
         End
         Begin VB.Image imgBuscarG 
            Height          =   240
            Index           =   91
            Left            =   1515
            Picture         =   "frmListado.frx":939F
            ToolTipText     =   "Buscar motivo baja"
            Top             =   240
            Width           =   240
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Motivo baja"
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
            Index           =   82
            Left            =   520
            TabIndex        =   577
            Top             =   0
            Width           =   1005
         End
         Begin VB.Label Label3 
            Caption         =   "Desde"
            Height          =   195
            Index           =   88
            Left            =   960
            TabIndex        =   576
            Top             =   240
            Width           =   465
         End
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   1
         Left            =   480
         TabIndex        =   553
         Top             =   5880
         Width           =   5895
         Begin VB.ComboBox cboTipoList 
            Height          =   315
            Left            =   1260
            Style           =   2  'Dropdown List
            TabIndex        =   554
            Tag             =   "Tipo Facturaci�n|N|N|||scaalb|tipofact||N|"
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo Listado"
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
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   555
            Top             =   240
            Width           =   4335
         End
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   45
         Left            =   1920
         TabIndex        =   212
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   45
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   252
         Text            =   "Text5"
         Top             =   1200
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   46
         Left            =   1920
         TabIndex        =   213
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   46
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   251
         Text            =   "Text5"
         Top             =   1560
         Width           =   3735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   51
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   250
         Text            =   "Text5"
         Top             =   4080
         Width           =   3735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   52
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   249
         Text            =   "Text5"
         Top             =   4440
         Width           =   3735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   48
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   238
         Text            =   "Text5"
         Top             =   2520
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   48
         Left            =   1920
         TabIndex        =   215
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   50
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   237
         Text            =   "Text5"
         Top             =   3480
         Width           =   3735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   49
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   236
         Text            =   "Text5"
         Top             =   3120
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   50
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   217
         Top             =   3480
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   49
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   216
         Top             =   3120
         Width           =   735
      End
      Begin VB.CommandButton cmdAceptarMante 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4080
         TabIndex        =   231
         Top             =   7080
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   9
         Left            =   5280
         TabIndex        =   235
         Top             =   7080
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   51
         Left            =   1920
         TabIndex        =   218
         Top             =   4080
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   52
         Left            =   1920
         TabIndex        =   219
         Top             =   4440
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   47
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   211
         Text            =   "Text5"
         Top             =   2160
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   47
         Left            =   1920
         TabIndex        =   214
         Top             =   2160
         Width           =   735
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   735
         Index           =   0
         Left            =   480
         TabIndex        =   355
         Top             =   4800
         Width           =   5415
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   53
            Left            =   1560
            TabIndex        =   358
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   54
            Left            =   3840
            TabIndex        =   357
            Top             =   360
            Width           =   1215
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   9
            Left            =   3555
            Picture         =   "frmListado.frx":94A1
            Top             =   360
            Width           =   240
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   8
            Left            =   1275
            Picture         =   "frmListado.frx":952C
            Top             =   360
            Width           =   240
         End
         Begin VB.Label Label3 
            Caption         =   "Desde"
            Height          =   195
            Index           =   44
            Left            =   720
            TabIndex        =   360
            Top             =   360
            Width           =   465
         End
         Begin VB.Label Label3 
            Caption         =   "Hasta"
            Height          =   195
            Index           =   45
            Left            =   3000
            TabIndex        =   359
            Top             =   360
            Width           =   420
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha Revisiones Efectuadas"
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
            Height          =   255
            Index           =   5
            Left            =   240
            TabIndex        =   356
            Top             =   120
            Width           =   4335
         End
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
         Index           =   28
         Left            =   1080
         TabIndex        =   255
         Top             =   1560
         Width           =   420
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
         Index           =   27
         Left            =   600
         TabIndex        =   254
         Top             =   960
         Width           =   420
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   41
         Left            =   1635
         Picture         =   "frmListado.frx":95B7
         Top             =   1200
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   43
         Left            =   1080
         TabIndex        =   253
         Top             =   1200
         Width           =   465
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   42
         Left            =   1635
         Picture         =   "frmListado.frx":96B9
         Top             =   1560
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   44
         Left            =   1635
         Picture         =   "frmListado.frx":97BB
         ToolTipText     =   "Buscar cliente"
         Top             =   2520
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   42
         Left            =   1080
         TabIndex        =   248
         Top             =   2160
         Width           =   465
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   46
         Left            =   1635
         Picture         =   "frmListado.frx":98BD
         ToolTipText     =   "Buscar agente"
         Top             =   3480
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   45
         Left            =   1635
         Picture         =   "frmListado.frx":99BF
         ToolTipText     =   "Buscar agente"
         Top             =   3120
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
         Index           =   26
         Left            =   600
         TabIndex        =   247
         Top             =   2880
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   41
         Left            =   1080
         TabIndex        =   246
         Top             =   3480
         Width           =   420
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   40
         Left            =   1080
         TabIndex        =   245
         Top             =   3120
         Width           =   465
      End
      Begin VB.Label Label7 
         Caption         =   "Informe de Mantenimientos"
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
         Left            =   600
         TabIndex        =   244
         Top             =   360
         Width           =   5775
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   39
         Left            =   1080
         TabIndex        =   243
         Top             =   4080
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   38
         Left            =   1080
         TabIndex        =   242
         Top             =   4440
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Contrato"
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
         Left            =   600
         TabIndex        =   241
         Top             =   3840
         Width           =   1155
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   47
         Left            =   1635
         Picture         =   "frmListado.frx":9AC1
         ToolTipText     =   "Buscar tipo contrato"
         Top             =   4080
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   48
         Left            =   1635
         Picture         =   "frmListado.frx":9BC3
         ToolTipText     =   "Buscar tipo contrato"
         Top             =   4440
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   43
         Left            =   1635
         Picture         =   "frmListado.frx":9CC5
         ToolTipText     =   "Buscar cliente"
         Top             =   2160
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
         Index           =   24
         Left            =   600
         TabIndex        =   240
         Top             =   1920
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
         Index           =   23
         Left            =   1080
         TabIndex        =   239
         Top             =   2520
         Width           =   420
      End
   End
End
Attribute VB_Name = "frmListado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ====  MODIFICACIONES  ==========================================
' ====  [16/09/2009] LAURA : A�adir el frame "FrameInvArtComp" para sacar listado articulos con componentes
' ====  [17/09/2009] LAURA : Cambiar conexion CONTAB si FAZ
' ================================================================


Public OpcionListado As Integer

    '==== Listados de ALMACEN ====
    '=============================
    ' 1 .- Listados Marcas.
    ' 2 .- Listado de Almacenes Propios
    ' 3 .- Listado de Tipos de Unidad
    ' 4 .- Listado de Tipos de Art�culos
    ' 5 .- Listado de Familias de art�culos
    
    ' 6 .- Listado de Art�culos
    ' 7 .- Informe de Traspaso de Almacenes
    ' 8 .- Informe de Movimientos de Almacen
    ' 9 .- Listado Busquedas de movimientos de Art�culos
    '10 .-
    
    '11 .- Listado de Articulos con componentes ' ====  [16/09/2009] LAURA
    '12 .- Listado Toma de Inventario Articulos
    '13 .- Listado de Diferencias de Inventario Articulos
    '14 .- Actualizar Diferencias de Inventario (No IMPRIME INFORME)
    '15 .- Listado de Articulos Inactivos.
    
    '16 .- Listado Valoracion de Stocks Inventariados
    '17 .- Listado Valoraci�n Stocks
    '18 .- Informe Stocks Maximos y Minimos
    '19 .- Informe de Stocks a una fecha
    
    '110 .- Listado de Ubicaciones
    
    
    
    
    '==== Listados de FACTURACION ====
    '=================================
    '20 .- Listado de Actividades de Clientes
    '21 .- Listado de Zonas de Clientes
    '22 .- Listado de Rutas de Asistencia
    '23 .- Listado de Formas de Env�o
    '24 .- Listado de Tarifas Ventas
    '25 .-
    
    '26 .-
    '27 .- Listado de Situaciones Especiales
    '28 .- Informe de Tarifas de Articulos
    '29 .- Informe de Promociones de Tarifas
    '30 .- Informe de Precios Especiales
    
    '31 .- Informe de Ofertas
    '32 .- Informe de Recordatorio de Ofertas
    '33 .- Informe de Valoraci�n de Ofertas
    '34 .- Informe de Ofertas Efectuadas
    '35 .- Informe Historico de Ofertas
    
    '36 .- Traspaso de Ofertas al Historico (NO IMPRIME INFORME)
    '37 .- Solicitar datos para pasar de Oferta a Pedido (NO IMPRIME INFORME)
    '38 .- Informe de Pedidos
    '39 .- Orden de Instalacion
    '40 .- Cartas Confirmacion de Pedidos
    
    '41 .- Informe de Pedidos por Articulo
    '42 .- Informe de Disponibilidad de Stocks
    '43 .- Generar Albaran desde Pedido (NO IMPRIME LISTADO)
    '44 .- Informe de Pedidos por Cliente
    '45 .- Informe de Albaran
    
    '46 .- Informe de Clientes Inactivos
    '47 .- Informe de Clientes
    '48 .- Informe de Altas de Nuevos Cliente
    '49 .- Informe de Albaranes por Articulo
    '50 .- Prevision de Facturacion de ALbaranes
    
    '51 .- Informe Incumplimiento Plazos de Entrega
    '52 .- Facturacion de Albaranes (NO IMPRIME LISTADO?)
    '53 .- Informe de Factura
    '54 .- Listado de Descuentos Familia/Marca
    
    '59 .- Informe de Factura ProForma
    '222 .- Informe de Factura Mostrador
    '223 .- Pedir datos para contabilizar facturas CLIENTES
    '224 .- Pedir datos para contabilizar facturas PROVEEDOR
    '225 .- Pedir datos para generar Facturas Rectificativas
    '226 .- Pedir datos para reimprimir Facturas
    '227 .- Informe estadistica Ventas por cliente
    '228 .- Informe estadistica Ventas por Trabajador
    '229 .- Informe estadistica Ventas por meses
    '230 .- Informe estadistica Ventas por familia
    '231 .- Informe detalle facturacion clientes
    
    '238 .- Confirmacion entrega Pedido
    '239 .- Hco de Pedidos de venta (Historico)
    '240 .- Informe Cierre de Caja del TPV
    
    '245 .- Informe control margenes tarifas
    '246 .- Informe Margen ventas por articulo
    '247 .- Correcci�n de errores y acutalizacion de tarifas
    
    
    'Abril 2008
    '248 .- Contabilizar facturas de tickets AGRUPADAS
    
    
    
    '==== Listados de COMPRAS ====
    '=============================
    '55 .- Informe de Pedido Proveedor
    '56 .- Inf. Historico Pedido Proveedor
    '57 .- Pasa Pedido a Albaran compras (NO IMPRIME LISTADO)
    '58 .- Listado de Proveedores
    
    
    '305 .- Listado Etiquetas de Proveedores
    '306 .- Listado Cartas a Proveedores
    '307 .- Listado Material pendiente de recibir
    '308 .- Listado Albaranes pendientes de facturar
    '309 .- Listado  Precios de Compra
    '310 .- Listado Compras por Proveedor
    '311 .- Listado Compras por Familia
    '312 .- Listado albaranes por proveedor
    
    
    '==== Listados de REPARACIONES ====
    '==================================
    '60 .- Informe de Numeros de Serie
    '61 .- Listado Motivos Pend. Rep.
    '62 .- Listado Resguardo Reparacion
    '63 .- Listado Reparaciones por D�a
    '64 .- Listado Reparaciones por Cliente
    '65 .- Listado motivos baja equipos
    
    '406 .- Listado Frecuencia de reparaciones
    '407 .- Sustituci�n N� de Serie
    '408 .- Informe Aviso de Averia
    '409 .- Listado Avisos de averia pendientes
    
    
    '==== Listados de ADMINISTRACION ====
    '====================================
    
    '501 .- Listado de Nominas y Gastos
    
    
    '==== Listados de MANTENIMIENTOS ====
    '==================================
    '70 .- Listado Mantenimiento
    '71 .- Listado Revisiones de Mantenimientos
    '72 .- Informe Fichas de Mantenimientos
    '73 .- Listado Altas de Mantenimientos
    '74 .- Prefacturaci�n Mantenimientos
    '75 .- Facturaci�n de Mantenimientos
    '76 .- IGUAL QUE EL 70 pero en ANULADOS
        
        
        
    '77 .- Informe te�rico de mantenimientos
    '78 .- Cartas de renovacion
    '79 .- Etiquetas manteimiento
    
    
    '==== Listados OTROS ====
    '==================================
    
    '80 .- Pasar Albaranes Ventas al historico (NO IMPRIME)
    '81 .- Pasar Pedidos Ventas al historico (NO IMPRIME)
       
           
    '82 .- Marcar facturar albaranes
    '83 .- Borre avisos cerrados
       
    
       
       
    '90 .- Etiquetas de Clientes
    '91 .- Cartas a Clientes
    
    '92 .- Informe de Gastos T�cnicos
    '93 .- Ticket del TPV
      
    '94 .- Etiquetas estanteria
    
    '95 .- Etiquetas de bultos
    '96 .- Frecuencias
    '97 .- Eliminar facturas
    '99 .- Traspaso a mantenimientos anulados
    
    
    'Marzo 2013
    '100 .- Listado extendido de almacen. Con stockmin,puntoped,stock ....
    
    'Octubre 2014
    '101 .-  Bultos PROVEEDOR. Mismo frame, distintos textos
    
    
    '512 .- Contabilizar una FACTURA unicamente   'JULIO 2010
    
    '513 .- Impresion etiquetaqs estanteria desde albaran de compra
    
    '514 .- Dada una recepcion de factura, una vez creada, mostrara
    '            el numregis asignado y los vencimientos
    
    
Public NumCod As String 'Para indicar cod. Traspaso,Movimiento, etc. que llama
                        'Para indicar n� oferta a imprimir

Public CadTag As String 'Cadena con el Tag del campo que se va a poner en D/H en los listados
                        'Se necesita si el tipo de codigo es texto

    
Private Conexion As Byte
'1.- Conexi�n a BD Ariges  2.- Conexi�n a BD Conta

Private HaDevueltoDatos As Boolean

Private WithEvents frmMtoAlPropios As frmAlmAlPropios
Attribute frmMtoAlPropios.VB_VarHelpID = -1
Private WithEvents frmMtoUbica As frmAlmUbicaciones 'Ubicaciones de Almacen
Attribute frmMtoUbica.VB_VarHelpID = -1
Private WithEvents frmMtoMarcas As frmAlmMarcas
Attribute frmMtoMarcas.VB_VarHelpID = -1
Private WithEvents frmMtoTUnidad As frmAlmTipoUnidad
Attribute frmMtoTUnidad.VB_VarHelpID = -1
Private WithEvents frmMtoTArticulo As frmAlmTipoArticulo
Attribute frmMtoTArticulo.VB_VarHelpID = -1
Private WithEvents frmMtoActiv As frmFacActividades
Attribute frmMtoActiv.VB_VarHelpID = -1
Private WithEvents frmMtoZonas As frmFacZonas
Attribute frmMtoZonas.VB_VarHelpID = -1
Private WithEvents frmMtoRutas As frmFacRutas
Attribute frmMtoRutas.VB_VarHelpID = -1
Private WithEvents frmMtoFEnvio As frmFacFormasEnvio
Attribute frmMtoFEnvio.VB_VarHelpID = -1
Private WithEvents frmMtoTarifas As frmFacTarifas
Attribute frmMtoTarifas.VB_VarHelpID = -1
Private WithEvents frmMtoSituac As frmFacSituaciones
Attribute frmMtoSituac.VB_VarHelpID = -1
Private WithEvents frmMtoFamilia As frmAlmFamiliaArticulo
Attribute frmMtoFamilia.VB_VarHelpID = -1
Private WithEvents frmMtoProveedor As frmComProveedores
Attribute frmMtoProveedor.VB_VarHelpID = -1
Private WithEvents frmMtoArticulos As frmAlmArticu2
Attribute frmMtoArticulos.VB_VarHelpID = -1
Private WithEvents frmMtoCartasOfe As frmFacCartasOferta
Attribute frmMtoCartasOfe.VB_VarHelpID = -1
Private WithEvents frmMtoTraba As frmAdmTrabajadores
Attribute frmMtoTraba.VB_VarHelpID = -1
Private WithEvents frmMtoClientes As frmFacClientes3
Attribute frmMtoClientes.VB_VarHelpID = -1
Private WithEvents frmMtoMotivos As frmRepMotivosPend
Attribute frmMtoMotivos.VB_VarHelpID = -1
Private WithEvents frmMtoAgentes As frmFacAgentesCom
Attribute frmMtoAgentes.VB_VarHelpID = -1
Private WithEvents frmMtoTiposCon As frmManTiposContrato
Attribute frmMtoTiposCon.VB_VarHelpID = -1
Private WithEvents frmMtoIncid As frmIncidencias
Attribute frmMtoIncid.VB_VarHelpID = -1

Private WithEvents frmB As frmBuscaGrid  'Busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmF As frmCal 'Calendario de Fechas
Attribute frmF.VB_VarHelpID = -1

'---- Variables para el INFORME ----
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private CadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para el frmImprimir
Private cadNomRPT As String 'Nombre del informe a Imprimir
Private conSubRPT As Boolean 'Si el informe tiene subreports
'-----------------------------------





'Para ademas de insertarlas en la conta, que las contabilice (pase a hsaldos)
'es decir, en el momento que inserta en cabfact tb insertaremos en hlinapu, hacabapu, hsaldos y hsaldosanal (si procede)










Dim indCodigo As Integer 'indice para txtCodigo

Dim codigo As String 'C�digo para FormulaSelection de Crystal Report
Dim Orden1 As String 'Campo de Ordenacion (por codigo) para Cristal Report
Dim Orden2 As String 'Campo de Ordenacion (por nombre) para Cristal Report

Dim PrimeraVez As Boolean
Dim indFrame As Single


Private Sub KEYpress(KeyAscii As Integer)
    Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 2, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub cboSituaAviso_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub





Private Sub cboStokFecha_KeyPress(KeyAscii As Integer)
 KEYpress KeyAscii
End Sub

Private Sub cboTipMov_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cboTipoList_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub






Private Sub cboVarios_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkCompo_Click()
    If Me.chkCompo.Value = 0 Then
        Label4(92).Caption = "Articulo"
    Else
        Label4(92).Caption = "Componentes"
    End If
    FrameAlmacenesListadoComponentes.visible = chkCompo.Value = 1
    FrameValorar3.visible = Not FrameAlmacenesListadoComponentes.visible
End Sub

Private Sub chkCompo_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkDtoFM_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkImpEtiq_Click(Index As Integer)
    If Index = 0 Then
        If chkImpEtiq(0).Value = 1 Then
            chkImpEtiq(1).Caption = "Stock minimo"
            chkImpEtiq(1).Value = 0
        Else
            chkImpEtiq(1).Caption = "P.V.P."
            chkImpEtiq(1).Value = 0
        End If
    End If
        
End Sub

Private Sub chkImprimeCodigoBarras_KeyPress(KeyAscii As Integer)
     KEYpress KeyAscii
End Sub



Private Sub chkProv_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkSitaucionArticulo2_KeyPress(Index As Integer, KeyAscii As Integer)
     KEYpress KeyAscii
End Sub

Private Sub cmbBulto_Click()
    PonerCamposDireccionBultos cmbBulto.ListIndex
End Sub

Private Sub cmbBulto_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmbProduccion_Click()
    If PrimeraVez Then Exit Sub
    PonerLabelsArticulosFrameVisible cmbProduccion.ListIndex = 1
End Sub



Private Sub cmdAceptar_Click(Index As Integer)
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim cadAux As String
Dim bol As Boolean
Dim bytPrecio As Byte

   InicializarVbles
   
   Select Case Index
   '========= Frame Listados =================================================
    
    ' ====  [16/09/2009] LAURA : Listado Articulos con componentes
    Case 0 ' Listado Articulos con componentes
        If Me.chkCompo.Value = 0 Then
            cadNomRPT = "rAlmArtCompon.rpt"
            conSubRPT = True
            codigo = "{sartic.codartic}"
        Else
            cadNomRPT = "rAlmArtCompVer.rpt"
            conSubRPT = False
            codigo = "{sarti1.codarti1}"
        End If
        Screen.MousePointer = vbHourglass
        DoEvents
        
        'A�adir el parametro de Empresa
        CadParam = "|pEmpresa=""" & vEmpresa.nomempre & """|"
        numParam = 1
        
        
        
        'Septiembre 2014. Herbelca
        If Me.chkCompo.Value = 1 Then
            cadAux = ""
          
            For bytPrecio = 1 To 3
                
                CadParam = CadParam & "Almacen" & bytPrecio & "=" & Val(txtCodigo(144 + bytPrecio).Text) & "|"
                numParam = numParam + 1
                'empieza en el 145
                If Trim(txtCodigo(144 + bytPrecio).Text) <> "" Then
                    cadAux = cadAux & "      " & Trim(txtCodigo(144 + bytPrecio)) & " " & txtNombre(144 + bytPrecio).Text
                End If
            Next
            If cadAux = "" Then cadAux = "ERROR ALMACENES"
            CadParam = CadParam & "pAlmacenes=""Almacenes:  " & Trim(cadAux) & """|"
            numParam = numParam + 1
        End If
        If Trim(txtCodigo(125).Text) <> "" Or Trim(txtCodigo(126).Text) <> "" Then
            cadFormula = CadenaDesdeHasta(txtCodigo(125).Text, txtCodigo(126).Text, codigo, "T")
            
            If cadFormula <> "" Then
                If cadFormula = "Error" Then Exit Sub
                cadAux = ""
                If txtCodigo(125).Text <> "" Then cadAux = "Desde: " & txtCodigo(125).Text & " " & txtNombre(125).Text
                If txtCodigo(126).Text <> "" Then
                    If cadAux <> "" Then cadAux = cadAux & "  -  "
                    cadAux = cadAux & " Hasta: " & txtCodigo(126).Text & " " & txtNombre(126).Text
                End If
                CadParam = CadParam & "pDesde=""" & cadAux & """|"
                numParam = numParam + 1
            End If
        End If
        
        'Solo los que tienen componentes
        If Me.chkCompo.Value = 0 Then AnyadirAFormula cadFormula, " {sartic.conjunto}=1"
        
        
        'Parametro Precio de Valoracion, elegir un Precio para realizar la valoracion: canstock * precio
        If Me.optPrecioMP3.Value Then bytPrecio = 1
        If Me.optPrecioMA3.Value Then bytPrecio = 2
        If Me.optPrecioUC3.Value Then bytPrecio = 3
        If Me.optPrecioStd3.Value Then bytPrecio = 4
        CadParam = CadParam & "pPrecio=" & bytPrecio & "|"
        numParam = numParam + 1
    ' ====
   
   
    Case 1 'Frame Listados
        If Me.Optcodigo.Value = True Then
            cadAux = Orden1
        Else
            cadAux = Orden2
        End If
        CadParam = "|pOrden=" & cadAux & "|"
        numParam = 1
        
        'A�adir el parametro de Empresa
        CadParam = CadParam & "pEmpresa=""" & vEmpresa.nomempre & """|"
        numParam = numParam + 1
        
        If Trim(txtCodigo(1).Text) <> "" Or Trim(txtCodigo(2).Text) <> "" Then
            'Cadena para seleccion Desde y Hasta
            If OpcionListado = 4 Or OpcionListado = 110 Then
                '4: Listado Tipos de Articulos, 110: List. Ubicaciones
                cadFormula = CadenaDesdeHasta(txtCodigo(1).Text, txtCodigo(2).Text, codigo, "T")
            Else
                cadFormula = CadenaDesdeHasta(txtCodigo(1).Text, txtCodigo(2).Text, codigo, "N")
            End If
            
            If cadFormula <> "" Then
                If cadFormula = "Error" Then Exit Sub
                cadAux = ""
                If txtCodigo(1).Text <> "" Then cadAux = "Desde: " & txtCodigo(1).Text & " " & txtNombre(1).Text
                If txtCodigo(2).Text <> "" Then
                    If cadAux <> "" Then cadAux = cadAux & "  -  "
                    cadAux = cadAux & " Hasta: " & txtCodigo(2).Text & " " & txtNombre(2).Text
                End If
                CadParam = CadParam & "pDesde=""" & cadAux & """|"
                numParam = numParam + 1
            End If
        End If
        
    '========= Frame Informes Almacen ========================================
    Case 2 'Frame Informes Almacen
        If OpcionListado = 7 Then '7: Traspaso Almacen
            indRPT = 1
            cadAux = "scatra"
            cadTitulo = "Informe Traspaso Almacenes"
        ElseIf OpcionListado = 8 Then '8: Movimientos Almacen
            indRPT = 3
            cadAux = "scamov"
            cadTitulo = "Informe Movimientos Almacen"
        End If
        
        CadParam = "|"
        If Not PonerParamEmpresa(CadParam, numParam) Then Exit Sub
        If PonerParamRPT2(indRPT, CadParam, numParam, cadNomRPT, pImprimeDirecto, pPdfRpt, pRptvMultiInforme) Then
            'Cadena para seleccion Desde y Hasta DOCUMENTO
            '----------------------------------------------
            If txtCodigo(3).Text <> "" Or txtCodigo(4).Text <> "" Then
                If Not PonerDesdeHasta(codigo, "N", 3, 4, "") Then Exit Sub
            End If
        
            If Not HayRegParaInforme(cadAux, cadSelect) Then Exit Sub
        End If
                       
                   
                   
    '========= Frame Listado Movimiento de Art�culos ========================
    Case 3 'Frame Listado Movimiento de Art�culos
        'Nombre fichero .rpt a Imprimir
        
        indRPT = 75
        If Not PonerParamRPT2(indRPT, CadParam, numParam, cadNomRPT, pImprimeDirecto, pPdfRpt, pRptvMultiInforme) Then cadNomRPT = "rAlmMovim.rpt"
        
        
        
        If Not PonerFormulaYParametrosInf9() Then Exit Sub
        'comprobar que hay datos para mostrar en el Informe
        cadAux = "smoval INNER JOIN sartic ON smoval.codartic=sartic.codartic "
        If Not HayRegParaInforme(cadAux, cadSelect) Then Exit Sub
        conSubRPT = True
    
    '========= Frame de Inventario ==========================================
    Case 4 'Frame de Inventario
        If Not ValidarCamposInventario Then Exit Sub
        If OpcionListado = 19 Then
            
            If chkProv2(0).Value = 1 Then
                cadNomRPT = "rAlmStocksFechaProv.rpt"
            Else
                cadNomRPT = "rAlmStocksFecha.rpt"
            End If
            
        Else
            'Nombre fichero .rpt a Imprimir
            If vParamAplic.InventarioxProv Then 'Se realiza inventario por Proveedor
                                                'Ordenar por: codprove, codfamia, codartic
                Select Case OpcionListado
                    Case 12: cadNomRPT = "rAlmInvenxProv.rpt"  'Listado Toma de Inventario
                    Case 13: cadNomRPT = "rAlmInvenxProvDif.rpt"  'Listado Diferencias
                    Case 15: cadNomRPT = "rAlmArtInactivos.rpt"
                    Case 16: cadNomRPT = "rAlmInvenxProvValoracion.rpt"  'Listado Valoracion Stocks Inventariados
                    Case 17: cadNomRPT = "rAlmValoracionxProv.rpt"  'Listado Valoracion Stocks (Por Proveedor)
                End Select
            Else 'Ordenar por Cod. Familia y no por Proveedor. Ordenar por: codfamia, codartic.
                Select Case OpcionListado
                    Case 12: cadNomRPT = "rAlmInventario.rpt"  'Listado Toma de Inventario
                    Case 13: cadNomRPT = "rAlmInventarioDif.rpt"  'Listado Diferencias
                    Case 15: cadNomRPT = "rAlmArtInactivos.rpt"
                    Case 16: cadNomRPT = "rAlmInvenValoracion.rpt"  'Listado Valoracion Stocks Inventariados
                    Case 17: cadNomRPT = "rAlmValoracion.rpt"  'Listado Valoracion Stocks)
                End Select
            End If
        End If
        Screen.MousePointer = vbHourglass
        DoEvents
        bol = PonerFormulaYParametrosInf12()
        Screen.MousePointer = vbDefault
        If Not bol Then Exit Sub
        
   End Select
    
       
   If OpcionListado = 14 Then 'Actualizar Inventario (NO IMPRIME INFORME)
        If Trim(txtCodigo(21).Text) <> "" Then
            'Quitar las llaves:{tabla.codigo} de la cadena consulta
            'para el FormulaSelection del informe Crystal Report y
            'Tendremos la clausula WHERE para insertar en la tabla:sinven
            cadAux = QuitarCaracterACadena(cadFormula, "{")
            cadFormula = QuitarCaracterACadena(cadAux, "}")
            If ActualizarInventario Then
                MsgBox "La Actualizaci�n de Inventario se ha realizado correctamente.", vbInformation
            End If
        Else
            MsgBox "El campo Trabajador debe tener valor", vbInformation
            PonerFoco txtCodigo(21)
            Exit Sub
        End If
        
   Else 'Listados
   
   

   
   
   
'        If OpcionListado = 19 Then cadFormula = ""
        If OpcionListado = 19 Then cadFormula = "({tmpstockfec.codusu} =" & vUsu.codigo & ")"
        
        LlamarImprimir Index = 2 'Movimientos almacen si tiene rpt personalizables

        'Realizar otras acciones segun el informe que llame
        Select Case OpcionListado
            Case 12 'Toma de Inventario
                If HaPulsadoElBotonDeImprimir Then
                    PrepararTomaInventario
                End If
            Case 7, 8 'Movimientos
                ActualizarImprimir
            Case 19
                DescargarDatosTMPStockFecha
        End Select
        
   End If
   Screen.MousePointer = vbDefault
End Sub


Private Sub PrepararTomaInventario()
Dim cadAux As String
Dim devuelve As String
    On Error GoTo ETomaInv
    
    

    
    
    If MsgBox("�Impresi�n correcta para Actualizar Inventario?", vbQuestion + vbYesNo) = vbYes Then
        'Quitar las llaves:{tabla.codigo} de la cadena consulta
        'para el FormulaSelection del informe Crystal Report y
        'Tendremos la clausula WHERE para insertar en la tabla:sinven
'                cadAux = QuitarCaracterACadena(cadFormula, "{")
'                cadFormula = QuitarCaracterACadena(cadAux, "}")
       If CrearTmpInventario(cadSelect) Then
            If InsertarInventario Then
                MsgBox "Puede pasar a realizar la Entrada de Inventario Real", vbInformation
            End If
       End If
       cadAux = "DROP TABLE IF EXISTS tmpInven "
       conn.Execute cadAux
    End If
    
ETomaInv:
    If Err.Number <> 0 Then Err.Clear
End Sub



Private Sub cmdAceptarArtic_Click()
'Listado de Articulos
Dim campo As String
Dim devuelve As String
Dim Opcion As Byte, numOp As Byte
Dim cadFrom As String

Dim PrevioArticulos As Boolean



    InicializarVbles
    
    'Si es informe=18 de Stocks Maximos y Minimos comprobar
    'que se ha seleccionado un almacen
    Select Case OpcionListado
    Case 18
        'If OpcionListado = 18 Then
        If txtCodigo(72).Text = "" Then
            MsgBox "Se debe seleccionar un Almacen para el informe.", vbInformation
            Exit Sub
        End If
        cadNomRPT = "rAlmStocksMaxMin.rpt"
        cadFrom = " salmac INNER JOIN sartic ON salmac.codartic=sartic.codartic "
    Case 247
        '
        Opcion = 0
        If vParamAplic.Produccion And Me.cmbProduccion.ListIndex = 1 Then Opcion = 1
        If Opcion = 0 Then
            If txtCodigo(107).Text = "" Or txtNombre(107) = "" Then
                MsgBox "Debe seleccionar una tarifa para el informe.", vbInformation
                Exit Sub
            End If
        Else
            'Corrector de precios de articulos con componentes
            txtCodigo(107).Text = ""
            txtNombre(107) = ""
        End If
    Case Else
        'El 6
        cadNomRPT = "rAlmListArticulos.rpt"  'Nombre fichero .rpt a Imprimir
        cadFrom = " sartic"
        CadParam = ""
        For Opcion = 0 To 3
            If Me.chkSitaucionArticulo2(Opcion).Value = 1 Then CadParam = CadParam & "O"
        Next
        If CadParam = "" Then
            MsgBox "Seleccione la situacion del articulo", vbExclamation
            Exit Sub
        End If
        Opcion = 0
        
        'Mayo 2015. Ya no necesitamos este msg, porque cuando NO son etiquetas el check es de PVP
'        If Me.chkImpEtiq(1).Value = 1 And Me.chkImpEtiq(0).Value = 0 Then
'            MsgBox "Opcion stock minimo solo para impresion de etiquetas", vbExclamation
'            chkImpEtiq(1).Value = 0
'        End If
    End Select
    
    '===================================================
    '============ PARAMETROS ===========================
    CadParam = "|"
    'Empresa
    CadParam = CadParam & "pEmpresa=""" & vParam.NombreEmpresa & """|"
    numParam = numParam + 1
    
    
    '====================================================
    '================= FORMULA ==========================
    'Cadena para seleccion  ALMACEN
    '--------------------------------------------
    If OpcionListado = 18 And txtCodigo(72).Text <> "" Then
        campo = "{salmac.codalmac}"
        cadFormula = campo & "= " & txtCodigo(72).Text
        
        
    Else
        'Es tarifa para la correccion
        If OpcionListado = 247 And txtCodigo(107).Text <> "" Then
            campo = "{slista.codlista}"
            cadFormula = campo & "= " & txtCodigo(107).Text
        End If
    End If
    
    
    'Cadena para seleccion D/H FAMILIA
    '--------------------------------------------
    devuelve = "pDHFamilia="""
        
    If OpcionListado = 6 Then
        'Listado articulos. SOLO rotacion
        If chkImpEtiq(2).Value = 1 Then
            campo = "{sartic.rotacion}=1"
            devuelve = devuelve & chkImpEtiq(2).Caption & "    "
            AnyadirAFormula cadSelect, campo
            AnyadirAFormula cadFormula, campo
        End If
    End If
    
    
    If txtCodigo(62).Text <> "" Or txtCodigo(63).Text <> "" Then
        'Parametro Desde/Hasta Familila
        devuelve = devuelve & "Familia: "
        campo = "{sartic.codfamia}"
        If Not PonerDesdeHasta(campo, "N", 62, 63, devuelve) Then Exit Sub
    Else
        
        CadParam = CadParam & devuelve & """|"
        numParam = numParam + 1
    End If
    
    'Cadena para seleccion D/H MARCA
    '--------------------------------------------
    If txtCodigo(64).Text <> "" Or txtCodigo(65).Text <> "" Then
        campo = "{sartic.codmarca}"
        'Parametro Desde/Hasta Marca
        devuelve = "pDHMarca=""Marca: "
        If Not PonerDesdeHasta(campo, "N", 64, 65, devuelve) Then Exit Sub
    End If
    
    'Cadena para seleccion D/H PROVEEDOR
    '--------------------------------------------
    If txtCodigo(66).Text <> "" Or txtCodigo(67).Text <> "" Then
        campo = "{sartic.codprove}"
        'Parametro Desde/Hasta Proveedor
        devuelve = "pDHProveedor=""Proveedor: "
        If Not PonerDesdeHasta(campo, "N", 66, 67, devuelve) Then Exit Sub
    End If
    
    'Cadena para seleccion D/H TIPO ARTICULO
    '--------------------------------------------
    cadTitulo = ""
    devuelve = ""
    If txtCodigo(68).Text <> "" Or txtCodigo(69).Text <> "" Then
        campo = "{sartic.codtipar}"
        'Parametro Desde/Hasta Tipo Articulo
        devuelve = "Tipo Articulo: "
        If Not PonerDesdeHasta(campo, "T", 68, 69, devuelve) Then Exit Sub
        If devuelve <> "" Then cadTitulo = devuelve
    End If
    
    
    If OpcionListado = 6 Then
        indCodigo = 0
        devuelve = ""
        If Me.chkSitaucionArticulo2(0).Value = 1 Then devuelve = "- NORMAL": indCodigo = indCodigo + 1
        If Me.chkSitaucionArticulo2(1).Value = 1 Then devuelve = devuelve & "- OBSOLETO": indCodigo = indCodigo + 1
        If Me.chkSitaucionArticulo2(2).Value = 1 Then devuelve = devuelve & "- BLOQUEADO": indCodigo = indCodigo + 1
        If Me.chkSitaucionArticulo2(3).Value = 1 Then devuelve = devuelve & "- CADUCADO": indCodigo = indCodigo + 1
        If indCodigo <> 4 Then
          cadTitulo = Trim(cadTitulo & "      Situacion: " & Mid(devuelve, 2))
        End If
    End If
    If cadTitulo <> "" Then
        devuelve = "pDHTipoArt=""" & cadTitulo & """|"
        CadParam = CadParam & devuelve
        numParam = numParam + 1
    End If
    
    'Cadena para seleccion D/H ARTICULO
    '--------------------------------------------
    If txtCodigo(70).Text <> "" Or txtCodigo(71).Text <> "" Then
        campo = "{sartic.codartic}"
        'Parametro Desde/Hasta Articulo
        devuelve = "pDHArticulo=""Articulo: "
        If Not PonerDesdeHasta(campo, "T", 70, 71, devuelve) Then Exit Sub
    End If
    
    
    
    'Obtener el parametro con el Orden del Informe
    '---------------------------------------------
    Select Case OpcionListado
    Case 6
    
        PrevioArticulos = False
        If chkImpEtiq(0).Value = 1 Then
            PrevioArticulos = True
        Else
            If chkImpEtiq(1).Value = 1 Then PrevioArticulos = True   'PVP
        End If
        If PrevioArticulos Then
        
            'A�adir a la formula el chk
            indCodigo = Abs(Me.chkSitaucionArticulo2(0).Value) + Abs(Me.chkSitaucionArticulo2(1).Value) + Abs(Me.chkSitaucionArticulo2(2).Value) + Abs(Me.chkSitaucionArticulo2(3).Value)
            If indCodigo < 4 Then
                'Ha seleccionado 2 o uno
                devuelve = ""
                If Me.chkSitaucionArticulo2(0).Value = 1 Then devuelve = ", 0"
                If Me.chkSitaucionArticulo2(1).Value = 1 Then devuelve = ", 1"
                If Me.chkSitaucionArticulo2(2).Value = 1 Then devuelve = ", 2"
                If Me.chkSitaucionArticulo2(3).Value = 1 Then devuelve = ", 3"
                devuelve = Mid(devuelve, 2)
                campo = "{sartic.codstatu} IN (" & devuelve & ")"
                AnyadirAFormula cadFormula, campo
                AnyadirAFormula cadSelect, campo
                
            End If
                
            'Solo con punto ped
            devuelve = ""  'En el frm.show solo cargara los del almacen con punto de pedido
            If chkImpEtiq(1).Value = 1 Then
                If chkImpEtiq(0).Value = 1 Then
                    Do
                        campo = InputBox("Seleccione un almacen?", "Almacen")
                        If campo <> "" Then
                            If Not IsNumeric(campo) Then
                                MsgBox "Almacen debe ser numerico", vbExclamation
                            Else
                                campo = DevuelveDesdeBD(conAri, "codalmac", "salmpr", "codalmac", campo)
                                If campo = "" Then
                                    MsgBox "No existe el almacen: ", vbExclamation
                                    campo = "N"
                                Else
                                    'OK
                                   
                                    devuelve = campo
                                    campo = ""
                                End If
                            End If
                        End If
                    Loop Until campo = ""
                    If devuelve = "" Then Exit Sub
                End If
            End If
        
            'Etiquetas. Lazanzaremos el mismo proceso que en etiquetas estanteria Punto de venta
            cadTitulo = "DELETE FROM tmpnseries WHERE codusu =" & vUsu.codigo
            conn.Execute cadTitulo
            cadTitulo = Replace(cadSelect, "{", "")
            cadTitulo = Replace(cadTitulo, "}", "")
            frmMensajes.cadWHERE2 = "#,##0.00"  'dos decimales
            frmMensajes.cadWhere = cadTitulo
            frmMensajes.vCampos = devuelve
            If chkImpEtiq(0).Value = 0 Then
                'PVPs
                frmMensajes.OpcionMensaje = 23
            Else
                'etiquetas
                frmMensajes.OpcionMensaje = 15
            End If
            frmMensajes.Show vbModal
        
        
            cadTitulo = "DELETE FROM tmpinformes WHERE codusu =" & vUsu.codigo
            conn.Execute cadTitulo
    
            'A�adire los tipos de IVA a esta tabla para los posibles links
            cadTitulo = "INSERT INTO tmpinformes(codusu,codigo1)  select " & vUsu.codigo & ",codigiva from tmpnseries,sartic"
            cadTitulo = cadTitulo & " WHERE codusu = " & vUsu.codigo & " AND tmpnseries.codartic=sartic.codartic"
            cadTitulo = cadTitulo & " GROUP BY codigiva"
            conn.Execute cadTitulo
            
            Espera 0.2
             'Abrimos los IVAS en conta
            Set miRsAux = New ADODB.Recordset
            cadTitulo = "Select codigo1 from tmpinformes WHERE codusu = " & vUsu.codigo
            miRsAux.Open cadTitulo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not miRsAux.EOF
                cadTitulo = DevuelveDesdeBD(conConta, "porceiva", "tiposiva", "codigiva", miRsAux!Codigo1)
                cadTitulo = TransformaComasPuntos(cadTitulo)
                cadTitulo = "UPDATE tmpinformes SET porcen1= " & cadTitulo
                cadTitulo = cadTitulo & " WHERE codusu = " & vUsu.codigo & " AND codigo1 = " & miRsAux!Codigo1
                conn.Execute cadTitulo
                miRsAux.MoveNext
            Wend
            miRsAux.Close
            Set miRsAux = Nothing
            
    
        
        
            'Es para imprimir etiquetas
            'rAlmArticulosPVP.rpt
            If chkImpEtiq(0).Value = 0 Then
                'OK. PVPV
                If Not PonerParamRPT2(80, CadParam, numParam, cadNomRPT, pImprimeDirecto, pPdfRpt, pRptvMultiInforme) Then cadNomRPT = "rAlmArticulosPVP.rpt"
                cadTitulo = "Articulos PVP IVA"
            Else
                If Not PonerParamRPT2(23, CadParam, numParam, cadNomRPT, pImprimeDirecto, pPdfRpt, pRptvMultiInforme) Then cadNomRPT = "rEtiArticulo.rpt"
                cadTitulo = "Etiquetas articulos gral."
                CadParam = "|pImprimeBarras=""1""|numerodecimales=2|"
                numParam = 2
            End If
            
            cadFrom = " tmpnseries   "
            cadSelect = "{tmpnseries.codusu} =" & vUsu.codigo
            cadFormula = "{tmpnseries.codusu} =" & vUsu.codigo
        Else
            
            
                    indCodigo = Abs(Me.chkSitaucionArticulo2(0).Value) + Abs(Me.chkSitaucionArticulo2(1).Value) + Abs(Me.chkSitaucionArticulo2(2).Value) + Abs(Me.chkSitaucionArticulo2(3).Value)
                    If indCodigo < 4 Then
                        'Ha seleccionado 2 o uno
                        devuelve = ""
                        If Me.chkSitaucionArticulo2(0).Value = 1 Then devuelve = ", 0"
                        If Me.chkSitaucionArticulo2(1).Value = 1 Then devuelve = ", 1"
                        If Me.chkSitaucionArticulo2(2).Value = 1 Then devuelve = ", 2"
                        If Me.chkSitaucionArticulo2(3).Value = 1 Then devuelve = ", 3"
                        devuelve = Mid(devuelve, 2)
                        campo = "{sartic.codstatu} IN "
                        AnyadirAFormula cadFormula, campo & "[" & devuelve & "]"
                        AnyadirAFormula cadSelect, campo & "(" & devuelve & ")"
                    End If
                
                    numOp = PonerGrupo(1, ListView2.ListItems(1).Text)
                    If numOp <> 0 Then Opcion = numOp
                    numOp = PonerGrupo(2, ListView2.ListItems(2).Text)
                    If numOp <> 0 Then Opcion = numOp
                    numOp = PonerGrupo(3, ListView2.ListItems(3).Text)
                    If numOp <> 0 Then Opcion = numOp
                    numOp = PonerGrupo(4, ListView2.ListItems(4).Text)
                    If numOp <> 0 Then Opcion = numOp
                    Opcion = Opcion - 1
                
                    Select Case Opcion
                        Case 1 'El group2 es el Proveedor
                            campo = "pTitulo1=""" & ListView2.ListItems(3).Text & """"
                            CadParam = CadParam & campo & "|"
                            numParam = numParam + 1
                            
                            campo = "pTitulo2=""" & ListView2.ListItems(4).Text & """"
                            CadParam = CadParam & campo & "|"
                            numParam = numParam + 1
                        Case 2 'El Group3 es el Proveedor
                            campo = "pTitulo1=""" & ListView2.ListItems(2).Text & """"
                            CadParam = CadParam & campo & "|"
                            numParam = numParam + 1
                            
                            campo = "pTitulo2=""" & ListView2.ListItems(4).Text & """"
                            CadParam = CadParam & campo & "|"
                            numParam = numParam + 1
                        Case 3, 0 'El Group4 es el Proveedor
                                  '0 'El Group1 es el Proveedor
                            campo = "pTitulo1=""" & ListView2.ListItems(2).Text & """"
                            CadParam = CadParam & campo & "|"
                            numParam = numParam + 1
                            
                            campo = "pTitulo2=""" & ListView2.ListItems(3).Text & """"
                            CadParam = CadParam & campo & "|"
                            numParam = numParam + 1
                            
                            If Opcion = 0 Then
                                campo = "pTitulo3=""" & ListView2.ListItems(4).Text & """"
                                CadParam = CadParam & campo & "|"
                                numParam = numParam + 1
                            End If
                    End Select
                   
                    
                    cadTitulo = "Listado de Art�culos"
                    campo = "pOrden=" & Opcion
                    CadParam = CadParam & campo & "|"
                    numParam = numParam + 1
            End If 'de etiqueta o listado
    Case 18
        ''ElseIf OpcionListado = 18 Then
        'filtrar ademas por solo articulos con control de stock
        campo = "{sartic.ctrstock}=1"
        AnyadirAFormula cadFormula, campo
    
    
        'Los articulos cuya situacion NO este cadaducado, es decir, NORMAL y BLOQUEADO
        campo = "{sartic.codstatu}<3"
        AnyadirAFormula cadFormula, campo
    
        'Filtrar ademas por stock<stockMin o stock>stockMax
        campo = "{salmac.canstock}"
        If Me.optStockMax Then
            cadFormula = cadFormula & " AND (" & campo & "> {salmac.stockmax})"
        Else
            'David G 30/01/2007
            If optPuntoPedido.Value Then
                cadFormula = cadFormula & " AND (" & campo & "< {salmac.puntoped})"
            Else
                cadFormula = cadFormula & " AND (" & campo & "< {salmac.stockmin})"
            End If
        End If
    
        'En pedidos
        CargaDatosEnPedidos
        CadParam = CadParam & "codusu= " & vUsu.codigo & "|"
        numParam = numParam + 1
    
        'A�adir el Parametro de Stocks Maximos o Minimos
        If Me.optStockMax.Value = True Then
            campo = "0"
        Else
            If optPuntoPedido.Value Then
                campo = "2"
            Else
                campo = "1"
            End If
        End If
        CadParam = CadParam & "pStockMax=" & campo & "|"
        numParam = numParam + 1
    Case 247

        'Correccion de importes
        '-------------------------------------------------------
        
        If BloqueoManual("CORRIGEPRECIOS", "1") Then
            
            
        
            'Mostrare el list
            cadSelect = QuitarCaracterACadena(cadFormula, "{")
            cadSelect = QuitarCaracterACadena(cadSelect, "}")
            frmMensajes.cadWhere = cadSelect
            
            frmMensajes.OpcionMensaje = 16
                ' CORRECCION DE PRECIOS DE ARTICULOS QUE TIENEN COMPONENTES
            If vParamAplic.Produccion And Me.cmbProduccion.ListIndex = 1 Then frmMensajes.OpcionMensaje = 20
                
            frmMensajes.vCampos = txtCodigo(107).Text
            frmMensajes.cadWHERE2 = Trim(Me.cmbDecimales.Text)
            'Por no utilizar otra variable
            NumRegElim = 0
            If Me.chkMinimoCorreg.Value = 1 Then NumRegElim = 1
            frmMensajes.Show vbModal
       
        End If
        DesBloqueoManual ("CORRIGEPRECIOS")
        Exit Sub
    End Select
    
    
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    cadSelect = cadFormula
    If Not HayRegParaInforme(cadFrom, cadSelect) Then Exit Sub
        
        
    
        
    LlamarImprimir False
End Sub


Private Sub cmdAceptarAviPtes_Click()
'409: Listado Avisos averias pendientes
Dim tabla As String
Dim campo As String, Cad As String

    InicializarVbles
    
    '===================================================
    '============ PARAMETROS ===========================
    'A�adir el nombre de la Empresa como parametro
    CadParam = "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    tabla = "scaavi"
    cadTitulo = "Listado Avisos de aver�as Pendientes"
    cadNomRPT = "rRepAvisosPtes.rpt"
    conSubRPT = False
    
    
    '===================================================
    '================= FORMULA =========================
    'Cadena para seleccion D/H RUTA
    '----------------------------------
    If txtCodigo(84).Text <> "" Or txtCodigo(85).Text <> "" Then
        campo = "{sclien.codrutas}"
        Cad = "pDHRuta=""Rutas: "
        If Not PonerDesdeHasta(campo, "N", 84, 85, Cad) Then Exit Sub
    End If



    'Cadena para seleccion SITUACION
    '----------------------------------
    Cad = "pDHSitua=""Situaci�n: "
    If Me.cboSituaAviso.ListIndex = -1 Or Me.cboSituaAviso.ListIndex = 0 Then
        Cad = Cad & "Todas" & """|"
    Else
        Cad = Cad & Me.cboSituaAviso.List(Me.cboSituaAviso.ListIndex) & """|"
        campo = "{" & tabla & ".situacio}=" & Me.cboSituaAviso.ListIndex - 1
        
        If Not AnyadirAFormula(cadFormula, campo) Then Exit Sub
        If Not AnyadirAFormula(cadSelect, campo) Then Exit Sub
    End If
    CadParam = CadParam & Cad
    numParam = numParam + 1


    'Cadena para seleccion D/H FECHA
    '----------------------------------
    If txtCodigo(82).Text <> "" Or txtCodigo(83).Text <> "" Then
        campo = "{scaavi.fechaavi}"
        Cad = "pDHFecha=""Fecha: "
        If Not PonerDesdeHasta(campo, "F", 82, 83, Cad) Then Exit Sub
    End If


    'Cadena para seleccion D/H RUTA
    '----------------------------------
    If txtCodigo(96).Text <> "" Or txtCodigo(97).Text <> "" Then
        campo = "{scaavi.codtecni}"
        Cad = "pDHTecni=""T�cnico: "
        If Not PonerDesdeHasta(campo, "N", 96, 97, Cad) Then Exit Sub
    End If



    '==============================================================
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    tabla = tabla & " INNER JOIN sclien ON " & tabla & ".codclien=sclien.codclien"
    If Not HayRegParaInforme(tabla, cadSelect) Then Exit Sub
    
    LlamarImprimir False
End Sub

Private Sub cmdAceptarDtosFM_Click()
'54: Listado de Descuentos Familia/Marca
'309: Listado precio compras
Dim campo As String, Cad As String
Dim tabla As String

    InicializarVbles
    
    'JUNIO 2014
    'Para un cliente puede tener dto metidos para "el " en sdtofm o puede venir de su actividad
    If OpcionListado = 54 And optFrDto(5).Value Then
        HacerListadoDtosCliente
        Exit Sub
    End If
    
    
    '===================================================
    '============ PARAMETROS ===========================
    'A�adir el nombre de la Empresa como parametro
    CadParam = "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
        
    If OpcionListado = 54 Then
        tabla = "sdtofm"
        conSubRPT = True
    ElseIf OpcionListado = 309 Then
        tabla = "slispr"
        cadTitulo = "Listado Precios de compra"
        cadNomRPT = "rComPrecios.rpt"
        conSubRPT = False
    End If
    
    '===================================================
    '================= FORMULA =========================
    'Cadena para seleccion D/H FAMILIA
    '----------------------------------
    Orden1 = ""
    If txtCodigo(75).Text <> "" Or txtCodigo(76).Text <> "" Then
        campo = "{" & tabla & ".codfamia}"
        If OpcionListado = 309 Then campo = "{sartic.codfamia}"
        Cad = "Familia: "
        If Not PonerDesdeHasta(campo, "N", 75, 76, Cad) Then Exit Sub
        Orden1 = Cad
    End If
    If OpcionListado = 309 Then
        If Me.chkVarios(4).Value = 1 Then
            'CABEL
            Set miRsAux = New ADODB.Recordset
            campo = ""
            miRsAux.Open "Select codfamia from sfamia where marcapropia=1", conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not miRsAux.EOF
                campo = campo & ", " & miRsAux!Codfamia
                miRsAux.MoveNext
            Wend
            miRsAux.Close
            Set miRsAux = Nothing
            
            If campo <> "" Then campo = Mid(campo, 2)
            campo = " {sartic.codfamia} IN [" & campo & "]"
                        
            If cadFormula <> "" Then campo = " AND " & campo
            
            'rpt
            cadFormula = cadFormula & campo
            'sql
            campo = Replace(campo, "[", "(")
            campo = Replace(campo, "]", ")")
            cadSelect = cadSelect & campo
            
            Orden1 = Trim(Orden1 & "   [** CABEL **]")
        End If
    End If
    If Orden1 <> "" Then
        Cad = "pDHFamilia=""" & Orden1 & " |"
        CadParam = CadParam & Cad
        numParam = numParam + 1
        Orden1 = ""
    End If
    



    If OpcionListado = 54 Then
        'Cadena para seleccion D/H CLIENTE
        '--------------------------------------------
        If txtCodigo(73).Text <> "" Or txtCodigo(74).Text <> "" Then
            campo = "{sdtofm.codclien}"
            Cad = "pDHCliente=""Cliente: "
            If Not PonerDesdeHasta(campo, "N", 73, 74, Cad) Then Exit Sub
            
            If Me.optFrDto(1).Value Then
                'Va a mostrar por actividad. NO debe poner desde hasta cliente
                MsgBox "Va a mostrar los datos por actividad. No debe poner D/H cliente", vbExclamation
                Exit Sub
            End If
        End If
    
    
        'Cadena para seleccion D/H MARCA
        '--------------------------------------------
        cadNomRPT = ""
        If txtCodigo(77).Text <> "" Or txtCodigo(78).Text <> "" Then
            campo = "{sdtofm.codmarca}"
            Cad = "Marca: "
            If Not PonerDesdeHasta(campo, "N", 77, 78, Cad) Then Exit Sub
            cadNomRPT = Cad
        End If
        
        If Me.cboVarios(0).ListIndex > 0 Then
            campo = "dtoesp"
            Cad = "0"                            'Dto especia
            If cboVarios(0).ListIndex = 1 Then Cad = "1"
                
                
            If cadSelect <> "" Then cadSelect = cadSelect & "  AND "
            If cadFormula <> "" Then cadFormula = cadFormula & " AND  "
            cadSelect = cadSelect & campo & " = " & Cad
            cadFormula = cadFormula & " ({sdtofm." & campo & "} = " & Cad & ")"
            
            Cad = "sin "                            'Dto especia
            If cboVarios(0).ListIndex = 1 Then Cad = " SOLO "
            cadNomRPT = Trim(cadNomRPT & "     (" & Cad & " dto especial)")
        End If
        
        Cad = "pDHMarca=""" & cadNomRPT & """"
        CadParam = CadParam & Cad & "|"
        numParam = numParam + 1
        cadNomRPT = ""
        
    End If
    'Cadena para seleccion D/H PROVEEDOR
    '--------------------------------------------
    Cad = "pDHProveedor="""
    If OpcionListado = 309 Then
           If Me.chkVarios(2).Value Then Cad = Cad & "[ROTACION]   "
    End If
    
    If txtCodigo(79).Text <> "" Or txtCodigo(80).Text <> "" Then
        If OpcionListado = 54 Then
            campo = "{sfamia.codprove}"
        Else
            campo = "{" & tabla & ".codprove}"
        End If
        
        Cad = Cad & "Proveedor: "
        If Not PonerDesdeHasta(campo, "N", 79, 80, Cad) Then Exit Sub
    Else
        If OpcionListado = 309 Then
            numParam = numParam + 1
            CadParam = CadParam & Cad & " ""|"
        End If
    End If
    Cad = ""
    
    '==============================================================
    If OpcionListado = 54 Then
        
        'En herbelca NO dejo continuar si no pone algun desde hasta
        If vParamAplic.AlmacenB > 90 Then
            If cadSelect = "" Then
                MsgBox "Escriba algun criterio de busqueda", vbExclamation
                Exit Sub
            End If
        End If
        
        If Me.optFrDto(0).Value Then
            cadNomRPT = "rFacDtosFM.rpt"
            campo = "codclien"
        ElseIf Me.optFrDto(1).Value Then
            cadNomRPT = "rFacDtosFMAct.rpt"
            campo = "codactiv"
        ElseIf Me.optFrDto(4).Value Then
            'Nuevo proveedor
            cadNomRPT = "rFacDtosFMprov.rpt"
            campo = "codclien"
        
        Else
            campo = ""
            If Me.optFrDto(2).Value Then
                cadNomRPT = "rFacDtosFMF.rpt"
            Else
                cadNomRPT = "rFacDtosFMM.rpt"
            End If
        End If
        If campo <> "" Then
            'dtofm
            If cadSelect <> "" Then cadSelect = cadSelect & "  AND "
            If cadFormula <> "" Then cadFormula = cadFormula & " AND  "
            cadSelect = cadSelect & campo & " > 0"
            cadFormula = cadFormula & " ({sdtofm." & campo & "}>0)"
        End If
    End If
    

    If OpcionListado = 54 Then
        If cadSelect <> "" Then cadSelect = cadSelect & "  AND "
        If cadFormula <> "" Then cadFormula = cadFormula & " AND  "
         cadSelect = cadSelect & " OcultarEnListDto = 0"
        cadFormula = cadFormula & " ({sprove.OcultarEnListDto} = 0)"
    End If
    
    '==============================================================
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    If OpcionListado = 309 Then
        tabla = tabla & " INNER JOIN sartic ON " & tabla & ".codartic=sartic.codartic"
    Else
        'dtos fm
        tabla = tabla & " INNER JOIN sfamia ON " & tabla & ".codfamia=sfamia.codfamia"
        
        'Julio 2013
        'OcultarEnListDto
        tabla = tabla & " INNER JOIN sprove ON `sfamia`.`codprove`=`sprove`.`codprove`"
        
    End If
    
    'Cotubre 2014
    If OpcionListado = 309 Then
        If Me.chkVarios(2).Value Then
            campo = "rotacion"
            If cadSelect <> "" Then cadSelect = cadSelect & "  AND "
            If cadFormula <> "" Then cadFormula = cadFormula & " AND  "
            cadSelect = cadSelect & campo & " =1 "
            cadFormula = cadFormula & " ({sartic." & campo & "}=1)"
        End If

    End If
    
    If Not HayRegParaInforme(tabla, cadSelect) Then Exit Sub
    
    
    
    If OpcionListado = 309 Then
        If Me.chkVarios(1).Value Then
            'Cargaremos tmpInformes con el sdtopm aplicado sobre el precio del articulo
            Orden1 = tabla
            HazCalculoPrecioNetoProve
            
            cadTitulo = "Precio neto proveedor"
            cadFormula = "({tmpinformes.codusu} = " & vUsu.codigo & ")"
            cadNomRPT = "rComPreciosNeto.rpt"
        End If
    End If
    LlamarImprimir False
End Sub


Private Sub cmdAceptarEst_Click()
'Estadistica margen ventas por articulo
Dim campo As String
Dim param As String
Dim tabla As String
Dim opcPrecio As String
Dim desPrecio As String

    InicializarVbles
    
    '===================================================
    '============ PARAMETROS ===========================
    'A�adir el nombre de la Empresa como parametro
    CadParam = CadParam & "|pNomEmpre=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    
    If chkMargen(0).Value = 0 Then
        'Para los que detalla fracion NO envio pDetalla
        CadParam = CadParam & "Detalla= " & chkMargen(1).Value & "|"
        numParam = numParam + 1
    End If
    
    ' Septiembre 2015
    ' Sin marcar (como estaba), marcado sobre las ventas(Herbelca)... y creo que es lo mas logico
    tabla = 1
    If chkMargen(4).Value = 1 Then tabla = "0"   'paremtro en rpt: MargenSobreCoste
    'Para los que detalla fracion NO envio pDetalla
    CadParam = CadParam & "MargenSobreCoste= " & tabla & "|"
    numParam = numParam + 1

    
    
    'Cadena para seleccion D/H fecha.   Lo metere junto con la seelccion del prico a copger
    '--------------------------------------------
    param = ""
    If txtCodigo(130).Text <> "" Or txtCodigo(131).Text <> "" Then
        campo = "{slifac.fecfactu}"
        param = "Fecha: "
        If Not PonerDesdeHasta(campo, "F", 130, 131, param) Then Exit Sub
        
    End If
    
    
    
    'Parametro Precio de Valoracion
    'elegir un Precio para realizar la valoracion
    '==================================================
    desPrecio = "Valoraci�n coste: "
    If Me.optPrecioMP2.Value Then
        opcPrecio = "{slifac.preciomp}" 'precio medio ponderado
        desPrecio = desPrecio & "Precio medio ponderado"
    ElseIf Me.optPrecioUC2.Value Then
        opcPrecio = "{slifac.preciouc}" 'precio ultima compra
        desPrecio = desPrecio & "Precio �ltima compra"
    ElseIf Me.optPrecioStd2.Value Then
        opcPrecio = "{slifac.preciost}" 'precio standard
        desPrecio = desPrecio & "Precio standard"
    End If
    
    'Mayo 2013
    If chkMargen(2).Value = 1 Then
          'Va a incluir los articulos de varios
          'Luevo en el SQL NO pongo nada pero lo indico en los campos d desd/hastas
          param = Trim(param & "      [Art. Varios]")
    Else
          AnyadirAFormula cadFormula, " {sartic.artvario}=0"
          AnyadirAFormula cadSelect, " sartic.artvario=0"
    End If
    
    CadParam = CadParam & "pCampo=" & opcPrecio & "|"
    'Le pong las fechas(si es k las han puesto)
    desPrecio = Trim(desPrecio & "          " & param)
    If chkMargen(4).Value = 1 Then desPrecio = desPrecio & "[% Sobre vta]"
    CadParam = CadParam & "pDesCampo=""" & desPrecio & """|"
    numParam = numParam + 2
    
    
    '===================================================
    '================= FORMULA =========================
    'Cadena para seleccion D/H familia
    '--------------------------------------------
    If txtCodigo(88).Text <> "" Or txtCodigo(89).Text <> "" Then
        campo = "{sartic.codfamia}"
        param = "pDHFamilia=""Familia: "
        If Not PonerDesdeHasta(campo, "N", 88, 89, param) Then Exit Sub
    End If
    
    'Cadena para seleccion D/H art�culo
    '--------------------------------------------
    If txtCodigo(90).Text <> "" Or txtCodigo(91).Text <> "" Then
        campo = "{slifac.codartic}"
        param = "pDHArticulo=""Art�culo: "
        If Not PonerDesdeHasta(campo, "T", 90, 91, param) Then Exit Sub
    End If
    
    
    
    
    
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    tabla = " slifac INNER JOIN sartic ON slifac.codartic=sartic.codartic "
    If Not HayRegParaInforme(tabla, cadSelect) Then Exit Sub
    
    If chkMargen(0).Value = 0 Then
        'El de antes
        cadNomRPT = "rFacEstMargen"
    Else
        If chkMargen(1).Value = 0 Then
            cadNomRPT = "rFacEstMargenDetaFra"
        Else
            cadNomRPT = "rFacEstMargenDetaFraArt"
        End If
    End If
    If Me.chkMargen(3).Value = 1 Then cadNomRPT = cadNomRPT & "P"
    
    cadNomRPT = cadNomRPT & ".rpt"
    
    LlamarImprimir False
     
End Sub

Private Sub cmdAceptarFichas_Click()
'Fichas de Mantenimientos
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim campo As String
Dim ParaClientemante As String
    
    If chkMante(4).Value = 0 Then
        'Enero 2011
        'No puede poner tipo contrato pq ahora estamos empezando el link por sserie, no por scaman
        If txtCodigo(57).Text <> "" Or txtCodigo(58).Text <> "" Then
            MsgBox "Solo puede poner tipo mantenimiento para la opcion INFORME COMPLETO", vbExclamation
            Exit Sub
        End If
    End If
    
    InicializarVbles
    
    
    '===================================================
    '============ PARAMETROS ===========================
    CadParam = "|"
    
    
    If chkMante(4).Value = 1 Then
        'Enero 2010
        'Informe completo
        indRPT = 38
    Else
        indRPT = 13
    End If
    If Not PonerParamRPT2(indRPT, CadParam, numParam, cadNomRPT, pImprimeDirecto, pPdfRpt, pRptvMultiInforme) Then Exit Sub
    'Ejercicio
    CadParam = CadParam & "pEjercicio=""" & txtCodigo(61).Text & """|"
    numParam = numParam + 1
    
    
    '===================================================
    '================= FORMULA =========================
    'Cadena para seleccion D/H CLIENTE
    '--------------------------------------------
    If txtCodigo(55).Text <> "" Or txtCodigo(56).Text <> "" Then
        'Desde mantenimientos
        If chkMante(4).Value = 1 Then
            campo = "{scaman.codclien}"
        Else
            'Saca los datos de los numeros de serie
            campo = "{sserie.codclien}"
        End If
        If Not PonerDesdeHasta(campo, "N", 55, 56, "") Then Exit Sub
    End If
    
    'Cadena para seleccion D/H TIPO CONTRATO
    '--------------------------------------------
    If txtCodigo(57).Text <> "" Or txtCodigo(58).Text <> "" Then
        campo = "{scaman.codtipco}"
        If Not PonerDesdeHasta(campo, "T", 57, 58, "") Then Exit Sub
    End If
    
    'Cadena para seleccion D/H N� Mantenimiento
    '--------------------------------------------
    If txtCodigo(59).Text <> "" Or txtCodigo(60).Text <> "" Then
        'Desde mantenimientos
        If chkMante(4).Value = 1 Then
            campo = "{scaman.nummante}"
        Else
            campo = "{sserie.nummante}"
        End If
        If Not PonerDesdeHasta(campo, "T", 59, 60, "") Then Exit Sub
    End If
    
    
    'Cadena para seleccion D/H RUTA
    '--------------------------------------------
    If txtCodigo(106).Text <> "" Or txtCodigo(108).Text <> "" Then
        campo = "{sclien.codrutas}"
        If Not PonerDesdeHasta(campo, "N", 106, 108, "") Then Exit Sub
    End If
    
    
    
    If chkMante(4).Value = 0 Then
        If cadFormula <> "" Then cadFormula = cadFormula & " AND "
        cadFormula = cadFormula & " {sserie.nummante} <> """""
        cadFormula = cadFormula & " AND {sserie.nummante} <> ""IRREPARABL"""
        cadFormula = cadFormula & " AND {sserie.nummante} <> ""BAJA"""
        cadFormula = cadFormula & " AND {sserie.nummante} <> ""S/MTO"""
        'RECOMPRADA
        cadFormula = cadFormula & " AND {sserie.nummante} <> ""RECOMPRA"""
        cadFormula = cadFormula & " AND {sserie.nummante} <> ""RECOMPRADA"""
        cadFormula = cadFormula & " AND {sserie.nummante} <> ""RECOMPRADO"""
        
        ',BAJA,S/MTO
    End If
    
    
    'Solo los que tienen mto contratado
    '
    '
    campo = cadFormula
    campo = Replace(campo, "sserie.", "scaman.")
    campo = Replace(campo, "{", "(")
    campo = Replace(campo, "}", ")")
    ParaClientemante = "Select scaman.codclien from scaman,sclien where scaman.codclien=sclien.codclien and " & campo & " GROUP BY 1"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open ParaClientemante, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    ParaClientemante = ""
    While Not miRsAux.EOF
        ParaClientemante = ParaClientemante & ", " & miRsAux!codClien
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    If ParaClientemante <> "" Then ParaClientemante = Mid(ParaClientemante, 2)
    
    
    
    
    '==============================================================
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    cadSelect = cadFormula
    If ParaClientemante <> "" Then
        cadSelect = cadSelect & " AND sclien.codclien IN (" & ParaClientemante & ")"
        cadFormula = cadFormula & " AND {sclien.codclien} IN [" & ParaClientemante & "]"
    End If
    ' ---- [30/10/2009] [LAURA]
'    campo = "   (`ariges2`.`scaman` `scaman` INNER JOIN `ariges2`.`sclien` `sclien` ON `scaman`.`codclien`=`sclien`.`codclien`) INNER JOIN `ariges2`.`stipco` `stipco` ON `scaman`.`codtipco`=`stipco`.`codtipco`"
    If chkMante(4).Value = 1 Then
        campo = "(scaman INNER JOIN sclien ON scaman.codclien=sclien.codclien) INNER JOIN stipco ON scaman.codtipco=stipco.codtipco"
    Else
        campo = " sserie INNER JOIN sclien ON sserie.codclien=sclien.codclien "
    End If
    
    
    
    
    ' ----
    
    If Not HayRegParaInforme(campo, cadSelect) Then Exit Sub
    
    'Si detalla articulos o no
    CadParam = CadParam & "ImprimeArticulo=" & Abs(Me.chkMante(1).Value) & "|"
    numParam = numParam + 1
    LlamarImprimir True
End Sub


Private Sub cmdAceptarMante_Click()
'Listado de Mantenimientos
Dim cadFrom As String
Dim campo As String, devuelve As String
Dim codigo  As String

    InicializarVbles
    cadFrom = ""
    
    Select Case OpcionListado
    Case 70, 76
        'comprobar que se ha seleccionado un Tipo de Informe
        If Me.cboTipoList.ListIndex = -1 Then Exit Sub
        'En funcion del valor seleccionado en Tipo Informe se abrira un listado diferente
        Select Case Me.cboTipoList.ListIndex
            Case 0 'Listado Equipos
                cadNomRPT = "rManListManEquipo"
            Case 1 'Listado Pagos
                cadNomRPT = "rManListManPago"
            Case 2 'Listado Importes Contrato
                cadNomRPT = "rManListManImporte"
        End Select
        
        cadTitulo = "Informe Mantenimientos"
        codigo = "scaman"
        If OpcionListado = 76 Then
            'ANULADOS    rManListManImporteAnu.rpt
            cadTitulo = cadTitulo & " Anulados"
            codigo = codigo & "a"
            cadNomRPT = cadNomRPT & "Anu"
        End If
        cadNomRPT = cadNomRPT & ".RPT"
    Case 71
        cadNomRPT = "rManListRevisiones.rpt"
        codigo = "scaman"
        cadTitulo = "Informe Revisiones"
    Case 78
    
        'PEque�a comprobacion.
        'Fecha obligatoria
        If txtCodigo(109).Text = "" Then
            MsgBox "Debe indicar la fecha", vbExclamation
            Exit Sub
        End If
    
    
        If Not PonerParamRPT2(21, CadParam, numParam, cadNomRPT, pImprimeDirecto, pPdfRpt, pRptvMultiInforme) Then Exit Sub
        codigo = "scaman"
    Case 79
        If Not PonerParamRPT2(45, CadParam, numParam, cadNomRPT, pImprimeDirecto, pPdfRpt, pRptvMultiInforme) Then Exit Sub
        codigo = "scaman"
    End Select
    cadFrom = "(" & codigo & " INNER JOIN sclien ON " & codigo & ".codclien=sclien.codclien) "
      
      
      
    '===================================================
    '============ PARAMETROS ===========================
    'A�adir el nombre de la Empresa como parametro
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
      
      
      
    '===================================================
    '================= FORMULA =========================
    'Cadena para seleccion ZONA
    '--------------------------------------------
    If txtCodigo(45).Text <> "" Or txtCodigo(46).Text <> "" Then
        campo = "{sclien.codzonas}"
'        'Parametro Desde/Hasta Zona
        devuelve = "pDHZona=""Zona: "
        If Not PonerDesdeHasta(campo, "N", 45, 46, devuelve) Then Exit Sub
    End If
    
    'Cadena para seleccion CLIENTE
    '--------------------------------------------
    If txtCodigo(47).Text <> "" Or txtCodigo(48).Text <> "" Then
        campo = "{" & codigo & ".codclien}"
        'Parametro Desde/Hasta Cliente
        devuelve = "pDHCliente=""Cliente: "
        If Not PonerDesdeHasta(campo, "N", 47, 48, devuelve) Then Exit Sub
    End If
    
    'Cadena para seleccion AGENTE
    '--------------------------------------------
    If txtCodigo(49).Text <> "" Or txtCodigo(50).Text <> "" Then
        campo = "{sclien.codagent}"
        'Parametro Desde/Hasta Agente
        devuelve = "pDHAgente=""Agente: "
        If Not PonerDesdeHasta(campo, "N", 49, 50, devuelve) Then Exit Sub
    End If
    
    'Cadena para seleccion TIPO CONTRATO y si la opcion es 70: ruta
    '--------------------------------------------
    Orden1 = ""
    If txtCodigo(51).Text <> "" Or txtCodigo(52).Text <> "" Then
        campo = "{" & codigo & ".codtipco}"
        'Parametro Desde/Hasta Tipo Contrato
        'devuelve = "pDHTipoCon=""Tipo Contrato: "
        devuelve = "Tipo Contrato: "
        If Not PonerDesdeHasta(campo, "T", 51, 52, devuelve) Then Exit Sub
        Orden1 = Replace(devuelve, """", "")
    End If
    If OpcionListado = 70 Then
        'Cadena para seleccion D/H RUTA
        '----------------------------------
        If txtCodigo(137).Text <> "" Or txtCodigo(138).Text <> "" Then
            campo = "{sclien.codrutas}"
            devuelve = "    Rutas: "
            If Not PonerDesdeHasta(campo, "N", 137, 138, devuelve) Then Exit Sub
        End If
        devuelve = Replace(devuelve, """", "")
        Orden1 = Trim(Orden1 & "    " & devuelve)
    End If
        
    If Orden1 <> "" Then
        CadParam = CadParam & "pDHTipoCon=""" & Orden1 & """|"
        numParam = numParam + 1
    End If
    
    
    
    'Motivo de baja. Solo para anulados
    If OpcionListado = 76 Then
        If txtCodigo(115).Text <> "" Or txtCodigo(116).Text <> "" Then
            campo = "{scamana.fechabaj}"
            devuelve = "pDHFecha=""Fecha: "
            If Not PonerDesdeHasta(campo, "F", 115, 116, devuelve) Then Exit Sub
        End If
    
        If txtCodigo(113).Text <> "" Or txtCodigo(114).Text <> "" Then
            campo = "{" & codigo & ".codincid}"
            'Parametro Desde/Hasta Cliente
            devuelve = "pDHMotivo=""Motivo anul.: "
            If Not PonerDesdeHasta(campo, "T", 113, 114, devuelve) Then Exit Sub
        End If
        
        
    ElseIf OpcionListado = 79 Then 'solo para Etiquetas
        'Cadena para seleccion ACTIVIDAD
        '--------------------------------------------
        If txtCodigo(127).Text <> "" Or txtCodigo(128).Text <> "" Then
            campo = "{sclien.codactiv}"
            'Parametro Desde/Hasta Actividad
            devuelve = "pDHActividad=""Actividad: "
            If Not PonerDesdeHasta(campo, "N", 127, 128, devuelve) Then Exit Sub
        End If
        
        'Cadena para seleccion COD. POSTAL
        '--------------------------------------------
         If txtCodigo(129).Text <> "" Then
            campo = "{sclien.codpobla}"
            'Parametro Desde/Hasta Actividad
            devuelve = "pCodPosta=""C. Postal: " & txtCodigo(129).Text
            AnyadirAFormula cadFormula, campo & "=" & DBSet(txtCodigo(129).Text, "T")
            AnyadirAFormula cadSelect, campo & "=" & DBSet(txtCodigo(129).Text, "T")
'            If Not PonerDesdeHasta(campo, "N", 127, 128, devuelve) Then Exit Sub
         End If
    ElseIf OpcionListado = 78 Then
        'Cadena para seleccion ACTIVIDAD
        '--------------------------------------------
        If txtCodigo(127).Text <> "" Or txtCodigo(128).Text <> "" Then
            campo = "{sclien.codactiv}"
            'Parametro Desde/Hasta Actividad
            devuelve = "pDHActividad=""Actividad: "
            If Not PonerDesdeHasta(campo, "N", 127, 128, devuelve) Then Exit Sub
        End If
    End If
    
    'Cadena para seleccion FECHA
    '--------------------------------------------
    If OpcionListado = 71 Then
        If txtCodigo(53).Text = "" Or txtCodigo(54).Text = "" Then
            MsgBox "Los campos Fecha Desde/Hasta deben tener valor", vbInformation
            Exit Sub
        End If
        If txtCodigo(53).Text <> "" Then
            devuelve = "Date(" & Year(txtCodigo(53).Text) & "," & Month(txtCodigo(53).Text) & "," & Day(txtCodigo(53).Text) & ")"
            'Parametro D/H Fecha
            If devuelve <> "" Then
                devuelve = "pDFecha=" & devuelve & "|"
                CadParam = CadParam & devuelve & """|"
                numParam = numParam + 1
            End If
        End If
        
        If txtCodigo(54).Text <> "" Then
            devuelve = "Date(" & Year(txtCodigo(54).Text) & "," & Month(txtCodigo(54).Text) & "," & Day(txtCodigo(54).Text) & ")"
            If devuelve <> "" Then
                devuelve = "pHFecha=" & devuelve & "|"
                CadParam = CadParam & devuelve & """|"
                numParam = numParam + 1
            End If
        End If
    End If
        
        

        
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    'cadSelect = cadFormula
    If Not HayRegParaInforme(cadFrom, cadSelect) Then Exit Sub
    
    
    'Esto lo hago siempre para gene temporales
    conn.Execute "DELETE from tmpnlotes where codusu =" & vUsu.codigo
    
    If OpcionListado = 79 Then

        ' ---- [30/10/2009] (LAURA) : agrupar por cliente y departamento
        devuelve = "Select scaman.codclien,nomclien,nifclien,scaman.coddirec,nomdirec"
        devuelve = devuelve & " FROM " & cadFrom
        devuelve = devuelve & " LEFT OUTER JOIN sdirec ON scaman.codclien=sdirec.codclien and scaman.coddirec=sdirec.coddirec"
        If cadSelect <> "" Then devuelve = devuelve & " WHERE " & cadSelect
        devuelve = devuelve & " group by scaman.codclien,scaman.coddirec"
        
        ' ---- [ANTES] Mostraremos los clientes para imprimirles etiquetas
'        If cadSelect <> "" Then
'            devuelve = " WHERE " & cadSelect
'        Else
'            devuelve = ""
'        End If
'        devuelve = "Select sclien.codclien,nomclien,nifclien FROM " & cadFrom & devuelve
'        devuelve = devuelve & " group by 1"
        ' ----
        
        
        
        NumRegElim = 0
        frmMensajes.cadWhere = devuelve
        frmMensajes.OpcionMensaje = 17 'Etiquetas clientes mantenimientos
        frmMensajes.Show vbModal
        If NumRegElim = 0 Then Exit Sub
    
        'cadFormula = "({tmpnlotes.codusu} =" & vUsu.Codigo & ")"
        cadFormula = "({tmpinformes.codusu} = " & vUsu.codigo & ")"
        conSubRPT = True
    End If
    devuelve = ""
    If OpcionListado = 78 Then
        'A�ado la fecha
        CadParam = CadParam & "|FechaImp=""" & txtCodigo(109).Text & """|"
        numParam = numParam + 1
    
    
    
        If Me.chkMante(2).Value Then devuelve = "EMAIL"
    End If
    
    If devuelve = "" Then
        LlamarImprimir True
    Else

        '------------------------------------------------------------
        'Envio por mail del desde hasta se  leccionado
        'Comprobaremos los mail, que todos tienen
        'FALTA###
        
        
       
       
        DoEvents
        If Me.optMante(0).Value Then
            devuelve = "1"
        Else
            devuelve = "2"
        End If
        
        devuelve = "Select maiclie" & devuelve & " as el_mail,nomclien,scaman.* "
        devuelve = devuelve & " FROM  scaman INNER JOIN sclien ON scaman.codclien=sclien.codclien"
        If cadSelect <> "" Then devuelve = devuelve & " AND " & cadSelect
        
        'INNER JOIN `ariges2`.`stipco` `stipco` ON `scaman`.`codtipco`=`stipco`.`codtipco`"
        Set miRsAux = New ADODB.Recordset
        miRsAux.Open devuelve, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    
        devuelve = ""
        NumRegElim = 0
        While Not miRsAux.EOF
            If IsNull(miRsAux!el_mail) Then
                devuelve = devuelve & "    - " & miRsAux!Nomclien & vbCrLf
            Else
                'INSERTAMOS
                NumRegElim = NumRegElim + 1
                codigo = "insert into tmpnlotes (codusu,numalbar,fechaalb,codprove,numlinea,codartic) values ("
                codigo = codigo & vUsu.codigo & ",1,'" & Format(txtCodigo(109).Text, FormatoFecha) & "'," & miRsAux!codClien & ","
                codigo = codigo & NumRegElim & ",'" & miRsAux!nummante & "')"
                conn.Execute codigo
            End If
            miRsAux.MoveNext
        Wend
        miRsAux.Close
    
        If NumRegElim = 0 Then
            MsgBox "No hay datos para poder enviar por email", vbExclamation
            Exit Sub
        End If
        
        
        If devuelve <> "" Then
            If Len(devuelve) > 500 Then devuelve = Mid(devuelve, 1, 500) & " ....."
            devuelve = "Clientes sin mail: " & vbCrLf & devuelve & "�Continuar?"
            If MsgBox(devuelve, vbQuestion + vbYesNo) = vbNo Then Exit Sub
        End If
        
        
        
        If Not PrepararCarpetasEnvioMail Then Exit Sub
            
        
        PonerTamnyosMail True
        frmPpal.visible = False
        'Voy arriesgar.
        'Confio en que no envien por mail mas de 32000 facturas (un integer)
        Label4(22).Caption = "Preparando datos"
        Me.PBMail.Max = CInt(NumRegElim)
        Me.PBMail.Value = 0
        
        
        
        NumRegElim = 0
        If GeneracionEnvioMail() Then NumRegElim = 1
            
    
        'Si ha ido todo bien entonces numregelim=1
        If NumRegElim = 1 Then
            'Procederemos a enviarlos por mail
            If Me.optMante(0).Value Then
                '1
                cadSelect = "1"  'de maiclie2
            Else
                cadSelect = "2"  'de maiclie1
            End If
            cadSelect = "Select nomclien,maiclie" & cadSelect
            cadSelect = cadSelect & " as email,tmpnlotes.* from tmpnlotes,sclien where codusu = " & vUsu.codigo & " and codclien=codprove"
        
            
            frmEMail.DatosEnvio = "Carta renovacion|Muchas gracias|" & Abs(chkMante(3).Value) & "|" & cadSelect & "|"
            frmEMail.Opcion = 5 'Multienvio de renovacion
            frmEMail.Show vbModal
            
            
            'Para tranquilizar las pantallas, borrar los ficheros generados
            'Confio en que no envien por mail mas de 32000 facturas (un integer)
            Label14(22).Caption = "Restaurando ...."
            Me.ProgressBarContab.visible = False
            Me.Refresh
            DoEvents
            Espera 1
            PrepararCarpetasEnvioMail
            Me.ProgressBarContab.visible = True
            
            
        End If
        
        
        
        
        'Es para evitar la cantidad de pantallas abriendose y cerrandose
        Me.visible = False
        PonerTamnyosMail False
        Espera 1
        Unload Me
        frmPpal.Show
    
        Screen.MousePointer = vbDefault
    
    
    End If
    
    
End Sub





Private Sub cmdAceptarNSerie_Click()
Dim campo As String
Dim Cad As String

    If txtCodigo(37).Text = "" Or txtCodigo(38).Text = "" Then 'And (txtCodigo(33).Text = "" Or txtCodigo(34).Text = "") Then
        MsgBox "Debe seleccionar un cliente para Imprimir.", vbInformation
        PonerFoco txtCodigo(37)
        Exit Sub
    End If
    
    InicializarVbles
    
    cadNomRPT = "rRepNumSerie.rpt"  'Informe Numeros de Serie Articulos
    cadTitulo = "Informe Num. Serie"
    
    '===================================================
    '============ PARAMETROS ===========================
    'A�adir el nombre de la Empresa como parametro
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    
    '===================================================
    '================= FORMULA =========================
    'Cadena para seleccion CLIENTE
    '--------------------------------------------
    If txtCodigo(37).Text <> "" Or txtCodigo(38).Text <> "" Then
        campo = codigo & ".codclien}"
        'Parametro Desde/Hasta Cliente
        Cad = "pDHCliente=""Cliente: "
        If Not PonerDesdeHasta(campo, "N", 37, 38, Cad) Then Exit Sub
    End If
    
    'Cadena para seleccion Del DEPARTAMENTO
    '--------------------------------------------
    If txtCodigo(39).Text <> "" Or txtCodigo(40).Text <> "" Then
        campo = codigo & ".coddirec}"
        'Parametro Desde/Hasta Direc/Dpto
        If vParamAplic.HayDeparNuevo = 1 Then
            Cad = "pDHDirec=""Dpto.: "
        ElseIf vParamAplic.HayDeparNuevo = 0 Then
            Cad = "pDHDirec=""Direc.: "
        Else
            Cad = "pDHDirec=""Obra: "
        End If
        If Not PonerDesdeHasta(campo, "N", 39, 40, Cad) Then Exit Sub
    End If
    
    'Cadena para seleccion N� CONTRATO
    '--------------------------------------------
    If txtCodigo(41).Text <> "" Or txtCodigo(42).Text <> "" Then
        campo = codigo & ".nummante}"
        'Parametro Desde/Hasta Cliente
        Cad = "pDHContrato=""N� Mantenimiento: "
        If Not PonerDesdeHasta(campo, "T", 41, 42, Cad) Then Exit Sub
    End If
    
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    cadSelect = cadFormula
    If Not HayRegParaInforme("sserie", cadSelect) Then Exit Sub
    
    
    
    LlamarImprimir False
    
End Sub


Private Sub cmdAceptarRepxClien_Click()
'Reparaciones por Cliente
Dim devuelve As String
Dim campo As String
Dim tabla As String

    InicializarVbles
    
    If OpcionListado = 406 Then 'Frecuencia de reparaciones
        tabla = "schrep"
    Else
        tabla = "scarep"
    End If
    'David Enero 2010
    tabla = "schrep"  'siempre va con el hco
    
    '===================================================
    '============ PARAMETROS ===========================
    'A�adir el nombre de la Empresa como parametro
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    
    '===================================================
    '================= FORMULA =========================
    'Cadena para seleccion Desde y Hasta CLIENTE
    '---------------------------------------------
    If txtCodigo(33).Text <> "" Or txtCodigo(34).Text <> "" Then
        campo = "{" & tabla & ".codclien}"
        devuelve = "pDHCliente=""Cliente: "
        If Not PonerDesdeHasta(campo, "N", 33, 34, devuelve) Then Exit Sub
    End If
    
    'Cadena para seleccion Desde y Hasta DIREC/DPTO
    '-----------------------------------------------
    If txtCodigo(35).Text <> "" Or txtCodigo(36).Text <> "" Then
        campo = "{" & tabla & ".coddirec}"
        If vParamAplic.HayDeparNuevo Then
            devuelve = "pDHDpto=""Departamento: "
        Else
            devuelve = "pDHDpto=""Direcci�n: "
        End If
        If Not PonerDesdeHasta(campo, "N", 35, 36, devuelve) Then Exit Sub
    End If
    
    
    'Cadena para seleccion Desde y Hasta FECHA
    '----------------------------------------------
    If Trim(txtCodigo(43).Text) <> "" Or Trim(txtCodigo(44).Text) <> "" Then
        'ANTES
        'campo = "{" & tabla & ".fecentre}"
        'Marzo 2010
        'Fecha reparacion la tengo en la fechaalb
        campo = "{" & tabla & ".fechaalb}"
        'If OpcionListado = 406 Then campo = "{" & tabla & ".fecrepar}"
        devuelve = "pDHFecha=""Fecha Rep.: "
        If Not PonerDesdeHasta(campo, "F", 43, 44, devuelve) Then Exit Sub
    End If
    
   'Comprobar si hay registros a Mostrar antes de abrir el Informe
   If Not HayRegParaInforme(tabla, cadSelect) Then Exit Sub
    
    
    If OpcionListado <> 406 Then
        cadTitulo = "Reparaciones por Cliente"
        cadNomRPT = "rRepReparxClien.rpt"
        conSubRPT = True
    Else
        cadTitulo = "Frecuencia de Reparaciones"
        cadNomRPT = "rRepFrecuencia.rpt"
        conSubRPT = True
        
        'N� de Reparaciones, A�adirlo como parametro
        '----------------------------------------------
        CadParam = CadParam & "pNumVeces=" & txtCodigo(0).Text & "|"
        numParam = numParam + 1
        
        On Error GoTo EFrecu
        'Insertar en la tabla temporal tmpInformes el total de reparaciones para cada
        'codartic, numserie para el criterio de seleccion introducid
        devuelve = "INSERT INTO tmpinformes(codusu,nombre1,nombre2,campo1) "
        devuelve = devuelve & "SELECT " & vUsu.codigo & ", codartic,numserie,count(numserie) as campo1 from schrep "
        devuelve = devuelve & " WHERE " & cadSelect
        devuelve = devuelve & " group by codartic,numserie"
        conn.Execute devuelve
        
        'Eliminamos de la tabla aquellos registros que no superen el n� de reparaciones introducido
        devuelve = "DELETE FROM tmpinformes where codusu=" & vUsu.codigo & " and campo1<=" & txtCodigo(0).Text
        conn.Execute devuelve
        
        'Volver a comprobar que hay registro a mostrar para ello miramos en la
        'tabla tmpInformes que supere el n� de reparaciones a mostrar
        cadSelect = "codusu=" & vUsu.codigo
        If Not HayRegParaInforme("tmpinformes", cadSelect) Then
            BorrarTempInformes
            Exit Sub
        End If
    End If
    
    LlamarImprimir False
    
    'Eliminar de la tabla temporal
    If OpcionListado = 406 Then BorrarTempInformes
    
EFrecu:
    If Err.Number <> 0 Then MuestraError Err.Number, "Obteniendo n� de reparaciones.", Err.Description
End Sub


Private Sub cmdAceptarRepxDia_Click()
'Reparaciones por Dia
Dim devuelve As String
Dim param As String
Dim TotalMante As Integer
Dim Rs As ADODB.Recordset
Dim fecha1 As String, fecha2 As String
Dim NomTabla As String
Dim bOk As Boolean

' ====  [17/09/2009] LAURA : Cambiar conexion CONTAB si FAZ
Dim ConexionContaOk As Boolean
Dim CambiaConta_ As Boolean
' ====

    InicializarVbles
    
    '===================================================
    '============ PARAMETROS ===========================
    'A�adir el nombre de la Empresa como parametro
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    
    Select Case OpcionListado
        Case 63
            'Codigo = "{scarep.fecentre}"
            codigo = "{scarep.fecrepar}"   '09/12/2010  Estan cambiadas en el form
            param = "pDHFecha=""Fecha Rep.: "
            NomTabla = "scarep"
            cadNomRPT = "rRepReparxDia.rpt"
            conSubRPT = True
            cadTitulo = "Reparaciones por d�a"
        Case 73
            'A�adir el parametro total Mantenim. si estamos en Informe de Altas
            devuelve = "SELECT DISTINCT COUNT(*) FROM scaman "
            Set Rs = New ADODB.Recordset
            Rs.Open devuelve, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not Rs.EOF Then
                TotalMante = Rs.Fields(0).Value
                CadParam = CadParam & "pTotalMante=" & TotalMante & "|"
                numParam = numParam + 1
            End If
            Rs.Close
            Set Rs = Nothing
            
            'A�adir el Total Mantenim. del Periodo anterior
            fecha1 = Day(txtCodigo(31).Text) & "/" & Month(txtCodigo(31).Text) & "/" & Year(txtCodigo(31).Text) - 1
            fecha2 = Day(txtCodigo(32).Text) & "/" & Month(txtCodigo(32).Text) & "/" & Year(txtCodigo(32).Text) - 1
            codigo = "scaman.fechaini"
            devuelve = CadenaDesdeHastaBD(fecha1, fecha2, codigo, "F")
            If devuelve <> "" And devuelve <> "Error" Then
                devuelve = "SELECT DISTINCT COUNT(*) FROM scaman WHERE " & devuelve
                Set Rs = New ADODB.Recordset
                Rs.Open devuelve, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If Not Rs.EOF Then
                    TotalMante = Rs.Fields(0).Value
                    CadParam = CadParam & "pTotalAnte=" & TotalMante & "|"
                    numParam = numParam + 1
                End If
                Rs.Close
                Set Rs = Nothing
            End If
            
            '================= FORMULA =========================
            codigo = "{scaman.fechaini}"
            param = "pDHFecha=""Fecha: "
            NomTabla = "scaman"
            cadNomRPT = "rManListAltas.rpt"
            cadTitulo = "Informe Altas Mantenimientos"
        
        Case 223
            param = ""
            If Me.OptClientes Then
                codigo = "{scafac.fecfactu}"
                NomTabla = "scafac"
            Else
                codigo = "{scafpc.fecrecep}"
                NomTabla = "scafpc"
            End If
    End Select
   
        
    '===================================================
    '================= FORMULA =========================
    
    '== Cadena para seleccion Desde y Hasta FECHA ==
    If OpcionListado = 223 Then
        'El usuario de B solo puede contabilzar facturas de B
        If Val(vUsu.AlmacenPorDefecto) = vParamAplic.AlmacenB Then
            If Me.cboTipMov.ListIndex < 1 Then
                MsgBox "Seleccione tipo factura a contabilizar", vbExclamation
                Exit Sub
            End If
        End If
    
        'comprobar que se han rellenado los dos campos de fecha
        'sino rellenar con fechaini o fechafin del ejercicio
        'que guardamos en vbles Orden1,Orden2
        
        'fechaini del ejercicio de la conta
        If txtCodigo(31).Text = "" Then txtCodigo(31).Text = Orden1
     
        'fecha fin del ejercicio de la conta
        If txtCodigo(32).Text = "" Then txtCodigo(32).Text = Orden2
     
        'Comprobar que el intervalo de fechas D/H esta dentro del ejercicio de la
        'contabilidad par ello mirar en la BD de la Conta los par�metros
        If Not ComprobarFechasConta(31) Then Exit Sub
        If Not ComprobarFechasConta(32) Then Exit Sub
    End If
    
    devuelve = CadenaDesdeHasta(txtCodigo(31).Text, txtCodigo(32).Text, codigo, "F", "Fecha Factura")
    If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
    'Parametro D/H Fecha
    If devuelve <> "" And param <> "" Then
        CadParam = CadParam & AnyadirParametroDH(param, 31, 32) & """|"
        numParam = numParam + 1
    End If
    
    
    '## LAURA 20/06/2008
    '## A�adir frame de selec. factuar en contabilizar
    '- cadena para select en BDatos
    cadSelect = CadenaDesdeHastaBD(txtCodigo(31).Text, txtCodigo(32).Text, codigo, "F")
    
    
    
    
    'DAVID###
    'Rep x dia. A�adimos desde hasta cliente
    If OpcionListado = 63 Then
        devuelve = CadenaDesdeHasta(txtCodigo(132).Text, txtCodigo(133).Text, "{scarep.codclien}", "N", "Cliente")
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
        'Parametro D/H Fecha
        If devuelve <> "" Then
            devuelve = ""
            CadParam = CadParam & AnyadirParametroDH("pDHCliente=""Cliente: ", 132, 133) & """|"
            numParam = numParam + 1
        End If
    
    End If
    
    '== Cadena para seleccion Desde y Hasta N�Factura ==
    If OpcionListado = 223 Then
        '- comprobar: si n� factura tienen valor tipoMov tb
        If txtCodigo(121).Text <> "" Or txtCodigo(122).Text <> "" Then
            If Me.cboTipMov.ListIndex = -1 Then
                MsgBox "Debe seleccionar el tipo de movimiento si quiere contabilizar Desde/Hasta N� Factura.", vbInformation
                Exit Sub
            End If
            
            If Trim(Me.cboTipMov.List(Me.cboTipMov.ListIndex)) = "" Then
                MsgBox "Debe seleccionar el tipo de movimiento si quiere contabilizar Desde/Hasta N� Factura.", vbInformation
                Exit Sub
            End If
            
            '- a�adir desde/hasta factura a cadena seleccion registros
            codigo = "{scafac.numfactu}"
            devuelve = CadenaDesdeHasta(txtCodigo(121).Text, txtCodigo(122).Text, codigo, "N", "N� Factura")
            If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
            'Parametro D/H n� factura
            If devuelve <> "" And param <> "" Then
                CadParam = CadParam & AnyadirParametroDH(param, 31, 32) & """|"
                numParam = numParam + 1
            End If
            ' a�adir a la formula de bd
            devuelve = CadenaDesdeHastaBD(txtCodigo(121).Text, txtCodigo(122).Text, codigo, "N")
            If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Sub
        End If
    
                
        '- a�adir tipo movimiento a cadena seleccion
        If Me.cboTipMov.ListIndex >= 0 Then
            If Trim(Me.cboTipMov.List(Me.cboTipMov.ListIndex)) <> "" Then
                codigo = "{scafac.codtipom}"
                devuelve = Mid(Me.cboTipMov.List(Me.cboTipMov.ListIndex), 1, 3)
                devuelve = codigo & "=" & DBSet(devuelve, "T")
                If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
                If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Sub
            End If
        End If
    End If

    '===================================================
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
'    cadSelect = CadenaDesdeHastaBD(txtCodigo(31).Text, txtCodigo(32).Text, Codigo, "F")
    If OpcionListado = 223 Then
        If cadSelect <> "" Then cadSelect = cadSelect & " AND "
        cadSelect = cadSelect & NomTabla & ".intconta=0 "
        
        
        'Nuevo 7 Abril 08
        'Hay un parametro que permite contbilizar los tickets agrupados (NO uno a uno)
        'para ello, a partir de los FTI crearemos los FTG (tickets agrupados)
        'y los FTI NO se contabilizaran
        If Me.OptProve.Tag = "" Then
            'Contabilizacion NORMAL. Viene del MENU contabilizar
            'Comprueblo de agrupar tickets o no
            If Me.OptClientes.Value Then
                If vParamAplic.ContabilizarTicketAgrupados Then
                    'Solo las de clientes
                    If Me.OptClientes.Value Then cadSelect = cadSelect & " AND scafac.codtipom <> 'FTI'"
                End If
                
                'Febrero 2011
                'Condiciones:
                '       1- NO ha seleccionado ningun tipo de movimiento
                '       2- El usuario NO es de B
                '  -->  pongo que codtipom<>'FAZ'
                If cboTipMov.ListIndex < 1 Then  'El 0 esta vacio
                     If Val(vUsu.AlmacenPorDefecto) <> vParamAplic.AlmacenB Then cadSelect = cadSelect & " AND scafac.codtipom <> 'FAZ'"
                End If
            End If
        Else
            'CONTABILZIACION DE LOS TICKETS AGRUPADOS
            'A�ado el tipom al cad select
            cadSelect = cadSelect & " AND scafac.codtipom = 'FTG'"
        End If
    End If
    
    If Not HayRegParaInforme(NomTabla, cadSelect) Then Exit Sub
    
    
    If OpcionListado <> 223 Then
        LlamarImprimir False
    Else
    
        ' ====  [17/09/2009] LAURA : Cambiar conexion CONTAB si FAZ
        If Me.OptProve.Tag = "" Then
            If Me.OptClientes.Value Then
                devuelve = "CLI"
            Else
                devuelve = "PRO"
            End If
        Else
            devuelve = "TIK"
        End If

        CambiaConta_ = False
        ConexionContaOk = True
        
        If devuelve = "CLI" Then
            'CLIENTES para tipos de factura FAZ, es decir, el B
            If Me.cboTipMov.ListIndex >= 0 Then
                If Trim(Me.cboTipMov.List(Me.cboTipMov.ListIndex)) <> "" Then
                    If Mid(Me.cboTipMov.List(Me.cboTipMov.ListIndex), 1, 3) = "FAZ" Then
                    '            If vUsu.TrabajadorB Then
                        If AbrirConexionConta(True) Then
                            CambiaConta_ = True
                            ConexionContaOk = True
                        Else
                            ConexionContaOk = False
                        End If
                    End If
                End If
            End If
        End If
            

        If ConexionContaOk Then
        ' ====
            '------------------------------------------------------------------------------
            '  LOG de acciones.                      5: Facturas compras
            Set LOG = New cLOG
            

            
            devuelve = "Contabilizar facturas " & devuelve & ":" & vbCrLf & NomTabla & vbCrLf & cadSelect
            LOG.Insertar 5, vUsu, devuelve
            Set LOG = Nothing
            '-----------------------------------------------------------------------------
            
        
            bOk = ContabilizarFacturas(NomTabla, cadSelect, ProgressBarContab, Me.lblProgess2(0), lblProgess2(1), False)
        
            TerminaBloquear
            'Eliminar la tabla TMP
            BorrarTMPFacturas
            'Desbloqueamos ya no estamos contabilizando facturas
            If Me.OptClientes.Value Then
                DesBloqueoManual ("VENCON") 'VENtas CONtabilizar
            Else
                DesBloqueoManual ("COMCON") 'COMpras CONtabilizar
            End If
            Me.FrameProgress.visible = False
'            If Me.FrameTipMov.visible Then
'                Me.FrameRepxDia.Height = 4400
'            Else
'                Me.FrameRepxDia.Height = 3500
'            End If
            Me.Height = 4750
            Me.Refresh
            If bOk Then Unload Me
        
        ' ====  [17/09/2009] LAURA : Cambiar conexion CONTAB si FAZ
        End If
        If CambiaConta_ Then AbrirConexionConta False
        ' ====
    End If
End Sub



Private Sub cmdAceptarSustNSerie_Click(Index As Integer)
'Sustitucion de un N� de Serie que este en garant�a por otro n� de serie.
Dim SQL As String
Dim Rs As ADODB.Recordset

    txtCodigo(81).Text = Trim(txtCodigo(81).Text)
    
    If txtCodigo(81).Text <> "" Then
        'Comprobar que el nuevo n� de serie no existe ya
        SQL = DevuelveDesdeBDNew(conAri, "sserie", "numserie", "numserie", txtCodigo(81).Text, "T", , "codartic", Me.CadTag, "T")
        If SQL <> "" Then
            MsgBox "Ya existe ese N� de serie.", vbExclamation
            Exit Sub
        End If
        
        On Error GoTo ESustNSerie
        conn.BeginTrans
        
        'Insertar un registro con ese n� de serie y todos los valores que tenga el
        'num serie que sustituye
        SQL = "SELECT codartic,codtipar,codclien,coddirec,tieneman,nummante,ultrepar,fingaran,codtipom,numfactu,fechavta,numalbar,numline1,codprove,numalbpr,fechacom,numline2 FROM sserie "
        SQL = SQL & " WHERE numserie=" & DBSet(NumCod, "T") & " AND codartic=" & DBSet(CadTag, "T")
        Set Rs = New ADODB.Recordset
        Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        SQL = ""
        If Not Rs.EOF Then
            SQL = "(" & DBSet(txtCodigo(81).Text, "T") & ", " & DBSet(Rs!codArtic, "T", "N") & "," & DBSet(Rs!codTipar, "T", "N") & ","
            SQL = SQL & DBSet(Rs!codClien, "N", "S") & "," & DBSet(Rs!CodDirec, "N", "S") & "," & DBSet(Rs!TieneMan, "N", "S") & ","
            SQL = SQL & DBSet(Rs!nummante, "T", "S") & "," & DBSet(Rs!ultrepar, "F", "S") & "," & DBSet(Rs!fingaran, "F", "S") & ","
            SQL = SQL & DBSet(Rs!codtipom, "T", "S") & "," & DBSet(Rs!NumFactu, "N", "S") & "," & DBSet(Rs!FechaVta, "F", "S") & ","
            SQL = SQL & DBSet(Rs!NumAlbar, "N", "S") & "," & DBSet(Rs!numline1, "N", "S") & "," & DBSet(Rs!Codprove, "N", "S") & ","
            SQL = SQL & DBSet(Rs!numalbPr, "T", "S") & "," & DBSet(Rs!fechaCom, "F", "S") & "," & DBSet(Rs!numline2, "N", "S") & ")"
        End If
        Rs.Close
        Set Rs = Nothing
        
        If SQL <> "" Then
            SQL = "INSERT INTO sserie (numserie,codartic,codtipar,codclien,coddirec,tieneman,nummante,ultrepar,fingaran,codtipom,numfactu,fechavta,numalbar,numline1,codprove,numalbpr,fechacom,numline2) VALUES " & SQL
            conn.Execute SQL
        
            'sustituir el campo numalbar del numserie viejo por 9999999
            'y poner en el campo "numsersu" en num. serie por el que se sustituye
            'limpiar campos del cliente
            SQL = "UPDATE sserie SET numalbar=9999999, numsersu=" & DBSet(txtCodigo(81).Text, "T")
            SQL = SQL & ", codclien=" & ValorNulo & ", coddirec=" & ValorNulo
            SQL = SQL & ", numfactu=" & ValorNulo
            SQL = SQL & " WHERE numserie=" & DBSet(NumCod, "T") & " AND codartic=" & DBSet(CadTag, "T")
            conn.Execute SQL
        End If
    Else
        MsgBox "Debe introducir el N� Serie por el que se sustituye.", vbInformation
        Exit Sub
    End If

ESustNSerie:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Sustituci�n N� Serie.", Err.Description
        conn.RollbackTrans
    Else
        conn.CommitTrans
        Unload Me
    End If
End Sub



Private Sub cmdAceptarTarif_Click()
Dim cadFrom As String

    InicializarVbles
   
   '========= Frame de Tarifas y Descuentos ===============================
    'Nombre fichero .rpt a Imprimir
    'Ordenar por: codtarifa, codfamia, codmarca, codartic
    Select Case OpcionListado
        Case 28: cadNomRPT = "rFacTarifasAlm.rpt"  'Listado Tarifas Articulos
        Case 29: cadNomRPT = "rFacPromociones.rpt"  'Listado Promociones
        Case 30: cadNomRPT = "rFacPreciosEsp.rpt"
        Case 245: cadNomRPT = "rFacTarifasMargen.rpt"
    End Select
    
    If Not PonerFormulaYParametrosInf28() Then Exit Sub
    

    
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    If cadFormula <> "" Or (OpcionListado = 245) Then
        cadFrom = codigo & " INNER JOIN sartic ON " & codigo & ".codartic=sartic.codartic "
    Else
        cadFrom = codigo
    End If
    
    'seleccionar solo los que tienen margen con error
    If OpcionListado = 245 Then
        If Me.chkMostrarErrores Then
            AnyadirAFormula cadSelect, " sartic.preciove <> sartic.preciouc + round((sartic.preciouc * if(isnull(sartic.margecom),0,sartic.margecom))/100,4)"
            AnyadirAFormula cadFormula, " {sartic.preciove} <> {sartic.preciouc} + round(({sartic.preciouc} * iif(IsNull({sartic.margecom}),0,{sartic.margecom}))/100,4)"
        End If
    End If
    
    
    If Not HayRegParaInforme(cadFrom, cadSelect) Then Exit Sub
    
    LlamarImprimir False
End Sub


Private Sub cmdActVtosFraPro_Click()
    
    'Para mostrar errores
    cadFormula = "" 'Fechas anteriores a fecha inicio ejercicio
    cadTitulo = DevuelveDesdeBDNew(conConta, "parametros", "FechaActiva", "1", "1")
    If cadTitulo = "" Then cadTitulo = vEmpresa.FechaIni
    
    For indCodigo = 0 To 4
        If Me.txtCodigo(149 + indCodigo).visible Then
            If Me.txtCodigo(149 + indCodigo).Text <> Me.txtCodigo(149 + indCodigo).Tag Then
                'OK. Ha cambiado una fecha de vencimiento
                If CDate(txtCodigo(149 + indCodigo).Text) < vEmpresa.FechaIni Then
                    cadFormula = cadFormula & txtCodigo(149 + indCodigo).Text & " -> Ejercicio cerrado" & vbCrLf
                    
                Else
                    If CDate(txtCodigo(149 + indCodigo).Text) < CDate(cadTitulo) Then cadFormula = cadFormula & txtCodigo(149 + indCodigo).Text & " -> Menor fecha activa" & vbCrLf
                End If
            End If
        End If
    Next
    
    If cadFormula <> "" Then
        cadFormula = "Error en fechas vencimiento: " & vbCrLf & vbCrLf & cadFormula
        MsgBox cadFormula, vbExclamation
        Exit Sub
    End If
        
    For indCodigo = 0 To 4
        If Me.txtCodigo(149 + indCodigo).visible Then
            If Me.txtCodigo(149 + indCodigo).Text <> Me.txtCodigo(149 + indCodigo).Tag Then
                codigo = "UPDATE spagop set fecefect = " & DBSet(txtCodigo(149 + indCodigo).Text, "F")
                codigo = codigo & " WHERE " & cmdActVtosFraPro.Tag & " AND numorden =" & Label3(120 + indCodigo).Tag
                ConnConta.Execute codigo
            End If
        End If
    Next
    Unload Me
End Sub

Private Sub cmdBajar_Click()
'Bajar el item seleccionado del listview2
    BajarItemList Me.ListView2
End Sub

Private Sub cmdCancel_Click(Index As Integer)
    Unload Me
End Sub


Private Sub cmdDeselTodos_Click()
Dim I As Byte

    For I = 1 To ListView1.ListItems.Count
        ListView1.ListItems(I).Checked = False
    Next I
End Sub

Private Sub cmdElimiaFacturas_Click()
Dim b As Boolean

'Igual hay que quitarlo


    'Proceso de borre de facturas
    If cmbEliFac.ListIndex < 0 Then Exit Sub
    
    
    
    'Tablas que voy a tener que borrar
    'Para que no se queden datos
    cadTitulo = String(60, "*") & vbCrLf & vbCrLf
    cadTitulo = cadTitulo & " Se eliminar�n los datos con fecha anterior a la solicitada de: " & vbCrLf
    cadTitulo = cadTitulo & " CLIENTES: " & vbCrLf
    cadTitulo = cadTitulo & Space(20) & "Albaranes, hco albaranes, ofertas, hco ofertas, pedidos, hco pedidos" & vbCrLf
    cadTitulo = cadTitulo & Space(20) & "facturas, hco facturas, ventas tpv, reparaciones, hco reparaciones, produccion" & vbCrLf & vbCrLf & vbCrLf
    cadTitulo = cadTitulo & " PRVEEDORES: " & vbCrLf
    cadTitulo = cadTitulo & Space(20) & "Albaranes, hco albaranes,  pedidos, hco pedidos, facturas, hco facturas " & vbCrLf & vbCrLf & vbCrLf
    
    codigo = cadTitulo & "El proceso es irreversible." & vbCrLf & vbCrLf & vbCrLf & "SEGURO QUE DESEA CONTINUAR?"
    
    'Reestablecer variables
    InicializarVbles
    cadTitulo = ""
    
    If MsgBox(codigo, vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    codigo = InputBox("Password seguridad")
    codigo = UCase(codigo)
    If codigo <> "ARIADNA" Then Exit Sub
    
    Label3(83).Caption = "Inicio del proceso del borre de facturas"
    Me.cmdElimiaFacturas.Enabled = False
    Me.Refresh
    Screen.MousePointer = vbHourglass
    Set miRsAux = New ADODB.Recordset
    'Conn.BeginTrans
    b = BorrarFacturas
    'Conn.RollbackTrans
    'Volvemos a dejarlo todo como estaba
    Set miRsAux = Nothing
    Orden1 = ""
    codigo = ""
    Label3(83).Caption = ""
    Me.cmdElimiaFacturas.Enabled = True
    Screen.MousePointer = vbDefault
    
    If b Then Unload Me
End Sub

Private Sub cmdEtiqBulto_Click()
Dim I As Integer

    CadParam = ""
    If OpcionListado = 95 Then
        If Me.txtClie.Text = "" Then CadParam = "Ponga el cliente"
    Else
        If Me.txtCodigo(148).Text = "" Or Me.txtNombre(148).Text = "" Then CadParam = "Ponga el proveedor"
    End If
    If CadParam <> "" Then
        MsgBox CadParam, vbExclamation
        Exit Sub
    End If
        
        
    
    If Val(txtBultos(1).Text) = 0 Then txtBultos(1).Text = "1"
    CadParam = "delete from tmpinformes where codusu =" & vUsu.codigo
    conn.Execute CadParam
       
    numParam = 0
    
    Orden2 = "INSERT INTO tmpinformes(codusu,codigo1,nombre1,nombre2,nombre3) VALUES (" & vUsu.codigo & ","
    If OpcionListado = 95 Then
        CadParam = "," & vParam.codigo & ",'" & DevNombreSQL(txtNombre(10).Text) & "')"
    Else
        CadParam = "," & vParam.codigo & ",'" & DevNombreSQL(txtNombre(148).Text) & "')"
    End If
    cadFormula = ""
    If txtBultos(7).Text <> "" Then
        'Lleva etiquetas en blanco
        For I = 1 To Val(txtBultos(7).Text)
            '           secuencia               'El cliente a blancos
            numParam = numParam + 1
            cadFormula = numParam & ",''"
            cadFormula = Orden2 & cadFormula & CadParam
            conn.Execute cadFormula
        Next I
    End If
    For I = 1 To Val(txtBultos(1).Text)
          '           secuencia               'El cliente a blancos
            numParam = numParam + 1
            cadFormula = numParam & ",'" & txtClie.Text & "'"
            cadFormula = Orden2 & cadFormula & CadParam
            conn.Execute cadFormula
            
    Next I
    cadFormula = ""
       
    'Como puede llevar saltos de linea
    Orden2 = SaltosDeLinea(txtBultos(0).Text)
    'Le pasare los datos
    CadParam = ""
    numParam = 0
    If PonerParamRPT2(19, CadParam, numParam, cadNomRPT, pImprimeDirecto, pPdfRpt, pRptvMultiInforme) Then
        Orden1 = "0"    'Ahora siempre es CERO. No tiene sentido pasar al rpt

        'Metemos los campos de direccion
        CadParam = CadParam & "Dom=""" & txtBultos(2).Text & """|"
        CadParam = CadParam & "Pob=""" & txtBultos(3).Text & """|"
        CadParam = CadParam & "Pro=""" & Trim(txtBultos(4).Text & "      " & txtBultos(5).Text) & """|"
        
        
        'Si lleva departamento lo metere
        cadSelect = ""
        If cmbBulto.ListIndex > 0 Then
            'Ha cogido departamento
            I = InStr(1, cmbBulto.Text, ":")
            If I = 0 Then
                'NO  deberia pasar nunca
                MsgBox "Error nombre departamento", vbExclamation
            Else
                cadSelect = Trim(Mid(cmbBulto.Text, 1, I - 1))
                cadSelect = Replace(cadSelect, """", "'")
            End If
        End If
        CadParam = CadParam & "nomdirec=""" & cadSelect & """|"
        
        'A�ado la direccion que se ve
        CadParam = CadParam & "DireccionAlternativa=0|"
        'cadParam = cadParam & "DireccionAlternativa=" & Orden1 & "|"
        CadParam = CadParam & "Texto= """ & Orden2 & """|"
        numParam = numParam + 2
        cadSelect = "codusu=" & vUsu.codigo
        cadFormula = "({tmpinformes.codusu} =" & vUsu.codigo & ")"
        LlamarImprimir True
        If Me.NumCod <> "" Then Unload Me
    End If
        
End Sub

'INTENTARE METERLO DENTRO DE OTRO PROC

'Abril 2010
'En una columna de tmpinforme voy a grabar el dto para la familia
'De moemnto pong la UNO a pi�on
'Veremos si hay que pedir datos o no. De momento esta a pi�on

Private Sub cmdEtiqEstanteria_Click()
'Estadistica margen ventas por articulo
Dim campo As String
Dim param As String
Dim tabla As String
Dim Rs As ADODB.Recordset
Dim Li As Collection
Dim I As Integer
Dim Dto As Currency
Dim Precio As Currency
Dim Codfamia As Integer

    InicializarVbles
    
    '===================================================
    '============ PARAMETROS ===========================
    'A�adir el nombre de la Empresa como parametro
    CadParam = CadParam & "|pNomEmpre=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    CadParam = CadParam & "|pImprimeBarras=""" & Abs(Me.chkImprimeCodigoBarras.Value) & """|"
    numParam = numParam + 1
    CadParam = CadParam & "|numerodecimales=" & Me.cboDecimal.List(cboDecimal.ListIndex) & "|"
    numParam = numParam + 1
    
    
    If OpcionListado = 94 Then
        'La normal. Impresion etque estanteria
        '===================================================
        '================= FORMULA =========================
        'Cadena para seleccion D/H familia
        '--------------------------------------------
        If txtCodigo(94).Text <> "" Or txtCodigo(95).Text <> "" Then
            campo = "{sartic.codfamia}"
            param = "pDHFamilia=""Familia: "
            If Not PonerDesdeHasta(campo, "N", 94, 95, param) Then Exit Sub
        End If
            
        'Cadena para seleccion D/H art�culo
        '--------------------------------------------
        If txtCodigo(92).Text <> "" Or txtCodigo(93).Text <> "" Then
            campo = "{sartic.codartic}"
            param = "pDHArticulo=""Art�culo: "
            If Not PonerDesdeHasta(campo, "T", 92, 93, param) Then Exit Sub
        End If
        
        'Cadena para seleccion D/H Fecha
        '--------------------------------------------
        If txtCodigo(123).Text <> "" Or txtCodigo(124).Text <> "" Then
            campo = "{sartic.ultfecpvp}"
            param = "pDHFecha=""Fecha: "
            If Not PonerDesdeHasta(campo, "F", 123, 124, param) Then Exit Sub
        End If
        
        If cadSelect <> "" Then cadSelect = cadSelect & " AND "
        cadSelect = cadSelect & " {sartic.codstatu} <=1 " 'normal y obsoleto
        
    Else
        'Desde el albaran
        cadSelect = "artvario=0 "
        cadSelect = cadSelect & " AND sartic.codstatu <= 1 "
        cadSelect = cadSelect & " and codartic in (Select codartic from slialp " & NumCod & ")"
    End If
    
    
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    tabla = " sartic  "
    If Not HayRegParaInforme(tabla, cadSelect) Then Exit Sub

    'Borro tmptemporal
    tabla = "DELETE FROM tmpinformes WHERE codusu =" & vUsu.codigo
    conn.Execute tabla
    
    'A�adire los tipos de IVA a esta tabla
    tabla = "INSERT INTO tmpinformes(codusu,codigo1)  select " & vUsu.codigo & ",codigiva from sartic"
    If cadSelect <> "" Then tabla = tabla & " WHERE " & cadSelect
    tabla = tabla & " GROUP BY codigiva"
    conn.Execute tabla
    
    
    
    
    
    'AHora desde conta cargo los % de IVA desde la conta
    Set Rs = New ADODB.Recordset
    tabla = "Select * from tmpinformes where codusu =" & vUsu.codigo
    Rs.Open tabla, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Set Li = New Collection
    While Not Rs.EOF
        Li.Add Val(Rs.Fields(1))
        Rs.MoveNext
    Wend
    Rs.Close
    
    
    '
    
    'Abrimos los IVAS en conta
    tabla = "Select codigiva,porceiva from tiposiva"
    Rs.Open tabla, ConnConta, adOpenKeyset, adLockOptimistic, adCmdText
    For I = 1 To Li.Count
        tabla = "codigiva = " & Li.item(I)
        Rs.Find tabla, , , 1
        If Rs.EOF Then
            MsgBox "Tipo de IVA no encontrado en la contabilidad" & tabla, vbExclamation
            Rs.Close
            Exit Sub
        Else
            tabla = "UPDATE tmpinformes SET porcen1 =" & TransformaComasPuntos(CStr(Rs!PorceIVA))
            tabla = tabla & " WHERE codusu =" & vUsu.codigo & " AND codigo1 = " & Rs!codigiva
            conn.Execute tabla
        End If
    Next I
    Rs.Close
    Set Li = Nothing
    
    
    'Borramos los datos de la tabla donde iran los articulos
    tabla = "DELETE FROM tmpnseries WHERE codusu =" & vUsu.codigo
    conn.Execute tabla
    I = Me.cboDecimal.List(cboDecimal.ListIndex)
    If I = 0 Then
        tabla = "0"
    Else
        tabla = "#,##0." & Mid("0000", 1, I)
    End If
    frmMensajes.cadWHERE2 = tabla
    frmMensajes.cadWhere = cadSelect
    frmMensajes.vCampos = ""  'estaopcion en etiquetas es para mostrar las del almacen con punto de pedido indicado
    frmMensajes.OpcionMensaje = 15
    frmMensajes.Show vbModal
    
    'Si ha devuelto seleccionados
    tabla = " tmpnseries   "
    cadFormula = " codusu =" & vUsu.codigo
    
    If Not HayRegParaInforme(tabla, cadFormula) Then Exit Sub
    
    
    'Para los articulos que hay que mostrar, si tienen dto hay que poner
    'cargalro
    If Me.chkDtoFM.Value = 1 Then
        'Cargo los dtos
        'A pi�on para ALZIRA
        tabla = "select * from sdtofm where codactiv=1 and codclien is null and codmarca is null and codfamia >=0 order by codfamia "
        Rs.Open tabla, conn, adOpenKeyset, adLockPessimistic, adCmdText
        
        If Not Rs.EOF Then
                tabla = "SELECT tmpinformes.codusu,`sartic`.`nomartic`, `sartic`.`preciove`, `tmpinformes`.`porcen1`, `sartic`.`codartic`,codfamia,codmarca,numlinea"
                tabla = tabla & " FROM   ((`tmpnseries` `tmpnseries` INNER JOIN `sartic` `sartic` ON `tmpnseries`.`codartic`=`sartic`.`codartic`)"
                tabla = tabla & " INNER JOIN `tmpinformes` `tmpinformes` ON (`sartic`.`codigiva`=`tmpinformes`.`codigo1`)"
                tabla = tabla & " AND (`tmpnseries`.`codusu`=`tmpinformes`.`codusu`)) Where tmpinformes.CodUsu = " & vUsu.codigo & " ORDER BY codfamia,codmarca"
                Set miRsAux = New ADODB.Recordset
                
                
                miRsAux.Open tabla, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                Codfamia = -1
                While Not miRsAux.EOF
                    
                    If Codfamia <> miRsAux!Codfamia Then
                        'Hay que buscar
                        I = 1
                    Else
                        I = 0
                    End If
                    
                    
                    If I = 1 Then
                        Codfamia = miRsAux!Codfamia
                        Dto = 0
                        Rs.MoveFirst
                        tabla = ""
                        While I = 1
                            If Rs!Codfamia = Codfamia Then
                                'OK. ESte es. No muevo
                                I = 0 'salga
                                Dto = Rs!dtoline1 + Rs!dtoline2
                            Else
                                If Rs!Codfamia > Codfamia Then Rs.MoveLast
                                Rs.MoveNext
                            End If
                            If Rs.EOF Then I = 0
                        Wend
                    End If
                    If Not Rs.EOF Then
                        'OK hay dto
                        
                        If Dto > 0 Then
                            Precio = DBLet(miRsAux!porcen1, "N")
                            Precio = (miRsAux!PrecioVe * Precio) / 100
                            Precio = Precio + miRsAux!PrecioVe
                            Precio = (Precio * Dto) / 100
                            
                            If Precio > 0 Then
                                tabla = Format(Precio, FormatoCantidad)
                                
                                tabla = "update tmpnseries set nummante = '" & tabla & "' WHERE codusu = " & vUsu.codigo
                                tabla = tabla & " AND codartic = " & DBSet(miRsAux!codArtic, "T")
                                tabla = tabla & " AND numlinea = " & miRsAux!numlinea
                                conn.Execute tabla
                            End If
                        End If
                        
                    End If
                    miRsAux.MoveNext
                Wend
                miRsAux.Close
                Set miRsAux = Nothing
        End If
        
        
        Rs.Close
    End If
    
    
    
    
    cadFormula = "({tmpnseries.codusu} =" & vUsu.codigo & ")"
    
    campo = ""
    If Not PonerParamRPT2(23, CadParam, numParam, campo, pImprimeDirecto, pPdfRpt, pRptvMultiInforme) Then
        cadNomRPT = "rEtiqEsta.rpt"
    Else
        cadNomRPT = campo
    End If
    
    LlamarImprimir True
    
    BorrarTempInformes
    
    'Borramos los datos de la tabla donde iran los articulos
    tabla = "DELETE FROM tmpnseries WHERE codusu =" & vUsu.codigo
    conn.Execute tabla
    
End Sub



Private Sub cmdFactAlbaranes_Click()
    codigo = "�Seguro que desea continuar?"
    If MsgBox(codigo, vbYesNo + vbQuestion) = vbNo Then Exit Sub
    If HacerSQLListado82_83 Then
        CadenaDesdeOtroForm = "OK"
        Unload Me
    End If
End Sub

Private Sub cmdFrecuencias_Click()
Dim campo As String

    ' ---- [06/11/2009] [LAURA] : corregir informe de frecuencias
    
    InicializarVbles
    
    '============ PARAMETROS ===========================
    'A�adir el nombre de la Empresa como parametro
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    'A�adir parametro es departamento o direccion
    CadParam = CadParam & "|pDpto=" & vParamAplic.HayDeparNuevo & "|"
    numParam = numParam + 1
    

    
    '================= FORMULA =========================
    'Cadena para seleccion D/H CLIENTE
    '----------------------------------
    If txtCodigo(98).Text <> "" Or txtCodigo(99).Text <> "" Then
        campo = "{scafre.codclien}"
        If Not PonerDesdeHasta(campo, "N", 98, 99, "pDHCliente=""Cliente: ") Then Exit Sub
    End If
    
    If Not HayRegParaInforme("scafre", cadSelect) Then Exit Sub
    
    
    If Me.OptFrecResumen.Value = True Then
        cadNomRPT = "rFrecuResum.rpt" 'Informe resumen
    Else
        cadNomRPT = "rFrecuFicha.rpt" 'Ficha
    End If
    
    cadTitulo = "Frecuencias"
    
'    conSubRPT = False
    
    LlamarImprimir False
    
    ' ----

'## ANTES
'        'Le pasare los datos
'    cadParam = ""
'    numParam = 0
'
'    If PonerParamRPT(19, cadParam, numParam, cadNomRPT) Then
'        Orden1 = "0"
'       ' If Me.optDirEnvio(1).Value Then Orden1 = "1"
'
'        'A�ado la direccion que se ve
'        cadParam = cadParam & "DireccionAlternativa=" & Orden1 & "|"
'        cadParam = cadParam & "Texto= """ & Orden2 & """|"
'        numParam = numParam + 2
'        cadSelect = "codusu=" & vUsu.Codigo
'
'        LlamarImprimir
'    End If
'##
End Sub


Private Sub cmdHcoMante_Click()
    codigo = ""
    For indCodigo = 110 To 112
        If txtCodigo(indCodigo).Text = "" Then codigo = codigo & "M"
        If indCodigo > 110 Then If txtNombre(indCodigo).Text = "" Then codigo = codigo & "M"
    Next indCodigo
    If codigo <> "" Then
        MsgBox "Rellene correctamente todos los datos", vbExclamation
        Exit Sub
    End If
    'CUATRO CAMPOS. El primero de control
    CadenaDesdeOtroForm = "OK|" & txtCodigo(110).Text & "|" & txtNombre(111).Text & "|" & txtCodigo(112).Text & "|"
    Unload Me
End Sub

'===================================================
'===================================================
' Informe teorico mantenimientos
Private Sub cmdManteTeorico_Click()
Dim cadFrom As String
Dim campo As String, devuelve As String
Dim codigo  As String

    InicializarVbles

    
    '===================================================
    '============ PARAMETROS ===========================
    'A�adir el nombre de la Empresa como parametro
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    
    
        cadNomRPT = "rManListTeorico.rpt"
    
        
        cadTitulo = "Informe Mantenimientos"
        codigo = "scaman"
    
    cadFrom = "(" & codigo & " INNER JOIN sclien ON " & codigo & ".codclien=sclien.codclien) "
      
    '===================================================
    '================= FORMULA =========================
    
    'Cadena para seleccion CLIENTE
    '--------------------------------------------
    If txtCodigo(102).Text <> "" Or txtCodigo(103).Text <> "" Then
        campo = "{" & codigo & ".codclien}"
        'Parametro Desde/Hasta Cliente
        devuelve = "pDHCliente=""Cliente: "
        If Not PonerDesdeHasta(campo, "N", 102, 103, devuelve) Then Exit Sub
    End If

    
    'Cadena para seleccion TIPO CONTRATO
    '--------------------------------------------
    If txtCodigo(104).Text <> "" Or txtCodigo(105).Text <> "" Then
        campo = "{" & codigo & ".codtipco}"
        'Parametro Desde/Hasta Tipo Contrato
        devuelve = "pDHTipoCon=""Tipo Contrato: "
        If Not PonerDesdeHasta(campo, "T", 104, 105, devuelve) Then Exit Sub
    End If
       
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    cadSelect = cadFormula
    If Not HayRegParaInforme(cadFrom, cadSelect) Then Exit Sub
    
    
    'Si  detalla o no
    CadParam = CadParam & "Detallar=" & Abs(Me.chkMante(0).Value) & "|"
    numParam = numParam + 1

    
    LlamarImprimir False
End Sub

Private Sub cmdSelTodos_Click()
Dim I As Byte

    For I = 1 To ListView1.ListItems.Count
        ListView1.ListItems(I).Checked = True
    Next I
End Sub



Private Sub cmdStockMin_Click()
    
    
    InicializarVbles
    
    
    cadFormula = ""
    
    
    Orden2 = ""
    'Cadena para seleccion D/H ALMACEN
    '--------------------------------------------
    If txtCodigo(139).Text <> "" Or txtCodigo(140).Text <> "" Then
        Orden1 = "{salmac.codalmac}"
        'Parametro Desde/Hasta Familila
        cadTitulo = "Almac�n: "
        If Not PonerDesdeHasta(Orden1, "N", 139, 140, cadTitulo) Then Exit Sub
        Orden2 = cadTitulo
    End If
    If Me.chkVarios(0).Value Then Orden2 = Orden2 & "            * Sin stock m�nimo"
    Orden2 = "|pDHZona=""" & Trim(Orden2) & """|"
    CadParam = CadParam & Orden2
    
    
    'Cadena para seleccion D/H FAMILIA
    '--------------------------------------------
    If txtCodigo(141).Text <> "" Or txtCodigo(142).Text <> "" Then
        Orden1 = "{sartic.codfamia}"
        'Parametro Desde/Hasta Familila
        cadTitulo = "pDHCliente=""Familia: "
        If Not PonerDesdeHasta(Orden1, "N", 141, 142, cadTitulo) Then Exit Sub
    End If
    
    
    'Cadena para seleccion D/H PROVEEDOR
    '--------------------------------------------
    If txtCodigo(143).Text <> "" Or txtCodigo(144).Text <> "" Then
        Orden1 = "{sartic.codprove}"
        'Parametro Desde/Hasta Proveedor
        cadTitulo = "pDHAgente=""Proveedor: "
        If Not PonerDesdeHasta(Orden1, "N", 143, 144, cadTitulo) Then Exit Sub
    End If
    
    
    Screen.MousePointer = vbHourglass
    
    
    If HacerInfrStockMinimo Then
        cadTitulo = "Listado Stock m�nimo"
        cadNomRPT = "rAlmStockMinimos.rpt"
        conSubRPT = False
        cadFormula = "{tmpInformes.codusu}=" & vUsu.codigo
        LlamarImprimir False
    End If
    Screen.MousePointer = vbDefault
    Label3(116).Caption = ""
End Sub

Private Sub cmdSubir_Click()
'Subir el item seleccionado del listview2 una posicion
    SubirItemList Me.ListView2
End Sub



Private Sub Form_Activate()
Dim IndiceFoco As Integer

    If PrimeraVez Then
        PrimeraVez = False
        IndiceFoco = -1
        Select Case OpcionListado
        Case 1, 2, 3, 4, 61, 20, 21, 22, 23, 24, 27, 58, 110
            '1:Listado de Marcas, 2:Almacenes Propios, 3:Tipos de Unidad
            '4:Tipos de Art�culos, 6:Art�culos
            '61:Motivos Pen. Rep
            '58:Proveedores, 110:Ubicaciones
             'PonerFoco txtCodigo(1)
             IndiceFoco = 1
        Case 6 '6: Informe de Articulos
            'PonerFoco txtCodigo(62)
            IndiceFoco = 62
        Case 7, 8 '7: Informe Traspaso Almacenes/Historico
                  '8: Informe Movimientos Almacen/Historico
            'PonerFoco txtCodigo(3)
            IndiceFoco = 3
        Case 9 'Informe Movimientos Art�culos
            'PonerFoco txtCodigo(5)
            IndiceFoco = 5
        Case 11     '11: Listado de Articulos con componentes ' ====  [16/09/2009] LAURA
            IndiceFoco = 125
            Orden1 = "nomalmac"
            codigo = DevuelveDesdeBD(conAri, "codalmac", "salmpr", "codalmac>0 AND 1", "1", , Orden1)
            txtCodigo(145).Text = codigo
            txtNombre(145).Text = Orden1
            
            
        Case 12, 13, 14, 15, 16, 17, 19
                        '12: Listado Toma de Inventario Articulos
                        '13: Listado Diferencias de Inventario Articulos
                        '14: Actualizar Diferencias de Inventario (No IMPRIME INFORME)
                        '15: Listado Articulos Inactivos
                        '16: Listado Valoracion de Stocks Inventariados
                        '17: Listado Valoraci�n Stocks
                        '19: Inf. Stocks a una Fecha
            'PonerFoco txtCodigo(13)
            IndiceFoco = 13
        Case 18      '18: Informe Stocks MAximos y Minimos
            'PonerFoco txtCodigo(72)
            IndiceFoco = 72
        Case 28, 29, 30 '28: Informe Tarifas de Articulos
                    '29: Informe Promociones
                    '30: Informe Precios Especiales
            'PonerFoco txtCodigo(23)
            IndiceFoco = 23
        Case 31, 73 '31: Informe Ofertas
                    '73: Listado Altas Mantenimientos
            'PonerFoco txtCodigo(31)
            IndiceFoco = 31
        Case 54 'Listado Descuentos Familia/ Marca
            'PonerFoco txtCodigo(73)
            IndiceFoco = 73
        Case 60 '60: Informe Reparacions - N� Series
            'PonerFoco txtCodigo(37)
            IndiceFoco = 37
        Case 63
            IndiceFoco = 131
            
        Case 73
            '63: Listado Reparaciones x d�a
            IndiceFoco = 31
        
        
        Case 223
        
            'FALTA### Quitar
            txtCodigo(31).Text = "22/03/2016"
            txtCodigo(32).Text = "22/03/2016"
        
        
            '223: Contabilizar facturas
            If Me.OptProve.Tag = "" Then
                'Contabilizacion normal clie/prov
                IndiceFoco = 31
            
            Else
                'TICKETS AGRUPADOS
                'Contabilizacion de facturas de tickets agrupadas. Lanzamos YA el proceso
                DoEvents
                cmdAceptarRepxDia_Click
                Me.Refresh
                Unload Me
                Exit Sub
            End If
        Case 246 '246: Informe margen ventas x articulo
            'PonerFoco txtCodigo(88)
            IndiceFoco = 130
        Case 64, 406 '64: Listado Reparaciones x Cliente
                     '406: List. Frecuencia de Reparaciones
            'PonerFoco txtCodigo(33)
            IndiceFoco = 33
        Case 70, 71, 76, 79 'Listado Mantenimientos
            'PonerFoco txtCodigo(45)
            IndiceFoco = 45
        Case 72 'Informe Fichas Mantenimientos
            'PonerFoco txtCodigo(55)
            IndiceFoco = 55
            
        Case 77
            'PonerFoco txtCodigo(102)
             IndiceFoco = 102
        Case 78
            'PonerFoco txtCodigo(109)
            IndiceFoco = 45
            
        Case 82, 83
            'Marca facturar a 1
            IndiceFoco = 119
        Case 94
            IndiceFoco = 94
            
            
        ' ---- [06/11/2009] [LAURA] : corregir informe de frecuencias
        Case 96 'Informe frecuencias
            IndiceFoco = 98
        ' ----
            
        Case 309 '309:Listado precios de compra
            'PonerFoco txtCodigo(79)
            IndiceFoco = 79
        Case 407 'Sustituci�n N� Serie
            'PonerFoco txtCodigo(81)
            IndiceFoco = 81
        Case 409 'List. Avisos de averias pendientes
            'PonerFoco txtCodigo(82)
            IndiceFoco = 82
        Case 95
            PonerFoco txtClie
            
        Case 99
            'PonerFoco txtCodigo(110)
            IndiceFoco = 110
        Case 247  'y Correccion de listados de precios tarias etc
             'PonerFoco txtCodigo(107)
            IndiceFoco = 107
             
        Case 512
            ContabilizarUnaFacturaProveedor
        
        Case 513
            'etquietas estanteria desde albaran compra
            PonerFocoCbo cboDecimal
        Case 514
            PonerDatosFacturaProveedorAcabadaRecepcionar
        Case 100
            PonerFoco Me.txtCodigo(139)
        End Select
        If IndiceFoco >= 0 Then PonerFoco txtCodigo(IndiceFoco)
    End If
    Screen.MousePointer = vbDefault
End Sub



Private Sub Form_Load()
Dim H As Integer, W As Integer

    'Icono del formulario
    Me.Icon = frmPpal.Icon

    PrimeraVez = True
    limpiar Me
    CargaIconosAyuda
    'Ocultar todos los Frames de Formulario
    FrameListado.visible = False
    FrameInfAlmacen.visible = False
    FrameMovArtic.visible = False
    FrameInventario.visible = False
    FrameTarifas.visible = False
    FrameRepNSerie.visible = False
    FrameRepxDia.visible = False
    FrameRepxClien.visible = False
    FrameMantenimientos.visible = False
    Me.FrameFichasMan2.visible = False
    FrameInfArticulos.visible = False
    FrameDtosFM.visible = False
    FrameRepSustNSerie.visible = False
    FrameListAvisosPtes.visible = False
    FrameEstMargenes.visible = False
    Me.FrameEtiqEstanteria.visible = False
    FrameBultos.visible = False
    Me.FrameFrecuencia.visible = False
    FrEliminarFacturas.visible = False
    FrameListMant2.visible = False
    FrameEnvioMail.visible = False
    FrameHcoMante.visible = False
    FrameAlbaranesMarcaFacturar.visible = False
    FrameConta1FRAPRO.visible = False
    Me.FrameInvArtComp.visible = False ' ====  [16/09/2009] LAURA : Listado Articulos con componentes
    FrameAlmacenStkMin.visible = False
    FrameFraProveedor.visible = False
    CommitConexion
    
    cadTitulo = ""
    cadNomRPT = ""
    
    Select Case OpcionListado
        Case 1 To 19, 247 'Listado de ALMACEN
            ListadosAlmacen H, W
        Case 100 To 199 'Listados de ALMACEN
            ListadosAlmacen H, W
        Case 20 To 30 'Listadod de FACTURACION
            ListadosFacturacion H, W
        Case 70 To 89 'Listados de MANTENIMIENTO
            ListadosMantenimiento H, W
        Case 245, 246 'Listados tarifas
            ListadosFacturacion H, W
        Case 300 To 390 'Listados de COMPRAS
            ListadosCompras H, W
        Case 407 To 490 'Listados de Reparaciones
            ListadosReparaciones H, W
    End Select
    
    
    Select Case OpcionListado
    
    'LISTADOS DE FACTURACION
    '-----------------------
        
    Case 54 '54: Listado Descuentos Familia/Marca
        H = 5775
        W = 6920
        PonerFrameVisible Me.FrameDtosFM, True, H, W
        ponerOptVisible True
        Me.Frame4.visible = True
        indFrame = 6
       ' txtCodigo(79).TabIndex = 318
       ' txtCodigo(80).TabIndex = 318
        cboVarios(0).visible = True
        optFrDto(5).Value = True
        
        'JUNIO 2014
        'Esta opcion es nueva. No debe poner nada
        
        
    Case 58 '58: listado Proveedores
        PonerFrameListadoVisible True, H, W
        Me.lblTitulo(1).Caption = "Listado Proveedores"
        indFrame = 1
        codigo = "{sprove.codprove}"
        Orden1 = "{sprove.codprove}"
        Orden2 = "{sprove.nomprove}"
        
        
    'LISTADOS DE REPARACIONES
    '-------------------------
    Case 60 '60: Informe N� Series
        H = 5415
        W = 6675
        PonerFrameVisible Me.FrameRepNSerie, True, H, W
        indFrame = 6
        codigo = "{sserie"
        
     Case 61, 65  'Listados de Motivos Pend. Rep.
        PonerFrameListadoVisible True, H, W
        Me.lblTitulo(1).Caption = "Listado de Motivos"
        indFrame = 1
        If OpcionListado = 61 Then
            codigo = "{smotre.codmotre}"
            Orden1 = "{smotre.codmotre}"
            Orden2 = "{smotre.nommotre}"
        Else
            codigo = "{smotba.codmotiv}"
            Orden1 = "{smotba.codmotiv}"
            Orden2 = "{smotba.desmotiv}"
        End If
        
    Case 63, 73, 223, 224, 248
                '63: Listado Reparaciones por D�a
                '73: Listado Altas Mantenimientos
                '223,224,248  Contabi facturas
                
        PonerFrameRepxDiaVisible True, H, W
        indFrame = 7
        If Me.OptProve.Tag = "" Then
            txtCodigo(31).Text = Format(Now, "dd/mm/yyyy")
            txtCodigo(32).Text = Format(Now, "dd/mm/yyyy")
        End If
        
        If OpcionListado = 223 Then
            Dim Cad As String
            If vParamAplic.ContabilizarTicketAgrupados Then
                Cad = "codtipom like 'FA%'"
            Else
                Cad = "codtipom like 'FA%' or codtipom='FTI'"
            End If
            Cad = "(" & Cad & " OR codtipom = 'FRT')"
            Cad = Cad & " and not isnull(letraser) and trim(letraser)<>''"
            
            'Febrero 2011
            'Solo los usuarios de B podran contabilizar fras de B
            If Val(vUsu.AlmacenPorDefecto) = vParamAplic.AlmacenB Then
                'Es usuario de B. Solo tienen b
                Cad = Cad & " and codtipom = 'FAZ'"
            Else
                'No ven el B
                Cad = Cad & " and codtipom <> 'FAZ'"
            End If
            CargarCombo_TipMov Me.cboTipMov, "stipom", "codtipom", "nomtipom", Cad, True
            
            'Si es usuario de B solo ha cargado el B
            If Val(vUsu.AlmacenPorDefecto) = vParamAplic.AlmacenB Then
                If cboTipMov.ListCount > 1 Then cboTipMov.ListIndex = 1
            End If
            
        End If
        
    Case 64, 406 'Listado Reparaciones por Cliente
                 '406: Listado Frecuencia de reparaciones
        H = 5415
        W = 6850
        PonerFrameVisible Me.FrameRepxClien, True, H, W
        indFrame = 8
        
        'txtCodigo(43).Text = Format(Now, "dd/mm/yyyy")
        txtCodigo(44).Text = Format(DateAdd("d", -1, Now), "dd/mm/yyyy")
        cadTitulo = "Reparaciones por Cliente"
        conSubRPT = False
        Me.Frame1.visible = (OpcionListado = 406)
        If OpcionListado = 406 Then
             cadTitulo = "Frecuencia de Reparaciones"
             Me.lblTitulo(8).Caption = "Frecuencia de Reparaciones"
             'Me.Label4(21).Caption = "Fecha Reparaci�n:"
             txtCodigo(0).Text = "1"
        End If
        
        
        
    Case 82, 83
        
        'LIstado etiquetas estanterias
        H = Me.FrameAlbaranesMarcaFacturar.Height
        W = FrameAlbaranesMarcaFacturar.Width
        PonerFrameVisible Me.FrameAlbaranesMarcaFacturar, True, H, W
        indFrame = 82
        If OpcionListado = 82 Then
            cadTitulo = "Poner marca facturaci�n"
            
        Else
            Label7(3).Caption = "Borre avisos cerrados"
        End If
        txtCodigo(117).visible = OpcionListado = 82
        txtCodigo(118).visible = OpcionListado = 82
        Frame7.visible = OpcionListado = 83
        conSubRPT = False
    Case 94, 513
        'LIstado etiquetas estanterias
        '513: es desde albaran compra
        FrameTapaEtiq.visible = OpcionListado = 513
        If OpcionListado = 513 Then
            Label3(111).Caption = CadTag
            CadTag = ""
            
            H = InStr(1, NumCod, " WHERE ")
            NumCod = Mid(NumCod, H)
            H = InStr(1, UCase(NumCod), " ORDER BY")
            NumCod = Mid(NumCod, 1, H)
            
            
        End If
        
        H = Me.FrameEtiqEstanteria.Height
        W = FrameEtiqEstanteria.Width
        PonerFrameVisible Me.FrameEtiqEstanteria, True, H, W
        indFrame = 94
        cadTitulo = "Etiq. estanteria"
        conSubRPT = False
        cboDecimal.ListIndex = 4
        
    Case 95, 101
    
        'LIstado etiquetas bultos
        H = Me.FrameBultos.Height
        W = FrameBultos.Width
        PonerFrameVisible Me.FrameBultos, True, H, W
        indFrame = 95
        cadTitulo = "Etiq. bultos"
        conSubRPT = False
        LimpiarTextosBultos
        Me.cmbBulto.Clear
        
        
        
        If OpcionListado = 95 Then
            '- Traer datos del Albaran: cliente, dpto, n� bultos
            If NumCod <> "" Then PonerCamposAlbaran
            conSubRPT = True
            Caption = "Cliente"
        Else
            Caption = "Proveedor"
            conSubRPT = False
        End If
        'Cliente
        txtClie.visible = conSubRPT
        Label4(61).visible = conSubRPT
        cmbBulto.visible = conSubRPT
        txtNombre(10).visible = conSubRPT
        imgBuscarG(75).visible = conSubRPT
        Label4(62).visible = conSubRPT
        optBultos(1).visible = conSubRPT
        optBultos(0).visible = conSubRPT
        optBultos(0).Caption = DevuelveTextoDepto(False)
        
        'Proveedor
        txtNombre(148).visible = Not conSubRPT
        txtCodigo(148).visible = Not conSubRPT
        imgBuscarG(148).visible = Not conSubRPT
        Label4(104).visible = Not conSubRPT
    Case 96
        
        H = Me.FrameFrecuencia.Height
        W = FrameFrecuencia.Width
        PonerFrameVisible Me.FrameFrecuencia, True, H, W
        indFrame = 96
        cadTitulo = "Etiq. bultos"
        conSubRPT = False
        HabilitarTextoCliente False
        
    Case 97
        H = Me.FrEliminarFacturas.Height
        W = Me.FrEliminarFacturas.Width
        PonerFrameVisible FrEliminarFacturas, True, H, W
        indFrame = 97
        cadTitulo = "Eliminar facturas"
        conSubRPT = False
        'Textos
        '--------------------------------------------------------------------
        Label11(0).Caption = "Este proceso es irreversible." & vbCrLf & " No deberia haber nadie trabajando en esta empresa y " & vbCrLf & _
            "deberia hacer una copia de seguridad."
        
        Label11(1).Caption = ""
        CargaFechasPosibleEliminacion
        
    Case 99
        
        H = Me.FrameHcoMante.Height
        W = Me.FrameHcoMante.Width
        PonerFrameVisible FrameHcoMante, True, H, W
        indFrame = 99
        cadTitulo = "Pasar a mantenimientos anulados"
        conSubRPT = False
        txtCodigo(110).Text = Format(Now, "dd/mm/yyyy")
    Case 512
        H = Me.FrameConta1FRAPRO.Height
        W = Me.FrameConta1FRAPRO.Width
        
        lblProvCon(2).Caption = RecuperaValor(NumCod, 1)
        lblProvCon(3).Caption = RecuperaValor(NumCod, 2)
        lblProvCon(0).Caption = ""
        lblProvCon(1).Caption = ""
        cadSelect = CStr(CadTag)
        CadTag = ""
        NumCod = ""
        PonerFrameVisible FrameConta1FRAPRO, True, H, W
        
    Case 514
        H = Me.FrameFraProveedor.Height
        W = Me.FrameFraProveedor.Width
        indFrame = 16
        PonerFrameVisible FrameFraProveedor, True, H, W
        Label4(106).visible = False 'mas de 5 vtos
        For indCodigo = 0 To 4
            PonerVisiblesParaVencimientosFraPro False
        Next
        
    End Select
    
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel(indFrame).Cancel = True
    Me.Width = W + 70
    Me.Height = H + 350
End Sub



Private Sub Form_Unload(Cancel As Integer)
    NumCod = ""
    CadTag = ""
End Sub



Private Sub frmF_Selec(vFecha As Date)
'Calendario de Fecha
    txtCodigo(indCodigo).Text = Format(vFecha, "dd/mm/yyyy")
End Sub


Private Sub frmMtoActiv_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Actividades de Clientes
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub frmMtoAgentes_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Agentes Comerciales
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoAlPropios_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Almacenes Propios
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub frmMtoArticulos_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Articulos
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoCartasOfe_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Cartas de Oferta
    txtCodigo(32).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(32).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub frmMtoClientes_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Clientes
    If indCodigo > 0 Then
        txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
        txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
    Else
        'EL 0 es para el listado de bultos
        Me.txtClie.Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
        txtClie_LostFocus
        
    End If

End Sub

Private Sub frmMtoFamilia_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Familia de Articulos
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoFEnvio_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Formas de Envio
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoIncid_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoMarcas_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Marcas de Art�culos
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoMotivos_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Marcas de Art�culos
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoProveedor_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Proveedores
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoRutas_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Rutas
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoSituac_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Situaciones
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoTarifas_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Tarifas
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoTArticulo_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Tipos de Art�culo
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoTiposCon_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Tipos de Contrato
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoTraba_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Proveedores
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoTUnidad_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Tipos de Unidad
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoUbica_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Ubicaciones de Almacen
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoZonas_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Zonas
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub imgayuda_Click(Index As Integer)
    Select Case Index
    Case 0
        codigo = "->Etiquetas." & vbCrLf
        codigo = codigo & "-Si marca con 'stock m�nimo' pedira un almacen y mostrara" & vbCrLf
        codigo = codigo & "los articulos que tenga valor para ese dato" & vbCrLf
        codigo = codigo & vbCrLf & "->PVP" & vbCrLf
        codigo = codigo & "-Listado PVP, con IVA" & vbCrLf
    Case 1
        codigo = "Mostrara los datos de stock minimo,maximo , punto de pedido y stock." & vbCrLf
        codigo = codigo & "Si marca sin 'stock m�nimo' mostrar� los articulos que tienen stock" & vbCrLf
        codigo = codigo & "y no tienen valor en el campo stock minimo" & vbCrLf
        
    Case 2
        codigo = "Descuentos familia marca." & vbCrLf & vbCrLf
        codigo = codigo & "CLIENTE - ACTIVIDAD: Mostrara todos los descuentos del cliente y los de la actividad que le corresponda."
        codigo = codigo & " No tendra en cuenta el resto de desde/hasta, solo cliente"
        codigo = codigo & vbCrLf & vbCrLf
        codigo = codigo & "Resto opciones: Mostrara los descuentos desde la tabla de descuentos familia marca, teniendo en cuenta"
        codigo = codigo & " la opcion del proveedor de ocultar en listados descuento"
    Case 3
        codigo = "El procentaje de margen puede calcularse sobre el coste o sobre las ventas"
     
    End Select
    
    MsgBox codigo, vbInformation
End Sub

Private Sub imgBuscar_Click(Index As Integer)
    Screen.MousePointer = vbHourglass
    imgBuscar(1).Tag = Index
    indCodigo = Index
    
    Select Case Index
    Case 1, 2 'FrameListado
        Select Case OpcionListado
            Case 1 'Listado de MARCAS
                AbrirFrmMarcas
                    
            Case 2 'Listado de ALMACENES Propios
                AbrirFrmAlmPropios
            
            Case 3  'Listado de Tipos de Unidad
                Set frmMtoTUnidad = New frmAlmTipoUnidad
                frmMtoTUnidad.DatosADevolverBusqueda = "0|1"
                frmMtoTUnidad.DeConsulta = True
                frmMtoTUnidad.Show vbModal
                Set frmMtoTUnidad = Nothing
            
            Case 4  'Listado de Tipos de Articulos
                AbrirFrmTipoArt

            Case 110 'Listado de Ubicaciones de Almacen
                Set frmMtoUbica = New frmAlmUbicaciones
                frmMtoUbica.DatosADevolverBusqueda = "0|1"
                frmMtoUbica.DeConsulta = True
                frmMtoUbica.Show vbModal
                Set frmMtoUbica = Nothing
        
            
            Case 20 'Listado de Actividades de Clientes
                AbrirFrmActividades
            
            Case 21 'Listado de Zonas de Clientes
                AbrirFrmZonas
            
            Case 22 'Listado de Rutas de Asistencia
                AbrirFrmRutas
                
'                Set frmMtoRutas = New frmFacRutas
'                frmMtoRutas.DatosADevolverBusqueda = "0|1"
'                frmMtoRutas.DeConsulta = True
'                frmMtoRutas.Show vbModal
'                Set frmMtoRutas = Nothing
            
            Case 23 'Listado de Formas de Env�o
                Set frmMtoFEnvio = New frmFacFormasEnvio
                frmMtoFEnvio.DatosADevolverBusqueda = "0|1"
                frmMtoFEnvio.DeConsulta = True
                frmMtoFEnvio.Show vbModal
                Set frmMtoFEnvio = Nothing
            
            Case 24 'Listado de Tarifas Venta
                AbrirFrmTarifas
            
            Case 27 'Listado de Situaciones Especiales
                Set frmMtoSituac = New frmFacSituaciones
                frmMtoSituac.DatosADevolverBusqueda = "0|1"
                frmMtoSituac.DeConsulta = True
                frmMtoSituac.Show vbModal
                Set frmMtoSituac = Nothing
                
            Case 58
                'DAVID
                indCodigo = Index
                Set frmMtoProveedor = New frmComProveedores
                frmMtoProveedor.DatosADevolverBusqueda = "0|1"
                frmMtoProveedor.Show vbModal
                Set frmMtoProveedor = Nothing
            Case 61 'Listado de Motivos Pend. Rep.
                Set frmMtoMotivos = New frmRepMotivosPend
                frmMtoMotivos.DatosADevolverBusqueda = "0|1"
                frmMtoMotivos.DeConsulta = True
                frmMtoMotivos.Show vbModal
                Set frmMtoMotivos = Nothing
        End Select
        
    Case 3, 4 'FrameInfAlmacen
            If OpcionListado = 7 Or OpcionListado = 8 Then
'            Case 7, 8 '7: Informe de Traspasos de Almacenes
                  '8: Informe de Movimientos de Almacen
                MandaBusquedaPrevia ""
            End If
    End Select
    
    PonerFoco Me.txtCodigo(Index)
    Screen.MousePointer = vbDefault
End Sub


Private Sub imgBuscarG_Click(Index As Integer)
'Buscar general: cada index llama a una tabla
    Screen.MousePointer = vbHourglass
    
    Select Case Index
        Case 0, 1, 6, 7, 35, 36, 43, 44, 49, 50, 75, 76, 77, 80, 81, 93, 94, 101, 102 'cod. CLIENTE
            Select Case Index
                Case 0, 1: indCodigo = Index + 73
                Case 6, 7: indCodigo = Index + 27
                Case 35, 36: indCodigo = Index + 20
                Case 43, 44: indCodigo = Index + 4
                Case 49, 50: indCodigo = Index - 12
                Case 75: indCodigo = 0
                Case 76, 77, 80, 81: indCodigo = Index + 22
                Case 93, 94: indCodigo = Index + 24
                Case 101, 102: indCodigo = Index + 31
            End Select
            AbrirFrmClientes
        
        Case 2, 3, 13, 14, 19, 20, 31, 32, 57, 58, 67, 68, 73, 74, 141, 142 'cod. FAMILIA
            Select Case Index
                Case 2, 3: indCodigo = Index + 73
                Case 13, 14: indCodigo = Index + 3
                Case 19, 20: indCodigo = Index + 43
                Case 31, 32: indCodigo = Index - 24
                Case 57, 58: indCodigo = Index - 32
                Case 67, 68, 73, 74: indCodigo = Index + 21
                Case 141, 142:  indCodigo = Index
            End Select
            Set frmMtoFamilia = New frmAlmFamiliaArticulo
            frmMtoFamilia.DatosADevolverBusqueda = "0|1"
            frmMtoFamilia.Show vbModal
            Set frmMtoFamilia = Nothing
            
            
        Case 90, 91, 92
            indCodigo = 22 + Index
            Set frmMtoIncid = New frmIncidencias
            frmMtoIncid.DatosADevolverBusqueda = "0|1|"
            frmMtoIncid.Show vbModal
            Set frmMtoIncid = Nothing
            
        Case 4, 5, 21, 22, 59, 60 'cod. MARCA
            Select Case Index
                Case 4, 5: indCodigo = Index + 73
                Case 21, 22: indCodigo = Index + 43
                Case 59, 60:  indCodigo = Index - 32
            End Select
            AbrirFrmMarcas
            
        Case 8, 9, 51, 52 'cod. Direc/DPTO
'            Select Case Index
'                Case 8, 9:
'                Case 51, 52: indCodigo = Index - 12
'            End Select
        
            If Index = 51 Or Index = 52 Then
                'Desde hsta departamento en Numserie
                'Si no teinen el mismo cliente NO pude ver dpto
                If txtCodigo(37).Text = "" And txtCodigo(38).Text = "" Then
                        MsgBox "Ponga un cliente", vbExclamation
                    
                ElseIf (txtCodigo(37).Text <> txtCodigo(38).Text) Then
                    MsgBox "No ha puesto el mismo cliente", vbExclamation
                Else
                    
                    indCodigo = 39 + (Index - 51)
                    MandaBusquedaPrevia "codclien = " & txtCodigo(37).Text

                End If
            End If
        Case 10, 18, 33, 34, 139, 140, 107, 108, 109 'cod. ALMACEN
            Select Case Index
                Case 10: indCodigo = Index + 3
                Case 18: indCodigo = Index + 54
                Case 33, 34: indCodigo = Index - 22
                Case 139, 140: indCodigo = Index
                Case 107, 108, 109: indCodigo = indCodigo + 38
                    
            End Select
            AbrirFrmAlmPropios
            
        Case 11, 12, 27, 28, 29, 30, 61, 62, 69, 70, 71, 72, 95, 98 'cod. ARTICULO
            ' ====  [16/09/2009] LAURA : Listado Articulos con componentes (a�ade index 95 y 98)
            Select Case Index
                Case 11, 12: indCodigo = Index + 3
                Case 27, 28: indCodigo = Index + 43
                Case 29, 30: indCodigo = Index - 24
                Case 61, 62: indCodigo = Index - 32
                Case 69, 70, 71, 72: indCodigo = Index + 21
                ' ====  [16/09/2009] LAURA : Listado Articulos con componentes (a�ade index 95 y 98)
                Case 95: indCodigo = 125
                Case 98: indCodigo = 126
                ' ====
            End Select
            Set frmMtoArticulos = New frmAlmArticu2
            'frmMtoArticulos.DatosADevolverBusqueda3 = "@1@" 'Abrimos en Modo Busqueda
            frmMtoArticulos.DesdeTPV = False
            frmMtoArticulos.Show vbModal
            Set frmMtoArticulos = Nothing
            
        Case 25, 26 'cod TIPO ARTICULO
            indCodigo = Index + 43
            AbrirFrmTipoArt

        Case 55, 56
            indCodigo = Index - 32
            If OpcionListado = 30 Then 'segun Informe mismo boton abre 2 distintas
               AbrirFrmClientes
            Else 'cod. TARIFA
                AbrirFrmTarifas
            End If
            
        Case 15, 16, 23, 24, 63, 64, 103, 104, 143, 144, 148 'cod. PROVEEDOR
            Select Case Index
                Case 15, 16: indCodigo = Index + 3
                Case 23, 24: indCodigo = Index + 43
                Case 63, 64: indCodigo = Index + 16
                Case 103, 104: indCodigo = Index + 31
                Case 143, 144, 148: indCodigo = Index
            End Select
            Set frmMtoProveedor = New frmComProveedores
            frmMtoProveedor.DatosADevolverBusqueda = "0|1"
            frmMtoProveedor.Show vbModal
            Set frmMtoProveedor = Nothing
            
            'Para el de etiquetas bulkto proveedor, fuerzo un lostfocus
            If Index = 148 Then txtCodigo_LostFocus Index
        Case 41, 42
            If Index <= 42 Then
                indCodigo = Index + 4
            Else
                '86,88
                indCodigo = Index + 20
            End If
            AbrirFrmZonas
            
        Case 17, 96, 97, 89 'cod. TRABAJADOR
            If Index = 89 Then
                indCodigo = 111
            Else
                indCodigo = 21
            End If
            Set frmMtoTraba = New frmAdmTrabajadores
            frmMtoTraba.DatosADevolverBusqueda = "0|1"
            frmMtoTraba.Show vbModal
            Set frmMtoTraba = Nothing
            
        Case 45, 46 'cod. AGENTE
            indCodigo = Index + 4
            Set frmMtoAgentes = New frmFacAgentesCom
            frmMtoAgentes.DatosADevolverBusqueda = "0|1"
            frmMtoAgentes.Show vbModal
            Set frmMtoAgentes = Nothing
            
        Case 37, 38, 47, 48, 82, 83 'cod. TIPO CONTRATO (= n� mantenimiento)
            Select Case Index
                Case 37, 38: indCodigo = Index + 20
                Case 47, 48: indCodigo = Index + 4
                Case 82, 83: indCodigo = Index + 22
            End Select
            Set frmMtoTiposCon = New frmManTiposContrato
            frmMtoTiposCon.DatosADevolverBusqueda = "0|1"
            frmMtoTiposCon.Show vbModal
            Set frmMtoTiposCon = Nothing
        
        Case 39, 40, 53, 54 'cod. N� CONTRATO (= n� mantenimiento)

        
        Case 84, 85, 86, 88, 105, 106 'RUTA DEL CLIENTE
            If Index <= 85 Then
                indCodigo = Index
            ElseIf Index <= 88 Then
                '86,88
                indCodigo = Index + 20
            Else
                '105,106 list mtos
                indCodigo = Index + 32
            End If
            
            AbrirFrmRutas
            
        ' ---- [30/10/2009] (LAURA) : Agrupar etiquetas mantenimiento por cliente, departamento
        Case 99, 100 'COD. ACTIVIDAD
            indCodigo = Index + 28
            AbrirFrmActividades
        ' ----
        Case 87
            indCodigo = 107
            AbrirFrmTarifas
    End Select
    PonerFoco txtCodigo(indCodigo)
    Screen.MousePointer = vbDefault
End Sub


Private Sub imgFecha_Click(Index As Integer)
   Screen.MousePointer = vbHourglass
   imgFecha(0).Tag = Index
   Set frmF = New frmCal
   frmF.Fecha = Now
   
   Select Case Index
        Case 0 'frameMovArtic
            indCodigo = 9
        Case 1 'frameMovArtic
            indCodigo = 10
        Case 2 'frameInventario (indFrame=4)
            indCodigo = 20
        Case 3 'frameInventario (indFrame=4)
            indCodigo = 22
        Case 4 'frameReparacionesxDia (indFrame=7)
            indCodigo = 31
        Case 5 'frameReparacionesxDia (indFrame=7)
            indCodigo = 32
        Case 6 'frameReparacionesxClien (indFrame=8)
            indCodigo = 43
        Case 7 'frameReparacionesxClien (indFrame=8)
            indCodigo = 44
        Case 8 'frameMAntenimientos
            indCodigo = 53
        Case 9 'frameMAntenimientos
            indCodigo = 54
        Case 10 'FrameListAvisosPtes
            indCodigo = 82
        Case 11 'FrameListAvisosPtes
            indCodigo = 83
        Case 13, 14
            indCodigo = Index + 102
        Case 15, 16
            indCodigo = Index + 104
        Case 17, 18
            indCodigo = Index + 106
        Case 19, 20
             indCodigo = Index + 111
        Case 21 To 25
            indCodigo = Index + 128
        Case 109
            indCodigo = 109
   End Select
   
   
   PonerFormatoFecha txtCodigo(indCodigo)
   If txtCodigo(indCodigo).Text <> "" Then frmF.Fecha = CDate(txtCodigo(indCodigo).Text)
   
   Screen.MousePointer = vbDefault
   frmF.Show vbModal
   Set frmF = Nothing
   PonerFoco txtCodigo(indCodigo)
End Sub




Private Sub ListView1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub optBultos_Click(Index As Integer)
    
    If Me.txtClie.Text <> "" Then
        'Limpiamos textos y cargamos las direcciones
        txtClie_LostFocus
    End If
End Sub

Private Sub OptClientes_Click()
    If Me.OptClientes.Value = True Then
        Label2(2).Caption = "Fecha Factura: "
    End If
    
    Me.FrameTipMov.visible = (OpcionListado = 223) And Me.OptClientes.Value = True
End Sub

Private Sub Optcodigo_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub optDirEnvio_Click(Index As Integer)
    If Index = 0 Then
        txtNombre(0).Text = RecuperaValor(txtNombre(0).Tag, 1)
    Else
        txtNombre(0).Text = RecuperaValor(txtNombre(0).Tag, 2)
    End If
End Sub

Private Sub optDirEnvio_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub



Private Sub optFrDto_KeyPress(Index As Integer, KeyAscii As Integer)
     KEYpress KeyAscii
End Sub

Private Sub OptNombre_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        PonerFocoBtn Me.cmdAceptar(1)
    ElseIf KeyAscii = 27 Then Unload Me  'ESC
    End If
End Sub

Private Sub OptProve_Click()
    If Me.OptProve.Value = True Then
        Label2(2).Caption = "Fecha Recepci�n: "
    End If
    
     Me.FrameTipMov.visible = (OpcionListado = 223) And Me.OptClientes.Value = True
    
End Sub


Private Sub txtBultos_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index <> 0 Then KEYpress KeyAscii
End Sub

Private Sub txtBultos_LostFocus(Index As Integer)
    If Index = 1 Or Index = 7 Then
        'Campos NUMERICOS
        txtBultos(Index).Text = Trim(txtBultos(Index).Text)
        If txtBultos(Index).Text <> "" Then
            If Not PonerFormatoEntero(txtBultos(Index)) Then
                txtBultos(Index).Text = ""
                PonerFoco txtBultos(Index)
            End If
        End If
    End If
End Sub

Private Sub txtClie_GotFocus()
    PonerFoco txtClie
    
End Sub

Private Sub txtClie_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtClie_LostFocus()
Dim Reestablecer As Boolean
Dim Clivario As Boolean
Dim Rs As ADODB.Recordset
Dim Ind As Integer
                Screen.MousePointer = vbHourglass
                txtClie.Text = Trim(txtClie.Text)
                Orden2 = ""
                Clivario = False
                If txtClie = "" Then
                    Reestablecer = True
                Else
                    If Not PonerFormatoEntero(txtClie) Then
                        Reestablecer = True
                    Else
                        cmbBulto.Clear
                        Set Rs = New ADODB.Recordset
                        codigo = "select nomclien,domclien,sclien.codpobla as cpos,sclien.pobclien,proclien,"
                        If Me.optBultos(1).Value Then
                            'Direcciones de envio
                            'nomdiren   domdiren     pobdiren   pobdiren    prodiren
                            codigo = codigo & "nomdiren  nomdirec ,domdiren  domdirec ,pobdiren pobdirec ,sdirenvio.codpobla  ,prodiren prodirec, coddiren CodDirec"
                            codigo = codigo & ",clivario from sclien left join sdirenvio on sclien.codclien=sdirenvio.codclien "
                            
                        
                        Else
                            'Departamentos
                            codigo = codigo & " nomdirec ,  domdirec ,pobdirec ,sdirec.codpobla  ,prodirec, CodDirec"
                            codigo = codigo & ",clivario from sclien left join sdirec on sclien.codclien=sdirec.codclien "
    
                        End If
                                
                        codigo = codigo & " WHERE sclien.codclien =" & txtClie.Text
                        codigo = codigo & " order by 6"   'nomdirec nomdiren
                        Rs.Open codigo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                        Orden1 = ""
                        
                        While Not Rs.EOF
                            'Meto primero la direccion de la ficha
                            If Orden1 = "" Then
                                cmbBulto.AddItem "Ppal:  " & DBLet(Rs.Fields(1), "T") & " - " & DBLet(Rs.Fields(3), "T")
                                txtBultos(2).Tag = DBLet(Rs.Fields(1), "T") & "|"
                                txtBultos(3).Tag = DBLet(Rs.Fields(3), "T") & "|"
                                txtBultos(4).Tag = DBLet(Rs.Fields(2), "T") & "|"
                                txtBultos(5).Tag = DBLet(Rs.Fields(4), "T") & "|"
                                txtBultos(6).Tag = "|"
                                Orden1 = "T"
                                
                                Orden2 = Rs!Nomclien
                                Clivario = DBLet(Rs!Clivario, "N") = 1
                            End If
                            'Las direcciones alternativas
                            If Not IsNull(Rs!domdirec) Or Not IsNull(Rs!domdirec) Then
                                'TIENE DIRECCION ALTERNATIVA
                                txtBultos(2).Tag = txtBultos(2).Tag & DBLet(Rs!domdirec, "T") & "|"
                                txtBultos(3).Tag = txtBultos(3).Tag & DBLet(Rs!pobdirec, "T") & "|"
                                txtBultos(4).Tag = txtBultos(4).Tag & DBLet(Rs!codpobla, "T") & "|"
                                txtBultos(5).Tag = txtBultos(5).Tag & DBLet(Rs!prodirec, "T") & "|"
                                txtBultos(6).Tag = txtBultos(6).Tag & "|"
'                                cmbBulto.AddItem "       " & DBLet(RS!domdirec, "T") & " - " & DBLet(RS!pobdirec, "T")
                                cmbBulto.AddItem DBLet(Rs!nomdirec, "T") & ":   " & DBLet(Rs!domdirec, "T") & " - " & DBLet(Rs!pobdirec, "T")
                                If Me.CadTag = CStr(DBLet(Rs!CodDirec, "N")) Then
                                    Ind = cmbBulto.ListCount - 1
                                End If
                            End If
                            Rs.MoveNext
                        Wend   '
                        If cmbBulto.ListCount > 0 Then
                            If Ind > 0 Then
                                cmbBulto.ListIndex = Ind
                            Else
                                cmbBulto.ListIndex = 0
                            End If
                            'PonerCamposDireccionBultos 0 'Lo hace el poner a 0 el list index
                        Else
                            Reestablecer = True
                        End If
                        Rs.Close
                        Set Rs = Nothing

                        
                    End If
                End If
                    'La direccion
                If Reestablecer Then
                    txtClie.Text = ""
                    'Hbilitamos o no
                    cmbBulto.Clear
                    LimpiarTextosBultos
                    txtNombre(10).Text = ""
                    Clivario = False
                Else
                    
                    txtNombre(10).Text = Orden2
                End If
                HabilitarTextoCliente Clivario
                
             Screen.MousePointer = vbDefault
    
End Sub

Private Sub HabilitarTextoCliente(Habilitar As Boolean)
    If Not Habilitar Then
        txtNombre(10).BackColor = &H80000018
    Else
        txtNombre(10).BackColor = &H80000005
    End If
    txtNombre(10).Locked = Not Habilitar
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
Dim tabla As String
Dim codCampo As String, NomCampo As String
Dim TipCampo As String, Formato As String
Dim Titulo As String
Dim EsNomCod As Boolean 'Si es campo Cod-Descripcion llama a PonerNombreDeCod


    'Quitar espacios en blanco por los lados
    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)

    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    EsNomCod = False
    If Index = 1 Or Index = 2 Then
    'el mismo frame ( y por tanto los mismos campos) se utilizan para distintos
    'informes. Seg�n de donde llamemos c�digo de una tabla u otra
        Select Case OpcionListado
            Case 1 'Listado MARCAS
                EsNomCod = True
                tabla = "smarca"
                codCampo = "codmarca"
                NomCampo = "nommarca"
                TipCampo = "N"
                Formato = "0000"
                Titulo = "Marca"
                
            Case 2 'Listado ALMACENES Propios
                EsNomCod = True
                tabla = "salmpr"
                codCampo = "codalmac"
                NomCampo = "nomalmac"
                TipCampo = "N"
                Formato = "000"
                Titulo = "Almacen Propio"
                
            Case 3 'Listado Tipos UNIDADES
                EsNomCod = True
                tabla = "sunida"
                codCampo = "codunida"
                NomCampo = "nomunida"
                TipCampo = "N"
                Formato = "00"
                Titulo = "Tipo Unidad"
                
            Case 4 'Listado Tipos Art�culos
                txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), 1, "stipar", "nomtipar", "codtipar", "Tipo de Art�culo", "T")
    
            Case 110 'Listado Ubicaciones Almacen
                txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, "subica", "nomubica", "codubica", "Ubicaciones Almacen", "T")
            
            
            Case 20 'Listado ACTIVIDADES de Clientes
                EsNomCod = True
                tabla = "sactiv"
                codCampo = "codactiv"
                NomCampo = "nomactiv"
                TipCampo = "N"
                Formato = "000"
                Titulo = "Actividad de Cliente"
            
            Case 21 'Listado ZONAS de Clientes
                EsNomCod = True
                tabla = "szonas"
                codCampo = "codzonas"
                NomCampo = "nomzonas"
                TipCampo = "N"
                Formato = "000"
                Titulo = "Zona de Cliente"
            
            Case 22 'Listado RUTAS de Asistencia
                EsNomCod = True
                tabla = "srutas"
                codCampo = "codrutas"
                NomCampo = "nomrutas"
                TipCampo = "N"
                Formato = "000"
                Titulo = "Ruta de Asistencia"
            
            Case 23 'Listado Formas de Env�o
                EsNomCod = True
                tabla = "senvio"
                codCampo = "codenvio"
                NomCampo = "nomenvio"
                TipCampo = "N"
                Formato = "000"
                Titulo = "Forma de Env�o"
            
            Case 24 'Listado Tarifas Venta
                EsNomCod = True
                tabla = "starif"
                codCampo = "codlista"
                NomCampo = "nomlista"
                TipCampo = "N"
                Formato = "000"
                Titulo = "Tarifa de Venta"
            
            Case 27 'Listado SITUACIONES Especiales
                EsNomCod = True
                tabla = "ssitua"
                codCampo = "codsitua"
                NomCampo = "nomsitua"
                TipCampo = "N"
                Formato = "00"
                Titulo = "Situaci�n Especial"
            
            Case 58 'Listado PROVEEDORES
                EsNomCod = True
                tabla = "sprove"
                codCampo = "codprove"
                NomCampo = "nomprove"
                TipCampo = "N"
                Formato = "000000"
                Titulo = "Proveedor"
            
            Case 61 'Listado MOTIVOS Pend. Rep.
                EsNomCod = True
                tabla = "smotre"
                codCampo = "codmotre"
                NomCampo = "nommotre"
                TipCampo = "N"
                Formato = "00"
                Titulo = "Motivos Pend. Rep."
                
            Case 65 'Listados NOTIVOS baja equipos
                EsNomCod = True
                tabla = "smotba"
                codCampo = "codmotiv"
                NomCampo = "desmotiv"
                TipCampo = "N"
                Formato = "00"
                Titulo = "Motivos Baja equipos"
        End Select
        
    ElseIf Index = 3 Or Index = 4 Then
         '7: Informe Traspaso Almacenes
         '8: Informe Movimientos Almacen
         txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "0000000")
    Else
        Select Case Index
        Case 0, 86, 87
            If txtCodigo(Index).Text <> "" Then
                PonerFormatoEntero txtCodigo(Index)
                If (Index = 86 Or Index = 87) Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000000")
            End If
            
        Case 5, 6, 14, 15, 29, 30, 70, 71, 90, 91, 92, 93, 125, 126 'Cod. ARTICULO
            ' ====  [16/09/2009] LAURA : a�ade index 125,126
            EsNomCod = True
            tabla = "sartic"
            codCampo = "codartic"
            NomCampo = "nomartic"
            TipCampo = "T"
            Titulo = "Art�culo"
            txtCodigo(Index).Text = UCase(txtCodigo(Index).Text)
        
        Case 7, 8, 16, 17, 25, 26, 62, 63, 75, 76, 88, 89, 94, 95, 141, 142 'Cod. FAMILIA
            EsNomCod = True
            tabla = "sfamia"
            codCampo = "codfamia"
            NomCampo = "nomfamia"
            TipCampo = "N"
            Formato = "0000"
            Titulo = "Familia"
        
        'FECHA Desde Hasta
        Case 9, 10, 20, 22, 31, 32, 43, 44, 53, 54, 82, 83, 109, 110, 115, 116, 119, 120, 123, 124, 130, 131
            If txtCodigo(Index).Text <> "" Then
                If Index = 22 And OpcionListado = 19 Then 'Este campo sera Hora y no Fecha
                    PonerFormatoHora txtCodigo(Index)
                Else
                    PonerFormatoFecha txtCodigo(Index)
                    If OpcionListado = 223 And txtCodigo(Index).Text <> "" Then
                        'Contabilizar facturas
                        If Not ComprobarFechasConta(Index) Then PonerFoco txtCodigo(Index)
                    End If
                End If
            End If
        Case 149 To 153
            'Fechas vencimiento frapro
            If txtCodigo(Index).Text = "" Then
                txtCodigo(Index).Text = txtCodigo(Index).Tag
            Else
                If Not EsFechaOK(txtCodigo(Index)) Then
                    txtCodigo(Index).Text = txtCodigo(Index).Tag
                Else
                    PonerFormatoFecha txtCodigo(Index)
                End If
            End If
        Case 11, 12, 13, 72, 139, 140, 145, 146, 147 'ALMACENES Propios
            EsNomCod = True
            tabla = "salmpr"
            codCampo = "codalmac"
            NomCampo = "nomalmac"
            TipCampo = "N"
            Formato = "000"
            Titulo = "Almacen Propio"
            
        Case 18, 19, 66, 67, 79, 80, 134, 135, 143, 144, 148 'PROVEEDOR
            EsNomCod = True
            tabla = "sprove"
            codCampo = "codprove"
            NomCampo = "nomprove"
            TipCampo = "N"
            Formato = "000000"
            Titulo = "Proveedor"
        
        Case 21, 96, 97, 111 'Cod. Operario/Trabajador
            EsNomCod = True
            tabla = "straba"
            codCampo = "codtraba"
            NomCampo = "nomtraba"
            TipCampo = "N"
            Formato = "0000"
            Titulo = "Trabajador"
        
        Case 23, 24, 107
            EsNomCod = True
            TipCampo = "N"
            If OpcionListado = 30 Then 'Precios Especiales
                tabla = "sclien"
                codCampo = "codclien"
                NomCampo = "nomclien"
                Formato = "000000"
                Titulo = "Cliente"
            Else   'Tarifas Precios
                tabla = "starif"
                codCampo = "codlista"
                NomCampo = "nomlista"
                Formato = "000"
                Titulo = "Tarifa de Venta"
            End If
        
        Case 27, 28, 64, 65, 77, 78 'MARCAS
            EsNomCod = True
            tabla = "smarca"
            codCampo = "codmarca"
            NomCampo = "nommarca"
            TipCampo = "N"
            Formato = "0000"
            Titulo = "Marca"
        
        Case 31 'N� de Oferta
            If txtCodigo(Index).Text = "" Then Exit Sub
            codCampo = DevuelveDesdeBDNew(conAri, "scapre", "numofert", "numofert", txtCodigo(Index).Text, "N")
            If codCampo = "" Then
                MsgBox "No existe el c�digo de Oferta: " & NumCod, vbInformation
                PonerFoco txtCodigo(Index)
            Else
                txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "0000000")
            End If
            
        Case 32, 43 'Carta de la Oferta
            EsNomCod = True
            tabla = "scartas"
            codCampo = "codcarta"
            NomCampo = "descarta"
            TipCampo = "N"
            Formato = "000"
            Titulo = "Cartas para Ofertas"
            
        Case 37, 38, 33, 34, 47, 48, 55, 56, 73, 74, 98, 99, 102, 103, 117, 118, 132, 133 'Cod. CLIENTE
            EsNomCod = True
            tabla = "sclien"
            codCampo = "codclien"
            NomCampo = "nomclien"
            TipCampo = "N"
            Formato = "000000"
            Titulo = "Cliente"
            
        Case 112, 113, 114
            EsNomCod = True
            tabla = "sincid"
            codCampo = "codincid"
            NomCampo = "nomincid"
            TipCampo = "T"
            'Formato = "0000"
            Titulo = "Incidencias"
            
        Case 39, 40, 35, 36 'Direcc./Dpto del Cliente
            If txtCodigo(Index).Text = "" Then
                txtNombre(Index).Text = ""
                Exit Sub
            End If
            txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000")
            'comprobar el departamento del cliente, cuando en el campo
            'Desde/Hasta se ha seleccionado un �nico cliente
            If Index = 39 Or Index = 40 Then
                If txtCodigo(37).Text <> txtCodigo(38).Text Then
                    MsgBox "Seleccionar dpto/direc. solo cuando se seleccione un �nico cliente.", vbInformation
                    txtCodigo(Index).Text = ""
                    Exit Sub
                End If
            ElseIf Index = 35 Or Index = 36 Then
                If txtCodigo(33).Text <> txtCodigo(34).Text Then
                    MsgBox "Seleccionar dpto/direc. solo cuando se seleccione un �nico cliente.", vbInformation
                    txtCodigo(Index).Text = ""
                    Exit Sub
                End If
            End If
            
            'Comprobar que el cliente seleccionado tiene esa direccion o dpto
            codCampo = DevuelveDesdeBDNew(conAri, "sdirec", "nomdirec", "codclien", txtCodigo(Index - 2).Text, "N", , "coddirec", txtCodigo(Index).Text, "N")
            txtNombre(Index).Text = codCampo 'Nombre direc. o dpto
            If codCampo = "" Then 'No existe el dpto
                'FALTA###
'                If vParamAplic.Departamento Then
'                    codCampo = " el Departamento "
'                Else
'                    codCampo = " la Direcci�n "
'                End If
                codCampo = "No existe" & codCampo & txtCodigo(Index).Text & " para el cliente: "
                codCampo = codCampo & txtCodigo(Index - 2).Text & " - " & txtNombre(Index - 2).Text
                MsgBox codCampo, vbInformation
                txtCodigo(Index).Text = ""
                PonerFoco txtCodigo(Index)
            End If
        
        Case 41, 42, 59, 60 'N� Contrato
'            If txtCodigo(Index).Text <> "" Then
'                txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "0000000")
'            End If

        Case 45, 46  'ZONAS del Cliente
            EsNomCod = True
            tabla = "szonas"
            codCampo = "codzonas"
            NomCampo = "nomzonas"
            TipCampo = "N"
            Formato = "000"
            Titulo = "Zonas de Clientes"
        
        Case 49, 50 'Cod. AGENTE
            EsNomCod = True
            tabla = "sagent"
            codCampo = "codagent"
            NomCampo = "nomagent"
            TipCampo = "N"
            Formato = "0000"
            Titulo = "Agente"
            
        Case 51, 52, 57, 58, 104, 105 'Tipos Contratos/MAntenimientos
            EsNomCod = True
            tabla = "stipco"
            codCampo = "codtipco"
            NomCampo = "nomtipco"
            TipCampo = "T"
            Titulo = "Tipos de Contratos"
            
        Case 61 'A�o Ejercicio
            If txtCodigo(Index).Text = "" Then Exit Sub
            If Not IsNumeric(txtCodigo(Index).Text) Then
                MsgBox "El Ejercicio debe ser un A�o", vbInformation
                Exit Sub
            End If
        
        Case 68, 69 'Tipos de Articulos
            EsNomCod = True
            tabla = "stipar"
            codCampo = "codtipar"
            NomCampo = "nomtipar"
            TipCampo = "T"
            Titulo = "Tipo de Articulo"
            
        Case 84, 85, 106, 108, 137, 138 'RUTAS del cliente
            EsNomCod = True
            tabla = "srutas"
            codCampo = "codrutas"
            NomCampo = "nomrutas"
            TipCampo = "N"
            Formato = "000"
            Titulo = "Ruta de Asistencia"
            
        Case 127, 128 'ACTIVIDADES del cliente
            EsNomCod = True
            tabla = "sactiv"
            codCampo = "codactiv"
            NomCampo = "nomactiv"
            TipCampo = "N"
            Formato = "000"
            Titulo = "Actividades"
            
        Case 121, 122 'N� Factura
            If PonerFormatoEntero(txtCodigo(Index)) Then
                
                
            End If
        Case 136
            If PonerFormatoDecimal(txtCodigo(Index), 4) Then
                
                
            End If
        End Select
    End If
    
    If EsNomCod Then

        If TipCampo = "N" Then
            If PonerFormatoEntero(txtCodigo(Index)) Then
                txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, tabla, NomCampo, codCampo, Titulo, TipCampo)
            
                If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, Formato)
                
                'Proveeedor etiquetas bulto, ademas del nobre, tiene que traer tb direccion codposta...
                If Index = 148 Then
                    If txtNombre(Index).Text <> "" Then
                        Set miRsAux = New ADODB.Recordset
                        miRsAux.Open "Select domprove,pobprove,codpobla,proprove from sprove where codprove=" & txtCodigo(Index).Text, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                        'NO PUEDE SER EOF
                        For numParam = 0 To 3
                            txtBultos(numParam + 2).Text = DBLet(miRsAux.Fields(CInt(numParam)), "T")
                        Next
                        miRsAux.Close
                        Set miRsAux = Nothing
                    Else
                        LimpiarTextosBultos
                    End If
                End If
                
            Else
                txtNombre(Index).Text = ""
            End If
        Else
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, tabla, NomCampo, codCampo, Titulo, TipCampo)
        End If
        
        If Index = 133 Then PonerFoco txtCodigo(31)
                
    End If
End Sub


Private Sub MandaBusquedaPrevia(cadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim Cad As String
Dim tabla As String
Dim Titulo As String

    'Llamamos a al form
    Cad = ""
    Conexion = conAri    'Conexi�n a BD: Ariges
    Select Case OpcionListado
        Case 7 'Traspaso de Almacenes
            Cad = Cad & "N� Trasp|scatra|codtrasp|N|0000000|40�Almacen Origen|scatra|almaorig|N|000|20�Almacen Destino|scatra|almadest|N|000|20�Fecha|scatra|fechatra|F||20�"
            tabla = "scatra"
            Titulo = "Traspaso Almacenes"
        Case 8 'Movimientos de Almacen
            Cad = Cad & "N� Movim.|scamov|codmovim|N|0000000|40�Almacen|scamov|codalmac|N|000|30�Fecha|scamov|fecmovim|F||30�"
            tabla = "scamov"
            Titulo = "Movimientos Almacen"
        Case 9, 12, 13, 14, 15, 16, 17 '9: Movimientos Articulos
                   '12: Inventario Articulos
                   '14:Actualizar Diferencias de Stock Inventariado
                   '16: Listado Valoracion stock inventariado
            Cad = Cad & "C�digo|sartic|codartic|T||30�Denominacion|sartic|nomartic|T||70�"
            tabla = "sartic"
            Titulo = "Articulos"
            
            
        Case 60
            If vParamAplic.HayDeparNuevo = 1 Then
                Titulo = "Dptos Cliente: "

            ElseIf vParamAplic.HayDeparNuevo = 0 Then
                 Titulo = "Direc. Cliente: "
            Else
                Titulo = "Obra Cliente: "
            End If
            Titulo = Titulo & txtCodigo(37).Text & " - " & txtNombre(37)
            Cad = Cad & "Codigo|sdirec|coddirec|N|000|15�"
            Cad = Cad & "Descripcion|sdirec|nomdirec|T||55�"
            tabla = "sdirec"
    End Select
          
    If Cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = Cad
        frmB.vTabla = tabla
        frmB.vSQL = cadB
        HaDevueltoDatos = False
        '###A mano
        'frmB.vDevuelve = "0|1|"
        frmB.vDevuelve = "0|1|"
        frmB.vTitulo = Titulo
        frmB.vselElem = 0
        frmB.vConexionGrid = Conexion
'        frmB.vBuscaPrevia = 1
        '#
        frmB.Show vbModal
        Set frmB = Nothing
        'Si ha puesto valores y tenemos que es formulario de busqueda entonces
        'tendremos que cerrar el form lanzando el evento
'        If HaDevueltoDatos Then
''            If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
''                cmdRegresar_Click
'        Else   'de ha devuelto datos, es decir NO ha devuelto datos
'            PonerFoco Text1(kCampo)
'        End If
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub frmB_Selecionado(CadenaDevuelta As String)
'Formulario para Busqueda

    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        Select Case OpcionListado
            Case 7, 8 'Informe Traspasos Almacen
                txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaDevuelta, 1), "0000000")
                PonerFoco txtCodigo(indCodigo)
            Case 9, 12, 13, 14, 15, 16, 17, 60 '9: Informe Movimiento Articulos
                                'Inventario Articulos
                                '14: Actualizar diferencias Stock Inventariado
                                '16: Listado Valoracion stock inventariado
                txtCodigo(indCodigo).Text = RecuperaValor(CadenaDevuelta, 1)
                txtNombre(indCodigo).Text = RecuperaValor(CadenaDevuelta, 2)
                PonerFoco txtCodigo(indCodigo)
            
            
                
        End Select
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerFrameListadoVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para los listados de los mantenimientos de tabla: smarca, stipar,...

    H = 4695
    W = 6555
    PonerFrameVisible Me.FrameListado, visible, H, W

    If visible = True Then
        Me.Optcodigo.Value = True
    End If
End Sub



Private Sub PonerFrameInventarioVisible2(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Pone el Frame de Inventario Visible y Ajustado al Formulario, y visualiza los controles
'necesario para cada Informe de Inventario
Dim VerOpcion As Boolean
    chkValorDesdeArticulo.visible = False
    If visible = True Then
        H = 6400
        W = 7995
        VerOpcion = (OpcionListado = 15) Or (OpcionListado = 16) Or (OpcionListado = 17) Or (OpcionListado = 19)
        
        If VerOpcion Then
            If (OpcionListado = 19) Then
                H = 7100
                Me.cmdAceptar(4).Top = 6460
            Else
                H = 6900
                Me.cmdAceptar(4).Top = 6360
            End If
            
        ElseIf OpcionListado = 13 Then
            H = 6200
            Me.cmdAceptar(4).Top = 5200
        End If
        Me.cmdCancel(4).Top = Me.cmdAceptar(4).Top
        
            
        PonerFrameVisible Me.FrameInventario, visible, H, W

                
        '======================================
        'Valorar con Precios
        If VerOpcion Then
            Me.FrameValorar.visible = VerOpcion
            Me.FrameValorar.Left = 240
            If OpcionListado = 17 Then
                Me.FrameValorar.Top = 4500
            Else
                Me.FrameValorar.Top = 5000
            End If
            Me.chkSinStock.visible = VerOpcion
                 
            
        End If
        
        
        chkSaltaPag.Caption = "Salta p�g. en Familia"
        If OpcionListado = 13 Then
            If vParamAplic.InventarioxProv Then chkSaltaPag.Caption = "Salta p�g. en proveedor"
            Me.FrameValorar.Top = 4400
            FrameValorar.visible = True
        End If
        '====================================
        'Poner el Trabajador
        VerOpcion = (OpcionListado = 14)
        Me.Label4(7).visible = VerOpcion
        Me.imgBuscarG(17).visible = VerOpcion
        Me.txtCodigo(21).visible = VerOpcion
        Me.txtNombre(21).visible = VerOpcion
'        If VerOpcion Then txtCodigo(21).TabIndex = 47
        Label3(109).Caption = ""
        
        '======================================
        'Fecha Listados
        If OpcionListado = 15 Then '15: Listado Articulos Inactivos
            Me.Label4(5).Caption = "Fecha Inactividad"
        ElseIf OpcionListado = 19 Then
            Me.Label4(5).Caption = "Fecha Stock"
        Else
            Me.Label4(5).Caption = "Fecha Inventario"
        End If
        
        VerOpcion = (OpcionListado = 12) Or (OpcionListado = 15) Or (OpcionListado = 16) Or (OpcionListado = 19)
        Me.Label4(5).visible = VerOpcion  'campo fecha
        Me.imgFecha(2).visible = VerOpcion
        Me.txtCodigo(20).visible = VerOpcion
        Frame8.visible = False
        'campo HAsta Fecha
        Me.Label4(8).visible = (OpcionListado = 16)
        chkValorDesdeArticulo.visible = (OpcionListado = 16) 'Or (OpcionListado = 17)
        'Si opcionlistado=19 este campo sera la hora
        Me.Label4(9).visible = (OpcionListado = 16) Or (OpcionListado = 19)
        
        If OpcionListado = 19 Then
            Me.Label4(9).Caption = "Hora"
            Me.Label4(9).Left = 4250
            Me.txtCodigo(22).Left = 4700
            'Mostraremos el frame8
            Frame8.visible = True
            Frame8.BorderStyle = 0
            Me.cboStokFecha.ListIndex = 0
            Me.txtCodigo(136).Text = "0"
        End If
        Me.imgFecha(3).visible = (OpcionListado = 16)
        Me.txtCodigo(22).visible = (OpcionListado = 16) Or (OpcionListado = 19)
        If OpcionListado = 16 Then
            Me.Label4(8).Left = 2280
            Me.imgFecha(2).Left = 2820
            Me.txtCodigo(20).Left = 3120
            Me.Label4(9).Left = 4680
            Me.imgFecha(3).Left = 5160
            Me.txtCodigo(22).Left = 5430
'            txtCodigo(22).TabIndex = 48
        End If
        
        
        '====================================
        'Activar o no los check de Opcion:
        VerOpcion = (OpcionListado = 12) Or (OpcionListado = 13) Or (OpcionListado = 16) Or (OpcionListado = 17) Or (OpcionListado = 19) Or OpcionListado = 15
                    '12: Toma de Inventario
                    '13: Listado Diferencias stock
        
        Me.FrameOpciones2.visible = VerOpcion
        If OpcionListado = 12 Then
            Me.FrameOpciones2.Top = 4800
            Me.FrameOpciones2.BorderStyle = 0
        Else
            Me.FrameOpciones2.Top = FrameValorar.Top
        End If
        Me.FrameOpciones2.Height = 1575
        If OpcionListado = 13 Then
            'Me.FrameOpciones2.Top = 4400
            Me.FrameOpciones2.Height = 495
            'Me.FrameOpciones.Left = 4500
            Me.FrameOpciones2.BorderStyle = 0
        End If
        Frame8.Top = FrameOpciones2.Top

        Me.chkSaltaPag.visible = VerOpcion
        Me.chkValorado.visible = (OpcionListado = 16) Or (OpcionListado = 17)

        
        VerOpcion = (OpcionListado = 12)
        'If VerOpcion Then Me.FrameOpciones2.Left = 700
        Me.chkImprimeStock.visible = VerOpcion
        Me.chkImprimeStock.Top = 600
        If VerOpcion Then Me.txtCodigo(20).Text = Date
    End If
End Sub



Private Sub PonerFrameTarifasVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Pone el Frame de Tarifas Visible y Ajustado al Formulario, y visualiza los controles
'necesario para cada Informe de Tarifas
Dim VerOpcion As Boolean

    H = 6375
    H = 7335
    If OpcionListado = 245 Then H = 5075
    W = 7635
    PonerFrameVisible Me.FrameTarifas, visible, H, W
    
    chkVarios(3).visible = False
    chkVarios(3).Value = 0
    If visible = True Then
        '====================================
        '28: Tarifas Precios 29: Promociones
        VerOpcion = (OpcionListado = 28) Or (OpcionListado = 29)
        Me.chkSaltaPagTarif.visible = VerOpcion
        Me.Label4(12).visible = VerOpcion
        
        chkVarios(3).visible = OpcionListado = 28 And vParamAplic.NumeroInstalacion = 2
        
        '====================================
        If OpcionListado = 30 Then Me.Label4(11).Caption = "Cliente"
        
        chkSoloRotacion.visible = (OpcionListado = 28)
        '245: Control margenes tarifas
        '==================================
        VerOpcion = OpcionListado = 245 Or OpcionListado = 28
        Me.cboDecimales.visible = VerOpcion
        Label4(88).visible = VerOpcion
        If VerOpcion Then cboDecimales.ListIndex = 2
        VerOpcion = (OpcionListado = 245)
        Me.chkMostrarErrores.visible = VerOpcion
        'Decimales
        
        If VerOpcion Then
            Me.chkMostrarErrores.Top = 4600
            Label4(88).Top = 4300
            cboDecimales.Top = 4600
            
            'no mostrar seleccion de marca D/H
            Me.Label4(13).visible = Not VerOpcion
            Me.Label3(13).visible = Not VerOpcion
            Me.Label3(14).visible = Not VerOpcion
            Me.imgBuscarG(59).visible = Not VerOpcion
            Me.imgBuscarG(60).visible = Not VerOpcion
            Me.txtCodigo(27).visible = Not VerOpcion
            Me.txtCodigo(28).visible = Not VerOpcion
            Me.txtNombre(27).visible = Not VerOpcion
            Me.txtNombre(28).visible = Not VerOpcion
            'subir seleccion Articulo D/H al sitio de la marca
            Me.Label4(14).Top = Me.Label4(13).Top
            Me.Label3(15).Top = Me.Label3(13).Top
            Me.Label3(16).Top = Me.Label3(14).Top
            Me.imgBuscarG(61).Top = Me.imgBuscarG(59).Top
            Me.imgBuscarG(62).Top = Me.imgBuscarG(60).Top
            Me.txtCodigo(29).Top = Me.txtCodigo(27).Top
            Me.txtCodigo(30).Top = Me.txtCodigo(28).Top
            Me.txtNombre(29).Top = Me.txtNombre(27).Top
            Me.txtNombre(30).Top = Me.txtNombre(28).Top
            Me.cmdAceptarTarif.Top = 4600
            Me.cmdCancel(indFrame).Top = Me.cmdAceptarTarif.Top
        End If
    End If
End Sub


Private Sub PonerFrameRepxDiaVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para los listados de las Reparaciones x dia, de tabla: scarep
    

    If OpcionListado = 223 Or OpcionListado = 224 Then
        H = 4400
        W = 6100
    Else
        H = 3500
        W = 6000
    End If
    FrameCliRepDia.visible = False
    PonerFrameVisible Me.FrameRepxDia, visible, H, W
    
    If visible = True Then
        Me.Caption = "AriGes"
'        Me.FrameContab.Enabled = False
'        Me.OptClientes.Enabled = False
        Me.FrameContab.visible = (OpcionListado = 223 Or OpcionListado = 224 Or OpcionListado = 248)
        Me.FrameTipMov.visible = (OpcionListado = 223)
        Me.FrameProgress.visible = False
        
        '-- alto del boton aceptar y cancelar
        
        If OpcionListado = 223 Or OpcionListado = 224 Then
            Me.cmdAceptarRepxDia.Top = 3800
        Else
            Me.cmdAceptarRepxDia.Top = 2800
        End If
        Me.cmdCancel(7).Top = Me.cmdAceptarRepxDia.Top
        
        Select Case OpcionListado
            Case 63
                Me.lblTitulo(0).Caption = "Reparaciones por D�a"
                Me.Label2(2).Caption = "Fecha Reparaci�n:"
                Frame2.Top = 1350
                FrameCliRepDia.visible = True
            Case 73
                Me.lblTitulo(0).Caption = "Altas de Mantenimientos"
                Me.Label2(2).Caption = "Fecha Mantenimiento:"
                Frame2.Top = 1350
            Case 223, 224, 248 'Pedir datos para contabilizar facturas
                Me.lblTitulo(0).Caption = "Contabilizar Facturas"
                Me.Label2(2).Caption = "Fecha Factura:"
                Frame2.Top = 1680
                Me.FrameTipMov.Top = 2650
                
                
                Me.OptProve.Tag = ""
                If OpcionListado = 248 Then
                    Me.OptProve.Tag = "TIK"  'Son las de tickets agrupados
                    OpcionListado = 223
                End If
                If OpcionListado = 224 Then
                    Me.OptProve.Value = True
                    OpcionListado = 223
                End If
        End Select
    End If
End Sub


Private Sub PonerFrameManteVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para los listados de los Mantenimientos, de tabla: scaman
Dim b As Boolean
        'Opciones: 70,71,78,79,76
    H = 7695
    W = 6875
    PonerFrameVisible Me.FrameMantenimientos, visible, H, W

    If visible = True Then
        b = (OpcionListado = 70)
        
        Me.cboTipoList.visible = b 'List. Mantenimientos
        Me.Label1(4).visible = b
        
        
        
        'List Revisiones Mantenimientos
        Me.Frame3(1).visible = (OpcionListado = 70) Or (OpcionListado = 76)
        Me.Frame3(0).visible = (OpcionListado = 71)
        Me.Frame3(2).visible = (OpcionListado = 78)
        
        Select Case OpcionListado
        Case 70
                Me.Label7(0).Caption = "Informe de Mantenimientos"
        Case 71
                Me.Label7(0).Caption = "Informe Revisiones Mantenimientos"
               ' Me.Frame3.Top = 4800
                Me.txtCodigo(53).TabIndex = 211
                Me.txtCodigo(54).TabIndex = 212
                
        Case 76
                Me.Label7(0).Caption = "Inf.  Mantenimientos ANULADOS"
                
                
        Case 78
                'Cartas de renovacvion
                Me.Label7(0).Caption = "Cartas de renovacion"
                Me.txtCodigo(109).Text = Format(Now, "dd/mm/yyyy")
        Case 79
                'Etiquetas de mantenimientos
                Me.Label7(0).Caption = "Etiquetas de mantenimientos"
        End Select
    End If
End Sub


Private Sub ponerFrameArticulosVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el informe de Articulos, de tabla: sartic
Dim b As Boolean



    'Hay una opcion mas que mostrara este frame. la 247. Correccion de de tarifas e importes en articulos
    FrameTapaINCORRECTO.visible = False
    chkMinimoCorreg.visible = False
    b = (OpcionListado = 6)
    chkImpEtiq(0).visible = b
    chkImpEtiq(1).visible = b
    Me.imgayuda(0).visible = b
    If b Then
        Me.Label9.Caption = "Informe de Articulos"
       
        W = 8595
    Else
        If OpcionListado = 18 Then
            Me.Label9.Caption = "Informe Stocks Maximos y Minimos"
            Label4(36).Caption = "Almac�n"
        Else
            'NUEVA OCPION:  247
            'Corregir tarifas y eso
            chkMinimoCorreg.visible = True
            Me.Label9.Caption = "Verificaci�n tarifas y P.V.P."
            FrameTapaINCORRECTO.visible = True
            Label4(36).Caption = "Tarifa"
            cmbDecimales.ListIndex = 0
        End If
        W = 7395
       
    End If
    H = 6820
    
    
    PonerFrameVisible Me.FrameInfArticulos, visible, H, W
    If visible = True Then
        'visible orden campos si opcionlistado=6
        Me.FrameOrden.visible = b
        Label4(36).visible = Not b

        Me.imgBuscarG(18).visible = Not b
        Me.txtCodigo(72).visible = Not b
        Me.txtNombre(72).visible = Not b
        
        'Visible Frame stocks Max Minimos si opcionlistado=18
        Me.optStockMax.Value = True
        Me.FrameStockMaxMin.visible = OpcionListado = 18
  
        FrameSituacionArticulo.visible = OpcionListado = 6
    
    
        'REajustes.
        'El articulo NO se muestra si la opcion es 247
        b = OpcionListado <> 247
        PonerLabelsArticulosFrameVisible b
        Label4(75).visible = Not b
        cmbDecimales.visible = Not b
        Label4(90).visible = Not b
        cmbDecimales.visible = Not b
    
    End If
End Sub


Private Sub PonerLabelsArticulosFrameVisible(Si As Boolean)
    Label4(38).visible = Si
    Label3(51).visible = Si
    imgBuscarG(27).visible = Si
    txtCodigo(70).visible = Si
    txtNombre(70).visible = Si
    Label3(54).visible = Si
    imgBuscarG(28).visible = Si
    txtCodigo(71).visible = Si
    txtNombre(71).visible = Si
    chkMinimoCorreg.visible = Not Si
    
End Sub


Private Sub CargarListView()
'Carga el List View del frame: frameMovimArtic
'con los parametros de la tabla: stipom (Tipos de Movimientos)
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim ItmX As ListItem

    'Los encabezados
    ListView1.ColumnHeaders.Clear
    ListView1.ColumnHeaders.Add , , "C�digo", 800
    ListView1.ColumnHeaders.Add , , "Descripci�n", 2250
    
    SQL = "select * from stipom where muevesto=1"
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        Set ItmX = ListView1.ListItems.Add
        ItmX.Text = Rs.Fields(0).Value
        ItmX.Checked = True
        ItmX.SubItems(1) = Rs.Fields(1).Value
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
End Sub



Private Sub CargarListViewOrden()
'Carga el List View del frame: frameInfArticulos
'para establecer el orden en que se van a mostrar los datos en el Informe
'Orden: Familia, MArca, Proveedor, Tipo de Articulo, Articulo
Dim ItmX As ListItem

    'Los encabezados
    ListView2.ColumnHeaders.Clear
    ListView2.ColumnHeaders.Add , , "Campo", 1600
    Set ItmX = ListView2.ListItems.Add
    ItmX.Text = "Familia"
    Set ItmX = ListView2.ListItems.Add
    ItmX.Text = "Marca"
    Set ItmX = ListView2.ListItems.Add
    ItmX.Text = "Proveedor"
    Set ItmX = ListView2.ListItems.Add
    ItmX.Text = "Tipo Articulo"
End Sub


Private Function PonerFormulaYParametrosInf9() As Boolean
Dim Cad As String
Dim todosMarcados As Boolean
Dim devuelve As String
Dim I As Byte

    PonerFormulaYParametrosInf9 = False
    InicializarVbles
    
    'Parametro EMPRESA
    CadParam = "|pNomEmpre=""" & vParam.NombreEmpresa & """|"
    numParam = 1
        
    '-- Cadena para seleccion Desde y Hasta ARTICULO
    If txtCodigo(5).Text <> "" Or txtCodigo(6).Text <> "" Then
        codigo = "{smoval.codartic}"
        devuelve = "pDHArticulo=""Art�culo: "
        If Not PonerDesdeHasta(codigo, "T", 5, 6, devuelve) Then Exit Function
    End If
                    
    '-- Cadena para seleccion Desde y Hasta FAMILIA
    If txtCodigo(7).Text <> "" Or txtCodigo(8).Text <> "" Then
        codigo = "{sartic.codfamia}"
        devuelve = "pDHFamilia=""Familia: "
        If Not PonerDesdeHasta(codigo, "N", 7, 8, devuelve) Then Exit Function
    End If
        
    '-- Cadena para seleccion Desde y Hasta ALMACEN
    If txtCodigo(11).Text <> "" Or txtCodigo(12).Text <> "" Then
        codigo = "{smoval.codalmac}"
        devuelve = "pDHAlmacen=""Almacen: "
        If Not PonerDesdeHasta(codigo, "N", 11, 12, devuelve) Then Exit Function
    End If
    
    
    '-- Cadena para seleccion Desde y Hasta CLIENTE/PROVEEDOR
    If txtCodigo(86).Text <> "" Or txtCodigo(87).Text <> "" Then
        codigo = "{smoval.codigope}"
        devuelve = "pDHOperario=""Cliente/Proveedor/Trab.: "
        If Not PonerDesdeHasta(codigo, "N", 86, 87, devuelve) Then Exit Function
    End If
    
        
'    cadSelect = QuitarCaracterACadena(cadFormula, "{")
'    cadSelect = QuitarCaracterACadena(cadSelect, "}")
        
    '=================================================
    '-- Cadena para seleccion Desde y Hasta FECHA
    If txtCodigo(9).Text <> "" Or txtCodigo(10).Text <> "" Then
        codigo = "{smoval.fechamov}"
        devuelve = "pDHFecha=""Fecha: "
        If Not PonerDesdeHasta(codigo, "F", 9, 10, devuelve) Then Exit Function
    End If
        
    '-- seleccionar los articulos que tienen control de stock
    codigo = "{sartic.ctrstock}=1"
    AnyadirAFormula cadFormula, codigo
    AnyadirAFormula cadSelect, codigo
        
        
    '-- Cadena de Seleccion TIPOS de MOVIMIENTOS
    codigo = "{smoval.detamovi}"
    devuelve = ""
    'Si todos seleccionados no a�adir la select
    todosMarcados = True
    I = 1
    While Not I > Me.ListView1.ListItems.Count And todosMarcados
        If Not Me.ListView1.ListItems(I).Checked Then todosMarcados = False
        I = I + 1
    Wend
    
    'si no estan todos seleccionados montar select de los seleccionados
    If Not todosMarcados Then
        Cad = ""
        devuelve = ""
        For I = 1 To Me.ListView1.ListItems.Count
            If Me.ListView1.ListItems(I).Checked Then
                If Cad = "" Then
                    Cad = Me.ListView1.ListItems(I).Text
                Else
                    Cad = Cad & ", " & Me.ListView1.ListItems(I).Text
                End If
                If devuelve = "" Then
                    devuelve = codigo & " = """ & Me.ListView1.ListItems(I).Text & """"
                Else
                    devuelve = devuelve & " or " & codigo & " = """ & Me.ListView1.ListItems(I).Text & """"
                End If
            End If
        Next I

        If devuelve <> "" Then 'Hay algun movimiento marcado
            If cadFormula <> "" Then
                cadFormula = cadFormula & " AND " & "(" & devuelve & ")"
'                devuelve = QuitarCaracterACadena(devuelve, "{")
'                devuelve = QuitarCaracterACadena(devuelve, "}")
                cadSelect = cadSelect & " AND " & "(" & devuelve & ")"
                CadParam = CadParam
            Else
                cadFormula = "(" & devuelve & ")"
'                devuelve = QuitarCaracterACadena(devuelve, "{")
'                devuelve = QuitarCaracterACadena(devuelve, "}")
                cadSelect = "(" & devuelve & ")"
            End If
            Cad = "pTiposMov=""Tipos Movimiento: " & Cad
            CadParam = CadParam & Cad & """|"
            numParam = numParam + 1
        Else 'Todos desmarcados
            Cad = ""
            For I = 1 To ListView1.ListItems.Count
                If Cad = "" Then
                    Cad = """" & ListView1.ListItems(I).Text & """"
                Else
                    Cad = Cad & ", """ & ListView1.ListItems(I).Text & """"
                End If
            Next I
            devuelve = codigo & " NOT IN [" & Cad & "]"
            Cad = codigo & " NOT IN (" & Cad & ")"
            Cad = QuitarCaracterACadena(Cad, "{")
            Cad = QuitarCaracterACadena(Cad, "}")
            If cadFormula = "" Then
                cadFormula = "(" & devuelve & ")"
                cadSelect = "(" & Cad & ")"
            Else
                cadFormula = cadFormula & " AND " & "(" & devuelve & ")"
                cadSelect = cadSelect & " AND " & "(" & Cad & ")"
            End If
        End If
    End If
    
    
    If cadFormula = "" Then
        MsgBox "Introduzca alg�n criterio de selecci�n para el Informe.", vbInformation
        Exit Function
    End If
    PonerFormulaYParametrosInf9 = True
    
End Function


Private Function PonerFormulaYParametrosInf12() As Boolean
Dim Cad As String, cadFrom As String
Dim devuelve As String
Dim ImprStock As String
Dim CodAux As String
Dim strValorado As String
Dim strSinStock As String
Dim bytPrecio As Byte

'    InicializarVbles
    CadParam = "|pNomEmpre=""" & vParam.NombreEmpresa & """|"
    numParam = numParam + 1
    cadFrom = ""
    devuelve = ""
    PonerFormulaYParametrosInf12 = False
    
    '===================================================
    '================= FORMULA =========================
    
    Select Case OpcionListado
        Case 12, 15, 16, 17, 19
            CodAux = "{salmac."
            cadFrom = "  salmac "
'        Case 15 'Listado articulos inactivos
'            CodAux = "{salmac."
'            cadFrom = "  (salmac LEFT OUTER JOIN smoval ON salmac.codartic=smoval.codartic AND salmac.codalmac=smoval.codalmac) "
'            cadFrom = "salmac"
        Case 13, 14
            CodAux = "{sinven."
            cadFrom = " sinven "
    End Select
    
    'Cadena para seleccion De ALMACEN
    '-----------------------------------
    codigo = CodAux & "codalmac}"
    If Trim(txtCodigo(13).Text) <> "" Then _
    devuelve = codigo & " = " & Val(txtCodigo(13).Text)
    If devuelve <> "" Then
        cadFormula = devuelve
        Cad = "pAlmacen= ""Almacen: " & Format(txtCodigo(13).Text, "000") & " " & txtNombre(13).Text
        
        If OpcionListado = 19 Then
            'QUE SALGA LA MARCA DE VARIOS
            If Me.chkProv2(2).Value = 1 Then Cad = Cad & " (VARIOS)"
        End If
        
        CadParam = CadParam & Cad & """|"
        numParam = numParam + 1
    End If
            
    'Cadena para seleccion Desde y Hasta ARTICULOS
    '----------------------------------------------
    If txtCodigo(14).Text <> "" Or txtCodigo(15).Text <> "" Then
        codigo = CodAux & "codartic}"
        devuelve = "pDHArticulo=""Articulo: "
        If Not PonerDesdeHasta(codigo, "T", 14, 15, devuelve) Then Exit Function
    End If
    
    'Cadena para seleccion Desde y Hasta FAMILIA
    '--------------------------------------------
    If txtCodigo(16).Text <> "" Or txtCodigo(17).Text <> "" Then
        Select Case OpcionListado
            Case 12, 15, 16, 17, 19: codigo = "{sartic.codfamia}"
            Case Else: codigo = "{sinven.codfamia}"
        End Select
        Cad = "pDHFamilia=""Familia: "
        If Not PonerDesdeHasta(codigo, "N", 16, 17, Cad) Then Exit Function
    End If
    cadFrom = cadFrom & " INNER JOIN sartic ON " & CodAux & "codartic=sartic.codartic "
    'Enero 2008
    'David
    cadFormula = cadFormula & " AND {sartic.ctrstock} = 1"
    
    'Enero 2009
    'David
    'Solo saldran los articulos que esten en situacion normal o bloqueados.
    'Los caducados NO salen
    cadFormula = cadFormula & " AND {sartic.codstatu} < 3"
    
    'Enero 2012
    'David
    'Los de varios si ha puesto la marca de que salgan
    If OpcionListado = 19 Then
        If Me.chkProv2(2).Value = 0 Then
            cadFormula = cadFormula & " AND {sartic.artvario} = 0"
        End If
    End If
    'Cadena para seleccion Desde y Hasta PROVEEDOR
    '----------------------------------------------
    If txtCodigo(18).Text <> "" Or txtCodigo(19).Text <> "" Then
        Select Case OpcionListado
            Case 12, 15, 16, 17, 19: codigo = "{sartic.codprove}"
            Case Else: codigo = "{sinven.codprove}"
        End Select
        Cad = "pDHProveedor=""Proveedor: "
        If Not PonerDesdeHasta(codigo, "N", 18, 19, Cad) Then Exit Function
    End If
    

    
    'Select para MySQL
    cadSelect = QuitarCaracterACadena(cadFormula, "{")
    cadSelect = QuitarCaracterACadena(cadSelect, "}")
    cadSelect = QuitarCaracterACadena(cadSelect, "_1")
    cadFrom = QuitarCaracterACadena(cadFrom, "{")
    
    'Cadena para seleccion Desde y Hasta FECHA
    '----------------------------------------------
    If (OpcionListado = 16) Then
        If txtCodigo(20).Text <> "" Or txtCodigo(22).Text <> "" Then
            'codigo = "{salmac.codartic}"
            codigo = CodAux & "fechainv}"
            devuelve = CadenaDesdeHasta(txtCodigo(20).Text, txtCodigo(22).Text, codigo, "F")
    
            If devuelve = "Error" Then Exit Function
            
            If Not AnyadirAFormula(cadFormula, devuelve) Then
                Exit Function
            ElseIf devuelve <> "" Then
                Cad = "pDHFecha=""Fecha: "
                If txtCodigo(20).Text <> "" Then _
                    Cad = Cad & "desde " & txtCodigo(20).Text
                If txtCodigo(22).Text <> "" Then _
                    Cad = Cad & "  hasta " & txtCodigo(22).Text
                    
                'Mayo
                If Me.chkValorado.Value = 1 Then
                    Cad = Cad & "      Precio desde "
                    If chkValorDesdeArticulo.Value = 1 Then
                        Cad = Cad & "articulo"
                    Else
                        Cad = Cad & " inventario"
                    End If
                End If
                CadParam = CadParam & Trim(Cad) & """|"
                numParam = numParam + 1
                'Para Comprobar si hay registros a Mostrar antes de abrir el Informe
                devuelve = "salmac.fechainv"
                devuelve = CadenaDesdeHastaBD(txtCodigo(20).Text, txtCodigo(22).Text, devuelve, "F")
                AnyadirAFormula cadSelect, devuelve
            Else
                'Si no hay fecha de inventario seleccionada coger solo
                'los articulos de los que se haya hecho inventario alguna vez
                devuelve = "not isnull({salmac.fechainv})"
                If Not AnyadirAFormula(cadFormula, devuelve) Then
                    Exit Function
                End If
                devuelve = "not isnull(salmac.fechainv)"
                AnyadirAFormula cadSelect, devuelve
            End If
        End If
    End If
    
    'Cadena de seleccion de FECHA de Inactividad
    '------------------------------------------------
    If OpcionListado = 15 Then '15: Listado de Articulos Inactivos
         If txtCodigo(20).Text <> "" Then _
            Cad = "pFechaInve=""" & txtCodigo(20).Text & """"
        
        'Poner en el parametro pListaArt la lista de Articulos que no tiene
        'un registro de movimiento en la smoval con fecha posterior a la
        'fecha de inactividad
        strValorado = ListaArtActivos(cadSelect, txtCodigo(20).Text)
        Cad = "pListaArtic=""" & strValorado & """|"
        CadParam = CadParam & Cad
        numParam = numParam + 1
        
        'A�adir a la formula de seleccion que no sea uno de la lista
        devuelve = " not (" & CodAux & "codartic} in {@pListaArtic})"
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Function
        
        strValorado = QuitarCaracterACadena(strValorado, "[")
        strValorado = QuitarCaracterACadena(strValorado, "]")
        devuelve = " not (salmac.codartic in (" & strValorado & "))"
        AnyadirAFormula cadSelect, devuelve
    End If
    
    'Cadena de seleccion de FECHA de Stocks a una Fecha
    '--------------------------------------------------
     If OpcionListado = 19 Then
        If txtCodigo(20).Text <> "" Then
            Cad = txtCodigo(20).Text
            'Hora
            If txtCodigo(22).Text <> "" Then _
                Cad = Cad & "  " & txtCodigo(22).Text
                
            CadParam = CadParam & "pFechaStock=""" & Cad & """|"
            numParam = numParam + 1
            
                            
            'Si lleva factor conversion y solo negativos o positivos
            Cad = ""
            devuelve = "0"
            If Me.cboStokFecha.ListIndex > 0 Then Cad = "S�lo " & Me.cboStokFecha.Text
            If txtCodigo(136).Text <> "" Then
                If Val(txtCodigo(136).Text) <> 0 Then
                    Cad = Cad & "       Inc." & txtCodigo(136).Text
                    devuelve = TransformaComasPuntos(ImporteFormateado(txtCodigo(136).Text))
                End If
            End If
            'Stocks a una fecha
            If OpcionListado = 19 Then
                If Me.chkProv2(2).Value = 1 Then
                    'Van a salir VARIOS tb
                    If Cad <> "" Then Cad = Cad & "         "
                    Cad = Cad & "ART. VARIOS"
                Else
                    'Los VARIOS no
                    cadFormula = cadFormula & " AND {sartic.artvario} =0"
                End If
            End If
            
            CadParam = CadParam & "pdhValora=""" & Cad & """|"
            numParam = numParam + 1
                
            'Incremento
            CadParam = CadParam & "Incremento=" & devuelve & "|"
            numParam = numParam + 1
            
            
            'Detalla
            CadParam = CadParam & "detalla=" & Abs(Me.chkProv2(1).Value) & "|"
            numParam = numParam + 1
            
        End If
     End If
     
    'Cadena para Seleccion de Articulos con Stock<>0
    '------------------------------------------------
    If OpcionListado = 16 Or OpcionListado = 17 Or OpcionListado = 15 Then
        If Me.chkSinStock.Value = 0 Then
            If OpcionListado = 16 Then
                devuelve = "{salmac.stockinv}<>0"
            Else
                devuelve = CodAux & "canstock}<>0"
            End If
            If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Function
            
            devuelve = QuitarCaracterACadena(devuelve, "{")
            devuelve = QuitarCaracterACadena(devuelve, "}")
            devuelve = QuitarCaracterACadena(devuelve, "_1")
            AnyadirAFormula cadSelect, devuelve
        End If
    ElseIf OpcionListado = 19 Then
         If Me.chkSinStock.Value Then
            ImprStock = "True"
        Else
            ImprStock = "False"
        End If
        CadParam = CadParam & "pSinStock=" & ImprStock & "|"
        numParam = numParam + 1
    End If
    
       
    '============================================
    '============= PARAMETROS ===================
    If OpcionListado = 12 Or OpcionListado = 15 Then
        '12: Toma de Inventario
        '15: Listado Articulos Inactivos
        CadParam = CadParam & "pFechaInve=""" & txtCodigo(20).Text & """|"
        numParam = numParam + 1
    End If
    
    If OpcionListado = 12 Then
        'Par�metro Imprime Stock (Si/No)
        If Me.chkImprimeStock.Value Then
            ImprStock = "True"
        Else
            ImprStock = "False"
        End If
        CadParam = CadParam & "pImprimeStock=" & ImprStock & "|"
        numParam = numParam + 1
        
'        'seleccionar para inventariar los articulos que no tienen control stock
'        devuelve = " {sartic.ctrstock} = 1 "
'        AnyadirAFormula cadFormula, devuelve
'        AnyadirAFormula cadSelect, devuelve
        'Laura 03/01/07
        If Not (InStr(cadFrom, "sartic") > 0) Then
            cadFrom = cadFrom & " INNER JOIN sartic ON " & CodAux & "codartic=sartic.codartic "
        End If
    End If
    
    If OpcionListado = 12 Or OpcionListado = 13 Or OpcionListado = 15 Or OpcionListado = 16 Or OpcionListado = 17 Or OpcionListado = 19 Then
        'Par�metro Salta Pag. en Familia (Si/No)
        If Me.chkSaltaPag.Value Then
            ImprStock = "True"
        Else
            ImprStock = "False"
        End If
        CadParam = CadParam & "pSaltaFamilia=" & ImprStock & "|"
        numParam = numParam + 1
    End If
    
    If OpcionListado = 16 Or OpcionListado = 17 Then '16: Valoraci�n de Stocks Inventariados
                                                     '17: Valoraci�n Stocks
        'Par�metro Valorado
        'Febrero 2012
        '---------------------------------------
        '  Si es un agente conectado, no dejaremos que vea valoracion
        If vParamAplic.AlmacenB > 1 Then
            'HErbelca
            If vUsu.CodigoAgente > 0 Then
                'Quito la marca de valorado
                If Me.chkValorado.Value = 1 Then chkValorado.Value = 0
            End If
        End If
        
        
        If Me.chkValorado.Value = 1 Then
            strValorado = "True"
        Else
            strValorado = "False"
        End If
        CadParam = CadParam & "pValorado=" & strValorado & "|"
        numParam = numParam + 1
        
        
        'Mayo 2013
        
        CadParam = CadParam & "pDesdeArticulo=" & Abs(Me.chkValorDesdeArticulo.Value) & "|"
        numParam = numParam + 1
        
    End If
    
    If (OpcionListado = 15) Or (OpcionListado = 16) Or (OpcionListado = 17) Or (OpcionListado = 19) Then
        'Parametro Precio de Valoracion, elegir un Precio para realizar la valoracion: canstock * precio
        If Me.optPrecioMP.Value Then bytPrecio = 1
        If Me.optPrecioMA.Value Then bytPrecio = 2
        If Me.optPrecioUC.Value Then bytPrecio = 3
        If Me.optPrecioStd.Value Then bytPrecio = 4
        CadParam = CadParam & "pPrecio=" & bytPrecio & "|"
        numParam = numParam + 1
    End If
    '=====================================================================
    
       
    'comprobar si hay registros para mostrar en el Informe antes de Abrirlo
    If Not HayRegParaInforme(cadFrom, cadSelect) Then Exit Function
    
    If OpcionListado = 19 Then
        cadSelect = "Select count(*) FROM " & cadFrom & " WHERE " & cadSelect
        cadSelect = Replace(cadSelect, "count(*)", "*")
        DescargarDatosTMPStockFecha
        If Not CargarTMPStockFecha(cadSelect, txtCodigo(20).Text, txtCodigo(22).Text, CByte(Me.cboStokFecha.ListIndex), Label3(109)) Then Exit Function
    End If
    
    If OpcionListado = 13 Then
        'Este listado va valorado
        strValorado = "preciouc"
        If Me.optPrecioMP.Value Then strValorado = "preciomp"
        If Me.optPrecioMA.Value Then strValorado = "precioma"
        If Me.optPrecioStd.Value Then strValorado = "preciost"
        CadParam = CadParam & "kprecio= """ & strValorado & """|"
        numParam = numParam + 1
        strValorado = ""
    End If
    
    
    
    'ENERO 2013
    'Antes de pasar a inventariar  ver sia hay datos inventariandose
'    If OpcionListado = 12 Then
'
'
'        devuelve = "salmac  INNER JOIN sartic ON salmac.codartic=sartic.codartic  "
'        CodAux = "sartic.ctrstock = 1 AND sartic.codstatu < 2  AND ( sartic.codartic,codalmac) IN (select codartic,codalmac from sinven) AND  salmac.codalmac"
'        CodAux = DevuelveDesdeBD(conAri, "count(*)", devuelve, CodAux, txtCodigo(13).Text)
'        If CodAux = "" Then CodAux = "0"
'
'        If Val(CodAux) > 0 Then MsgBox "Existen articulos(" & CodAux & ")    YA inventariandose", vbExclamation
'
'
'
'    End If
    

    PonerFormulaYParametrosInf12 = True
End Function



Private Function PonerFormulaYParametrosInf28() As Boolean
'Informes de Descuentos y Tarifas
Dim Cad As String
Dim cadCodigo As String
Dim Aux As String

    CadParam = "|pNomEmpre=""" & vParam.NombreEmpresa & """|"
    numParam = 1
    
    PonerFormulaYParametrosInf28 = False
    
    '===================================================
    '================= FORMULA =========================
    
    'Cadena para seleccion Desde y Hasta TARIFA o D/H CLIENTE
    '--------------------------------------------
     If txtCodigo(23).Text <> "" Or txtCodigo(24).Text <> "" Then
        If OpcionListado = 30 Then 'Precios Especiales Cliente
            cadCodigo = codigo & ".codclien}"
            Cad = "pDHCliente=""Cliente: "
        Else
            cadCodigo = codigo & ".codlista}"
            Cad = "pDHTarifa=""Tarifa: "
        End If
        If Not PonerDesdeHasta(cadCodigo, "N", 23, 24, Cad) Then Exit Function
    End If
            
            
    'Cadena para seleccion Desde y Hasta FAMILIA
    '--------------------------------------------
    Aux = ""
    If txtCodigo(25).Text <> "" Or txtCodigo(26).Text <> "" Then
        cadCodigo = "{sartic.codfamia}"
        Cad = "Familia: "
        If Not PonerDesdeHasta(cadCodigo, "N", 25, 26, Cad) Then Exit Function
        Aux = Cad
    End If
    If OpcionListado = 28 Then
        If Me.chkVarios(3).Value = 1 Then
            'CABEL
            Set miRsAux = New ADODB.Recordset
            cadCodigo = ""
            miRsAux.Open "Select codfamia from sfamia where marcapropia=1", conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not miRsAux.EOF
                cadCodigo = cadCodigo & ", " & miRsAux!Codfamia
                miRsAux.MoveNext
            Wend
            miRsAux.Close
            Set miRsAux = Nothing
            
            If cadCodigo <> "" Then cadCodigo = Mid(cadCodigo, 2)
            cadCodigo = " {sartic.codfamia} IN [" & cadCodigo & "]"
                        
            If cadFormula <> "" Then cadCodigo = " AND " & cadCodigo
            
            'rpt
            cadFormula = cadFormula & cadCodigo
            'sql
            cadCodigo = Replace(cadCodigo, "[", "(")
            cadCodigo = Replace(cadCodigo, "]", ")")
            cadSelect = cadSelect & cadCodigo
            
            Aux = Trim(Aux & "   [** CABEL **]")
        End If
    End If
    If Aux <> "" Then
        Cad = "pDHFamilia=""" & Aux & " |"
        CadParam = CadParam & Cad
        numParam = numParam + 1
        Aux = ""
    End If
    
    If OpcionListado <> 245 Then
    
        
    
        'Cadena para seleccion Desde y Hasta MARCA
        '--------------------------------------------
        Aux = ""
        If txtCodigo(27).Text <> "" Or txtCodigo(28).Text <> "" Then
            cadCodigo = "{sartic.codmarca}"
            Cad = "Marca: "
            If Not PonerDesdeHasta(cadCodigo, "N", 27, 28, Cad) Then Exit Function
            Aux = Cad
        End If
        
        If txtCodigo(134).Text <> "" Or txtCodigo(135).Text <> "" Then
            cadCodigo = "{sartic.codprove}"
            Cad = "   Prov : "
            If Not PonerDesdeHasta(cadCodigo, "N", 134, 135, Cad) Then Exit Function
            Aux = Aux & Cad
        End If
        
        
        'ROTACION, para el listado 28
        If OpcionListado = 28 Then
            'No entrara en el de abajo
            If Me.chkSoloRotacion.Value = 1 Then
                If cadFormula <> "" Then
                    cadFormula = cadFormula & " AND "
                    cadSelect = cadSelect & " AND "
                End If
                cadFormula = cadFormula & "{sartic.rotacion}=1"
                cadSelect = cadSelect & " sartic.rotacion=1 "
                
                'Texto para el report
                If Len(Aux) > 75 Then
                    'LO pongo al principio
                    Aux = "[ROT] " & Aux
                Else
                    'Lo pongo al final
                    Aux = Aux & " [ROTACION]"
                End If
            End If
            
        End If
        
        
        
        If Aux <> "" Then
            Cad = "pDHMarca=""" & Trim(Aux) & """|"
            CadParam = CadParam & Cad
            numParam = numParam + 1
        End If
    End If
            
            
            
            
            
            
            
    'Cadena para seleccion Desde y Hasta ARTICULOS
    '----------------------------------------------
    If txtCodigo(29).Text <> "" Or txtCodigo(30).Text <> "" Then
        cadCodigo = codigo & ".codartic}"
        Cad = "pDHArticulo=""Articulo: "
        If Not PonerDesdeHasta(cadCodigo, "T", 29, 30, Cad) Then Exit Function
    End If
 
 
 
 
 
 
 
    '=====================================================================
    '====   PARAMETROS
    If (OpcionListado = 28 Or OpcionListado = 29) Then
        'Par�metro Salta Pag. en Familia (Si/No)
        If Me.chkSaltaPagTarif.Value = 1 Then
            Cad = "True"
        Else
           Cad = "False"
        End If
        CadParam = CadParam & "pSaltaFamilia=" & Cad & "|"
        numParam = numParam + 1
    End If
       
    If OpcionListado = 245 Then
        'Par�metro mostrar solo tarifas con errores (Si/No)
        Cad = Abs(Val(Me.chkMostrarErrores.Value))
        CadParam = CadParam & "Suprimr=" & Cad & "|"
        numParam = numParam + 1
        'Decimales
    End If
    
    If OpcionListado = 245 Or OpcionListado = 28 Then
        If cboDecimales.ListIndex < 0 Then
            MsgBox "Seleccione decimales", vbExclamation
            Exit Function
        End If
        Cad = (cboDecimales.ItemData(Me.cboDecimales.ListIndex))
        CadParam = CadParam & "Decimales=" & Cad & "|"
        numParam = numParam + 1
    End If
       
    PonerFormulaYParametrosInf28 = True
End Function


Private Function AnyadirParametroDH(Cad As String, indD As Byte, indH As Byte) As String
On Error Resume Next
    
     If txtCodigo(indD).Text <> "" Then
        Cad = Cad & "desde " & txtCodigo(indD).Text
        If txtNombre(indD).Text <> "" Then Cad = Cad & " - " & txtNombre(indD).Text
    End If
    If txtCodigo(indH).Text <> "" Then
        Cad = Cad & "  hasta " & txtCodigo(indH).Text
        If txtNombre(indH).Text <> "" Then Cad = Cad & " - " & txtNombre(indH).Text
    End If
    
    AnyadirParametroDH = Cad
    If Err.Number <> 0 Then Err.Clear
End Function


Private Function InsertarInventario() As Boolean
'Inserta en la Tabla:sinven los articulos seleccionados para realizar Inventario
'Inserta en la Tabla Hist.: shinve los datos que habia de inventario
'Adem�s Actualiza la Tabla:salmac los campos:fechainv, horainve, statusin
Dim SQL As String, ADonde As String
Dim Rs As ADODB.Recordset
Dim hora As Date

On Error GoTo EInventario:
   
'   If CrearTmpInventario(cadSelect) Then
   

        'Aqui empieza transaccion
        conn.BeginTrans
    
          
    
'        'Insertar en la tabla de Hist�rico: shinve
'        'Pasamos al Hist. los datos que habia antes de hacer nuevo inventario
'        ADonde = "Insertando datos en Hist�rico. Tabla: shinve"
'        SQL = "INSERT INTO shinve (codartic,codalmac,fechainv,horainve,existenc) "
'        SQL = SQL & " SELECT salmac.codartic, salmac.codalmac, salmac.fechainv,salmac.horainve,salmac.stockinv "
'        SQL = SQL & " FROM salmac INNER JOIN sartic ON salmac.codartic=sartic.codartic "
'        SQL = SQL & " WHERE " & cadFormula
'        'si no se ha inventariado antes no lo pasamos al historico
'        SQL = SQL & " AND not isnull(salmac.fechainv) "
'        Conn.Execute SQL
'
        
        'Insertar en la tabla de Hist�rico: shinve
        'Pasamos al Hist. los datos que habia antes de hacer nuevo inventario
        ADonde = "Insertando datos en Hist�rico. Tabla: shinve"
        SQL = "INSERT INTO shinve (codartic,codalmac,fechainv,horainve,existenc) "
        SQL = SQL & " SELECT codartic,codalmac,fechainv,horainve,stockinv "
'        SQL = SQL & " FROM salmac INNER JOIN sartic ON salmac.codartic=sartic.codartic "
        SQL = SQL & " FROM tmpInven "
'        SQL = SQL & " WHERE " & cadFormula
        'si no se ha inventariado antes no lo pasamos al historico
        SQL = SQL & " WHERE not isnull(fechainv) "
        '--- Laura 03/01/2006
        SQL = SQL & " AND fechainv<>'0000-00-00' AND date(horainve)<>'0000-00-00' "
        '---
        conn.Execute SQL
        
        
        
        
        
        hora = Format(txtCodigo(20).Text & " " & Time, "yyyy-mm-dd hh:mm:ss")
        
'        'Insertamos en la Tabla sinven
'        ADonde = "Insertando datos en Inventario Real. Tabla: sinven"
'        SQL = "INSERT INTO sinven (codartic, codalmac, codfamia, codprove, fechainv, horainve, existenc) "
'        SQL = SQL & "SELECT salmac.codartic, salmac.codalmac, sartic.codfamia, sartic.codprove," & DBSet(txtCodigo(20).Text, "F") & " as fechainv," & DBSet(hora, "FH") & " as horainve, 0 as existenc "
'        SQL = SQL & " FROM salmac INNER JOIN sartic ON salmac.codartic=sartic.codartic "
'        SQL = SQL & " WHERE " & cadFormula
'        'Insertamos los articulos que tiene control de stock
'        SQL = SQL & " AND sartic.ctrstock=1"
'        Conn.Execute SQL

        'Insertamos en la Tabla sinven
        ADonde = "Insertando datos en Inventario Real. Tabla: sinven"
        SQL = "INSERT INTO sinven (codartic, codalmac, codfamia, codprove, fechainv, horainve, existenc) "
        SQL = SQL & "SELECT codartic, codalmac, codfamia, codprove," & DBSet(txtCodigo(20).Text, "F") & " as fechainv,"
        SQL = SQL & "" & DBSet(hora, "FH") & " as horainve, stockinv  as existenc "
'        SQL = SQL & " FROM salmac INNER JOIN sartic ON salmac.codartic=sartic.codartic "
        SQL = SQL & " FROM tmpInven "
'        SQL = SQL & " WHERE " & cadFormula
        'Insertamos los articulos que tiene control de stock
'        SQL = SQL & " AND sartic.ctrstock=1"
        conn.Execute SQL


        
        
'        SQL = "SELECT salmac.codartic, salmac.codalmac, sartic.codfamia, sartic.codprove "
'        SQL = SQL & "FROM salmac INNER JOIN sartic ON salmac.codartic=sartic.codartic "
'        SQL = SQL & " WHERE " & cadFormula
        
        SQL = "SELECT tmpInven.codartic, tmpInven.codalmac, tmpInven.codfamia, tmpInven.codprove "
        SQL = SQL & ",preciomp,precioma,preciouc,preciost"
        SQL = SQL & " FROM tmpInven left join sartic ON tmpInven.codartic=sartic.codartic "
    
        Set Rs = New ADODB.Recordset
        Rs.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        While Not Rs.EOF
            
          
            
            
            
            
            'Actualizamos la tabla salmac ponemos statusin=1 para indicar que se
            'esta realizando inventario y bloquear los articulos para que no se puedan
            'realizar movimientos, traspasos, etc.
            'Actualizamos la Tabla: salmac los campos: fechainv, horainve
            ADonde = "Actualizando datos en Articulos x Almacen"
            SQL = "UPDATE salmac SET fechainv='" & Format(txtCodigo(20).Text, "yyyy-mm-dd") & "', "
            SQL = SQL & " horainve='" & Format(hora, "yyyy-mm-dd hh:mm:ss") & "', " & "statusin=1 "
            'MAYO 2013 preciomp,precioma,preciouc,preciost
            SQL = SQL & ", preciompin = " & DBSet(Rs!PrecioMP, "N")
            SQL = SQL & ", preciomain = " & DBSet(Rs!precioma, "N")
            SQL = SQL & ", precioucin = " & DBSet(Rs!precioUC, "N")
            SQL = SQL & ", preciostin = " & DBSet(Rs!preciost, "N")
            
            'SEPTIEMBRE 2013
            'Incializar stock al inventariar
            SQL = SQL & ", stockinv="
            If vParamAplic.IncializarStockEnInventario Then
                SQL = SQL & " 0"  'stockinv=0 inicializamos el stock del articulo
            Else
                SQL = SQL & " if(canstock>0,canstock,0)"
            End If
            SQL = SQL & " WHERE codartic=" & DBSet(Rs.Fields(0).Value, "T") & " AND "
            SQL = SQL & "codalmac=" & Rs.Fields(1).Value
            conn.Execute SQL
            Rs.MoveNext
        Wend
    
        Rs.Close
        Set Rs = Nothing
'    Else
'        Exit Function
'    End If
    
EInventario:
    If Err.Number <> 0 Then
         'Hay error , almacenamos y salimos
          SQL = "Insertando Datos de Inventario." & vbCrLf & "--------------------------------------" & vbCrLf
          SQL = SQL & ADonde
          MuestraError Err.Number, SQL, Err.Description
        conn.RollbackTrans
        InsertarInventario = False
    Else
        InsertarInventario = True
        conn.CommitTrans
    End If
End Function


Private Function CrearTmpInventario(cadFormula As String) As Boolean
Dim SQL As String
Dim b As Boolean

    On Error GoTo ECrearInv
    
    b = False
    SQL = "CREATE TEMPORARY TABLE tmpInven ( "
    SQL = SQL & "codartic varchar(16) NOT NULL default '0', "
    SQL = SQL & "codalmac smallint(3) unsigned NOT NULL default '0', "
    SQL = SQL & "codfamia smallint(4) unsigned NOT NULL default '0', "
    SQL = SQL & "codprove int(6) unsigned NOT NULL default '0', "
    SQL = SQL & "fechainv date NOT NULL default '0000-00-00', "
    SQL = SQL & "horainve datetime NOT NULL default '0000-00-00 00:00:00', "
    SQL = SQL & "stockinv decimal(12,2) NOT NULL default '0.00')"
    conn.Execute SQL
    b = True
    
    
    'Seleccionar todos los registros que vamos a inventariar, insertarlos en la TMP
    'y trabajar con estos valores
    SQL = "SELECT salmac.codartic, salmac.codalmac, sartic.codfamia, sartic.codprove,salmac.fechainv,salmac.horainve,"
    'ANTES Estaba este salmac.stockinv  "
    If vParamAplic.IncializarStockEnInventario Then
        '0
        SQL = SQL & "0"
    Else
        SQL = SQL & "if(canstock>0,canstock,0)"
    End If
    SQL = SQL & " FROM salmac INNER JOIN sartic ON salmac.codartic=sartic.codartic "
    SQL = SQL & " WHERE " & cadFormula
    SQL = SQL & " AND sartic.ctrstock=1"

    SQL = " INSERT INTO tmpInven " & SQL
    conn.Execute SQL
    
    
    
ECrearInv:
    If Err.Number <> 0 Then
        SQL = " DROP TABLE IF EXISTS tmpInven;"
        conn.Execute SQL
        b = False
        'Err.Clear
        MuestraError Err.Number, "Crear temporal inventario.", Err.Description
    End If
    CrearTmpInventario = b
End Function






Private Function ActualizarInventario() As Boolean
'-----------------------------------------------------------------
'* Modifica en la Tabla: salmac los campos: cansotck, fechainv, horainve,statusin de los articulos seleccionados
'y les asigna los valores de los campos: existenc, fechainv, horainve, false de la tabla: sinven
'* Elimina de la Tabla: sinven los registros seleccinados para actualizar
'* Inserta Movimientos de Articulos en la Tabla: smoval
'-------------------------------------------------------------------
Dim SQL As String, ADonde As String
Dim Rs As ADODB.Recordset
Dim DevStock As String
Dim CanStock As Long, Diferencia As Long
Dim vTipoMov As CTiposMov
'Dim CodTipoMov As String * 3
Dim NumMovim As Long, numlinea As Long
Dim LetraSerie As String * 1
Dim CadValues As String
Dim bol As Boolean
    
    On Error Resume Next
    
    'Obtener Registros de la Tabla:sinven de los que se va a actualizar el Stock
    SQL = "SELECT sinven.* "
    
    'DAVID ENERO 2008
    'SQL = SQL & " FROM sinven "
    SQL = SQL & " FROM sinven  INNER JOIN sartic ON sinven.codartic=sartic.codartic"
    
    SQL = SQL & " WHERE " & cadFormula
    

    bol = True
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Rs.EOF Then
        bol = False
        ActualizarInventario = False
        MsgBox "No existen Registros en la Tabla: sinven para Actualizar Inventario.", vbExclamation
        Rs.Close
        Set Rs = Nothing
        Exit Function
    End If
    
    
    'Obtener el contador para los movimientos del Almacen que se esta inventariando
    'A cada registro de la tabla sinven se le asignar� un numero de linea.
    '----------------------------------------------------------------------------
    Set vTipoMov = New CTiposMov
'    CodTipoMov = "REG"
    If vTipoMov.Leer("DFI") Then  'Se han cargado correctamente los valores de la clase
        'Obtener el contador para el codigo de Movimiento
        LetraSerie = vTipoMov.LetraSerie
        NumMovim = vTipoMov.ConseguirContador("DFI")
        numlinea = 1
        bol = True
    Else
        bol = False
    End If
    
    If Not bol Then
        Set vTipoMov = Nothing
        Exit Function
    End If
    
   
    On Error GoTo EActualizarInven:
    'Aqui empieza la transaccion
    conn.BeginTrans
    
    While Not Rs.EOF And bol 'Para cada registro de la tabla sinven
    
        'Introducir Movimiento de Entrada/Salida si hay diferencia entre el
        'Stock del Sistema y el Stock Real Inventariado.
        '------------------------------------------------------------------
        ADonde = "Introduciendo Movimiento de Entrada/Salida. Tabla: smoval."
        DevStock = DevuelveDesdeBDNew(conAri, "salmac", "canstock", "codartic", Rs!codArtic, "T", , "codalmac", Rs!codAlmac, "N")
        If DevStock <> "" Then
            CanStock = CLng(DevStock)
            Diferencia = Rs!existenc - CanStock
            If Diferencia <> 0 Then 'Insertar Movimiento de Entrada/Salida en Almacen
                CadValues = DBSet(Rs!codArtic, "T") & ", " & Rs!codAlmac & ", '" & Format(Rs!FechaINV, "yyyy-mm-dd") & "', '"
                CadValues = CadValues & Format(Rs!horainve, "yyyy-mm-dd hh:mm:ss") & "', "
                bol = InsertarMovimArticulos(CadValues, Rs!codArtic, Diferencia, LetraSerie, NumMovim, numlinea)
                numlinea = numlinea + 1
            Else
                bol = True
            End If
        Else
            bol = False
        End If
        
        'Actualizamos la Tabla: salmac
        '           salmac.canstock := existencia Real
        '           salmac.statusin := false (desbloqueamos los articulos )
        '---------------------------------------
        If bol Then
            ADonde = "Actualizando Stock de Articulos en Almacen. Tabla: salmac."
            SQL = "UPDATE salmac SET canstock=" & DBSet(Rs!existenc, "N") & ", statusin=0"
            SQL = SQL & " WHERE codartic=" & DBSet(Rs!codArtic, "T") & " AND codalmac=" & Rs!codAlmac
            conn.Execute SQL
        End If

        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    
    
    If bol Then
'        'Pasamos la tabla de inventario real sinven al historico: shinve
'        'antes de eliminarla
'        ADonde = "Pasando registros de Inventario al Hist�rico: shinve."
'        SQL = "INSERT INTO shinve (codartic,codalmac,fechainv,horainve,existenc) "
'        SQL = SQL & "SELECT codartic,codalmac,fechainv,horainve,existenc "
'        SQL = SQL & " FROM sinven WHERE " & cadFormula
'        Conn.Execute SQL
    
        'Eliminamos los registros seleccionados de la Tabla: sinven
        '----------------------------------------------------------
        ADonde = "Eliminando registros de Inventario. Tabla: sinven."
       ' SQL = "DELETE FROM sinven "
  
        'DAVID ENERO 2008
        SQL = "DELETE sinven.* FROM sinven  INNER JOIN sartic ON"
        SQL = SQL & " sinven.codartic=sartic.codartic WHERE " & cadFormula
        conn.Execute SQL
        
        
        'Incrementamos el contador para el Tipo De Movimiento
        '-----------------------------------------------------
        ADonde = "Actualizando el contador ."
        'bol = vTipoMov.IncrementarContador(
        vTipoMov.IncrementarContador ("DFI")
    End If
    Set vTipoMov = Nothing
        
EActualizarInven:
    If Err.Number <> 0 Or Not bol Then
         'Hay error , almacenamos y salimos
          SQL = "Actualizar Inventario." & vbCrLf & "----------------------------" & vbCrLf
          SQL = SQL & ADonde
          MuestraError Err.Number, SQL, Err.Description
          conn.RollbackTrans
          ActualizarInventario = False
          Set vTipoMov = Nothing
    Else
        ActualizarInventario = True
        conn.CommitTrans
    End If
End Function


Private Function InsertarMovimArticulos(CadValues As String, codArtic As String, cantidad As Long, LetraSerie As String, NumMovim As Long, numlinea As Long) As Boolean
Dim vImporte As Single, vPrecioVenta As String
Dim tipoMov As Byte
Dim SQL As String
On Error Resume Next
         
        'Obtener el precio de venta del articulo
         vPrecioVenta = DevuelveDesdeBDNew(conAri, "sartic", "preciomp", "codartic", codArtic, "T")
        If vPrecioVenta <> "" Then
            vImporte = cantidad * CSng(vPrecioVenta)
        Else
            vImporte = 0
        End If
        
        'Tipo de Movimiento de Almacen
        If cantidad > 0 Then 'Insertar Movimiento de Entrada en Almacen
            tipoMov = 1
        ElseIf cantidad < 0 Then 'Insertar Movimiento de Salida en Almacen
            tipoMov = 0
        End If
        
        SQL = "INSERT INTO smoval (codartic, codalmac, fechamov, horamovi, tipomovi, detamovi, cantidad, impormov, codigope, letraser, document, numlinea) "
        SQL = SQL & " VALUES (" & CadValues & tipoMov & ", '" & "DFI" & "', " & DBSet(cantidad, "N") & ", " & DBSet(vImporte, "N") & ", " & Val(txtCodigo(21).Text) & ", '"
        SQL = SQL & LetraSerie & "', " & NumMovim & ", " & numlinea & ")"
        conn.Execute SQL
        
        If Err.Number <> 0 Then
             'Hay error , almacenamos y salimos
            InsertarMovimArticulos = False
        Else
            InsertarMovimArticulos = True
        End If
    
End Function


Private Function ValidarCamposInventario() As Boolean
'Comprobar que los campos requeridos tienen valor antes de abrir el listado
Dim b As Boolean

        b = True
        '- campo almacen debe tener valor
        If Trim(txtCodigo(13).Text) = "" Then
             MsgBox "El campo Almacen debe tener valor.", vbInformation
             PonerFoco txtCodigo(13)
             b = False
        End If
    
        '- fecha de inventario debe tener valor
        If b Then
            If (OpcionListado = 12 Or OpcionListado = 15 Or OpcionListado = 19) And Trim(txtCodigo(20).Text) = "" Then
                 MsgBox "El campo Fecha debe tener valor.", vbInformation
                 PonerFoco txtCodigo(20)
                 b = False
            End If
        End If
        
        'informe 19: stocks a una fecha
        'la fecha tiene que ser < a fecha hoy
        If OpcionListado = 19 And txtCodigo(20).Text <> "" Then
            If Not CDate(txtCodigo(20).Text) < CDate(Format(Now, "dd/mm/yyyy")) Then
                MsgBox "La fecha stock tiene que ser anterior a la fecha de hoy.", vbInformation
                PonerFoco txtCodigo(20)
                b = False
            End If
        End If
        If OpcionListado = 19 And txtCodigo(103).Text <> "" Then
            If Not IsNumeric(txtCodigo(103).Text) Then
                MsgBox "Campo incremento incorrecto", vbExclamation
                txtCodigo(103).Text = "2"
                PonerFoco txtCodigo(103)
                b = False
            End If
        End If
        If b Then
            If OpcionListado = 16 Then
                If Me.chkValorado.Value = 0 And Me.chkValorDesdeArticulo.Value = 1 Then
                    MsgBox "No esta marcada la opcion de valorar. NO mostrar� valoraci�n alguna", vbExclamation
                End If
            End If
        End If
        ValidarCamposInventario = b
End Function



Private Function ListaArtActivos(cadWhere As String, FechaIn As String) As String
Dim Rs As ADODB.Recordset
Dim SQL As String
Dim Lista As String
'Devuelve una cadena con la concatenacion de todos los articulos que
'no debe seleccionar ya que si tienen movimientos con fecha posterior
'a FechaIn.
'ej:    "[""00000004"", ""00000033""]"

    Lista = "["
    
    SQL = "SELECT distinct smoval.codartic from smoval "
    If InStr(cadWhere, "sartic") > 0 Then SQL = SQL & " INNER JOIN sartic ON smoval.codartic=sartic.codartic "
    SQL = SQL & " WHERE " & Replace(cadWhere, "salmac", "smoval")
    If cadWhere <> "" Then SQL = SQL & " AND "
    SQL = SQL & " fechamov>='" & Format(FechaIn, FormatoFecha) & "' "
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not Rs.EOF
'        lista = lista & """" & RS.Fields(0).Value & """"
        Lista = Lista & DBSet(Rs.Fields(0).Value, "T")
        Rs.MoveNext
        If Not Rs.EOF Then Lista = Lista & ", "
    Wend
    Lista = Lista & "]"
    ListaArtActivos = Lista
    Rs.Close
    Set Rs = Nothing
End Function



Private Sub ActualizarImprimir()
Dim I As Long
Dim Desde As Long, Hasta As Long
Dim SQL As String

    Select Case OpcionListado
    Case 7  'TRASPASO ALMACEN
        If frmVisReport.EstaImpreso = True Then
        'Desde=-1 si estamos en Historico ya que aqui no se introducen valores Desde/Hasta
            If Trim(txtCodigo(3).Text) <> "" Then Desde = CLng(txtCodigo(3).Text)
            If Trim(txtCodigo(4).Text) <> "" Then Hasta = CLng(txtCodigo(4).Text)
            For I = Desde To Hasta
                SQL = "UPDATE scatra SET situacio=1" 'Impreso
                SQL = SQL & " WHERE codtrasp=" & I
                conn.Execute SQL
            Next I
        End If
        
    Case 8  'MOVIMIENTO ALMACEN
        If frmVisReport.EstaImpreso = True Then
           'Desde=-1 si estamos en Historico ya que aqui no se introducen valores Desde/Hasta
           If Trim(txtCodigo(3).Text) <> "" Then Desde = CLng(txtCodigo(3).Text)
           If Trim(txtCodigo(4).Text) <> "" Then Hasta = CLng(txtCodigo(4).Text)
           For I = Desde To Hasta
                SQL = "UPDATE scamov SET situacio=1"
                SQL = SQL & " WHERE codmovim=" & I
                conn.Execute SQL
           Next I
        End If
    End Select
End Sub


Private Sub CargarComboTipoList()
'### Combo Tipo Facturaci�n
'Cargaremos el combo, o bien desde una tabla o con valores fijos o como
'se quiera, la cuestion es cargarlo
' El estilo del combo debe de ser 2 - Dropdown List
' Si queremos que este ordenado, o lo ordenamos por la sentencia sql
' o marcamos la opcion sorted del combo
'1-Equipos, 2-Pagos, 3-Importes Contrato

    Me.cboTipoList.Clear
    cboTipoList.AddItem "Equipos"
    cboTipoList.ItemData(cboTipoList.NewIndex) = 1

    cboTipoList.AddItem "Pagos"
    cboTipoList.ItemData(cboTipoList.NewIndex) = 2

    cboTipoList.AddItem "Importes Contrato"
    cboTipoList.ItemData(cboTipoList.NewIndex) = 3

End Sub




Private Sub CargarComboSituacion()
'### Combo Tipo Facturaci�n
'Cargaremos el combo, o bien desde una tabla o con valores fijos o como
'se quiera, la cuestion es cargarlo
' El estilo del combo debe de ser 2 - Dropdown List
' Si queremos que este ordenado, o lo ordenamos por la sentencia sql
' o marcamos la opcion sorted del combo
'0-Abierta, 1-En Reparacion, 2-Pendiente, 3-Cerrado

    Me.cboSituaAviso.Clear
    
    cboSituaAviso.AddItem "-- Todas --"
    cboSituaAviso.ItemData(cboSituaAviso.NewIndex) = 0
    
    cboSituaAviso.AddItem "Abierta"
    cboSituaAviso.ItemData(cboSituaAviso.NewIndex) = 1

    cboSituaAviso.AddItem "En reparaci�n"
    cboSituaAviso.ItemData(cboSituaAviso.NewIndex) = 2
    
    cboSituaAviso.AddItem "Pendiente"
    cboSituaAviso.ItemData(cboSituaAviso.NewIndex) = 3
    
    cboSituaAviso.AddItem "Cerrado"
    cboSituaAviso.ItemData(cboSituaAviso.NewIndex) = 4

End Sub


Private Sub InicializarVbles()
    cadFormula = ""
    cadSelect = ""
    CadParam = ""
    numParam = 0
'    cadTitulo = ""
'    cadNomRPT = ""
    conSubRPT = False
    pPdfRpt = ""
    pRptvMultiInforme = 0
End Sub


Private Function PonerDesdeHasta(campo As String, Tipo As String, indD As Byte, indH As Byte, param As String) As Boolean
Dim devuelve As String
Dim Cad As String

    PonerDesdeHasta = False
    devuelve = CadenaDesdeHasta(txtCodigo(indD).Text, txtCodigo(indH).Text, campo, Tipo)
    If devuelve = "Error" Then Exit Function
    If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Function
    
    If Tipo <> "F" Then
        'Fecha para Crystal Report
        If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Function
    Else
        'Fecha para la Base de Datos
        Cad = CadenaDesdeHastaBD(txtCodigo(indD).Text, txtCodigo(indH).Text, campo, Tipo)
        If Not AnyadirAFormula(cadSelect, Cad) Then Exit Function
    End If
    
    If devuelve <> "" Then
        If param <> "" Then
            'Parametro Desde/Hasta
            CadParam = CadParam & AnyadirParametroDH(param, indD, indH) & """|"
            numParam = numParam + 1
        End If
        PonerDesdeHasta = True
    End If
End Function


Private Sub LlamarImprimir(PonerNombrePDF As Boolean)
    With frmImprimir
        .FormulaSeleccion = cadFormula
        .OtrosParametros = CadParam
        .NumeroParametros = numParam

        .SoloImprimir = False
        .EnvioEMail = False
        .Opcion = OpcionListado
        .Titulo = cadTitulo
        .NombreRPT = cadNomRPT
        .NombrePDF = ""
        .SeleccionaRPTCodigo = pRptvMultiInforme
        If PonerNombrePDF Then .NombrePDF = pPdfRpt
        .ConSubInforme = conSubRPT
        .Show vbModal
    End With
End Sub


Private Function PonerGrupo(numGrupo As Byte, cadgrupo As String) As Byte
Dim campo As String
Dim NomCampo As String

    campo = "pGroup" & numGrupo & "="
    NomCampo = "pGroup" & numGrupo & "Name="
    PonerGrupo = 0
    
    Select Case cadgrupo
        Case "Familia"
            CadParam = CadParam & campo & "{sartic.codfamia}" & "|"
            If numGrupo = 1 Then
                CadParam = CadParam & NomCampo & " ""FAMILIA: "" & " & " totext({sartic.codfamia},""0000"") & " & """  """ & " & {sfamia.nomfamia}" & "|"
            Else
                CadParam = CadParam & NomCampo & " totext({sartic.codfamia},""0000"") & " & """ """ & " & {sfamia.nomfamia}" & "|"
            End If
            numParam = numParam + 1
        Case "Marca"
            CadParam = CadParam & campo & "{sartic.codmarca}" & "|"
            If numGrupo = 1 Then
                CadParam = CadParam & NomCampo & " ""MARCA: "" & " & " totext({sartic.codmarca},""0000"") & " & """  """ & " & {smarca.nommarca}" & "|"
            Else
                CadParam = CadParam & NomCampo & " totext({sartic.codmarca},""0000"") & " & """ """ & " & {smarca.nommarca}" & "|"
            End If
            numParam = numParam + 1
        Case "Proveedor"
            CadParam = CadParam & campo & "{sartic.codprove}" & "|"
            If numGrupo = 1 Then
                CadParam = CadParam & NomCampo & " ""PROVEEDOR: "" & " & " totext({sartic.codprove},""000000"") & " & """  """ & " & {sprove.nomprove}" & "|"
            Else
                CadParam = CadParam & NomCampo & " totext({sartic.codprove},""000000"") & " & """ """ & " & {sprove.nomprove}" & "|"
            End If
            numParam = numParam + 1
            PonerGrupo = numGrupo
        Case "Tipo Articulo"
            CadParam = CadParam & campo & "{sartic.codtipar}" & "|"
            If numGrupo = 1 Then
                CadParam = CadParam & NomCampo & " ""TIPO ARTICULO: "" & " & " {sartic.codtipar} & " & """  """ & " & {stipar.nomtipar}" & "|"
            Else
                CadParam = CadParam & NomCampo & " {sartic.codtipar} & " & """ """ & " & {stipar.nomtipar}" & "|"
            End If
            numParam = numParam + 1
    End Select

'Case "Familia"
'            cadParam = cadParam & "pGroup1=" & "{sartic.codfamia}" & "|"
'            cadParam = cadParam & "pGroup1Name= ""FAMILIA: "" & " & " totext({sartic.codfamia},""0000"") & " & """  """ & " & {sfamia.nomfamia}" & "|"
'            numParam = numParam + 1
'            Select Case ListView2.ListItems(2).Text
'                Case "Marca"
'                    cadParam = cadParam & "pGroup2=" & "{sartic.codmarca}" & "|"
'                    cadParam = cadParam & "pGroup2Name= ""MARCA: "" & " & " totext({sartic.codmarca},""0000"") & " & """  """ & " & {smarca.nommarca}" & "|"
'                    numParam = numParam + 1
'                    If ListView2.ListItems(3).Text = "Proveedor" Then
'                        Opcion = 1
'                    Else
'                        Opcion = 2
'                    End If
'                Case "Proveedor"
'                Case "Tipo Articulo"
'            End Select
End Function



Private Sub AbrirFrmActividades(Optional indice As Integer)
    Set frmMtoActiv = New frmFacActividades
    frmMtoActiv.DatosADevolverBusqueda = "0|1|"
    frmMtoActiv.DeConsulta = True
    frmMtoActiv.Show vbModal
    Set frmMtoActiv = Nothing
End Sub



Private Sub AbrirFrmMarcas()
    Set frmMtoMarcas = New frmAlmMarcas
    frmMtoMarcas.DatosADevolverBusqueda = "0|1"
    frmMtoMarcas.DeConsulta = True
    frmMtoMarcas.Show vbModal
    Set frmMtoMarcas = Nothing
End Sub


Private Sub AbrirFrmAlmPropios()
    Set frmMtoAlPropios = New frmAlmAlPropios
    frmMtoAlPropios.DatosADevolverBusqueda = "0|1"
    frmMtoAlPropios.DeConsulta = True
    frmMtoAlPropios.Show vbModal
    Set frmMtoAlPropios = Nothing
End Sub


Private Sub AbrirFrmZonas()
    Set frmMtoZonas = New frmFacZonas
    frmMtoZonas.DatosADevolverBusqueda = "0|1"
    frmMtoZonas.DeConsulta = True
    frmMtoZonas.Show vbModal
    Set frmMtoZonas = Nothing
End Sub


Private Sub AbrirFrmRutas()
    Set frmMtoRutas = New frmFacRutas
    frmMtoRutas.DatosADevolverBusqueda = "0|1"
    frmMtoRutas.DeConsulta = True
    frmMtoRutas.Show vbModal
    Set frmMtoRutas = Nothing
End Sub


Private Sub AbrirFrmTarifas()
'tarifas venta
    Set frmMtoTarifas = New frmFacTarifas
    frmMtoTarifas.DatosADevolverBusqueda = "0|1"
    frmMtoTarifas.Show vbModal
    Set frmMtoTarifas = Nothing
End Sub


Private Sub AbrirFrmTipoArt()
'Tipos de Articulos
    Set frmMtoTArticulo = New frmAlmTipoArticulo
    frmMtoTArticulo.DatosADevolverBusqueda = "0|1"
    frmMtoTArticulo.DeConsulta = True
    frmMtoTArticulo.Show vbModal
    Set frmMtoTArticulo = Nothing
End Sub

Private Sub AbrirFrmClientes()
'Clientes
    Set frmMtoClientes = New frmFacClientes3
    frmMtoClientes.DatosADevolverBusqueda = "0|1"
    frmMtoClientes.Show vbModal
    Set frmMtoClientes = Nothing
End Sub


Private Function ComprobarFechasConta(Ind As Integer) As Boolean
'comprobar que el periodo de fechas a contabilizar esta dentro del
'periodo de fechas del ejercicio de la contabilidad
Dim FechaIni As String, FechaFin As String
Dim Cad As String
Dim Rs As ADODB.Recordset
    
On Error GoTo EComprobar

    ComprobarFechasConta = False
    
    If txtCodigo(Ind).Text <> "" Then
        FechaIni = "Select fechaini,fechafin From parametros"
        Set Rs = New ADODB.Recordset
        Rs.Open FechaIni, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
    
        If Not Rs.EOF Then
            FechaIni = DBLet(Rs!FechaIni, "F")
            '## LAURA 19/06/2008
'            FechaFin = DBLet(RS!FechaFin, "F") + 365
'            FechaFin = DateAdd("d", 365, DBLet(RS!FechaFin, "F"))
            FechaFin = DateAdd("yyyy", 1, DBLet(Rs!FechaFin, "F"))
            '##
            
            'nos guardamos los valores
            Orden1 = FechaIni
            Orden2 = FechaFin
        
            If Not EntreFechas(FechaIni, txtCodigo(Ind).Text, FechaFin) Then
                 Cad = "El per�odo de contabilizaci�n debe estar dentro del ejercicio:" & vbCrLf & vbCrLf
                 Cad = Cad & "    Desde: " & FechaIni & vbCrLf
                 Cad = Cad & "    Hasta: " & FechaFin
                 MsgBox Cad, vbExclamation
                 txtCodigo(Ind).Text = ""
            Else
                ComprobarFechasConta = True
            End If
        End If
        Rs.Close
        Set Rs = Nothing
    Else
        ComprobarFechasConta = True
    End If
    
EComprobar:
    If Err.Number <> 0 Then MuestraError Err.Number, "Comprobar Fechas", Err.Description
End Function



Private Function ContabilizarFacturas(cadTabla As String, cadWhere As String, ByRef PGB As ProgressBar, ByRef LblPg0 As Label, LblPg1 As Label, DesdeGenerarFraProveedor As Boolean) As Boolean
'Contabiliza Facturas de Clientes o de Proveedores
Dim SQL As String
Dim b As Boolean
Dim tmpErrores As Boolean 'Indica si se creo correctamente la tabla de errores
Dim CCoste2 As Byte
        '0.- Si devuelve la funcion el 0 habra CC sin confgurar en trabaja
        '1.- Todos los CC son el mismo
        '2.- Mas de un CC. Hay que agrupar
Dim AuxD As String

    ContabilizarFacturas = False

    If cadTabla = "scafac" Then
        SQL = "VENCON" 'contabilizar facturas de venta
    ElseIf cadTabla = "scafpc" Then
        SQL = "COMCON" 'contabilizar facturas de compra
    End If

    'Bloquear para que nadie mas pueda contabilizar
    DesBloqueoManual (SQL)
    If Not BloqueoManual(SQL, "1") Then
        MsgBox "No se pueden Contabilizar Facturas. Hay otro usuario contabilizando.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If


     'comprobar que se han rellenado los dos campos de fecha
     'sino rellenar con fechaini o fechafin del ejercicio
     'que guardamos en vbles Orden1,Orden2
     If txtCodigo(31).Text = "" Then
        txtCodigo(31).Text = vEmpresa.FechaIni  'fechaini del ejercicio de la conta
     End If

     If txtCodigo(32).Text = "" Then
        txtCodigo(32).Text = vEmpresa.FechaFin  'fecha fin del ejercicio de la conta
     End If


     
     'Comprobar que el intervalo de fechas D/H esta dentro del ejercicio de la
     'contabilidad par ello mirar en la BD de la Conta los par�metros
     If Not ComprobarFechasConta(32) Then Exit Function
     
     

    'La comprobacion solo lo hago para facturas nuestras, ya que mas adelante
    'el programa hara cdate(text1(31) cuando contabilice las facturas y dara error de tipos
    If cadTabla = "scafac" Then
        If Me.txtCodigo(31).Text = "" Then
            MsgBox "Fecha inicio incorrecta", vbExclamation
            Exit Function
        End If
    End If
    
    
    
    'comprobar si existen en Ariges facturas anteriores al periodo solicitado
    'sin contabilizar.
    If Me.txtCodigo(31).Text <> "" Then 'anteriores a fechadesde
        SQL = "SELECT COUNT(*) FROM " & cadTabla
        If cadTabla = "scafac" Then
            SQL = SQL & " WHERE fecfactu <"
        ElseIf cadTabla = "scafpc" Then
            SQL = SQL & " WHERE fecrecep <"
        End If
        SQL = SQL & DBSet(txtCodigo(31), "F") & " AND intconta=0 "
        
        
        'Si contabiliza tickets agrupados
        'SOLO PARA CLIENTES, obviamente
        If cadTabla = "scafac" Then
            If OptProve.Tag = "" Then
                If vParamAplic.ContabilizarTicketAgrupados Then SQL = SQL & " AND codtipom <>'FTI' "
            Else
                SQL = SQL & " AND scafac.codtipom  = 'FTG' "
            End If
            
            '## LAURA 20/06/2008
            If Trim(Me.cboTipMov.List(Me.cboTipMov.ListIndex)) <> "" Then
                SQL = SQL & " AND scafac.codtipom = " & DBSet(Mid(Me.cboTipMov.List(Me.cboTipMov.ListIndex), 1, 3), "T")
            Else
                
                If Val(vUsu.AlmacenPorDefecto) <> vParamAplic.AlmacenB Then SQL = SQL & " AND scafac.codtipom <> 'FAZ'"
                   
            End If
        End If
        
        If RegistrosAListar(SQL) > 0 Then
            If MsgBox("Hay Facturas anteriores sin contabilizar. " & "�Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2) <> vbYes Then
                Exit Function
            End If
        End If
    End If
    
    
'    'Laura: 11/10/2006 bloquear los registros q vamos a contabilizar
'    If Not BloqueaRegistro(cadTabla, cadWhere) Then
'        MsgBox "No se pueden Contabilizar Facturas. Hay registros bloqueados.", vbExclamation
'        Screen.MousePointer = vbDefault
'        Exit Sub
'    End If
    
    
    '==========================================================
    'REALIZAR COMPROBACIONES ANTES DE CONTABILIZAR FACTURAS
    '==========================================================
    
'    Me.lblProgess(0).Caption = "Comprobaciones: "
'    CargarProgres Me.ProgressBar1, 100
        
    'Cargar tabla TEMP con las Facturas que vamos a Trabajar
    b = CrearTMPFacturas(cadTabla, cadWhere)
    If Not b Then Exit Function
            
            
    'Laura: 11/10/2006 bloquear los registros q vamos a contabilizar
'    TerminaBloquear
    SQL = cadTabla & " INNER JOIN tmpFactu ON " & cadTabla
    If cadTabla = "scafac" Then
        SQL = SQL & ".codtipom=tmpFactu.codtipom AND "
    Else
        SQL = SQL & ".codprove=tmpFactu.codprove AND "
    End If
    SQL = SQL & cadTabla & ".numfactu=tmpFactu.numfactu AND " & cadTabla & ".fecfactu=tmpFactu.fecfactu "
    If Not BloqueaRegistro(SQL, cadWhere) Then
        MsgBox "No se pueden Contabilizar Facturas. Hay registros bloqueados.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If
            
            
    '---- Preparamos la pantalla de Contabilizar
    'Visualizar la barra de Progreso
    
    If Not DesdeGenerarFraProveedor Then
        
        If Me.FrameTipMov.visible Then
            Me.FrameRepxDia.Height = 6100
            Me.FrameProgress.Top = 4400
        Else
            Me.FrameRepxDia.Height = 5100
            Me.FrameProgress.Top = 3350
        End If
        Me.Height = Me.FrameRepxDia.Height
        Me.FrameProgress.visible = True
    End If
    
    Me.Refresh
            
    LblPg0.Caption = "Comprobaciones: "
    CargarProgres PGB, 100
        
        
    'comprobar que todas las LETRAS SERIE existen en la contabilidad y en Ariges
    '--------------------------------------------------------------------------
    IncrementarProgres PGB, 10
    If cadTabla = "scafac" Then
        LblPg1.Caption = "Comprobando letras de serie ..."
        LblPg1.Refresh
        b = ComprobarLetraSerie(cadTabla)
    End If
    IncrementarProgres PGB, 10
    Me.Refresh
    If Not b Then Exit Function
    
    
    'comprobar que no haya N� FACTURAS en la contabilidad para esa fecha
    'que ya existan
    '-----------------------------------------------------------------------
    If cadTabla = "scafac" Then
        LblPg1.Caption = "Comprobando N� Facturas en contabilidad ..."
        LblPg1.Refresh
        SQL = "anofaccl>=" & Year(txtCodigo(31).Text) & " AND anofaccl<= " & Year(txtCodigo(32).Text)
        b = ComprobarNumFacturas_new(cadTabla, SQL)
    End If
    IncrementarProgres PGB, 20
    Me.Refresh
    If Not b Then Exit Function
    
    
    'comprobar que todas las CUENTAS de los distintos clientes que vamos a
    'contabilizar existen en la Conta: sclien.codmacta IN (conta.cuentas.codmacta)
    '-----------------------------------------------------------------------------
    LblPg1.Caption = "Comprobando Cuentas Contables en contabilidad ..."
    LblPg1.Refresh
    b = ComprobarCtaContable_new(cadTabla, 1)
    IncrementarProgres PGB, 20
    Me.Refresh
    If Not b Then Exit Function
    
    
    'comprobar que todas las CUENTAS de venta de la familia de los articulos que vamos a
    'contabilizar existen en la Conta: sfamia.ctaventa IN (conta.cuentas.codmacta)
    '-----------------------------------------------------------------------------
    If cadTabla = "scafac" Then
        LblPg1.Caption = "Comprobando Cuentas Ctbles Ventas en contabilidad ..."
    Else
        LblPg1.Caption = "Comprobando Cuentas Ctbles Compras en contabilidad ..."
    End If
    LblPg1.Refresh
    b = ComprobarCtaContable_new(cadTabla, 2)
    IncrementarProgres PGB, 20
    Me.Refresh
    If Not b Then Exit Function
    
    
    
    If Me.OptProve.Tag <> "" Then
        'TIKETS. Voy a comprobar las cuentas de la familia
        LblPg1.Caption = "Comprobando Cuentas Ctbles tickets ..."
        LblPg1.Refresh
        
        
    End If
    
    
    'comprobar que todos las TIPO IVA de las distintas fecturas que vamos a
    'contabilizar existen en la Conta: scafac.codigiv1,codigiv2,codigiv3 IN (conta.tiposiva.codigiva)
    '--------------------------------------------------------------------------
    LblPg1.Caption = "Comprobando Tipos de IVA en contabilidad ..."
    LblPg1.Refresh
    b = ComprobarTiposIVA(cadTabla)
    IncrementarProgres PGB, 10
    Me.Refresh
    If Not b Then Exit Function
    
    'comprobar si hay contabilidad ANALITICA: conta.parametros.autocoste=1
    'y verificar que las cuentas de sfamia.ctaventa empiezan por el digito
    'de conta.parametros.grupogto o conta.parametros.grupovta
    'obtener el centro de coste del usuario para insertarlo en linfact
    
    If vEmpresa.TieneAnalitica Then
       LblPg1.Caption = "Comprobando Contabilidad Anal�tica ..."
       LblPg1.Refresh
       b = ComprobarCtaContable_new(cadTabla, 3)
       If Not b Then Exit Function
       
       '(si tiene anal�tica requiere un centro de coste para insertar en conta.linfact)
       b = cadTabla = "scafac"
       If b And OptProve.Tag <> "" Then
        'NUEVO
        'CONTABUILZIACION AGRUPADA DE TIKETS
        
            CCoste2 = ComprobarCCosteTikAgrupado(cadWhere)
       Else
            CCoste2 = ComprobarCCoste(cadWhere, b)
       End If
       If CCoste2 = 0 Then Exit Function 'Error comprobando CCs
       
    Else
        'No tiene analitica, NO agrupamos por codtraba
        CCoste2 = 0
    End If
    IncrementarProgres PGB, 10
    Me.Refresh
    
    If Me.OptProve.Tag <> "" Then
        LblPg1.Caption = "Comprobando Ctas facmilias TICKETS ..."   'FTG
        b = ComprobarCtaContable_new(cadTabla, 4)
        If Not b Then Exit Function
    End If
    
    
    'Comprobamos, si es factura proveedore, que si el tipoprove = 3 (REA)
    'entonces tiene que existir el paremetro aplicacion codret
    If cadTabla = "scafpc" Then
        If vParamAplic.CtaReten = "" Then
            SQL = "SELECT COUNT(*) FROM scafpc,sprove WHERE scafpc.codprove = sprove.codprove and tipprove=3"
            If cadWhere <> "" Then SQL = SQL & " AND " & cadWhere
            If RegistrosAListar(SQL) > 0 Then
                MsgBox "Existen facturas SOCIOS proveedor con cta. retencion y no esta configurada", vbExclamation
                Exit Function
            End If
        
        
            'Neuvo 29Mayo 2008
            ' Cualquier factura puede llevar retencion. Necesito que la cuenta de retencion este configurada
            SQL = "SELECT COUNT(*) FROM scafpc  WHERE  tiporet=0 and impret<>0"
            If cadWhere <> "" Then SQL = SQL & " AND " & cadWhere
            If RegistrosAListar(SQL) > 0 Then
                MsgBox "Existen facturas proveedor con retencion y no esta configurada", vbExclamation
                Exit Function
            End If
         End If
        
       
        'Veremos si va a contabilizar fras proveedor INTRACOM. Si es asin
        SQL = "SELECT COUNT(*) FROM scafpc,sprove WHERE scafpc.codprove = sprove.codprove and tipprove=1"
        
        'ABRIL 2016
        'YA no hace la ceracion de la AUTOFACTURA
        If False Then
            If cadWhere <> "" Then SQL = SQL & " AND " & cadWhere
            If RegistrosAListar(SQL) > 0 Then
                'Hay facturas intracomunitarias.  Veremos si esta bien configuarado el programa
                If vParamAplic.CtaContabIntracom <> "" Then
                    'NO ha puesto IVA intracom
                    If vParamAplic.IvaIntracomAdicional = 0 Then
                        MsgBox "No esta en parametros el IVA intracom para las facturas adicionales", vbExclamation
                        Exit Function
                    End If
                    
                    'Ha puesto cta contable para las dos facturas "extras". Veamos si existe
                    SQL = DevuelveDesdeBDNew(conConta, "cuentas", "codmacta", "codmacta", vParamAplic.CtaContabIntracom, "N")
                    If SQL = "" Then
                        MsgBox "No existe la cuenta contable para las facturas intracom", vbExclamation
                        Exit Function
                    End If
                    
                    SQL = "" 'todo ok
                    
                End If
            End If
            
        End If
    End If
    
    LblPg1.Caption = "Fechas contabilizacion"
    LblPg1.Refresh
    b = NuevasComprobacionesContabilizacion(cadTabla = "scafpc", cadWhere)
    If Not b Then Exit Function
    
    
    
    
    '===========================================================================
    'CONTABILIZAR FACTURAS
    '===========================================================================
    LblPg0.Caption = "Contabilizar Facturas: "
    CargarProgres PGB, 10
    LblPg1.Caption = "Insertando Facturas en Contabilidad..."
    Me.Refresh
    
    
    
    
    '------------------------------------------------------------------------------
    '  LOG de acciones
    Set LOG = New cLOG
    LOG.Insertar 3, vUsu, "Contabilizar facturas: " & vbCrLf & cadTabla & vbCrLf & cadWhere
    Set LOG = Nothing
    '-----------------------------------------------------------------------------
    
    
    
    '---- Crear tabla TEMP para los posible errores de facturas
    tmpErrores = CrearTMPErrFact(cadTabla)
    
    
    
    'Modificacion 11 Abril 2011
    '-----------------------------------------------------------
    ' Mosrtaremos para cada factura de PROVEEDOR
    ' que numregis le ha asignado
    SQL = "DELETE FROM tmpinformes WHERE codusu = " & vUsu.codigo
    conn.Execute SQL
    
    
    '---- Pasar las Facturas a la Contabilidad
    b = PasarFacturasAContab(cadTabla, CCoste2)
    
    
    
    '---- Mostrar ListView de posibles errores (si hay)
    If Not b Then
        If tmpErrores Then
            'Cargar un listview con la tabla TEMP de Errores y mostrar
            'las facturas que fallaron
            frmMensajes.OpcionMensaje = 10
            frmMensajes.Show vbModal
        Else
            MsgBox "No pueden mostrarse los errores.", vbInformation
        End If
    Else
        'Para la facturacion de TICKTS agrupada NO mostramos el mensaje de OK
        If Me.OptProve.Tag = "" Then
            If cadTabla = "scafac" Then MsgBox "El proceso ha finalizado correctamente.", vbInformation
        End If
    End If
    
    'Este bien o mal, si son proveedores abriremos el listado
    'Imprimimiremos un listado de contabilizacion de facturas
    '------------------------------------------------------
    If cadTabla <> "scafac" Then
        If DesdeGenerarFraProveedor Then
            
            'Esto lo paso dentro del FRMCOMFACTURAR, para poder mostrar otra pantalla con los datos
            

            
        Else
        
            If NumRegistros("Select count(*) from tmpinformes where codusu = " & vUsu.codigo) > 0 Then
                InicializarVbles
                CadParam = "|pEmpresa=""" & vEmpresa.nomempre & """|"
                numParam = numParam + 1
                
                CadParam = "|pDHFecha=""" & vUsu.Nombre & "   Hora: " & Format(Now, "hh:mm") & """|"
                numParam = numParam + 1
                cadFormula = "({tmpinformes.codusu} =" & vUsu.codigo & ")"
                cadNomRPT = "rContabPRO.rpt"
                conSubRPT = False
                cadTitulo = "Listado contabilizacion FRAPRO"
                
                LlamarImprimir True
            End If
        End If
    End If
    '---- Eliminar tabla TEMP de Errores
    BorrarTMPErrFact
    ejecutar "DELETE FROM tmpinformes WHERE codusu =" & vUsu.codigo, False
    ContabilizarFacturas = True
End Function





'Ccoste
'   0: No tendra analitica
'   1: Solo hay un CC que tratar. NO agruparemos por trabajador
'   2: Mas de un CC. Agruparemos por trabajador
Private Function PasarFacturasAContab(cadTabla As String, miCC As Byte) As Boolean
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim b As Boolean
Dim I As Integer
Dim NumFactu As Integer
Dim Codigo1 As String
Dim ContabilizacionAgrupadaTickets As Boolean

'ENERO 2009
Dim cContaFra As cContabilizarFacturas



    On Error GoTo EPasarFac

    PasarFacturasAContab = False
    
    
    'Si escontailizacion de facturas de tickets agrupados
    ContabilizacionAgrupadaTickets = False
    If Me.OptProve.Tag <> "" Then ContabilizacionAgrupadaTickets = True
    
    Set Rs = New ADODB.Recordset
    
    
    
    '---- Obtener el total de Facturas a Insertar en la contabilidad
    SQL = "SELECT count(*) "
    SQL = SQL & " FROM " & cadTabla & " INNER JOIN tmpFactu "
    If cadTabla = "scafac" Then
        Codigo1 = "codtipom"
    Else
        Codigo1 = "codprove"
    End If
    SQL = SQL & " ON " & cadTabla & "." & Codigo1 & "=tmpFactu." & Codigo1
    SQL = SQL & " AND " & cadTabla & ".numfactu=tmpFactu.numfactu AND " & cadTabla & ".fecfactu=tmpFactu.fecfactu "
    
    
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        NumFactu = Rs.Fields(0)
    Else
        NumFactu = 0
    End If
    Rs.Close
    Set Rs = Nothing


    'Enero 2009
    '------------------------------------------------------------
    Set cContaFra = New cContabilizarFacturas
    
    If Not cContaFra.EstablecerValoresInciales(ConnConta) Then
        'NO ha establcedio los valores de la conta.  Le dejaremos seguir, avisando que
        ' obviamente, no va a contabilizar las FRAS
        SQL = "Si continua, las facturas se insertaran en el registro, pero no ser�n contabilizadas" & vbCrLf
        SQL = SQL & "en este momento. Deber�n ser contabilizadas desde el ARICONTA" & vbCrLf & vbCrLf
        SQL = SQL & Space(50) & "�Continuar?"
        If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Function
    End If
    
    
    
    


    '---- Pasar cada una de las facturas seleccionadas a la Conta
    If NumFactu > 0 Then
    
        Set Rs = New ADODB.Recordset
    
        CargarProgres Me.ProgressBarContab, NumFactu
        
        
        'PreComproabacion de los asientos
        If cContaFra.RealizarContabilizacion Then
            SQL = "Select min(fecfactu) from tmpfactu"
            Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not Rs.EOF Then
                If Not cContaFra.PreComprobacionNumeroAsiento(Rs.Fields(0), NumFactu) Then
                    
                    'Para que la ventana siguiente muestr bien el error
                    SQL = "Insert into tmpErrFac(codtipom,numfactu,fecfactu,error) VALUES ("
                    SQL = SQL & "'',0,'" & Format(Rs.Fields(0), FormatoFecha) & "','Error contadores')"
                    
                    conn.Execute SQL
                    Rs.Close
                    Err.Raise 6, , "Comprobacion numeros asiento"
                End If
            End If
            Rs.Close
        End If
        
        'seleccinar todas las facturas que hemos insertado en la temporal (las que vamos a contabilizar)
        SQL = "SELECT * "
        SQL = SQL & " FROM tmpFactu "
            

        Rs.Open SQL, conn, adOpenStatic, adLockPessimistic, adCmdText
        I = 1

        b = True
   
   
   
   
   
   
   
        
        'pasar a contabilidad cada una de las facturas seleccionadas
        While Not Rs.EOF
        
            'Segun sea cli o pro
            If cadTabla = "scafac" Then
                SQL = cadTabla & "." & Codigo1 & "=" & DBSet(Rs.Fields(0), "T") & " AND scafac.numfactu=" & Rs!NumFactu
                SQL = SQL & " and scafac.fecfactu=" & DBSet(Rs!FecFactu, "F")
                If PasarFactura(SQL, miCC, ContabilizacionAgrupadaTickets, cContaFra) = False And b Then b = False
            Else
                SQL = cadTabla & "." & Codigo1 & "=" & DBSet(Rs.Fields(0), "N") & " and scafpc.numfactu=" & DBSet(Rs!NumFactu, "T")
                SQL = SQL & " and scafpc.fecfactu=" & DBSet(Rs!FecFactu, "F")
                If PasarFacturaProv(SQL, miCC, Orden2, cContaFra) = False And b Then b = False
            End If
            
            '---- Laura 26/10/2006
            'Al pasar cada factura al hacer el commit desbloqueamos los registros
            'que teniamos bloqueados y los volvemos a bloquear
            'Laura: 11/10/2006 bloquear los registros q vamos a contabilizar
            SQL = cadTabla & " INNER JOIN tmpFactu ON " & cadTabla & "." & Codigo1 & "=tmpFactu." & Codigo1 & " AND " & cadTabla & ".numfactu=tmpFactu.numfactu AND " & cadTabla & ".fecfactu=tmpFactu.fecfactu "
            If Not BloqueaRegistro(SQL, cadTabla & "." & Codigo1 & "=tmpFactu." & Codigo1 & " AND " & cadTabla & ".numfactu=tmpFactu.numfactu AND " & cadTabla & ".fecfactu=tmpFactu.fecfactu") Then
'                MsgBox "No se pueden Contabilizar Facturas. Hay registros bloqueados.", vbExclamation
'                Screen.MousePointer = vbDefault
'                Exit Sub
            End If
            '----
            
            IncrementarProgres Me.ProgressBarContab, 1
            Me.lblProgess2(1).Caption = "Insertando Facturas en Contabilidad...   (" & I & " de " & NumFactu & ")"
            Me.Refresh
            I = I + 1
            Rs.MoveNext   'Siguiente factura
        Wend
        
        'Veremos si ha dado error la contabilizacion de factiras
        If cContaFra.TieneErrores Then cContaFra.MuestraErroresContabilizacion
        
        
        Rs.Close
        Set Rs = Nothing
    End If
    
EPasarFac:
    If Err.Number <> 0 Then b = False
    Set cContaFra = Nothing
    If b Then
        PasarFacturasAContab = True
    Else
        PasarFacturasAContab = False
    End If
End Function



Private Sub ListadosAlmacen(H As Integer, W As Integer)
    'LISTADOS DE ALMACENES
    '---------------------
    Select Case OpcionListado
        Case 1   'Listados de Marcas
            PonerFrameListadoVisible True, H, W
            Me.lblTitulo(1).Caption = "Listado de Marcas"
            indFrame = 1
            codigo = "{smarca.codmarca}"
            Orden1 = "{smarca.codmarca}"
            Orden2 = "{smarca.nommarca}"
            cadTitulo = "Listado Marcas"
            cadNomRPT = "rAlmMarcas.rpt"
            conSubRPT = False
            
        Case 2   'Listado de Almacenes Propios
            PonerFrameListadoVisible True, H, W
            Me.lblTitulo(1).Caption = "Listado de Almacenes"
            indFrame = 1
            codigo = "{salmpr.codalmac}"
            Orden1 = "{salmpr.codalmac}"
            Orden2 = "{salmpr.nomalmac}"
            cadTitulo = "Listado Almacenes Propios"
            cadNomRPT = "rAlmAPropios.rpt"
            conSubRPT = False
            
        Case 3   'Listado de Tipos de Unidad
            PonerFrameListadoVisible True, H, W
            Me.lblTitulo(1).Caption = "Listado Tipos de Unidad"
            indFrame = 1
            codigo = "{sunida.codunida}"
            Orden1 = "{sunida.codunida}"
            Orden2 = "{sunida.nomunida}"
            cadTitulo = "Listado Tipos de Unidad"
            cadNomRPT = "rAlmTUnidad.rpt"
            conSubRPT = False
            
        Case 4   'Listado de Tipos de Art�culos
            PonerFrameListadoVisible True, H, W
            Me.lblTitulo(1).Caption = "Listado Tipos de Art�culos"
            indFrame = 1
            codigo = "{stipar.codtipar}"
            Orden1 = "{stipar.codtipar}"
            Orden2 = "{stipar.nomtipar}"
            txtCodigo(1).Tag = CadTag
            txtCodigo(2).Tag = CadTag
            cadTitulo = "Listado Tipos de Art�culos"
            cadNomRPT = "rAlmTArticulo.rpt"
            conSubRPT = False
            
        Case 6    'Listado de Art�culo
            ponerFrameArticulosVisible True, H, W
            CargarListViewOrden
            codigo = "{sartic"
            indFrame = 11
           
            
            
        Case 110   'Listados Ubicaciones Almacen
            PonerFrameListadoVisible True, H, W
            Me.lblTitulo(1).Caption = "Listado Ubicaciones Almacen"
            indFrame = 1
            codigo = "{subica.codubica}"
            Orden1 = "{subica.codubica}"
            Orden2 = "{subica.nomubica}"
            cadTitulo = "Listado Ubicaciones Almacen"
            cadNomRPT = "rAlmUbica.rpt"
            conSubRPT = False
            
        Case 18, 247 'Informe Stocks Maximos y Minimos   'OPCION: 247 es este tb
            ponerFrameArticulosVisible True, H, W
            codigo = "{salmac"
            indFrame = 11
            cmbProduccion.ListIndex = 0
            cmbProduccion.visible = vParamAplic.Produccion
            Label4(90).visible = vParamAplic.Produccion
            
        Case 7, 8 '7: Informe de Traspasos de Almacen
                  '8: Informe de Movimientos de Almacen
            If OpcionListado = 7 Then
                Me.lblTitulo(2).Caption = "Informe Traspaso de Almacen"
                Me.Label2(1).Caption = "N� Traspaso"
                codigo = "{scatra.codtrasp}"
            Else
                Me.lblTitulo(2).Caption = "Informe Movimientos de Almacen"
                Me.Label2(1).Caption = "N� Movimiento"
                codigo = "{scamov.codmovim}"
            End If
            H = 3495
            W = 5835
            PonerFrameVisible Me.FrameInfAlmacen, True, H, W
            indFrame = 2
            If NumCod <> "" Then
                txtCodigo(3).Text = NumCod
                txtCodigo(4).Text = NumCod
            End If
            
        Case 9 'Informe Movimiento Art�culos
            W = 10700
            H = 5775
            PonerFrameVisible Me.FrameMovArtic, True, H, W
            indFrame = 3
            codigo = "{smoval.codartic}"
            cadTitulo = "Informe Movimientos Articulos"
            conSubRPT = True
            CargarListView
            
        ' ====  [16/09/2009] LAURA : Listado Articulos con componentes
        Case 11
            W = Me.FrameInvArtComp.Width
            H = Me.FrameInvArtComp.Height
            PonerFrameVisible Me.FrameInvArtComp, True, H, W
            codigo = "{sartic.codartic}"
            cadTitulo = "Listado Art�culos con Componentes"
        ' ====
            
        Case 12 '12: Listado Toma de Inventario Articulos
            PonerFrameInventarioVisible2 True, H, W
            indFrame = 4
            Me.chkImprimeStock.visible = True
            Me.lbltituloInven.Caption = "Listado Toma de Inventario Articulos"
            cadTitulo = "Toma Inventario Articulos"
            'codigo = "{salmac.codalmac}"
            
        Case 13 '13: Listado Diferencias de Inventario Articulos
            PonerFrameInventarioVisible2 True, H, W
            indFrame = 4
            Me.lbltituloInven.Caption = "Listado Diferencias de Inventario Articulos"
            'codigo = "{sinven.codalmac}"
            cadTitulo = "Diferencias Inventario Articulos"
            
        Case 14 '14: Actualizar Direfencias Inventario (NO IMPRIME INFORME)
            PonerFrameInventarioVisible2 True, H, W
            indFrame = 4
            Me.lbltituloInven.Caption = "Actualizar Diferencias de Inventario de Articulos"
            Me.Caption = "Inventario de Articulos"
            
        Case 15 '15: Listado de Articulos Inactivos
            PonerFrameInventarioVisible2 True, H, W
            indFrame = 4
            Me.lbltituloInven.Caption = "Listado Articulos Inactivos"
            cadTitulo = "Listado Articulos Inactivos"
    
        Case 16 '16 .- Listado Valoracion de Stocks Inventariados
            PonerFrameInventarioVisible2 True, H, W
            indFrame = 4
            Me.lbltituloInven.Caption = "Listado Valoraci�n Stocks Inventariados"
            cadTitulo = "Listado Valoraci�n Stocks Inventariados"
            
        Case 17 '17 .- Listado Valoraci�n Stocks
            PonerFrameInventarioVisible2 True, H, W
            indFrame = 4
            Me.lbltituloInven.Caption = "Listado Valoraci�n Stocks"
            cadTitulo = "Listado Valoraci�n Stocks"
            
        Case 19 '19 .- Inf. Stocks a una Fecha
            PonerFrameInventarioVisible2 True, H, W
            indFrame = 4
            Me.lbltituloInven.Caption = "Informe Stocks a una Fecha"
            cadTitulo = "Stocks a una Fecha"
        Case 100
            H = Me.FrameAlmacenStkMin.Height
            W = Me.FrameAlmacenStkMin.Width
            PonerFrameVisible FrameAlmacenStkMin, True, H, W
            Label3(116).Caption = ""
    End Select
End Sub


Private Sub ListadosFacturacion(H As Integer, W As Integer)
    Select Case OpcionListado
        Case 20    'Listado de Actividades de Clientes
            PonerFrameListadoVisible True, H, W
            Me.lblTitulo(1).Caption = "Listado Actividades de Clientes"
            indFrame = 1
            codigo = "{sactiv.codactiv}"
            Orden1 = "{sactiv.codactiv}"
            Orden2 = "{sactiv.nomactiv}"
            cadTitulo = "Listado Actividades de Clientes"
            cadNomRPT = "rFacActividades.rpt"
            
        Case 21    'Listado de Zonas de Clientes
            PonerFrameListadoVisible True, H, W
            Me.lblTitulo(1).Caption = "Listado Zonas de Clientes"
            indFrame = 1
            codigo = "{szonas.codzonas}"
            Orden1 = "{szonas.codzonas}"
            Orden2 = "{szonas.nomzonas}"
            cadTitulo = "Listado Zonas de Clientes"
            cadNomRPT = "rFacZonas.rpt"
        
        Case 22    'Listado de Rutas de Asistencia
            PonerFrameListadoVisible True, H, W
            Me.lblTitulo(1).Caption = "Listado Rutas de Asistencia"
            indFrame = 1
            codigo = "{srutas.codrutas}"
            Orden1 = "{srutas.codrutas}"
            Orden2 = "{srutas.nomrutas}"
            cadTitulo = "Listado Rutas de Asistencia"
            cadNomRPT = "rFacRutas.rpt"
            
        Case 23     'Listado de Tipos de Formas de Env�o
            PonerFrameListadoVisible True, H, W
            Me.lblTitulo(1).Caption = "Listado Formas de Env�o"
            indFrame = 1
            codigo = "{senvio.codenvio}"
            Orden1 = "{senvio.codenvio}"
            Orden2 = "{senvio.nomenvio}"
            cadTitulo = "Listado Formas de Envio"
            cadNomRPT = "rFacEnvio.rpt"
            
        Case 24    'Tarifas Venta
            PonerFrameListadoVisible True, H, W
            Me.lblTitulo(1).Caption = "Listado Tarifas Venta"
            indFrame = 1
            codigo = "{starif.codlista}"
            Orden1 = "{starif.codlista}"
            Orden2 = "{starif.nomlista}"
            cadTitulo = "Listado Tarifas Venta"
            cadNomRPT = "rFacTarifasVen.rpt"
            
        Case 27     'Situaciones Especiales
            PonerFrameListadoVisible True, H, W
            Me.lblTitulo(1).Caption = "Situaciones Especiales"
            indFrame = 1
            codigo = "{ssitua.codsitua}"
            Orden1 = "{ssitua.codsitua}"
            Orden2 = "{ssitua.nomsitua}"
            cadTitulo = "Listado Situaciones Especiales"
            cadNomRPT = "rFacSituaciones.rpt"
            
        Case 28    '28: Informe de Tarifas de Precios
            PonerFrameTarifasVisible True, H, W
            Me.lblTituloTarif.Caption = "Informe Tarifas de Art�culos"
            codigo = "{slista"
            indFrame = 5
            cadTitulo = "Listado Tarifas Articulos"
            
        Case 29  '29: Informe Promociones
            PonerFrameTarifasVisible True, H, W
            Me.lblTituloTarif.Caption = "Informe Promociones Tarifas"
            codigo = "{spromo"
            indFrame = 5
            cadTitulo = "Listado Promociones de Tarifas"
            
        Case 30 '30: Informe Precios Especiales
            PonerFrameTarifasVisible True, H, W
            Me.lblTituloTarif.Caption = "Informe Precios Especiales Art�culos"
            codigo = "{sprees"
            indFrame = 5
            cadTitulo = "Listado Precios Especiales"
            
        Case 245, 247 '245: Informe control margenes tarifas
            indFrame = 5
            PonerFrameTarifasVisible True, H, W
            Me.lblTituloTarif.Caption = "Informe Control Margenes de Tarifas"
            codigo = "{slista"
            cadTitulo = "Listado Control Margenes Tarifas"
            cboDecimales.ListIndex = 4
        Case 246 '246: Informe margen ventas x articulo
            indFrame = 15
            H = 5300
            W = 7820
            PonerFrameVisible Me.FrameEstMargenes, True, H, W
            cadTitulo = "Listado Margen ventas por art�culo"
    End Select
End Sub


Private Sub ListadosMantenimiento(H As Integer, W As Integer)
'=============================================
'==== Listados de MANTENIMIENTOS

    Select Case OpcionListado
        Case 70, 71, 76, 78, 79 'Listado Mantenimientos
            FrameManteAnu.visible = False
            Me.FrameManteActi.visible = False
            PonerFrameManteVisible True, H, W
            FrameRuta.visible = False
            Select Case OpcionListado
            Case 70, 76
                CargarComboTipoList
                If OpcionListado = 70 Then
                    Me.cboTipoList.ListIndex = 0
                    FrameRuta.BorderStyle = 0
                    FrameRuta.visible = True
                    FrameRuta.Top = 4740
                    
                Else
                    FrameManteAnu.visible = True
                    Me.cboTipoList.ListIndex = 2
                End If
                cadTitulo = "Informe de Mantenimientos"
                conSubRPT = False
            Case 71
                    txtCodigo(53).Text = Format(Now, "dd/mm/yyyy")
                    txtCodigo(54).Text = Format(Now, "dd/mm/yyyy")
                    cadTitulo = "Informe Revisiones Mantenimientos"
                    conSubRPT = True
            Case 78
                FrameManteActi.visible = True
            Case 79
                FrameManteActi.visible = True
            End Select
            indFrame = 9
            
        Case 72 'Informe Fichas de Mantenimiento
            H = 5295
            W = 7395
            PonerFrameVisible Me.FrameFichasMan2, True, H, W
            txtCodigo(61).Text = Year(Now) 'Ejercicio
            indFrame = 10
            cadTitulo = "Informe Fichas Mantenimientos"
            conSubRPT = True
        Case 77
            'Informe teorico
            H = FrameListMant2.Height
            W = FrameListMant2.Width
            PonerFrameVisible FrameListMant2, True, H, W
            indFrame = 77
    End Select
End Sub



Private Sub ListadosCompras(H As Integer, W As Integer)
'=============================================
'==== Listados de COMPRAS

    Select Case OpcionListado
        Case 309 '309: Listado precios de compra
            H = 4450
            W = 6920
            PonerFrameVisible Me.FrameDtosFM, True, H, W
            ponerOptVisible False
            Me.Frame4.visible = True
            Me.Frame4.Top = 840
            Me.Frame5.visible = False
            Me.Frame6.visible = False
            
            chkVarios(1).visible = True
            chkVarios(1).Top = 3300
            chkVarios(2).visible = True
            chkVarios(2).Top = chkVarios(1).Top
            chkVarios(4).visible = vParamAplic.NumeroInstalacion = 2
            chkVarios(4).Top = chkVarios(1).Top
            chkVarios(4).Value = 0
            Label4(103).Top = chkVarios(1).Top + 320
            Label4(103).Left = chkVarios(1).Left
            Label4(103).visible = True
            
            Me.cmdAceptarDtosFM.Top = 3500
            Me.cmdCancel(12).Top = Me.cmdAceptarDtosFM.Top
            indFrame = 6
    End Select
End Sub



Private Sub ListadosReparaciones(H As Integer, W As Integer)
'=============================================
'==== Listados de REPARACIONES

    Select Case OpcionListado
        Case 407 'Sustituci�n Num. serie
            H = 3700
            W = 5720
            PonerFrameVisible Me.FrameRepSustNSerie, True, H, W
            Me.lblNumSerie(0).Caption = "N� Serie:   " & NumCod
            Me.lblNumSerie(1).Caption = "Art�culo:   " & Me.CadTag
            Me.Caption = "Numeros de Serie"
            indFrame = 13
            
        Case 409 '409: Listado de avisos de averia pendientes
            H = FrameListAvisosPtes.Height + 120
            W = FrameListAvisosPtes.Width + 120
            PonerFrameVisible Me.FrameListAvisosPtes, True, H, W
            CargarComboSituacion
            indFrame = 14
            Me.cboSituaAviso.ListIndex = 0
    End Select
End Sub




'---------------------------------------------------
'Para los bultos
Private Sub LimpiarTextosBultos()
Dim I As Integer
    For I = 2 To 6
        Me.txtBultos(I).Text = ""
        Me.txtBultos(I).Tag = ""
    Next I
End Sub



Private Sub PonerCamposDireccionBultos(indice As Integer)
Dim I As Integer

    'El indice mara el listindex del combo, por lo tanto sera indice + 1
    For I = 2 To 6
        Me.txtBultos(I).Text = RecuperaValor(Me.txtBultos(I).Tag, indice + 1)
    Next I
End Sub


Private Sub PonerCamposAlbaran()
'Informe Etiquetas Bultos
'si en NumCod se ha pasado el n� de un Albaran cargar por defectos valores
Dim SQL As String
Dim Rs As ADODB.Recordset

    On Error GoTo ErrAlb
    
    '1) -- Buscar en la tabla de ALBARANES: PED -> ALV
    SQL = "SELECT codclien,coddirec, sum(numbultos) as totBultos"
    SQL = SQL & " FROM scaalb c INNER JOIN slialb l ON c.numalbar=l.numalbar and c.codtipom=l.codtipom"
    SQL = SQL & " WHERE c.numalbar=" & NumCod & " and c.codtipom='ALV'"
    SQL = SQL & " GROUP by c.numalbar,c.codtipom"
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    If Not Rs.EOF Then
        txtClie.Text = Rs!codClien
    
        CadTag = DBLet(Rs!CodDirec, "T")
        
        txtBultos(1).Text = DBLet(Rs!totbultos, "N")
        
        txtClie_LostFocus
    End If
    
    Rs.Close
    Set Rs = Nothing
    
    '2) Buscar en la tabla de FACTURAS PED -> FAV
    If txtClie.Text = "" Then
         'Comprobar en FACTURAS: x si se pasa de PED -> FAC
        SQL = "SELECT codclien,coddirec, sum(numbultos) as totBultos "
        SQL = SQL & " FROM (scafac c INNER JOIN scafac1 a ON c.numfactu=a.numfactu and c.codtipom=a.codtipom and c.fecfactu=a.fecfactu)"
        SQL = SQL & " INNER JOIN slifac l ON a.numfactu=l.numfactu and a.codtipom=l.codtipom and a.fecfactu=l.fecfactu and a.numalbar=l.numalbar and a.codtipoa=l.codtipoa"
        SQL = SQL & " WHERE a.numalbar=" & NumCod & " and a.codtipoa='ALV'"
        SQL = SQL & " GROUP BY a.numalbar,a.codtipoa"
        
        Set Rs = New ADODB.Recordset
        Rs.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        
        If Not Rs.EOF Then
            txtClie.Text = Rs!codClien
        
            CadTag = DBLet(Rs!CodDirec, "T")
            
            txtBultos(1).Text = DBLet(Rs!totbultos, "N")
            
            txtClie_LostFocus
        End If
        Rs.Close
        Set Rs = Nothing
    End If
    
    Exit Sub
    
ErrAlb:
    MuestraError Err.Number, "Poner campos Albaran.", Err.Description
End Sub



'-----------------------------------------------------------------------------
'-----------------------------------------------------------------------------
'
'   Borre de facturas
'
'
'   Borraremos las tablas de facturas , albaranes, hcos....
'
Private Sub CargaFechasPosibleEliminacion()
Dim F As Date
Dim F2 As Date
    Set miRsAux = New ADODB.Recordset
    cmbEliFac.Clear
    codigo = "select min(fecfactu) from scafac"
    miRsAux.Open codigo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    F2 = DateAdd("yyyy", -5, CDate("01/01/" & Year(Now)))

    codigo = F2
    If Not miRsAux.EOF Then
        If Not IsNull(miRsAux.Fields(0)) Then codigo = miRsAux.Fields(0)
    End If
    miRsAux.Close
    Set miRsAux = Nothing
    codigo = "31/12/" & Year(CDate(codigo))
    
    While CDate(codigo) < F2
        
        cmbEliFac.AddItem "     " & Format(CDate(codigo), "dd/mm/yyyy")
        codigo = CStr(DateAdd("yyyy", 1, CDate(codigo)))
    
    Wend
    If cmbEliFac.ListCount > 0 Then cmbEliFac.ListIndex = 0
End Sub

Private Function BorrarFacturas() As Boolean
Dim FechaBorre As Date



    On Error GoTo EBorraFac
    BorrarFacturas = False
    
    FechaBorre = CDate(Trim(Me.cmbEliFac.List(cmbEliFac.ListIndex)))
    
    'Compruebo si estaban todas las facturas contabilizadas
    '------------------------------------------------------
    codigo = "Select count(*) from scafac where fecfactu<='" & Format(FechaBorre, FormatoFecha) & "' and intconta = 0"
    miRsAux.Open codigo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    If Not miRsAux.EOF Then NumRegElim = DBLet(miRsAux.Fields(0), "N")
    miRsAux.Close
    
    If NumRegElim > 0 Then
        MsgBox "Existen " & NumRegElim & " facturas sin contabilizar en esas fechas", vbExclamation
        Exit Function
    End If
    
        
    'lo mismo para proeedores
    codigo = "Select count(*) from scafpc where fecfactu<='" & Format(FechaBorre, FormatoFecha) & "' and intconta = 0"
    miRsAux.Open codigo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    If Not miRsAux.EOF Then NumRegElim = DBLet(miRsAux.Fields(0), "N")
    miRsAux.Close
    
    If NumRegElim > 0 Then
        MsgBox "Existen " & NumRegElim & " facturas de proveedores sin contabilizar en esas fechas", vbExclamation
        Exit Function
    End If
        
        
        
    '------------------------------------------------------------------------------
    '  LOG de acciones
    Set LOG = New cLOG
    LOG.Insertar 1, vUsu, "Borre facturas: " & Format(FechaBorre, "dd/mm/yyyy")
    Set LOG = Nothing
    '-----------------------------------------------------------------------------
    
    
    
    '   Lo dicho. LAS TABLAS son las indcadas above (jeje arriba)
    '   La fecha la manda fecfactu
    codigo = "slifac|scafac1|svenci|srecom|scafac|"
    For NumRegElim = 1 To 5
        Orden1 = RecuperaValor(codigo, CInt(NumRegElim))
        Label3(83).Caption = "Tabla CLI: " & Orden1
        Orden1 = "DELETE FROM " & Orden1 & " WHERE  fecfactu<='" & Format(FechaBorre, FormatoFecha) & "'"
        Me.Refresh
        DoEvents
        conn.Execute Orden1
    Next NumRegElim
    
    '---------------------------------------------------------------------------------
    'Albarananes CLIENTES.
    '--
    codigo = "scaalb|schalb|slialb|slhalb|"
    For NumRegElim = 1 To 2
        Orden1 = RecuperaValor(codigo, CInt(NumRegElim))
        Label3(83).Caption = "Albaranes(L): " & Orden1
        Label3(83).Refresh
        Orden1 = "Select * from " & Orden1
        Orden1 = Orden1 & " WHERE  fechaalb<='" & Format(FechaBorre, FormatoFecha) & "'"
        miRsAux.Open Orden1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        Orden1 = RecuperaValor(codigo, NumRegElim + 2)
        Orden1 = "DELETE FROM " & Orden1 & " WHERE codtipom = '"
        While Not miRsAux.EOF
            conn.Execute Orden1 & miRsAux!codtipom & "'  AND numalbar = " & miRsAux!NumAlbar
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        'Borramos las cabceeras
        Orden1 = RecuperaValor(codigo, CInt(NumRegElim))
        Label3(83).Caption = "Albaranes(C): " & Orden1
        Label3(83).Refresh
        Orden1 = "DELETE from " & Orden1
        Orden1 = Orden1 & " WHERE  fechaalb<='" & Format(FechaBorre, FormatoFecha) & "'"
        conn.Execute Orden1
    Next
    DoEvents
    
    
    '---------------------------------------------------------------------------------
    'Pedidos CLIENTES.
    '--
    codigo = "scaped|schped|sliped|slhped|"
    For NumRegElim = 1 To 2
        Orden1 = RecuperaValor(codigo, CInt(NumRegElim))
        Label3(83).Caption = "Pedidos(L): " & Orden1
        Label3(83).Refresh
        Orden1 = "Select * from " & Orden1
        Orden1 = Orden1 & " WHERE  fecpedcl<='" & Format(FechaBorre, FormatoFecha) & "'"
        miRsAux.Open Orden1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Orden1 = RecuperaValor(codigo, NumRegElim + 2)
        Orden1 = "DELETE FROM " & Orden1 & " WHERE numpedcl = "
        While Not miRsAux.EOF
            conn.Execute Orden1 & miRsAux!numpedcl
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        'Cabce
        Orden1 = RecuperaValor(codigo, CInt(NumRegElim))
        Label3(83).Caption = "Pedidos(C): " & Orden1
        Label3(83).Refresh
        Orden1 = "DELETE from " & Orden1
        Orden1 = Orden1 & " WHERE  fecpedcl<='" & Format(FechaBorre, FormatoFecha) & "'"
        conn.Execute Orden1
    Next
    DoEvents
    
    
    '---------------------------------------------------------------------------------
    'ofertas CLIENTES.
    '--
    codigo = "scapre|schpre|slipre|slhpre|"
    For NumRegElim = 1 To 2
        Orden1 = RecuperaValor(codigo, CInt(NumRegElim))
        Label3(83).Caption = "Ofertas(L): " & Orden1
        Label3(83).Refresh
        Orden1 = "Select * from " & Orden1
        Orden1 = Orden1 & " WHERE  fecofert<='" & Format(FechaBorre, FormatoFecha) & "'"
        miRsAux.Open Orden1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Orden1 = RecuperaValor(codigo, NumRegElim + 2)
        Orden1 = "DELETE FROM " & Orden1 & " WHERE numofert = "
        While Not miRsAux.EOF
            conn.Execute Orden1 & miRsAux!NumOfert
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        
        
        Orden1 = RecuperaValor(codigo, CInt(NumRegElim))
        Label3(83).Caption = "Ofertas(C): " & Orden1
        Label3(83).Refresh
        Orden1 = "DELETE from " & Orden1
        Orden1 = Orden1 & " WHERE  fecofert <='" & Format(FechaBorre, FormatoFecha) & "'"
        conn.Execute Orden1
        
    Next
    DoEvents
    
    
    codigo = "scarep|schrep|slirep|"
    For NumRegElim = 1 To 2
        Orden1 = RecuperaValor(codigo, CInt(NumRegElim))
        Label3(83).Caption = "Reparaciones: " & Orden1
        Label3(83).Refresh
        Orden1 = "Select * from " & Orden1
        Orden1 = Orden1 & " WHERE  fecrepar<='" & Format(FechaBorre, FormatoFecha) & "'"
        If NumRegElim = 1 Then
            'Lineas de reparacion solo hay en scarep
            'En shrep no hay lineas
            miRsAux.Open Orden1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            Orden1 = RecuperaValor(codigo, NumRegElim + 2)
            Orden1 = "DELETE FROM " & Orden1 & " WHERE numrepar = "
            While Not miRsAux.EOF
                conn.Execute Orden1 & miRsAux!numrepar
                miRsAux.MoveNext
            Wend
            miRsAux.Close
            
        End If
        Orden1 = RecuperaValor(codigo, CInt(NumRegElim))
        Label3(83).Caption = "Ofertas(C): " & Orden1
        Label3(83).Refresh
        Orden1 = "DELETE from " & Orden1
        Orden1 = Orden1 & " WHERE  fecrepar <='" & Format(FechaBorre, FormatoFecha) & "'"
        conn.Execute Orden1
        
    Next
    DoEvents
    
    
    'TPV
    Label3(83).Caption = "TPV"
    Label3(83).Refresh
    Orden1 = " WHERE  fecventa <='" & Format(FechaBorre, FormatoFecha) & "'"
    conn.Execute "DELETE FROM sliven " & Orden1
    conn.Execute "DELETE FROM scaven " & Orden1

    
    'PRODUCCION
    Label3(83).Caption = "Produccion"
    Label3(83).Refresh
    Orden1 = "Select * from sordprod WHERE  feccreacion<='" & Format(FechaBorre, FormatoFecha) & "'"
    miRsAux.Open Orden1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Orden1 = "DELETE FROM sliordpr WHERE codigo = "
    While Not miRsAux.EOF
        conn.Execute Orden1 & miRsAux!codigo
        miRsAux.MoveNext
    Wend
    miRsAux.Close

    Orden1 = "DELETE from sordprod WHERE  feccreacion <='" & Format(FechaBorre, FormatoFecha) & "'"
    conn.Execute Orden1

    Me.Refresh
    DoEvents
    
    '---------------------------------------------------------------------------------
    'Facturas proveedor
    '--
    codigo = "slifpc|scafpa|scafpc|"
    For NumRegElim = 1 To 3
        Orden1 = RecuperaValor(codigo, CInt(NumRegElim))
        Label3(83).Caption = "Tabla PRO: " & Orden1
        Orden1 = "DELETE FROM " & Orden1 & " WHERE  fecfactu<='" & Format(FechaBorre, FormatoFecha) & "'"
        conn.Execute Orden1
    Next NumRegElim
    
    
    
    
    codigo = "slhalp|slialp|scaalp|schalp|"
    For NumRegElim = 1 To 4
        Orden1 = RecuperaValor(codigo, CInt(NumRegElim))
        Label3(83).Caption = "Albaranes prov: " & Orden1
        Label3(83).Refresh
        Orden1 = "DELETE from " & Orden1 & " WHERE  fechaalb<='" & Format(FechaBorre, FormatoFecha) & "'"
        conn.Execute Orden1
        
    Next
    DoEvents
    
    
    
    
    '-----------------------------------------------
    'Pedidos proveedor
    '--
    codigo = "scappr|schppr|slippr|slhppr|"
    For NumRegElim = 1 To 2
        Orden1 = RecuperaValor(codigo, CInt(NumRegElim))
        Label3(83).Caption = "Pedidos prov.(L): " & Orden1
        Label3(83).Refresh
        Orden1 = "Select * from " & Orden1
        Orden1 = Orden1 & " WHERE  fecpedpr<='" & Format(FechaBorre, FormatoFecha) & "'"
        miRsAux.Open Orden1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Orden1 = RecuperaValor(codigo, NumRegElim + 2)
        Orden1 = "DELETE FROM " & Orden1 & " WHERE numpedpr = "
        While Not miRsAux.EOF
            conn.Execute Orden1 & miRsAux!numpedpr
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        
        
        Orden1 = RecuperaValor(codigo, CInt(NumRegElim))
        Label3(83).Caption = "Pedidos prov.(C): " & Orden1
        Label3(83).Refresh
        Orden1 = "DELETE from " & Orden1
        Orden1 = Orden1 & " WHERE  fecpedpr <='" & Format(FechaBorre, FormatoFecha) & "'"
        conn.Execute Orden1
        
    Next
    Me.Refresh
    DoEvents
    
    'slhmov slhtra
    Label3(83).Caption = "Hco movimientos"
    Label3(83).Refresh
    Orden1 = "DELETE FROM slhmov WHERE  fecmovim<='" & Format(FechaBorre, FormatoFecha) & "'"
    conn.Execute Orden1
    
    Label3(83).Caption = "Hco traspasos"
    Label3(83).Refresh
    Orden1 = "DELETE FROM slhtra WHERE  fechamov<='" & Format(FechaBorre, FormatoFecha) & "'"
    conn.Execute Orden1
    
    
    'Ahora me cargo los movimientos en la smoval
    Label3(83).Caption = "Movimientos"
    Label3(83).Refresh
    Orden1 = "DELETE FROM smoval WHERE  fechamov<='" & Format(FechaBorre, FormatoFecha) & "'"
    conn.Execute Orden1
    
    'Inventario
    Label3(83).Caption = "Hco inventario"
    Label3(83).Refresh
    Orden1 = "DELETE FROM shinve WHERE  fechainv<='" & Format(FechaBorre, FormatoFecha) & "'"
    conn.Execute Orden1
    
    
    BorrarFacturas = True
    Exit Function
EBorraFac:
    MuestraError Err.Number
End Function


'Envio -EMAIL

Private Sub PonerTamnyosMail(peque As Boolean)
    If peque Then
        Me.Height = Me.FrameEnvioMail.Height + 60
        Me.Width = Me.FrameEnvioMail.Width
    Else
        Me.Height = Me.FrameMantenimientos.Height
        Me.Width = Me.FrameMantenimientos.Width
    End If
    Me.Height = Me.Height + 420
    Me.Width = Me.Width + 120
    Me.FrameEnvioMail.visible = peque
    DoEvents
    Me.Refresh
End Sub





Private Function GeneracionEnvioMail() As Boolean
Dim m As CParamRpt

    On Error GoTo EGeneracionEnvioMail
    GeneracionEnvioMail = False

    Set m = New CParamRpt
    If m.Leer(21) = 1 Then
        Set m = Nothing
        Exit Function
    End If
    
    cadSelect = "Select * from tmpnlotes where codusu =" & vUsu.codigo & " ORDER BY codalmac,numalbar,codprove"
    miRsAux.Open cadSelect, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
    While Not miRsAux.EOF
        
        If Dir(App.Path & "\docum.pdf", vbArchive) <> "" Then Kill App.Path & "\docum.pdf"
    
        Label14(22).Caption = "Mantenimiento: " & miRsAux!codArtic & " Cliente: " & miRsAux!Codprove
        Label14(22).Refresh
        
'
        cadFormula = "({scaman.nummante}='" & miRsAux!codArtic & "') "
        cadFormula = cadFormula & " AND ({scaman.codclien}=" & miRsAux!Codprove & ") "


        With frmImprimir
            .FormulaSeleccion = cadFormula
            .OtrosParametros = CadParam
            .NumeroParametros = numParam
            .SoloImprimir = False
            .EnvioEMail = True
            .NombreRPT = m.Documento
            .Opcion = 78  'Carta renovacion manteniientos
            .Titulo = ""
            .Show vbModal
        End With
    
                    
        'Subo el progress bar
        Label14(22).Caption = "Generando PDF"
        Label14(22).Refresh
        Me.PBMail.Value = Me.PBMail.Value + 1
        If (Me.PBMail.Value Mod 25) = 24 Then
            Me.Refresh
            DoEvents
            Espera 1
        End If
        Me.Refresh
        DoEvents
        
        
        
        'FileCopy App.Path & "\docum.pdf", App.Path & "\temp\" & RS!NumAlbar & Format(RS!codProve, "0000000") & Format(RS!codArtic, "0000000") & Format(RS!FechaAlb, "yymmdd") & ".pdf"
        FileCopy App.Path & "\docum.pdf", App.Path & "\temp\" & Format(miRsAux!Codprove, "0000000") & ".pdf"
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set m = Nothing
    GeneracionEnvioMail = True
    Exit Function
EGeneracionEnvioMail:
       MuestraError Err.Number
End Function



Private Function HacerSQLListado82_83() As Boolean
    
On Error GoTo EHacerSQLListado82_83
    
    
    HacerSQLListado82_83 = False
    InicializarVbles


    If OpcionListado = 82 Then
        'Hacer UPDATE de scaalb
        codigo = "UPDATE scaalb set factursn = 1 "
        If NumCod <> "" Then cadSelect = " codtipom ='" & NumCod & "'"
        
        CadParam = "fechaalb"
        cadFormula = CadenaDesdeHastaBD(txtCodigo(117).Text, txtCodigo(118).Text, "codclien", "N")
        If cadFormula <> "" Then
            If cadSelect <> "" Then cadSelect = cadSelect & " AND "
            cadSelect = cadSelect & cadFormula
        End If
        

    Else
        'Hacer borrar avisos
        codigo = "DELETE FROM scaavi"
        cadSelect = " situacio = 3"
        CadParam = "fechaavi"
    End If
    
    cadFormula = CadenaDesdeHastaBD(txtCodigo(119).Text, txtCodigo(120).Text, CadParam, "F")
    If cadFormula <> "" Then
        If cadSelect <> "" Then cadSelect = cadSelect & " AND "
        cadSelect = cadSelect & cadFormula
    End If
    
    If cadSelect <> "" Then cadSelect = " WHERE " & cadSelect
    codigo = codigo & cadSelect
    conn.Execute codigo
    
    If OpcionListado = 83 Then MsgBox "Proceso finalizado", vbExclamation
    
    HacerSQLListado82_83 = True
    Exit Function
EHacerSQLListado82_83:
    MuestraError Err.Number
End Function







Private Function NuevasComprobacionesContabilizacion(Proveedores As Boolean, ByVal SQL As String) As Boolean
Dim RT As ADODB.Recordset
Dim C As String
Dim F As Date
Dim fin As Boolean
Dim ComprobacionFechaMenor As Boolean
Dim cControlFra As CControlFacturaContab
    
    On Error GoTo ENuevasComprobacionesContabilizacion
    NuevasComprobacionesContabilizacion = False
    
    
    
    Set cControlFra = New CControlFacturaContab
        'Tenemos que comprobar la fecha factura
    Set RT = New ADODB.Recordset
    ComprobacionFechaMenor = False

    If Proveedores Then
        C = "select fecrecep from scafpc WHERE " & SQL
        C = C & " GROUP BY fecrecep ORDER BY fecrecep"
    Else
        C = "Select fecfactu from scafac WHERE " & SQL
        C = C & " GROUP BY fecfactu ORDER BY fecfactu"
    End If
    
    
    RT.Open C, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    fin = False
    While Not fin
        F = RT.Fields(0)
        C = cControlFra.FechaCorrectaContabilizazion(ConnConta, F)
        If C <> "" Then
            fin = True
        Else
            C = cControlFra.FechaCorrectaIVA(ConnConta, F)
            If C <> "" Then
                fin = True
            Else
                If Proveedores Then
                    'Solo compruebo una vez
                    If Not ComprobacionFechaMenor Then
                        If cControlFra.FechaRecepMenorQueProveedor(ConnConta, F) Then
                            C = "Factura contabilizada con fecha de recepci�n menor que ya existentes en contabilidad."
                            C = C & vbCrLf & vbCrLf & "�Continuar?"
                            If MsgBox(C, vbQuestion + vbYesNo) = vbYes Then
                                C = ""
                            Else
                                C = "Proceso cancelado por el usuario"
                            End If
                        End If
                        ComprobacionFechaMenor = True
                    End If
                End If
            End If
        End If
        RT.MoveNext
        If Not fin Then fin = RT.EOF
    Wend
    RT.Close
    
    If C <> "" Then
        C = C & "(" & F & ")"
        MsgBox C, vbExclamation
    Else
        NuevasComprobacionesContabilizacion = True
    End If
    
    
ENuevasComprobacionesContabilizacion:
    If Err.Number <> 0 Then MuestraError Err.Number, "Nueva Comprobacion Contabilizacion"
    Set RT = Nothing
    Set cControlFra = Nothing
End Function

Private Sub ponerOptVisible(Vis As Boolean)

        Me.optFrDto(0).visible = Vis
        Me.optFrDto(1).visible = Vis
        Me.optFrDto(2).visible = Vis
        Me.optFrDto(3).visible = Vis
End Sub




Private Sub ContabilizarUnaFacturaProveedor()
    Screen.MousePointer = vbHourglass
        OpcionListado = 223
        'Algunos valores para despues
        Me.OptProve.Tag = ""
        OptProve.Value = True
        txtCodigo(31).Text = "": txtCodigo(32).Text = ""  'Pongo las fechas vacias POR si acaso
       ContabilizarFacturas "scafpc", cadSelect, Me.pg1, Me.lblProvCon(0), lblProvCon(1), True
        
            lblProvCon(0).Caption = ""
            lblProvCon(1).Caption = "Finalizando proceso"
            Me.Refresh
            TerminaBloquear
            'Eliminar la tabla TMP
            BorrarTMPFacturas
            DesBloqueoManual ("COMCON") 'COMpras CONtabilizar

     
    Screen.MousePointer = vbDefault
    Unload Me
End Sub




Private Sub CargaDatosEnPedidos()
Dim s As String
Dim R As ADODB.Recordset
Dim J As Integer
    s = "DELETE FROM tmpsliped where codusu = " & vUsu.codigo
    conn.Execute s
    
    
    'PEdidos cliente
    s = cadSelect
    s = Replace(s, "{", "(")
    s = Replace(s, "}", ")")
    
    If s <> "" Then s = " AND " & s
    s = "select codalmac,sliped.codartic,sum(cantidad) c from sliped,sartic WHERE sliped.codartic=sartic.codartic" & s
    s = s & " group by codalmac,codartic"
    Set R = New ADODB.Recordset
    R.Open s, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    J = 0
    s = ""
    While Not R.EOF
        J = J + 1
        '(codusu,numpedcl,numlinea,codalmac,codartic,nomartic,cantidad,importel)
        s = s & ", (" & vUsu.codigo & ",1," & J & "," & R!codAlmac & "," & DBSet(R!codArtic, "T")
        s = s & ",''," & DBSet(R!C, "N") & ",0)"
        If Len(s) > 3500 Then
            s = Mid(s, 2)
            s = "INSERT INTO tmpsliped (codusu,numpedcl,numlinea,codalmac,codartic,nomartic,cantidad,importel) VALUES " & s
            conn.Execute s
            s = ""
        End If
        R.MoveNext
    Wend
    R.Close
    
    If s <> "" Then
            s = Mid(s, 2)
            s = "INSERT INTO tmpsliped (codusu,numpedcl,numlinea,codalmac,codartic,nomartic,cantidad,importel) VALUES " & s
            conn.Execute s
            s = ""
    End If
    
    
  
    'pedidos prov
    
    s = cadSelect
    s = Replace(s, "{", "(")
    s = Replace(s, "}", ")")
    
    If s <> "" Then s = " AND " & s
    s = "select codalmac,slippr.codartic,sum(cantidad) c from slippr,sartic WHERE slippr.codartic=sartic.codartic" & s
    s = s & " group by codalmac,codartic"
    Set R = New ADODB.Recordset
    R.Open s, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    J = 0
    s = ""
    While Not R.EOF
        J = J + 1
        '(codusu,numpedcl,numlinea,codalmac,codartic,nomartic,cantidad,importel)
        s = s & ", (" & vUsu.codigo & ",2," & J & "," & R!codAlmac & "," & DBSet(R!codArtic, "T")
        s = s & ",'',0," & DBSet(R!C, "N") & ")"
        If Len(s) > 3500 Then
            s = Mid(s, 2)
            s = "INSERT INTO tmpsliped (codusu,numpedcl,numlinea,codalmac,codartic,nomartic,cantidad,importel) VALUES " & s
            conn.Execute s
            s = ""
        End If
        R.MoveNext
    Wend
    R.Close
    
    If s <> "" Then
            s = Mid(s, 2)
            s = "INSERT INTO tmpsliped (codusu,numpedcl,numlinea,codalmac,codartic,nomartic,cantidad,importel) VALUES " & s
            conn.Execute s
            s = ""
    End If
    
    
    
    
    
    
    Set R = Nothing
    
End Sub


Private Sub CargaIconosAyuda()
Dim Ima As Image
    On Error Resume Next 'mejor que no diera errores, pero bien, tampoco vamos a enfadarnos
    For Each Ima In Me.imgayuda
        Ima.Picture = frmPpal.imgListComun.ListImages(46).Picture
    Next
    Err.Clear
End Sub




Private Function HacerInfrStockMinimo() As Boolean

    On Error GoTo eHacerInfrStockMinimo
    
    codigo = "DELETE FROM tmpinformes WHERE codusu = " & vUsu.codigo
    conn.Execute codigo
    
    
    codigo = " FROM sartic,salmac WHERE sartic.codartic=salmac.codartic AND ctrstock=1 AND artvario  =0  "
    'El desde hasta
    If cadSelect <> "" Then
        cadTitulo = Replace(cadSelect, "{", "")
        cadTitulo = Replace(cadTitulo, "}", "")
        codigo = codigo & " AND " & cadTitulo
    End If
    
    If Me.chkVarios(0).Value = 0 Then
        'Normal. Listado de stcok minimos
        codigo = codigo & " AND stockmin>0"
        
        
    Else
        'Los que no tiene minimo y tienen stock
        codigo = codigo & " AND COALESCE(stockmin,0)<=0 and canstock>0"
        
    End If
    codigo = "SELECT " & vUsu.codigo & ",codprove,codfamia,codalmac,sartic.codartic,nomartic,stockmin,stockmax,puntoped,if(canstock<0,0,canstock) stock " & codigo
    
    'codusu,campo2,codigo1,campo1,nombre1,nombre2,
    'codalmac,stockmin,stockmax,puntoped,canstock,numorden
    codigo = "INSERT INTO tmpinformes(codusu,codigo1,campo1,campo2,nombre1,nombre2,importe1,importe2,importe3,importe4) " & codigo
    
    Label3(116).Caption = "Insertando en BD"
    Label3(116).Refresh
    conn.Execute codigo
    DoEvents
    
    'Actualizamos la familia
    Set miRsAux = New ADODB.Recordset
    codigo = "Select campo1 from tmpinformes WHERE codusu = " & vUsu.codigo & " GROUP BY 1"
    Set miRsAux = New ADODB.Recordset
    Label3(116).Caption = "Leer familias"
    Label3(116).Refresh
    miRsAux.Open codigo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    While Not miRsAux.EOF
        NumRegElim = NumRegElim + 1
        codigo = DevuelveDesdeBD(conAri, "nomfamia", "sfamia", "codfamia", miRsAux!campo1)
        Label3(116).Caption = codigo
        Label3(116).Refresh
        codigo = "UPDATE tmpinformes SET nombre3=" & DBSet(codigo, "T") & " WHERE codusu =" & vUsu.codigo & " AND campo1 = " & miRsAux!campo1
        conn.Execute codigo
        miRsAux.MoveNext
        If (NumRegElim Mod 10) = 0 Then DoEvents
    Wend
    miRsAux.Close
   
   
    'marzo 2014
    'A�adir pedidos clientes de ese almacen
    
    If NumRegElim > 0 Then
        Label3(116).Caption = "Pedidos clientes"
        Label3(116).Refresh
        DoEvents
        
        codigo = "select codalmac,codartic,sum(cantidad) cuantos from sliped where (codalmac,codartic) IN"
        codigo = codigo & "(select campo2,nombre1 from tmpinformes where codusu=" & vUsu.codigo & " ) group by 1,2 ORDER BY 1,2"
        miRsAux.Open codigo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            Label3(116).Caption = miRsAux!codAlmac & " " & miRsAux!codArtic
            Label3(116).Refresh
            
            
            codigo = "UPDATE tmpinformes SET importe5=" & DBSet(DBLet(miRsAux!Cuantos, "N"), "N")
            codigo = codigo & " WHERE codusu =" & vUsu.codigo & " AND campo2 = " & miRsAux!codAlmac
            codigo = codigo & " AND nombre1 = " & DBSet(miRsAux!codArtic, "T")
            conn.Execute codigo
   
   
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        
    End If
   
    HacerInfrStockMinimo = NumRegElim > 0
    Label3(116).Caption = "Mostrar informe"
    Label3(116).Refresh
    
    
eHacerInfrStockMinimo:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set miRsAux = Nothing
End Function


Private Sub HazCalculoPrecioNetoProve()
Dim Rs As ADODB.Recordset

    Label4(103).Caption = "Preparando datos"
    Label4(103).Refresh
    BorrarTempInformes
        
    codigo = "SELECT slispr.*,nomartic,codfamia,codmarca from " & Orden1
    If cadSelect <> "" Then codigo = codigo & " WHERE " & cadSelect
    codigo = codigo & " ORDER BY codprove,codfamia,codmarca"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open codigo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    codigo = ""
    Set Rs = New ADODB.Recordset
    NumRegElim = 0
    'Para no tener que hacer Un monton de selects a sdtomp
    cadSelect = ""
    
    While Not miRsAux.EOF
        Label4(103).Caption = miRsAux!NomArtic
        Label4(103).Refresh
        
        NumRegElim = NumRegElim + 1
        
        
        Orden2 = Format(miRsAux!Codprove, "0000000") & Format(miRsAux!Codfamia, "00000") & Format(miRsAux!codmarca, "00000")
        
        If Orden2 <> cadSelect Then
            
            'Del modulo   vPrecio.ObtenerDescuentos2
            Orden1 = "SELECT dtoline1,dtoline2,rap1,rap2 FROM sdtomp "
            Orden1 = Orden1 & " WHERE codprove=" & miRsAux!Codprove & " AND codfamia=" & miRsAux!Codfamia
            Orden1 = Orden1 & " AND (codmarca is null or codmarca=" & miRsAux!codmarca & ")"
            Orden1 = Orden1 & " and (fechadto<= '" & Format(Now, FormatoFecha) & "') ORDER BY codmarca desc"
            
            If cadSelect <> "" Then Rs.Close
            Rs.Open Orden1, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
            
            cadSelect = Orden2
        End If
        
        '                                                                           rap1  rap2
        'codusu,codigo1,campo1 ,campo2,nombre1,nombre2,importe1,porcen1,porcen2,importe4,importe5
        '1er trozo del insert
        codigo = codigo & ", (" & vUsu.codigo & "," & NumRegElim & "," & miRsAux!Codprove & "," & miRsAux!Codfamia
        codigo = codigo & "," & DBSet(miRsAux!codArtic, "T") & "," & DBSet(miRsAux!NomArtic, "T") & ","
   
        cadTitulo = miRsAux!precioac
        If Not IsNull(miRsAux!fechanue) Then
            If miRsAux!fechanue <= Now Then cadTitulo = DBLet(miRsAux!precionu, "N")
        End If
        codigo = codigo & TransformaComasPuntos(cadTitulo) & ","
   
   
        
        If Not Rs.EOF Then
            If miRsAux!dtopermi = 0 Then
                codigo = codigo & "0,0,"
            Else
                codigo = codigo & TransformaComasPuntos(CStr(DBLet(Rs!dtoline1, "N"))) & "," & TransformaComasPuntos(CStr(DBLet(Rs!dtoline2, "N"))) & ","
            End If
            codigo = codigo & TransformaComasPuntos(CStr(DBLet(Rs!Rap1, "N"))) & "," & TransformaComasPuntos(CStr(DBLet(Rs!Rap2, "N"))) & ")"
            
        Else
            codigo = codigo & "0,0,0,0)"
        End If
       

        
        

        
        If (NumRegElim Mod 20) = 0 Then InsertaEnTmpHazCalculo
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    Set Rs = Nothing
    If codigo <> "" Then InsertaEnTmpHazCalculo
    

    Label4(103).Caption = "Calculando neto"
    Label4(103).Refresh
    'Seguntipo dto
    If vParamAplic.TipoDtos = 1 Then
        codigo = "UPDATE tmpinformes Set importeb3=importeb1 * ((100 - porcen1) / 100) WHERE codusu =" & vUsu.codigo
        conn.Execute codigo
        Espera 0.25
        codigo = "UPDATE tmpinformes Set importeb2=importeb3 * ((100 - porcen2) / 100) WHERE codusu =" & vUsu.codigo
        conn.Execute codigo
    Else
        codigo = "UPDATE tmpinformes Set importeb2=importeb1 * ((100 - (porcen1+porcen2)) / 100) WHERE codusu =" & vUsu.codigo
        conn.Execute codigo
    End If
    Label4(103).Caption = ""
End Sub


Private Sub InsertaEnTmpHazCalculo()
    codigo = Mid(codigo, 2)
    codigo = "" & codigo
    'codusu,codigo1,campo1 ,campo2,nombre1,nombre2,importe1,porcen1,porcen2
    codigo = "INSERT INTO tmpinformes(codusu,codigo1,campo1 ,campo2,nombre1,nombre2,importeb1,porcen1,porcen2,importe4,importe5) VALUES " & codigo
    conn.Execute codigo
    codigo = ""
End Sub



Private Sub HacerListadoDtosCliente()



    
    'Vaciamos
    conn.Execute "DELETE FROM tmpinformes WHERE codusu =" & vUsu.codigo
    
    
    'Solo tendra en cuenta el desde hasta cliente.
    'Sacara todos los dtos propios y los de su actividad(si es que estan dados de alta)
    codigo = "SELECT codclien,codactiv FROM sclien WHERE codclien>=0 "
    If txtCodigo(73).Text <> "" Then codigo = codigo & " AND codclien >=" & txtCodigo(73).Text
    If txtCodigo(74).Text <> "" Then codigo = codigo & " AND codclien <=" & txtCodigo(74).Text
    Set miRsAux = New ADODB.Recordset
    NumRegElim = 0
    miRsAux.Open codigo, conn, adOpenKeyset, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        NumRegElim = NumRegElim + 1
        miRsAux.MoveNext
    Wend
 
    
    If NumRegElim = 0 Then
        MsgBox "No existe datos para mostrar", vbExclamation
    Else
        If NumRegElim > 4 Then
            If MsgBox("El proceso puede llevar mucho tiempo. �Continuar?", vbQuestion + vbYesNo) = vbNo Then NumRegElim = 0
        End If
    End If
    
    Screen.MousePointer = vbHourglass
    If NumRegElim > 0 Then
        DoEvents
        miRsAux.MoveFirst
        
        While Not miRsAux.EOF
            codigo = "insert into tmpinformes(codusu,codigo1,campo1,campo2,importe1,importe2,fecha1,porcen1)"
            codigo = codigo & " select " & vUsu.codigo & ",codclien,codfamia,codmarca,dtoline1,dtoline2,fechadto,0"
            codigo = codigo & " from sdtofm where codclien=" & miRsAux!codClien
            conn.Execute codigo
            Espera 0.2
    
    
            'Los que vienen de descuento
            codigo = "insert into tmpinformes(codusu,codigo1,campo1,campo2,importe1,importe2,fecha1,porcen1)"
            codigo = codigo & " select " & vUsu.codigo & "," & miRsAux!codClien & ",codfamia,null,dtoline1,"
            codigo = codigo & " dtoline2,fechadto,1 from sdtofm where codactiv=" & miRsAux!codactiv & " and not codfamia in ("
            codigo = codigo & " select campo1 from tmpinformes where codusu =" & vUsu.codigo & " and codigo1=" & miRsAux!codClien & " and campo2 is null)"
            conn.Execute codigo
            
            miRsAux.MoveNext
        
        Wend
    End If
    miRsAux.Close
    
    'Si tiene alguno de MARCA
    codigo = "Select campo2 from tmpinformes WHERE codusu =" & vUsu.codigo & " AND campo2>=0 GROUP BY 1"
    miRsAux.Open codigo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        codigo = DevuelveDesdeBD(conAri, "nommarca", "smarca", "codmarca", miRsAux.Fields(0))
        codigo = "UPDATE tmpinformes SET nombre2=" & DBSet(codigo, "T") & " WHERE codusu =" & vUsu.codigo & " AND campo2 = " & miRsAux.Fields(0)
        conn.Execute codigo
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    
    
    codigo = "UPDATE tmpinformes SET nombre3='Activ.' WHERE codusu =" & vUsu.codigo & " AND porcen1>=1"
    conn.Execute codigo
    
    
    cadTitulo = "Descuento cliente / actividad"
    cadFormula = "({tmpinformes.codusu} = " & vUsu.codigo & ")"
    cadNomRPT = "rFacDtoCliACtiv.rpt"
    
    LlamarImprimir False
    
    
    
    Screen.MousePointer = vbDefault
End Sub




Private Sub PonerDatosFacturaProveedorAcabadaRecepcionar()
          'Si genera la factura y la contabiliza (UNA A UNA),
            numParam = CByte(vbCritical)
            cadFormula = "importe1"
            CadParam = DevuelveDesdeBD(conAri, "codigo1", "tmpinformes", "codusu", vUsu.codigo, "N", cadFormula)
             Orden1 = ""
            If CadParam <> "" Then
                If CadParam = "0" Then
                    CadParam = ""
                    Orden1 = ""
                Else
                    Orden1 = DevuelveDesdeBD(conAri, "importeb5", "tmpinformes", "codusu", vUsu.codigo, "N")
                    'If Orden1 <> "" And Val(Orden1) <> 0 Then
                    '    Orden1 = vbCrLf & "Asiento: " & Val(Orden1)
                    'Else
                    '    Orden1 = ""
                    'End If
                    
                    
                End If
                
            End If
            
            Me.txtimporte(5).Text = CadParam
            txtimporte(6).Text = Orden1
          
           
            CadParam = "Numero de registro: " & CadParam & vbCrLf
                If cadFormula <> "importe1" Then
                    cadFormula = DevuelveDesdeBD(conAri, "codmacta", "sprove", "codprove", CStr(Val(CCur(cadFormula))))
                    Orden1 = ""
                Else
                    cadFormula = ""
                End If
                
                If cadFormula = "" Then
                    Orden1 = "Error en cuenta contable proveedor"
                Else
                        
                
                    Orden1 = DevuelveDesdeBD(conAri, "nombre1", "tmpinformes", "codusu", vUsu.codigo, "N")
                    numParam = InStr(1, Orden1, "@")
                    If numParam = 0 Then
                        Orden1 = "Err en campo temporal(nombre1)"
                        numParam = CByte(vbCritical)
                    Else
                    
                    
                        Set miRsAux = New ADODB.Recordset
                    
                        If vParamAplic.ContabilidadNueva Then
                            Orden1 = "numfactu=" & DBSet(Trim(Mid(Orden1, 1, numParam - 1)), "T") & " AND fecfactu = " & DBSet(Trim(Mid(Orden1, numParam + 1)), "F")
                            Orden1 = "codmacta = '" & cadFormula & "' AND " & Orden1
                        Else
                            Orden1 = "numfactu=" & DBSet(Trim(Mid(Orden1, 1, numParam - 1)), "T") & " AND fecfactu = " & DBSet(Trim(Mid(Orden1, numParam + 1)), "F")
                            Orden1 = "ctaprove = '" & cadFormula & "' AND " & Orden1
                        End If
                        
                        cmdActVtosFraPro.Tag = Orden1
                        
                        Orden1 = "Select * from " & IIf(vParamAplic.ContabilidadNueva, "pagos", "spagop") & " WHERE " & Orden1 & " ORDER BY numorden"
                        miRsAux.Open Orden1, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
                        Orden1 = ""
                        indCodigo = 0
                        While Not miRsAux.EOF
                            
                            
                            
                            If indCodigo < 5 Then
                                PonerVisiblesParaVencimientosFraPro True
                                Me.txtimporte(indCodigo).Text = Format(miRsAux!impefect, FormatoImporte)
                                Me.txtCodigo(149 + indCodigo).Text = Format(miRsAux!fecefect, "dd/mm/yyyy")
                                'Lo pongo en el tag
                                Me.txtCodigo(149 + indCodigo).Tag = Me.txtCodigo(149 + indCodigo).Text
                                Label3(120 + indCodigo).Tag = miRsAux!numorden
                            Else
                                Label4(106).visible = True 'mas de 5 vtos
                            End If
                            indCodigo = indCodigo + 1
                            miRsAux.MoveNext
                            Orden1 = "OK"
                        Wend
                        miRsAux.Close
                        Set miRsAux = Nothing
                    
                        If Orden1 = "" Then
                            numParam = CByte(vbCritical)
                            Orden1 = "Error. No se han encontrado los vencimientos"
                        Else
                            Orden1 = "" 'para que no de msgbox
                            PonerFoco txtCodigo(149)
                        End If
                    End If
                End If
                

            
        If Orden1 <> "" Then MsgBox Orden1, numParam
            
End Sub




Private Sub PonerVisiblesParaVencimientosFraPro(visible As Boolean)

        Me.txtimporte(indCodigo).visible = visible
        Label3(120 + indCodigo).visible = visible
        Me.txtCodigo(149 + indCodigo).visible = visible
        imgFecha(21 + indCodigo).visible = visible
End Sub
