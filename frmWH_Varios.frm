VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmWH_Varios 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7125
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameFechaRechazo 
      Height          =   1695
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   3495
      Begin VB.CommandButton cmdRechazar 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   1200
         TabIndex        =   9
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox txtFecha 
         Height          =   375
         Index           =   1
         Left            =   2040
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   1
         Left            =   2280
         TabIndex        =   10
         Top             =   960
         Width           =   975
      End
      Begin VB.Label lblPedirfecha 
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   1545
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   1
         Left            =   1680
         Picture         =   "frmWH_Varios.frx":0000
         Top             =   360
         Width           =   240
      End
   End
   Begin VB.Frame FramePasarAclientes 
      Height          =   3975
      Left            =   120
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   5895
      Begin VB.CommandButton cmdVerClientePotencial 
         Height          =   495
         Left            =   5040
         Picture         =   "frmWH_Varios.frx":008B
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Modificar datos cliente"
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdPasarCliente 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3360
         TabIndex        =   24
         Top             =   3360
         Width           =   975
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   375
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   23
         Text            =   "Text2"
         Top             =   960
         Width           =   5295
      End
      Begin VB.ComboBox cboTipoIVA 
         Height          =   315
         ItemData        =   "frmWH_Varios.frx":0A8D
         Left            =   240
         List            =   "frmWH_Varios.frx":0A8F
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   2640
         Width           =   1455
      End
      Begin VB.ComboBox cboFacturacion 
         Height          =   315
         Left            =   3840
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   2640
         Width           =   1695
      End
      Begin VB.ComboBox cboAlbaran 
         Height          =   315
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   2640
         Width           =   1695
      End
      Begin VB.ComboBox cboFP 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   1800
         Width           =   5295
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   3
         Left            =   4560
         TabIndex        =   14
         Top             =   3360
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo IVA"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   22
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Facturación"
         Height          =   255
         Index           =   4
         Left            =   3840
         TabIndex        =   21
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Valorar Albaran con"
         Height          =   255
         Index           =   18
         Left            =   1920
         TabIndex        =   20
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Forma de pago"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   16
         Top             =   1560
         Width           =   1065
      End
      Begin VB.Label Label2 
         Caption         =   "Pasar a clientes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   435
         Index           =   0
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   2805
      End
   End
   Begin VB.Frame FrameDocCliente 
      Height          =   2895
      Left            =   480
      TabIndex        =   38
      Top             =   0
      Visible         =   0   'False
      Width           =   7575
      Begin VB.ComboBox cboEGDA 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   42
         Top             =   1680
         Width           =   4575
      End
      Begin VB.CommandButton cmdNuevoDesdeExpediente 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   5040
         TabIndex        =   43
         Top             =   2400
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   6
         Left            =   6240
         TabIndex        =   44
         Top             =   2400
         Width           =   975
      End
      Begin VB.TextBox txtFecha 
         Height          =   375
         Index           =   3
         Left            =   6000
         TabIndex        =   41
         Text            =   "Text1"
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   2
         Left            =   120
         Locked          =   -1  'True
         OLEDropMode     =   1  'Manual
         TabIndex        =   39
         Text            =   "Text1"
         Top             =   960
         Width           =   5655
      End
      Begin VB.Label Label1 
         Caption         =   "E.G.D.A"
         Height          =   195
         Index           =   13
         Left            =   120
         TabIndex        =   70
         Top             =   1440
         Width           =   765
      End
      Begin VB.Label Label2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   435
         Index           =   2
         Left            =   120
         TabIndex        =   46
         Top             =   240
         Width           =   7005
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   3
         Left            =   6600
         Picture         =   "frmWH_Varios.frx":0A91
         Top             =   720
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha"
         Height          =   195
         Index           =   2
         Left            =   6000
         TabIndex        =   45
         Top             =   720
         Width           =   465
      End
      Begin VB.Label Label1 
         Caption         =   "Fichero"
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   40
         Top             =   720
         Width           =   525
      End
      Begin VB.Image imgDir 
         Height          =   240
         Index           =   2
         Left            =   840
         Picture         =   "frmWH_Varios.frx":0B1C
         Top             =   720
         Width           =   240
      End
   End
   Begin VB.Frame FrameTitulares 
      Height          =   2655
      Left            =   1080
      TabIndex        =   60
      Top             =   0
      Width           =   6495
      Begin VB.CommandButton cmdTitularidad 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   4200
         TabIndex        =   63
         Top             =   1920
         Width           =   975
      End
      Begin VB.ComboBox cboTitularidad 
         Height          =   315
         ItemData        =   "frmWH_Varios.frx":0C1E
         Left            =   4680
         List            =   "frmWH_Varios.frx":0C28
         TabIndex        =   62
         Text            =   "Combo1"
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox txtTitularidad 
         Height          =   315
         Left            =   240
         TabIndex        =   61
         Text            =   "Text3"
         Top             =   1200
         Width           =   4335
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   11
         Left            =   5280
         TabIndex        =   64
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Titularidad"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   435
         Index           =   3
         Left            =   240
         TabIndex        =   65
         Top             =   360
         Width           =   4425
      End
   End
   Begin VB.Frame FrameContratoCli 
      Height          =   3255
      Left            =   1080
      TabIndex        =   71
      Top             =   3600
      Visible         =   0   'False
      Width           =   6135
      Begin VB.CommandButton cmdContratoCliente 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3840
         TabIndex        =   83
         Top             =   2640
         Width           =   975
      End
      Begin VB.Frame FrameModContrato 
         Caption         =   "Contestación "
         Height          =   975
         Left            =   120
         TabIndex        =   78
         Top             =   1320
         Width           =   5895
         Begin VB.OptionButton optContrato 
            Caption         =   "Aceptado"
            Height          =   255
            Index           =   1
            Left            =   4200
            TabIndex        =   82
            Top             =   480
            Width           =   1455
         End
         Begin VB.OptionButton optContrato 
            Caption         =   "Rechazado"
            Height          =   255
            Index           =   0
            Left            =   2760
            TabIndex        =   81
            Top             =   480
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.TextBox txtFecha 
            Height          =   315
            Index           =   6
            Left            =   1200
            TabIndex        =   79
            Text            =   "Text1"
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label Label3 
            Caption         =   "Fecha"
            Height          =   195
            Index           =   5
            Left            =   240
            TabIndex        =   80
            Top             =   480
            Width           =   465
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   6
            Left            =   840
            Picture         =   "frmWH_Varios.frx":0C42
            Top             =   480
            Width           =   240
         End
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   4
         Left            =   240
         Locked          =   -1  'True
         OLEDropMode     =   1  'Manual
         TabIndex        =   75
         Text            =   "Text1"
         Top             =   1800
         Width           =   5655
      End
      Begin VB.TextBox txtFecha 
         Height          =   315
         Index           =   5
         Left            =   1800
         TabIndex        =   73
         Text            =   "Text1"
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   2
         Left            =   4920
         TabIndex        =   72
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "C"
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
         Left            =   840
         TabIndex        =   77
         Top             =   240
         Width           =   4365
      End
      Begin VB.Image imgDir 
         Height          =   240
         Index           =   4
         Left            =   960
         Picture         =   "frmWH_Varios.frx":0CCD
         Top             =   1560
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Fichero"
         Height          =   195
         Index           =   14
         Left            =   240
         TabIndex        =   76
         Top             =   1560
         Width           =   525
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   5
         Left            =   1440
         Picture         =   "frmWH_Varios.frx":0DCF
         Top             =   840
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "F. Presentacion"
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   74
         Top             =   840
         Width           =   1785
      End
   End
   Begin VB.Frame FrameActuacion 
      Height          =   6135
      Left            =   120
      TabIndex        =   47
      Top             =   0
      Width           =   7335
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   3
         Left            =   240
         Locked          =   -1  'True
         OLEDropMode     =   1  'Manual
         TabIndex        =   54
         Text            =   "Text1"
         Top             =   3360
         Width           =   6855
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   4
         Left            =   240
         TabIndex        =   49
         Text            =   "Text1"
         Top             =   2640
         Width           =   1215
      End
      Begin VB.TextBox txtObra 
         Height          =   1695
         Index           =   2
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   48
         Text            =   "frmWH_Varios.frx":0E5A
         Top             =   600
         Width           =   6855
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   9
         Left            =   6120
         TabIndex        =   53
         Top             =   5520
         Width           =   975
      End
      Begin VB.CommandButton cmdActuacion 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   4920
         TabIndex        =   52
         Top             =   5520
         Width           =   975
      End
      Begin VB.TextBox txtimporte 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   1680
         TabIndex        =   50
         Text            =   "Text1"
         Top             =   2640
         Width           =   1095
      End
      Begin VB.TextBox txtimporte 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   3000
         TabIndex        =   51
         Text            =   "Text1"
         Top             =   2640
         Width           =   1095
      End
      Begin MSComctlLib.ListView lwTrabajadores 
         Height          =   1215
         Left            =   2160
         TabIndex        =   67
         Tag             =   "Actuaciones"
         Top             =   3960
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   2143
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Codigo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Trabajador"
            Object.Width           =   6068
         EndProperty
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   0
         Left            =   1320
         Picture         =   "frmWH_Varios.frx":0E60
         ToolTipText     =   "Añadir trabajador"
         Top             =   3960
         Width           =   240
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   1
         Left            =   1680
         Picture         =   "frmWH_Varios.frx":1862
         ToolTipText     =   "Eliminar trabajador"
         Top             =   3960
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Trabajadores"
         Height          =   195
         Index           =   12
         Left            =   240
         TabIndex        =   66
         Top             =   3960
         Width           =   930
      End
      Begin VB.Image imgDir 
         Height          =   240
         Index           =   3
         Left            =   960
         Picture         =   "frmWH_Varios.frx":2264
         Top             =   3120
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Fichero"
         Height          =   195
         Index           =   11
         Left            =   240
         TabIndex        =   59
         Top             =   3120
         Width           =   525
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha"
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   58
         Top             =   2400
         Width           =   465
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   4
         Left            =   840
         Picture         =   "frmWH_Varios.frx":2366
         Top             =   2400
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Actuacion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   10
         Left            =   240
         TabIndex        =   57
         Top             =   360
         Width           =   930
      End
      Begin VB.Label Label1 
         Caption         =   "Horas"
         Height          =   195
         Index           =   8
         Left            =   3000
         TabIndex        =   56
         Top             =   2400
         Width           =   525
      End
      Begin VB.Label Label1 
         Caption         =   "Importe"
         Height          =   195
         Index           =   9
         Left            =   1680
         TabIndex        =   55
         Top             =   2400
         Width           =   525
      End
   End
   Begin VB.Frame FrameNuevaObra 
      Height          =   5895
      Left            =   120
      TabIndex        =   26
      Top             =   0
      Visible         =   0   'False
      Width           =   7215
      Begin VB.ComboBox cboPresentador 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   69
         Top             =   2640
         Width           =   3135
      End
      Begin VB.CommandButton cmdNuevaObra 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   4920
         TabIndex        =   37
         Top             =   5400
         Width           =   975
      End
      Begin VB.TextBox txtObra 
         Height          =   1695
         Index           =   1
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   34
         Text            =   "frmWH_Varios.frx":23F1
         Top             =   3480
         Width           =   6855
      End
      Begin VB.TextBox txtFecha 
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   32
         Text            =   "Text1"
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox txtObra 
         Height          =   375
         Index           =   0
         Left            =   1440
         MaxLength       =   100
         TabIndex        =   31
         Text            =   "Text3"
         Top             =   1800
         Width           =   5535
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   1
         Left            =   120
         Locked          =   -1  'True
         OLEDropMode     =   1  'Manual
         TabIndex        =   29
         Text            =   "Text1"
         Top             =   960
         Width           =   6855
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   4
         Left            =   6000
         TabIndex        =   28
         Top             =   5400
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Tipo"
         Height          =   255
         Left            =   120
         TabIndex        =   68
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Descripción"
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   36
         Top             =   3240
         Width           =   930
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre"
         Height          =   195
         Index           =   5
         Left            =   1440
         TabIndex        =   35
         Top             =   1560
         Width           =   555
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   33
         Top             =   1560
         Width           =   525
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   2
         Left            =   720
         Picture         =   "frmWH_Varios.frx":23F7
         Top             =   1560
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Fichero"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   30
         Top             =   720
         Width           =   525
      End
      Begin VB.Image imgDir 
         Height          =   240
         Index           =   1
         Left            =   840
         Picture         =   "frmWH_Varios.frx":2482
         Top             =   720
         Width           =   240
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   435
         Index           =   1
         Left            =   240
         TabIndex        =   27
         Top             =   240
         Width           =   6825
      End
   End
   Begin VB.Frame FrameInsertarDocumento 
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   7335
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   4800
         TabIndex        =   6
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox txtFecha 
         Height          =   375
         Index           =   0
         Left            =   5880
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   0
         Left            =   120
         Locked          =   -1  'True
         OLEDropMode     =   1  'Manual
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   600
         Width           =   5655
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   0
         Left            =   6120
         TabIndex        =   1
         Top             =   1320
         Width           =   975
      End
      Begin VB.Image imgDir 
         Height          =   240
         Index           =   0
         Left            =   720
         Picture         =   "frmWH_Varios.frx":2584
         Top             =   360
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha"
         Height          =   195
         Index           =   0
         Left            =   5880
         TabIndex        =   5
         Top             =   360
         Width           =   465
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   6480
         Picture         =   "frmWH_Varios.frx":2686
         Top             =   360
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Fichero"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmWH_Varios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public Opcion As Integer
    '
    '0.-    Insertar nueva propuesta/contraot
    '1.-    FRechazo oferta/contrato
    '2.-    OFERTA/contrato aceptad@
    
    '3.-    Pasar a CLIENTES
    
    '4.-    Nueva OBRA/expediente
    '5.-    Modificar OBRA expediente
    
    
    '6-     Nuevo PI
    '7-     Nuevo presentacion gestoras de derechos de autor
    '9-     Accion comercial (PUEDE LLEVAR DOCUMENTOS)
    '10     MODIFICAR acccion comercial
    
    '11     Titularidad
    '12     Modificar titularidad
    
    
    '13     Rechazar PRI o SGD
    '14     Aceptar PRI o sgd
    
    '------------ Contrato desde Cliente, no desde potencial
    '15     Nuevo
    '16     Modificar
    
    
Public ExtraData2 As String


Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmT As frmBasico2 'frmAdmTrabajadores
Attribute frmT.VB_VarHelpID = -1
    

Dim SQL As String
Dim PrimVez As Boolean
Dim DesdeAceptar As Boolean  'Para cargar o no la variable Cadenadesdeotroform
    
    
Private Sub cmdAceptar_Click()
    'Guaradaremos la extension
    
    SQL = ""
    If Me.Text1(0).Text = "" Then
    
            'Si es propuesta comercial dejamos que pase VACIA
            If RecuperaValor(ExtraData2, 3) = "1" Then 'propuesta comercial
            
            Else
                SQL = SQL & "-Nombre fichero vacio"
            End If
    Else
        If Dir(Text1(0).Text, vbArchive) = "" Then
                SQL = SQL & "- No existe el fichero"
        Else
            CadenaDesdeOtroForm = ExtensionSoportada(Text1(0).Text, True)   'EXTENSION
            If Len(CadenaDesdeOtroForm) > 10 Then SQL = SQL & "-" & CadenaDesdeOtroForm
        End If
    End If
    If Me.txtFecha(0).Text = "" Then
        SQL = SQL & vbCrLf & "-Fecha "
    Else
        If CDate(RecuperaValor(ExtraData2, 2)) > CDate(txtFecha(0).Text) Then SQL = SQL & vbCrLf & "-Fecha debe ser mayor que " & RecuperaValor(ExtraData2, 2)
    End If
    If SQL <> "" Then
        MsgBox "Errores: " & vbCrLf & SQL, vbExclamation
        Exit Sub
    End If
    
    
    'OK. VALE. INSERTAMOS el archivo
    SQL = RecuperaValor(ExtraData2, 3)
    If SQL = 1 Then
        'PROPUESTA COMERCIAL  PROPUESTA COMERCIAL   PROPUESTA COMERCIAL
        SQL = DevuelveDesdeBD(conAri, "max(idPropComer)", "whoexpedientepotprocomer", "codclien", RecuperaValor(ExtraData2, 1))
        If SQL = "" Then SQL = "0"
        SQL = Val(SQL) + 1
        'Cliente , PropuestaComercial , Id , ArchivoOrigen
        If CopiaArchivoWHOSE(CLng(RecuperaValor(ExtraData2, 1)), True, CLng(SQL), Text1(0).Text) Then
            'insert into whoexpedientepotcontrato(codclien,idPropComer,f_preprop)
            SQL = "(" & RecuperaValor(ExtraData2, 1) & "," & SQL & "," & DBSet(txtFecha(0).Text, "F") & ",'" & UCase(CadenaDesdeOtroForm) & "')"
            SQL = "insert into whoexpedientepotprocomer(codclien,idPropComer,f_preprop,extension) VALUES " & SQL
            If Not ejecutar(SQL, False) Then
                MsgBox "ERROR CRITICO insertando en BD", vbExclamation
            End If
            Unload Me
        End If
    Else
        SQL = DevuelveDesdeBD(conAri, "max(idcontrato)", "whoexpedientepotcontrato", "codclien", RecuperaValor(ExtraData2, 1))
        If SQL = "" Then SQL = "0"
        SQL = Val(SQL) + 1
        'Cliente , PropuestaComercial , Id , ArchivoOrigen
        If CopiaArchivoWHOSE(CLng(RecuperaValor(ExtraData2, 1)), False, CLng(SQL), Text1(0).Text) Then
            'insert into whoexpedientepotcontrato(codclien,idPropComer,f_preprop)
            SQL = "(" & RecuperaValor(ExtraData2, 1) & "," & SQL & "," & DBSet(txtFecha(0).Text, "F") & ",'" & LCase(CadenaDesdeOtroForm) & "')"
            SQL = "insert into whoexpedientepotcontrato(codclien,idcontrato,f_precont,extension ) VALUES " & SQL
            If Not ejecutar(SQL, False) Then
                MsgBox "ERROR CRITICO insertando en BD", vbExclamation
            End If
            Unload Me
        End If
    
    
    
    End If




End Sub

Private Sub cmdActuacion_Click()
Dim Aux As String
Dim b As Boolean
Dim idActua As Integer

    SQL = ""
    If Me.txtFecha(4).Text = "" Then SQL = "-Fecha"
    If Me.txtObra(2).Text = "" Then SQL = SQL & vbCrLf & "-Observaciones"
    If Me.Text1(3).Text = "" Then
        CadenaDesdeOtroForm = ""
    Else
        CadenaDesdeOtroForm = ExtensionSoportada(Text1(3).Text, True)    'EXTENSION
        If Len(CadenaDesdeOtroForm) > 10 Then SQL = SQL & "-" & CadenaDesdeOtroForm
    End If
    
    If SQL <> "" Then
        MsgBox "Falta: " & vbCrLf & SQL, vbExclamation
        Exit Sub
    End If
    
    
        
    
    'ACTUACIONES
    'whoobrascliactua (expediente,anoexp,IdActua,f_preact,extension,importe,horas,observa)
    If Opcion = 9 Then
        'NUEVO      NUEVO   NUEVO   NUEVO      NUEVO   NUEVO    NUEVO      NUEVO   NUEVO
        SQL = "anoexp = " & RecuperaValor(ExtraData2, 3) & " AND expediente "
        SQL = DevuelveDesdeBD(conAri, "max(IdActua)", "whoobrascliactua", SQL, RecuperaValor(ExtraData2, 2))
        If SQL = "" Then SQL = "0"
        idActua = Val(SQL) + 1
        
        'Monto la cadena NOMBRE PI
        If Me.Text1(3).Text <> "" Then
            'INSERTA ARCHIVO
            Aux = Format(RecuperaValor(ExtraData2, 2), "000000") & RecuperaValor(ExtraData2, 3) & Format(idActua, "000")
            Aux = "\" & RecuperaValor(ExtraData2, 1) & "\ACTUA\" & Aux & "." & CadenaDesdeOtroForm
            b = CopiaObraWHOSE(Aux, Text1(3).Text)
        Else
            'Como puede no adjuntar archivo a la
            b = True
        
        End If
        
        If b Then
            'whoobrasclipi( expediente anoexp idPI f_prePI fcontesta aceptado extension)
            SQL = "(" & RecuperaValor(ExtraData2, 2) & "," & RecuperaValor(ExtraData2, 3) & "," & idActua & "," & DBSet(txtFecha(4).Text, "F") & ",'" & UCase(CadenaDesdeOtroForm) & "'"
            SQL = SQL & "," & DBSet(Me.txtimporte(0).Text, "N", "S") & "," & DBSet(Me.txtimporte(1).Text, "N", "S") & "," & DBSet(Me.txtObra(2).Text, "T", "S") & ")"
            SQL = "INSERT INTO whoobrascliactua (expediente,anoexp,IdActua,f_preact,extension,horas,importe,observa) VALUES " & SQL
            If Not ejecutar(SQL, False) Then
                CadenaDesdeOtroForm = ""
                MsgBox "ERROR CRITICO insertando en BD", vbExclamation
                
            Else
                'Metemos los trabajadores
                InsertarTrabajadoresActuacion idActua
            
                'Lo metemos en el crm del cliente
                'scrmacciones(usuario,fechora,codclien,agente,codtraba,estado,tipo,medio,observaciones)
                SQL = "Expediente: " & Format(RecuperaValor(ExtraData2, 2), "00000") & "/" & RecuperaValor(ExtraData2, 3) & vbCrLf
                SQL = SQL & "ID: " & Format(idActua, "000") & "          Fecha actuación:" & Format(Me.txtFecha(4), "dd/mm/yyyy") & vbCrLf
                If Me.txtimporte(0).Text <> "" Or Me.txtimporte(1).Text <> "" Then
                    If Me.txtimporte(0).Text <> "" Then SQL = SQL & "Horas: " & Me.txtimporte(0).Text & "      "
                    If Me.txtimporte(1).Text <> "" Then SQL = SQL & "Importe: " & Me.txtimporte(1).Text
                    SQL = SQL & vbCrLf
                End If
                'Fichero
                If Text1(3).Text <> "" Then SQL = SQL & "Fichero: " & Text1(3).Text & " #Nom. origen" & vbCrLf
                SQL = SQL & vbCrLf & vbCrLf & txtObra(2).Text
                'Los trabajadores
                AñadirObservaCRMtrabajadores
                
                'Ya estan las observaciones
                'scrmacciones(usuario,fechora,codclien,agente,codtraba,estado,tipo,medio,observaciones)
                SQL = ",0,21,'Otros'," & DBSet(SQL, "T") & ")"
                CadenaDesdeOtroForm = PonerTrabajadorConectado("")
                If CadenaDesdeOtroForm = "" Then CadenaDesdeOtroForm = DevuelveDesdeBD(conAri, "codtraba", "straba", "1", "1") 'el primer trabajador
                SQL = "," & CadenaDesdeOtroForm & SQL
                CadenaDesdeOtroForm = DevuelveDesdeBD(conAri, "codagent", "sagent", "1", "1") 'el primer agente
                SQL = " VALUES ('" & vUsu.Login & "',now()," & RecuperaValor(ExtraData2, 1) & "," & CadenaDesdeOtroForm & SQL
                
                SQL = "INSERT INTO scrmacciones(usuario,fechora,codclien,agente,codtraba,estado,tipo,medio,observaciones) " & SQL
                If Not ejecutar(SQL, False) Then MsgBox "Error insertando en CRM", vbExclamation
                CadenaDesdeOtroForm = "OK"
            End If
            Unload Me
        Else
            CadenaDesdeOtroForm = ""
        End If

    Else
        'MODIFICAR
        
        idActua = RecuperaValor(ExtraData2, 4)
        
        SQL = "UPDATE whoobrascliactua SET f_preact=" & DBSet(txtFecha(4).Text, "F")
        SQL = SQL & ", horas=" & DBSet(Me.txtimporte(0).Text, "N", "S") & ", importe= " & DBSet(Me.txtimporte(1).Text, "N", "S")
        SQL = SQL & ",observa= " & DBSet(Me.txtObra(2).Text, "T", "S")
        SQL = SQL & " WHERE expediente=" & RecuperaValor(ExtraData2, 2) & " AND anoexp=" & RecuperaValor(ExtraData2, 3)
        SQL = SQL & " AND IdActua=" & idActua
        
        If Not ejecutar(SQL, True) Then
            CadenaDesdeOtroForm = ""
            MsgBox "ERROR CRITICO updateando en BD", vbExclamation
            
        Else
            
            'Metemos los trabajadores
            InsertarTrabajadoresActuacion idActua
            
            Set miRsAux = New ADODB.Recordset
            SQL = "Select * from scrmacciones WHERE tipo=21  AND codclien=" & RecuperaValor(ExtraData2, 1)
            SQL = SQL & " AND observaciones like '%Expediente: " & Format(RecuperaValor(ExtraData2, 2), "00000") & "/" & RecuperaValor(ExtraData2, 3)
            SQL = SQL & "%ID: " & Format(idActua, "000") & "%'"
            
            miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not miRsAux.EOF Then
                'HA encontrado la entrada en el CRM
                SQL = "Expediente: " & Format(RecuperaValor(ExtraData2, 2), "00000") & "/" & RecuperaValor(ExtraData2, 3) & vbCrLf
                SQL = SQL & "ID: " & Format(idActua, "000") & "          Fecha actuación:" & Format(Me.txtFecha(4), "dd/mm/yyyy") & vbCrLf
                If Me.txtimporte(0).Text <> "" Or Me.txtimporte(1).Text <> "" Then
                    If Me.txtimporte(0).Text <> "" Then SQL = SQL & "Horas: " & Me.txtimporte(0).Text & "      "
                    If Me.txtimporte(1).Text <> "" Then SQL = SQL & "Importe: " & Me.txtimporte(1).Text
                    SQL = SQL & vbCrLf
                End If
                'Fichero
                If Text1(3).Text <> "" Then SQL = SQL & "Fichero: " & Text1(3).Text & " #Nom. origen" & vbCrLf
                SQL = SQL & vbCrLf & vbCrLf & txtObra(2).Text
                'Los trabajadores
                AñadirObservaCRMtrabajadores
                
                SQL = "UPDATE scrmacciones SET observaciones=" & DBSet(SQL, "T", "S")
                SQL = SQL & " WHERE usuario =" & DBSet(miRsAux!Usuario, "T")
                SQL = SQL & " AND fechora =" & DBSet(miRsAux!fechora, "FH")
                SQL = SQL & " AND codclien =" & DBSet(miRsAux!codClien, "T")
                SQL = SQL & " AND Tipo =21"
                
                If Not ejecutar(SQL, False) Then MsgBox "Error actualizando en CRM", vbExclamation
                
            End If
            CadenaDesdeOtroForm = "OK"
            miRsAux.Close
            Unload Me
        End If
    
    End If
    
End Sub

Private Sub AñadirObservaCRMtrabajadores()
Dim J As Integer
    If Me.lwTrabajadores.ListItems.Count = 0 Then Exit Sub
    
    SQL = SQL & vbCrLf & vbCrLf & vbCrLf
    SQL = SQL & "Trabajadores: " & vbCrLf
    For J = 1 To lwTrabajadores.ListItems.Count
        SQL = SQL & "        .-" & lwTrabajadores.ListItems(J).Text & " - " & Me.lwTrabajadores.ListItems(J).SubItems(1)
    Next
    
End Sub


Private Sub cmdCancel_Click(Index As Integer)
    Unload Me
    If Opcion = 1 Then CadenaDesdeOtroForm = ""
    
    
    If Opcion > 5 Then CadenaDesdeOtroForm = ""
End Sub



Private Sub cmdContratoCliente_Click()
Dim codCli As Long
    '-----------------
    'Contrato cliente
    SQL = ""
    If Me.txtFecha(5).Text = "" Then SQL = SQL & "-Fecha presentacion" & vbCrLf
    If Opcion = 15 Then
        'nuevo
        If Text1(4).Text = "" Then
            SQL = SQL & "-Archivo adjunto" & vbCrLf
        Else
            CadenaDesdeOtroForm = ExtensionSoportada(Text1(4).Text, True)     'EXTENSION
            If Len(CadenaDesdeOtroForm) > 10 Then SQL = SQL & "-" & CadenaDesdeOtroForm & vbCrLf
        End If
    Else
        If Me.txtFecha(6).Text = "" Then SQL = SQL & "-Fecha contestación" & vbCrLf
        
    End If
    If SQL <> "" Then
        SQL = "Errores : " & vbCrLf & vbCrLf & SQL
        MsgBox SQL, vbExclamation
        Exit Sub
    End If
    
    
    'Nuevo
    If Opcion = 15 Then
        CadenaDesdeOtroForm = RecuperaValor(ExtraData2, 1)
        CadenaDesdeOtroForm = DevuelveDesdeBD(conAri, "max(idcontrato)", "whoexpedienteclicontrato", "codclien", CadenaDesdeOtroForm)
        NumRegElim = Val(CadenaDesdeOtroForm) + 1
        
        
        
        
        'OK vamos a meter el archivo en la BD y en la estrcutura de fichhero
        'whoobrascli expediente,anoexp,codclien,nombre,fecaltobra,descripcion,extension
        Screen.MousePointer = vbHourglass
        codCli = RecuperaValor(ExtraData2, 1)
        CadenaDesdeOtroForm = ExtensionSoportada(Text1(4).Text, False)
        SQL = "CON" & Format(NumRegElim, "00000")
        SQL = "\" & Format(codCli, "000000") & "\CONTRATO\" & SQL & "." & CadenaDesdeOtroForm
        If CopiaObraWHOSE(SQL, Text1(4).Text) Then
            SQL = "INSERT INTO whoexpedienteclicontrato(codclien,idcontrato,f_precont,extension) VALUES ("
            SQL = SQL & codCli & "," & NumRegElim & "," & DBSet(Me.txtFecha(5).Text, "F") & ","
            SQL = SQL & DBSet(CadenaDesdeOtroForm, "T") & ")"
            If Not ejecutar(SQL, False) Then
                SQL = "CON" & Format(NumRegElim, "00000")
                SQL = "\" & Format(ExtraData2, "000000") & "\CONTRATO\" & SQL & "." & CadenaDesdeOtroForm
                EliminarUnFichero SQL
            End If
        Else
            NumRegElim = 0
        End If
        Screen.MousePointer = vbDefault
        DesdeAceptar = True
        Unload Me
    Else
        
        SQL = "UPDATE whoexpedienteclicontrato SET f_precont = " & DBSet(Me.txtFecha(5).Text, "F")
        
        If Me.optContrato(0).Value Then
            SQL = SQL & ", f_rechazocon = " & DBSet(Me.txtFecha(6).Text, "F")
            SQL = SQL & ", f_aceptado = NULL"
        Else
            SQL = SQL & ", f_aceptado  = " & DBSet(Me.txtFecha(6).Text, "F")
            SQL = SQL & ", f_rechazocon = NULL"
        End If
        'codclien idcontrato
        SQL = SQL & " WHERE codclien =" & RecuperaValor(ExtraData2, 1)
        SQL = SQL & " AND idcontrato =" & RecuperaValor(ExtraData2, 2)
        If Not ejecutar(SQL, False) Then
            If Opcion = 15 Then
                SQL = "CON" & Format(NumRegElim, "00000")
                SQL = "\" & Format(ExtraData2, "000000") & "\CONTRATO\" & SQL & "." & CadenaDesdeOtroForm
                EliminarUnFichero SQL
            End If
        Else
            DesdeAceptar = True
            Unload Me
        End If
    
        
    
    
    
    
    End If
    
End Sub


Private Sub EliminarUnFichero(Nombre As String)
    On Error Resume Next
    Kill Nombre
    If Err.Number <> 0 Then
        MsgBox "El programa continuará. Avise soporte tecnico", vbExclamation
        Err.Clear
    End If
End Sub

Private Sub cmdNuevaObra_Click()
    SQL = ""
    If Me.Text1(1).Text = "" Then SQL = SQL & vbCrLf & "-Archivo"
    If Me.txtFecha(2).Text = "" Then SQL = SQL & vbCrLf & "-Fecha"
    If Me.txtObra(0).Text = "" Then SQL = SQL & vbCrLf & "-Descripcion"
    If Me.cboPresentador.ListIndex < 0 Then SQL = SQL & vbCrLf & "-Tipo presentador"
        
    
    If SQL <> "" Then
        MsgBox "Faltan campos:" & vbCrLf & SQL, vbExclamation
        Exit Sub
    End If


    If Opcion = 4 Then
        
        NuevaObra
    Else
        ModificarObra
    End If
    If NumRegElim > 0 Then
        
        DesdeAceptar = True
        CadenaDesdeOtroForm = NumRegElim & "|" & Year(Me.txtFecha(2).Text) & "|"
        Unload Me
    End If
End Sub

Private Sub NuevaObra()
    
        
        
    CadenaDesdeOtroForm = DevuelveDesdeBD(conAri, "max(expediente)", "whoobrascli", "anoexp", CStr(Year(CDate(Me.txtFecha(2).Text))))
    NumRegElim = Val(CadenaDesdeOtroForm) + 1
    CadenaDesdeOtroForm = Format(NumRegElim, "000000") & " / " & CStr(Year(CDate(Me.txtFecha(2).Text)))
    
    
    CadenaDesdeOtroForm = "Va a crear el expediente :" & vbCrLf & CadenaDesdeOtroForm
    If MsgBox(CadenaDesdeOtroForm, vbQuestion + vbYesNo) <> vbYes Then Exit Sub

    
    'OK vamos a meter el archivo en la BD y en la estrcutura de fichhero
    'whoobrascli expediente,anoexp,codclien,nombre,fecaltobra,descripcion,extension
    Screen.MousePointer = vbHourglass
    CadenaDesdeOtroForm = ExtensionSoportada(Text1(1).Text, False)
    SQL = Format(NumRegElim, "000000") & CStr(Year(CDate(txtFecha(2).Text)))
    SQL = "\" & Format(ExtraData2, "000000") & "\OBRA\" & SQL & "." & CadenaDesdeOtroForm
    If CopiaObraWHOSE(SQL, Text1(1).Text) Then
        SQL = "INSERT INTO whoobrascli (expediente,anoexp,codclien,nombre,fecaltobra,descripcion,extension,tipoPresentador) VALUES ("
        SQL = SQL & NumRegElim & "," & Year(Me.txtFecha(2).Text) & "," & ExtraData2 & "," & DBSet(Me.txtObra(0).Text, "T")
        SQL = SQL & "," & DBSet(Me.txtFecha(2).Text, "F") & "," & DBSet(Me.txtObra(1).Text, "T", "S") & ",'" & CadenaDesdeOtroForm
        SQL = SQL & "'," & Me.cboPresentador.ItemData(Me.cboPresentador.ListIndex) & ")"
        If Not ejecutar(SQL, True) Then
            MsgBox "Error insertando en BD. Llame a soporte tecnico " & vbCrLf & SQL, vbExclamation
            NumRegElim = 0
'        Else
'            'Ya esta creada
'            CadenaDesdeOtroForm = NumRegElim
'            Unload Me
        End If
    Else
        NumRegElim = 0
    End If
    Screen.MousePointer = vbDefault
    
End Sub


Private Sub ModificarObra()

    NumRegElim = 0
    SQL = "UPDATE whoobrascli SET fecaltobra = " & DBSet(txtFecha(2).Text, "F")
    SQL = SQL & ", nombre =  " & DBSet(txtObra(0).Text, "T")
    SQL = SQL & ", Descripcion =  " & DBSet(txtObra(1).Text, "T")
    SQL = SQL & ", tipoPresentador=" & Me.cboPresentador.ItemData(Me.cboPresentador.ListIndex)
    
    SQL = SQL & " Where codclien = " & RecuperaValor(ExtraData2, 1)
    SQL = SQL & " AND anoexp = " & RecuperaValor(ExtraData2, 3) & " AND expediente = " & RecuperaValor(ExtraData2, 2)
    If ejecutar(SQL, False) Then NumRegElim = RecuperaValor(ExtraData2, 2)

End Sub


Private Sub cmdNuevoDesdeExpediente_Click()
Dim Aux As String
        
    SQL = ""
    If Me.Text1(2).Text = "" Then
        SQL = SQL & "-Nombre fichero vacio"
    Else
        If Dir(Text1(2).Text, vbArchive) = "" Then
            SQL = SQL & "-No existe el fichero"
        Else
            CadenaDesdeOtroForm = ExtensionSoportada(Text1(2).Text, True)   'EXTENSION
            If Len(CadenaDesdeOtroForm) > 10 Then SQL = SQL & "-Extension no soportada"
        End If
    End If
    If Me.txtFecha(3).Text = "" Then
        SQL = SQL & vbCrLf & "-Fecha "
    Else
        If Opcion = 6 Then
            If CDate(RecuperaValor(ExtraData2, 4)) > CDate(txtFecha(3).Text) Then SQL = SQL & vbCrLf & "-Fecha debe ser mayor que " & RecuperaValor(ExtraData2, 4)
        Else
            If Me.cboEGDA.ListIndex < 0 Then SQL = SQL & vbCrLf & "-Seleccione una empresa gestora de derechos"
            If Not IsDate(txtFecha(3).Text) Then SQL = SQL & vbCrLf & "-Fecha incorrecta"
        End If
        
    End If
    If SQL <> "" Then
        MsgBox "Errores: " & vbCrLf & SQL, vbExclamation
        CadenaDesdeOtroForm = ""
        Exit Sub
    End If
    
    
    'Para las SGED, es mas complicado. Debemos ver que para esa EGDA o
    '   .- No hay documento dado de alta
    '   .- Si hay uno, no puede o estar aceptado
    If Opcion = 7 Then
        SQL = "select * from whoobrasclisgd where expediente=" & RecuperaValor(ExtraData2, 2) & " and  anoexp = " & RecuperaValor(ExtraData2, 3)
        SQL = SQL & " AND SGD = " & Me.cboEGDA.ItemData(cboEGDA.ListIndex) & " order by f_preSGD desc,fcontesta desc"
        Set miRsAux = Nothing
        Set miRsAux = New ADODB.Recordset
        miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If miRsAux.EOF Then
            'OK. PERFECTo. Es el primer documento presentado a esta compañia
            SQL = ""
        Else
            
            If IsNull(miRsAux!fcontesta) Then
                SQL = "Falta contestar presentacion anterior"
            Else
                If miRsAux!aceptado = 1 Then
                    SQL = "Presentacion anterior aceptada"
                Else
                
                    'La fecha no puede ser menor que la fecha de contestacion y presentacion anterior
                    If CDate(txtFecha(3).Text) < miRsAux!f_preSGD Then
                        SQL = "No puede ser menor que una presentacion anterior: " & miRsAux!f_preSGD
                    Else
                        If CDate(txtFecha(3).Text) < miRsAux!fcontesta Then
                            SQL = "No puede ser menor que una contestacion anterior: " & miRsAux!fcontesta
                        Else
                            SQL = "" 'OK. Otra presentacion
                        End If
                    End If
                    
                End If
            End If
        End If
        miRsAux.Close
        If SQL <> "" Then
            MsgBox SQL, vbExclamation
            Exit Sub
        End If
    
        
    End If
    
    If Opcion = 6 Then
        'Presentacion en propiedad intelectual
        Aux = "anoexp =" & RecuperaValor(ExtraData2, 3) & " AND expediente "
        SQL = DevuelveDesdeBD(conAri, "max(idPI)", "whoobrasclipi", Aux, RecuperaValor(ExtraData2, 2))
        If SQL = "" Then SQL = "0"
        SQL = Val(SQL) + 1
        
        'Monto la cadena NOMBRE PI
        Aux = Format(RecuperaValor(ExtraData2, 2), "000000") & RecuperaValor(ExtraData2, 3) & Format(SQL, "000")
        Aux = "\" & RecuperaValor(ExtraData2, 1) & "\PI\" & Aux & "." & CadenaDesdeOtroForm
        'CopiaObraWHOSE
        
        If CopiaObraWHOSE(Aux, Text1(2).Text) Then
            'whoobrasclipi( expediente anoexp idPI f_prePI fcontesta aceptado extension)
            SQL = "(" & RecuperaValor(ExtraData2, 2) & "," & RecuperaValor(ExtraData2, 3) & "," & SQL & "," & DBSet(txtFecha(3).Text, "F") & ",'" & UCase(CadenaDesdeOtroForm) & "')"
            SQL = "insert into whoobrasclipi( expediente ,anoexp ,idPI ,f_prePI ,extension) VALUES " & SQL
            If Not ejecutar(SQL, False) Then
                CadenaDesdeOtroForm = ""
                MsgBox "ERROR CRITICO insertando en BD", vbExclamation
            End If
            Unload Me
        Else
            CadenaDesdeOtroForm = ""
        End If
    
    Else
        'SOCIENDADE GESTION DERECHOS AUTOR
        Aux = "anoexp =" & RecuperaValor(ExtraData2, 3) & " AND SGD = " & Me.cboEGDA.ItemData(cboEGDA.ListIndex) & " AND expediente"
        SQL = DevuelveDesdeBD(conAri, "max(IdPres)", "whoobrasclisgd", Aux, RecuperaValor(ExtraData2, 2))
        If SQL = "" Then SQL = "0"
        SQL = Val(SQL) + 1
        
        'Monto la cadena NOMBRE SGDA
        Aux = Format(RecuperaValor(ExtraData2, 2), "000000") & RecuperaValor(ExtraData2, 3) & Format(SQL, "000") & Format(Me.cboEGDA.ItemData(cboEGDA.ListIndex), "00")
        Aux = "\" & RecuperaValor(ExtraData2, 1) & "\EGD\" & Aux & "." & CadenaDesdeOtroForm
        'CopiaObraWHOSE
        
        If CopiaObraWHOSE(Aux, Text1(2).Text) Then
            'whoobrasclisgd(expediente,anoexp,SGD,IdPres,f_preSGD,extension)
            SQL = "(" & RecuperaValor(ExtraData2, 2) & "," & RecuperaValor(ExtraData2, 3) & "," & Me.cboEGDA.ItemData(cboEGDA.ListIndex) & "," & SQL
            SQL = SQL & "," & DBSet(txtFecha(3).Text, "F") & ",'" & UCase(CadenaDesdeOtroForm) & "')"
            SQL = "insert into whoobrasclisgd(expediente,anoexp,SGD,IdPres,f_preSGD,extension) VALUES " & SQL
            If Not ejecutar(SQL, False) Then
                CadenaDesdeOtroForm = ""
                MsgBox "ERROR CRITICO insertando en BD", vbExclamation
            End If
            Unload Me
        Else
            CadenaDesdeOtroForm = ""
        End If
    
    End If
    
End Sub

Private Sub cmdPasarCliente_Click()
    Dim b As Boolean
    
    'Pasar de cliente potencial a cliente
    SQL = ""
    If Me.cboFP.ListIndex < 0 Then SQL = "Seleccion forma de pago"
        
    
    If SQL <> "" Then
        MsgBox SQL, vbExclamation
        Exit Sub
    End If
    
    
    'Seguro que desea....
    
    
    
    'NUmregelim
    
    SQL = DevuelveDesdeBD(conAri, "max(codclien)", "sclien", "1", "1")
    NumRegElim = Val(SQL) + 1
    
    
    'Vere si ya existe la estructura
    If YaExisteExtructuraCliente(NumRegElim) Then
        MsgBox "No puede continuar. Ya existe la estructura para el cliente: " & NumRegElim, vbExclamation
        Exit Sub
    End If
    
    
    'Comprobamos la cuenta, que no exista
    SQL = "43." & NumRegElim
    SQL = RellenaCodigoCuenta(SQL)
    CadenaDesdeOtroForm = DevuelveDesdeBD(conConta, "codmacta", "cuentas", "codmacta", SQL, "T")
    If CadenaDesdeOtroForm <> "" Then
        MsgBox "Ya existe la cuenta contable: " & SQL
        Exit Sub
    End If
        
        
    If MsgBox("Seguro que desea crear el cliente?", vbYesNoCancel + vbQuestion) <> vbYes Then Exit Sub
    
    'GEneramos la estructura
    Screen.MousePointer = vbHourglass
    b = False
    If TratarExtructuraClienteConArchivos(True, CLng(ExtraData2), NumRegElim) Then
        conn.BeginTrans
        If InsertarClienteDesdePotencial Then
            conn.CommitTrans
            b = True
            
            Volver_A_Cargar_Datos = True  'para que refresque los datos en el LW de seleccion
            
            'Borramos la estrucutura anterior de potenciales
            TratarExtructuraClienteConArchivos False, CLng(ExtraData2), 0
        Else
            conn.RollbackTrans
        End If
    End If
    Screen.MousePointer = vbDefault
    
    If b Then Unload Me
        
End Sub


Private Function InsertarClienteDesdePotencial() As Boolean

On Error GoTo eInsertarClienteDesdePotencial

    InsertarClienteDesdePotencial = False

    SQL = "INSERT INTO sclien(codclien,nomclien,nomcomer,domclien,codpobla,pobclien,proclien,nifclien,wwwclien,fechamov,fechaalt,"
    SQL = SQL & " codactiv , CodEnvio, codzonas, codrutas, perclie1, telclie1, faxclie1, maiclie1, perclie2, telclie2, faxclie2, maiclie2, observac,"
    SQL = SQL & " codagent,codforpa,codmacta,clivario,tipoiva,tipofact,albarcon,periodof,codtarif,dtoppago,dtognral,promocio,codsitua,referobl,cliabono,pasclien)"
    SQL = SQL & " SELECT " & NumRegElim & ",nomclien,nomcomer,domclien,codpobla,pobclien,proclien,nifclien,wwwclien,NULL," & DBSet(Now, "F")
    SQL = SQL & " ,codactiv , CodEnvio, codzonas, codrutas, perclie1, telclie1, faxclie1, maiclie1, perclie2, telclie2, faxclie2, maiclie2, observac,"
    CadenaDesdeOtroForm = DevuelveDesdeBD(conAri, "min(codagent)", "sagent", "1", "1")
    'Cad = Cad & " codagent,codforpa,codmacta,clivario,tipoiva,tipofact,albarcon,periodof,numrepet,
    SQL = SQL & CadenaDesdeOtroForm & "," & Me.cboFP.ItemData(cboFP.ListIndex)
    'Codmacta
    CadenaDesdeOtroForm = "43." & NumRegElim
    CadenaDesdeOtroForm = RellenaCodigoCuenta(CadenaDesdeOtroForm)
    SQL = SQL & ",'" & CadenaDesdeOtroForm & "',0," & Me.cboTipoIVA.ItemData(cboTipoIVA.ListIndex) & ","
    SQL = SQL & Me.cboFacturacion.ItemData(cboFacturacion.ListIndex) & "," & Me.cboAlbaran.ItemData(cboAlbaran.ListIndex) & ",1,"
    'codtarif,dtoppago,dtognral,promocio,codsitua,referobl,cliabono,pasclien)"
    CadenaDesdeOtroForm = DevuelveDesdeBD(conAri, "min(codsitua)", "ssitua", "1", "1")
    SQL = SQL & vParamAplic.CodTarifa & ",0,0,1," & CadenaDesdeOtroForm & ",0,1,nifclien "
    SQL = SQL & " FROM sclipot WHERE codclien = " & ExtraData2
    conn.Execute SQL
    
    'Pasamos tb personas de contacto
    '                     codclien
    CadenaDesdeOtroForm = ",id,nombre,dpto,cargo,telefono,ext,maidirec,movil,observa"
    SQL = "INSERT INTO scliendp(codclien" & CadenaDesdeOtroForm & ") SELECT " & NumRegElim & CadenaDesdeOtroForm
    SQL = SQL & " FROM sclipotdp WHERE codclien=" & ExtraData2
    conn.Execute SQL
    
    
    
    'Pasaremos los procesos
    'EXPEDIENTE CLIENTES
    CadenaDesdeOtroForm = ",fecAceptPropComer,fecAceptContrato "
    SQL = "INSERT INTO whoexpedienteCLI(codclien" & CadenaDesdeOtroForm & ") SELECT " & NumRegElim & CadenaDesdeOtroForm
    SQL = SQL & " FROM whoexpedientepot WHERE codclien=" & ExtraData2
    conn.Execute SQL
    
    
    
    'CONTRATO
    CadenaDesdeOtroForm = ", idcontrato, f_precont, f_rechazocon, extension"
    SQL = "INSERT INTO whoexpedienteCLIcontrato(codclien" & CadenaDesdeOtroForm & ") SELECT " & NumRegElim & CadenaDesdeOtroForm
    SQL = SQL & " FROM whoexpedientepotcontrato WHERE codclien=" & ExtraData2
    conn.Execute SQL
    
    'PROPUESTAS
    CadenaDesdeOtroForm = ",idPropComer,f_preprop,f_rechazoprop,extension "
    SQL = "INSERT INTO whoexpedienteCLIprocomer(codclien" & CadenaDesdeOtroForm & ") SELECT " & NumRegElim & CadenaDesdeOtroForm
    SQL = SQL & " FROM whoexpedientepotprocomer WHERE codclien=" & ExtraData2
    conn.Execute SQL
    
    
    'fECHA DE ALTA COMO cliente potencial
    CadenaDesdeOtroForm = DevuelveDesdeBD(conAri, "fechaalt", "sclipot", "codclien", ExtraData2)
    If CadenaDesdeOtroForm <> "" Then
        SQL = "UPDATE whoexpedientecli set  fecAltaPotencial=" & DBSet(CadenaDesdeOtroForm, "F") & " WHERE codclien =" & NumRegElim
        ejecutar SQL, False
    End If
    
    'DELETES
    conn.Execute "DELETE FROM whoexpedientepot WHERE codclien=" & ExtraData2
    conn.Execute "DELETE FROM whoexpedientepotcontrato WHERE codclien=" & ExtraData2
    conn.Execute "DELETE FROM whoexpedientepotprocomer WHERE codclien=" & ExtraData2
    conn.Execute "DELETE FROM sclipotdp WHERE codclien=" & ExtraData2
    conn.Execute "DELETE FROM sclipot WHERE codclien=" & ExtraData2
    
    
    '
    SQL = "43." & NumRegElim
    SQL = RellenaCodigoCuenta(SQL)
    If Not InsertarCuentaCble(SQL, CStr(NumRegElim)) Then MsgBox "Se ha producido un error insertando la cuenta: " & Text1(1).Text & ". El proceso continua. Avise soporte técnico", vbExclamation
        


    'JULIO 2014
    'Puede meter los contratos que deseen
    'Para ello tenemos que guardar la fecha de aceptacion del contrato cuando lo pasamos a CLIENTE
    SQL = DevuelveDesdeBD(conAri, "fecAceptContrato", "whoexpedientecli", "codclien", CStr(NumRegElim))
    If SQL <> "" Then
        'Siempre deberia ser <>''
        SQL = "UPDATE whoexpedienteclicontrato SET f_aceptado=" & DBSet(SQL, "F")
        SQL = SQL & " WHERE codclien = " & NumRegElim & " AND f_rechazocon is null"
        ejecutar SQL, False
    End If
    
 


    InsertarClienteDesdePotencial = True
    
    Exit Function
eInsertarClienteDesdePotencial:
    MuestraError Err.Number, SQL, Err.Description
    
End Function


Private Sub cmdRechazar_Click()
    If Trim(Me.txtFecha(1).Text) = "" Then Exit Sub
    
    SQL = ""
    If Opcion = 13 Or Opcion = 14 Then
        If CDate(Me.txtFecha(1).Text) < CDate(RecuperaValor(ExtraData2, 1)) Then SQL = RecuperaValor(ExtraData2, 1)
    Else
        If CDate(Me.txtFecha(1).Text) < CDate(RecuperaValor(ExtraData2, 2)) Then SQL = RecuperaValor(ExtraData2, 2)
    End If
    
    If SQL <> "" Then
        MsgBox "No puede ser menor a " & SQL, vbExclamation
        Exit Sub
    End If
    DesdeAceptar = True
    CadenaDesdeOtroForm = txtFecha(1).Text
    Unload Me
End Sub

Private Sub cmdTitularidad_Click()
            
    If txtTitularidad.Text = "" Or Me.cboTitularidad.Text = "" Then
        MsgBox "Campos obligados", vbExclamation
        Exit Sub
    End If
           
    
    ''whotitularidadcli (   codclien,idTitularidad ,nombre ,relacion )
    If Opcion = 11 Then
        SQL = DevuelveDesdeBD(conAri, "max(idTitularidad)", "whotitularidadcli", "codclien", RecuperaValor(ExtraData2, 1))
        If SQL = "" Then SQL = "0"
        SQL = Val(SQL) + 1
        SQL = "INSERT INTO whotitularidadcli (codclien,idTitularidad ,nombre ,relacion ) VALUES (" & RecuperaValor(ExtraData2, 1) & "," & SQL
        SQL = SQL & "," & DBSet(Me.txtTitularidad.Text, "T") & "," & DBSet(Me.cboTitularidad.Text, "T") & ")"
        
    Else
        'UPDATE
        SQL = "UPDATE whotitularidadcli SET nombre = " & DBSet(Me.txtTitularidad.Text, "T") & ",relacion = " & DBSet(Me.cboTitularidad.Text, "T")
        SQL = SQL & " WHERE codclien =" & RecuperaValor(ExtraData2, 1) & " AND idTitularidad =" & RecuperaValor(ExtraData2, 4)
        
    End If
    If ejecutar(SQL, False) Then
        CadenaDesdeOtroForm = "OK"
        Unload Me
    Else
        CadenaDesdeOtroForm = ""
    End If
    
End Sub

Private Sub cmdVerClientePotencial_Click()
  
    frmFacClienPot.DatosADevolverBusqueda = ExtraData2
    frmFacClienPot.Show vbModal

    CargaDatosCliente False
End Sub

Private Sub Form_Activate()
    
    If PrimVez Then
        PrimVez = False
        If Opcion = 3 Then CargaDatosCliente True
        If Opcion = 5 Then CargaDatosObra
        If Opcion = 10 Then CargaDatosActuacion
        If Opcion = 11 Or Opcion = 12 Then PonTitularidad
        
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim Indice As Integer
Dim Propu As Byte

    PrimVez = True
    DesdeAceptar = False
    limpiar Me
    Me.Icon = frmPPalWhose.Icon
    FrameInsertarDocumento.visible = False
    FrameFechaRechazo.visible = False
    FrameNuevaObra.visible = False
    FrameActuacion.visible = False
    FrameTitulares.visible = False
    FrameContratoCli.visible = False
    
    Indice = Opcion
    Select Case Opcion
    Case 0
        Propu = RecuperaValor(ExtraData2, 3) = 1
        If Propu Then
            Caption = "Nueva propuesta comercial"
        Else
            Caption = "Nuevo contrato"
        End If
        Me.txtFecha(0).Text = Format(Now - 1, "dd/mm/yyyy")
        PonerFrameVisible FrameInsertarDocumento
    Case 1, 2
        Propu = RecuperaValor(ExtraData2, 3) = 1
        If Opcion = 1 Then
            If Propu Then
                Caption = "Rechazo propuesta comercial"
            Else
                Caption = "Rechazo contrato"
            End If
            lblPedirfecha.Caption = "Fecha rechazo"
            lblPedirfecha.ForeColor = vbRed
        Else
            If Propu Then
                Caption = "Rechazo propuesta comercial"
            Else
                Caption = "Rechazo contrato"
            End If
            lblPedirfecha.Caption = "Fecha aceptacion"
            lblPedirfecha.ForeColor = vbBlue
        End If
        Me.txtFecha(1).Text = Format(Now, "dd/mm/yyyy")
        PonerFrameVisible FrameFechaRechazo
        Indice = 1
    
    
    
    Case 3
        Caption = "Traspasar"
        PonerFrameVisible FramePasarAclientes
                
        'Combos
        CargarComboAlbaran
        CargarComboTipoIVA
        CargarComboFacturacion
        
        CargarCombo_Tabla cboFP, "sforpa", "codforpa", "nomforpa"
        
        
    Case 4, 5
        
        Caption = "EXPEDIENTE"
        PonerFrameVisible FrameNuevaObra
        
        Label2(1).Caption = "obra / expediente"
        If Opcion = 4 Then
            Label2(1).Caption = "Nueva " & Label2(1).Caption
            Label2(1).ForeColor = vbGreen
        Else
            Label2(1).Caption = "Modificar " & Label2(1).Caption
            Label2(1).ForeColor = vbRed
        End If
        imgDir(1).visible = Opcion = 4
        Me.Text1(1).Enabled = Opcion = 4
        Indice = 4
        CargarCombo_Tabla Me.cboPresentador, "whorelacioncliobra", "idRelacion", "desRelacion"
    Case 6, 7
        Caption = "Expediente: " & RecuperaValor(ExtraData2, 1) & " -- " & Format(RecuperaValor(ExtraData2, 2), "0000") & " / " & RecuperaValor(ExtraData2, 3)
        If Opcion = 6 Then
            Label2(2).Caption = "Propiedad intelectual"
        Else
            Label2(2).Caption = "Sociedades gestión derechos autor"
            CargaEGDA
        End If
        PonerFrameVisible FrameDocCliente
        
        Indice = 6
        
        Label1(13).visible = Opcion = 7
        Me.cboEGDA.visible = Opcion = 7
    Case 9, 10
        'ACtuaciones
        PonerFrameVisible FrameActuacion
        
        imgDir(3).visible = Opcion = 9   'de momento NO dejo insertar
        BloquearTxt Me.Text1(3), Opcion = 10
        If Opcion = 9 Then Me.Text1(3).Locked = True
        
        Caption = "Actuacion"
        
        
        lwTrabajadores.Tag = 0 'SI hay cambios
        Indice = 9
        
    Case 11, 12
        
        PonerFrameVisible FrameTitulares
        Indice = 11
        
    Case 13, 14
        If Opcion = 13 Then
            Caption = "Rechazo " & RecuperaValor(ExtraData2, 2)
            lblPedirfecha.Caption = "Fecha rechazo"
            lblPedirfecha.ForeColor = vbBlack
        Else
            Caption = "ACEPTACION " & RecuperaValor(ExtraData2, 2)
            lblPedirfecha.Caption = "Fecha aceptacion"
            lblPedirfecha.ForeColor = vbBlack
        End If
        Me.txtFecha(1).Text = Format(Now, "dd/mm/yyyy")
        PonerFrameVisible FrameFechaRechazo
        Indice = 1
        
        
    Case 15, 16
        Caption = "Contrato"
        FrameModContrato.BorderStyle = 0
        
        If Opcion = 15 Then
            txtFecha(5).Text = Format(Now, "dd/mm/yyyy")
            lblTitulo(0).Caption = "Nuevo contrato"
        Else
            lblTitulo(0).Caption = "Modificar contrato"
            
            'Fecha presentacion
            txtFecha(5).Text = RecuperaValor(ExtraData2, 5)
            txtFecha(6).Text = RecuperaValor(ExtraData2, 4)   'F rechaxo
            SQL = RecuperaValor(ExtraData2, 3)
            Me.optContrato(CInt(Val(SQL))).Value = True
        End If
        FrameModContrato.visible = Opcion = 16
        
        PonerFrameVisible FrameContratoCli
        Indice = 2
    End Select
    Me.cmdCancel(Indice).Cancel = True
End Sub

Private Sub CargaEGDA()
    CargarCombo_Tabla cboEGDA, "whoEGDA ", "idempresa", "NombreSGD ", ""
    cboEGDA.ListIndex = -1
End Sub

Private Sub PonerFrameVisible(ByRef F As Frame)
    F.Top = 0
    F.Left = 120
    F.visible = True
    
    Height = F.Height + 420
    Width = F.Width + 240
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not DesdeAceptar Then
        If Opcion = 4 Or Opcion = 5 Then CadenaDesdeOtroForm = ""
        If Opcion = 13 Or Opcion = 14 Or Opcion = 15 Then CadenaDesdeOtroForm = ""
    End If
End Sub

Private Sub frmC_Selec(vFecha As Date)
    SQL = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmT_DatoSeleccionado(CadenaSeleccion As String)
    SQL = CadenaSeleccion
End Sub

Private Sub Image3_Click(Index As Integer)
    If Index < 2 Then
        'Trabajadores
        If Index = 0 Then
            SQL = ""
'            Set frmT = New frmAdmTrabajadores
'            frmT.DatosADevolverBusqueda = "0|1|" 'Poner Modo Busqueda
'            frmT.Show vbModal
            Set frmT = New frmBasico2
            AyudaTrabajadores frmT, Text1(Index)
            Set frmT = Nothing
            
            If SQL <> "" Then InsertaEnLwTrabajadores
                
        
        Else
            If Me.lwTrabajadores.SelectedItem Is Nothing Then Exit Sub
            
            SQL = "Desea quitar de la actuacion al trabajador: " & Me.lwTrabajadores.SelectedItem.SubItems(1) & "?"
            If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then Exit Sub
            
            SQL = "DELETE"
        End If
    End If
    Me.lwTrabajadores.Tag = 1
End Sub

Private Sub imgDir_Click(Index As Integer)
    
    cd1.FileName = ""
    cd1.InitDir = "c:\"
    cd1.CancelError = False
    If Index = 0 Then
        cd1.Filter = "Adobe PDF (*.pdf)|*.pdf|MS Office WORD (*.doc)|*.doc|MS Office WORD 2007|*.docx"
        cd1.FilterIndex = 0
    End If
    cd1.ShowOpen
    If cd1.FileName = "" Then Exit Sub
    
    Text1(Index).Text = cd1.FileName
    
End Sub

Private Sub imgFecha_Click(Index As Integer)
   SQL = ""
   Set frmC = New frmCal
   frmC.Fecha = Now
   If txtFecha(Index).Text <> "" Then
        If IsDate(txtFecha(Index).Text) Then frmC.Fecha = CDate(txtFecha(Index).Text)
   End If
   frmC.Show vbModal
   Set frmC = Nothing
   If SQL <> "" Then txtFecha(Index).Text = SQL
End Sub



Private Sub Text1_OLEDragDrop(Index As Integer, data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Dim V
    NumRegElim = 0
    For Each V In data.Files
        Debug.Print V
        Text1(Index).Text = V
        NumRegElim = NumRegElim + 1
        
    Next V
    If NumRegElim > 1 Then MsgBox "Solo se contempla un archivo", vbExclamation
        
    'Cargariamos el Visor
    
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
        SQL = txtFecha(Index).Text
        If EsFechaOK(SQL) Then
            txtFecha(Index).Text = SQL
        Else
            MsgBox "Fecha con formato incorrecto: " & txtFecha(Index).Text, vbExclamation
            txtFecha(Index).Text = ""
            PonerFoco txtFecha(Index)
        End If
    End If
    
End Sub

'*********************************************************************
'
' Pasar a clientes
'

Private Sub CargaDatosCliente(EsInicio As Boolean)
    
    Me.cmdPasarCliente.visible = False
    SQL = "Select * from sclipot WHERE codclien = " & ExtraData2
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If miRsAux.EOF Then
        MsgBox "Error obteniendo cliente potencial", vbExclamation
        SQL = ""
    Else
        'OK. DATOS OBLIGATORIOS
        SQL = ""
        Me.Text2.Text = miRsAux!NomClien
        CadenaDesdeOtroForm = "nomcomer|domclien|codpobla|pobclien|proclien|nifclien|codactiv|codenvio|codzonas|codrutas|"
        For NumRegElim = 1 To 10
            ExtraData2 = RecuperaValor(CadenaDesdeOtroForm, CInt(NumRegElim))
            If IsNull(miRsAux.Fields(ExtraData2)) Then SQL = SQL & vbCrLf & "-" & ExtraData2
        Next
        ExtraData2 = miRsAux!codClien 'reestablezco
    End If
    miRsAux.Close
    Set miRsAux = Nothing
    
    If SQL <> "" Then
        If EsInicio Then
            cmdVerClientePotencial_Click
            Exit Sub
        Else
            MsgBox "Campos obligatorios" & vbCrLf & SQL, vbExclamation
        End If
    Else
        Me.cmdPasarCliente.visible = True
    End If

    
End Sub



Private Sub CargarComboAlbaran()
'### Combo Valorar Albaran con
'Cargaremos el combo, o bien desde una tabla o con valores fijos o como
'se quiera, la cuestion es cargarlo
' El estilo del combo debe de ser 2 - Dropdown List
' Si queremos que este ordenado, o lo ordenamos por la sentencia sql
' o marcamos la opcion sorted del combo
'0-Todo, 1-Cantidad y Precio, 2-Cantidad

    cboAlbaran.Clear
    cboAlbaran.AddItem "Todo"
    cboAlbaran.ItemData(cboAlbaran.NewIndex) = 0

    cboAlbaran.AddItem "Cantidad y Precio"
    cboAlbaran.ItemData(cboAlbaran.NewIndex) = 1

    cboAlbaran.AddItem "Cantidad"
    cboAlbaran.ItemData(cboAlbaran.NewIndex) = 2
    cboAlbaran.ListIndex = 0
End Sub


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
    
    cboFacturacion.ListIndex = 1
End Sub


Private Sub CargarComboTipoIVA()
'### Combo Tipo de IVA a Aplicar
'Cargaremos el combo, o bien desde una tabla o con valores fijos o como
'se quiera, la cuestion es cargarlo
' El estilo del combo debe de ser 2 - Dropdown List
' Si queremos que este ordenado, o lo ordenamos por la sentencia sql
' o marcamos la opcion sorted del combo
'0-Normal, 1-Con Recargo de Equivalencia, 2-Exento de IVA

    Me.cboTipoIVA.Clear
    cboTipoIVA.AddItem "Normal"
    cboTipoIVA.ItemData(cboTipoIVA.NewIndex) = 0

    cboTipoIVA.AddItem "Recargo Equivalencia"
    cboTipoIVA.ItemData(cboTipoIVA.NewIndex) = 1

    cboTipoIVA.AddItem "Exento de IVA"
    cboTipoIVA.ItemData(cboTipoIVA.NewIndex) = 2

    cboTipoIVA.AddItem "Intracomunitario"
    cboTipoIVA.ItemData(cboTipoIVA.NewIndex) = 3
    
    'Junio 2012 Reducido
    cboTipoIVA.AddItem "Reducido"
    cboTipoIVA.ItemData(cboTipoIVA.NewIndex) = 4
    cboTipoIVA.ListIndex = 0

End Sub




Private Sub CargaDatosObra()
    
    
    SQL = "Select * from whoobrascli WHERE codclien = " & RecuperaValor(ExtraData2, 1)
    SQL = SQL & " AND anoexp = " & RecuperaValor(ExtraData2, 3) & " AND expediente = " & RecuperaValor(ExtraData2, 2)
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If miRsAux.EOF Then
        MsgBox "Error obteniendo datos expediente", vbExclamation
    Else
        'OK. DATOS OBLIGATORIOS
        Me.txtFecha(2).Text = miRsAux!fecaltobra
        Me.txtObra(0).Text = miRsAux!Nombre
        Me.Text1(1).Text = Format(miRsAux!expediente, "000000") & miRsAux!anoexp & "." & miRsAux!Extension
        Me.txtObra(1).Text = DBLet(miRsAux!Descripcion, "T")
        
        SituarCombo Me.cboPresentador, DBLet(miRsAux!tipoPresentador, "N")
    End If
    miRsAux.Close
    Set miRsAux = Nothing
End Sub


Private Sub CargaDatosActuacion()
    
    
    SQL = "Select * from whoobrascliactua WHERE idactua = " & RecuperaValor(ExtraData2, 4)
    SQL = SQL & " AND anoexp = " & RecuperaValor(ExtraData2, 3) & " AND expediente = " & RecuperaValor(ExtraData2, 2)
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If miRsAux.EOF Then
        MsgBox "Error obteniendo datos expediente", vbExclamation
    Else
        'OK. DATOS OBLIGATORIOS
        ''whoobrascliactua (expediente,anoexp,IdActua,f_preact,extension,importe,horas,observa)
        Me.txtFecha(4).Text = miRsAux!f_preact
        If DBLet(miRsAux!Extension, "T") <> "" Then
            SQL = Format(RecuperaValor(ExtraData2, 2), "0000000") & RecuperaValor(ExtraData2, 3) & Format(RecuperaValor(ExtraData2, 4), "000")
            Text1(3).Text = SQL & "." & miRsAux!Extension
        End If
        Me.txtObra(2).Text = DBLet(miRsAux!observa, "T")
        txtimporte(0).Text = "": txtimporte(1).Text = ""
        If Not IsNull(miRsAux!Horas) Then Me.txtimporte(0).Text = Format(miRsAux!Horas, FormatoImporte)
        If Not IsNull(miRsAux!Importe) Then Me.txtimporte(1).Text = Format(miRsAux!Importe, FormatoImporte)
        
    End If
    miRsAux.Close
    
    'OK
    
    If txtFecha(4).Text <> "" Then
        'OK. Actuacion correcta
        SQL = "Select whoobrascliactuatrab.codtraba,straba.nomtraba  from whoobrascliactuatrab,straba WHERE "
        SQL = SQL & " whoobrascliactuatrab.codtraba=straba.codtraba AND  "
        SQL = SQL & " idactua = " & RecuperaValor(ExtraData2, 4)
        SQL = SQL & " AND anoexp = " & RecuperaValor(ExtraData2, 3) & " AND expediente = " & RecuperaValor(ExtraData2, 2)
        miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        While Not miRsAux.EOF
            SQL = miRsAux!CodTraba & "|" & miRsAux!NomTraba & "|"
            miRsAux.MoveNext
        Wend
    
        miRsAux.Close
    End If
    
    
    
    Set miRsAux = Nothing
End Sub

'A partir del SQL.  codtraba|nomtraba|
Private Sub InsertaEnLwTrabajadores()
    On Error GoTo eInsertaEnLwTrabajadores
    lwTrabajadores.ListItems.Add , "T" & Format(RecuperaValor(SQL, 1), "00000"), Format(RecuperaValor(SQL, 1), "00000")
    lwTrabajadores.ListItems(lwTrabajadores.ListItems.Count).SubItems(1) = RecuperaValor(SQL, 2)
    
    Exit Sub
eInsertaEnLwTrabajadores:
    MuestraError Err.Number, Err.Description

End Sub

Private Sub txtImporte_GotFocus(Index As Integer)
    ConseguirFoco txtimporte(Index), 3
End Sub

Private Sub txtImporte_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtImporte_LostFocus(Index As Integer)
    txtimporte(Index).Text = Trim(txtimporte(Index).Text)
    If txtimporte(Index).Text = "" Then Exit Sub
    Select Case Index
    Case 0, 1
        If Not PonerFormatoDecimal(txtimporte(Index), 1) Then txtimporte(Index).Text = ""
    
    End Select
End Sub


Private Sub PonTitularidad()
    'Igual hay que cargar el combo cboTitularidad


    If Opcion = 11 Then
        'NUEVO
        cboTitularidad.Text = ""
        Me.txtTitularidad.Text = ""
    Else
        cboTitularidad.Text = RecuperaValor(ExtraData2, 3)
        Me.txtTitularidad.Text = RecuperaValor(ExtraData2, 2)
    End If
    
    
End Sub


Private Sub InsertarTrabajadoresActuacion(idActua As Integer)

    If Opcion = 10 Then
        SQL = "DELETE FROM whoobrascliactuatrab "
        SQL = SQL & " WHERE expediente=" & RecuperaValor(ExtraData2, 2) & " AND anoexp=" & RecuperaValor(ExtraData2, 3)
        SQL = SQL & " AND IdActua=" & idActua
        conn.Execute SQL
        Espera 0.5
        
    End If
        
    
        
    
        
    SQL = ""
    For NumRegElim = 1 To Me.lwTrabajadores.ListItems.Count
        'whoobrascliactuatrab expediente anoexp IdActua codtraba
        SQL = SQL & ", (" & RecuperaValor(ExtraData2, 2) & "," & RecuperaValor(ExtraData2, 3)
        SQL = SQL & "," & idActua & "," & lwTrabajadores.ListItems(NumRegElim).Text & ")"
        
    Next
    
    If SQL <> "" Then
        SQL = Mid(SQL, 2)
        SQL = "INSERT INTO whoobrascliactuatrab(expediente ,anoexp ,IdActua ,codtraba) VALUES " & SQL
        ejecutar SQL, False
    
    End If
    
End Sub
