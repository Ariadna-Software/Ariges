VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMensajes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mensajes"
   ClientHeight    =   7620
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14160
   Icon            =   "frmMensajes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   14160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame FrameAcercaDe 
      BorderStyle     =   0  'None
      Height          =   3855
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   6375
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cambios version: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   240
         Left            =   120
         TabIndex        =   89
         Top             =   2040
         Width           =   1875
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Haga click en este enlace  para ver los cambios"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   1920
         MousePointer    =   3  'I-Beam
         TabIndex        =   88
         ToolTipText     =   "Haga click para seguir enlace"
         Top             =   2040
         Width           =   4410
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pasaje Ventura Feli�, 13 entlo.izquierdo 2�"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   195
         Left            =   2400
         TabIndex        =   10
         Top             =   3000
         Width           =   3855
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "46007 - VALENCIA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   195
         Left            =   3000
         TabIndex        =   9
         Top             =   3285
         Width           =   2535
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tfno:  902 88 88 78  -  96 380 55 79"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   195
         Left            =   2685
         TabIndex        =   8
         Top             =   3555
         Width           =   3165
      End
      Begin VB.Label Label27 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   195
         Left            =   4215
         TabIndex        =   7
         Top             =   3480
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Image Image2 
         Height          =   735
         Left            =   0
         Picture         =   "frmMensajes.frx":000C
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   1920
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Version"
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
         Left            =   -120
         TabIndex        =   6
         Top             =   1260
         Width           =   4155
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   81.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1725
         Left            =   4260
         TabIndex        =   5
         Top             =   0
         Width           =   1350
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "ARIGES"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   915
         Left            =   1080
         TabIndex        =   4
         Top             =   300
         Width           =   3495
      End
   End
   Begin VB.Frame FrameArticulosAgrupados 
      Height          =   6015
      Left            =   0
      TabIndex        =   68
      Top             =   1080
      Visible         =   0   'False
      Width           =   9375
      Begin VB.Frame FrameSelecArtAgrupado 
         Height          =   3255
         Left            =   1440
         TabIndex        =   75
         Top             =   1320
         Width           =   6495
         Begin VB.TextBox txtNoEditable 
            Alignment       =   1  'Right Justify
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
            Height          =   375
            Index           =   5
            Left            =   4560
            Locked          =   -1  'True
            TabIndex        =   81
            Text            =   "6"
            Top             =   2400
            Width           =   1335
         End
         Begin VB.TextBox txtNoEditable 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   4
            Left            =   2280
            Locked          =   -1  'True
            TabIndex        =   80
            Text            =   "5"
            Top             =   2400
            Width           =   1335
         End
         Begin VB.TextBox txtNoEditable 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   480
            Locked          =   -1  'True
            TabIndex        =   79
            Text            =   "4"
            Top             =   2400
            Width           =   1335
         End
         Begin VB.TextBox txtNoEditable 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   480
            Locked          =   -1  'True
            TabIndex        =   78
            Text            =   "3"
            Top             =   1560
            Width           =   5295
         End
         Begin VB.TextBox txtNoEditable 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   2280
            Locked          =   -1  'True
            TabIndex        =   77
            Text            =   "2"
            Top             =   720
            Width           =   1695
         End
         Begin VB.TextBox txtNoEditable 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   480
            Locked          =   -1  'True
            TabIndex        =   76
            Text            =   "1"
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label Label8 
            Caption         =   "TOTAL"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   4680
            TabIndex        =   87
            Top             =   2040
            Width           =   855
         End
         Begin VB.Label Label8 
            Caption         =   "PVP"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   2400
            TabIndex        =   86
            Top             =   2040
            Width           =   735
         End
         Begin VB.Label Label8 
            Caption         =   "Uds"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   480
            TabIndex        =   85
            Top             =   2040
            Width           =   735
         End
         Begin VB.Label Label8 
            Caption         =   "Descripci�n"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   2
            Left            =   480
            TabIndex        =   84
            Top             =   1200
            Width           =   1245
         End
         Begin VB.Label Label8 
            Caption         =   "Referencia"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   1
            Left            =   2280
            TabIndex        =   83
            Top             =   360
            Width           =   1170
         End
         Begin VB.Label Label8 
            Caption         =   "Id"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   480
            TabIndex        =   82
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.TextBox txtArtAgrupado 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   70
         Text            =   "1"
         Top             =   5400
         Width           =   735
      End
      Begin VB.CommandButton cmdArtAgrupado 
         Caption         =   "&Salir"
         Height          =   375
         Index           =   1
         Left            =   8160
         TabIndex        =   72
         Top             =   5520
         Width           =   975
      End
      Begin VB.CommandButton cmdArtAgrupado 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   0
         Left            =   7080
         TabIndex        =   71
         Top             =   5520
         Width           =   975
      End
      Begin MSComctlLib.ListView lwArticulosAgrupados 
         Height          =   4455
         Left            =   480
         TabIndex        =   69
         Top             =   840
         Width           =   8625
         _ExtentX        =   15214
         _ExtentY        =   7858
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Caja"
            Object.Width           =   2187
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripcion"
            Object.Width           =   7408
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Referencia"
            Object.Width           =   2187
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "PVP"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label7 
         Caption         =   "CAJAS"
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
         Height          =   495
         Index           =   3
         Left            =   480
         TabIndex        =   74
         Top             =   5400
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Art�culos agrupados"
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
         Height          =   495
         Index           =   2
         Left            =   480
         TabIndex        =   73
         Top             =   240
         Width           =   5295
      End
   End
   Begin VB.Frame FrameErrorCC 
      Height          =   6135
      Left            =   6000
      TabIndex        =   64
      Top             =   960
      Width           =   6495
      Begin VB.TextBox txtCCError 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4335
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   67
         Text            =   "frmMensajes.frx":042A
         Top             =   840
         Width           =   5895
      End
      Begin VB.CommandButton cmdSalirCC 
         Caption         =   "&Salir"
         Height          =   375
         Left            =   5040
         TabIndex        =   66
         Top             =   5520
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Errores centro de coste"
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
         Height          =   495
         Index           =   0
         Left            =   240
         TabIndex        =   65
         Top             =   360
         Width           =   5295
      End
   End
   Begin VB.Frame FrameNSeries 
      Height          =   5000
      Left            =   480
      TabIndex        =   11
      Top             =   120
      Width           =   6975
      Begin VB.CommandButton cmdSelTodos 
         Caption         =   "&Todos"
         Height          =   315
         Left            =   720
         TabIndex        =   22
         Top             =   4320
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.CommandButton cmdDeselTodos 
         Caption         =   "&Ninguno"
         Height          =   315
         Left            =   1680
         TabIndex        =   21
         Top             =   4320
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5040
         TabIndex        =   20
         Top             =   4320
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptarNSeries 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3960
         TabIndex        =   12
         Top             =   4320
         Width           =   975
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   3015
         Left            =   720
         TabIndex        =   13
         Top             =   840
         Width           =   3945
         _ExtentX        =   6959
         _ExtentY        =   5318
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
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
      Begin VB.Label Label7 
         Caption         =   "Empresas en el sistema"
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
         Height          =   495
         Index           =   1
         Left            =   720
         TabIndex        =   30
         Top             =   240
         Visible         =   0   'False
         Width           =   5295
      End
   End
   Begin VB.Frame FrameEtiqEstant 
      Height          =   7455
      Left            =   0
      TabIndex        =   31
      Top             =   -120
      Width           =   8535
      Begin VB.CommandButton cmdEtiqEstan 
         Caption         =   "Imprimir"
         Height          =   375
         Index           =   1
         Left            =   5520
         TabIndex        =   34
         Top             =   6960
         Width           =   1215
      End
      Begin VB.CommandButton cmdEtiqEstan 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   0
         Left            =   6960
         TabIndex        =   33
         Top             =   6960
         Width           =   1215
      End
      Begin MSComctlLib.ListView ListView3 
         Height          =   6495
         Left            =   240
         TabIndex        =   32
         Top             =   360
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   11456
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Descripcion"
            Object.Width           =   6703
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Precio"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Familia"
            Object.Width           =   4410
         EndProperty
      End
      Begin VB.Label Label6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   63
         Top             =   6960
         Width           =   4095
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   1
         Left            =   600
         Picture         =   "frmMensajes.frx":0430
         Top             =   6960
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   0
         Left            =   240
         Picture         =   "frmMensajes.frx":057A
         Top             =   6960
         Width           =   240
      End
   End
   Begin VB.Frame FrameEmail 
      Height          =   6975
      Left            =   3600
      TabIndex        =   50
      Top             =   0
      Width           =   8055
      Begin VB.TextBox txmemail 
         Height          =   315
         Index           =   4
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   61
         Text            =   "Text2"
         Top             =   2160
         Width           =   2535
      End
      Begin VB.CommandButton cmdEmail 
         Caption         =   "Salir"
         Height          =   375
         Left            =   6720
         TabIndex        =   60
         Top             =   6480
         Width           =   975
      End
      Begin VB.TextBox txmemail 
         Height          =   3555
         Index           =   3
         Left            =   360
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   58
         Text            =   "frmMensajes.frx":06C4
         Top             =   2760
         Width           =   7335
      End
      Begin VB.TextBox txmemail 
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   56
         Text            =   "Text2"
         Top             =   2160
         Width           =   4815
      End
      Begin VB.TextBox txmemail 
         Height          =   315
         Index           =   1
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   52
         Text            =   "Text2"
         Top             =   1560
         Width           =   7335
      End
      Begin VB.TextBox txmemail 
         Height          =   315
         Index           =   0
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   51
         Text            =   "Text2"
         Top             =   960
         Width           =   7335
      End
      Begin VB.Label Label5 
         Caption         =   "Fecha"
         Height          =   255
         Index           =   4
         Left            =   5280
         TabIndex        =   62
         Top             =   1920
         Width           =   2535
      End
      Begin VB.Label Label5 
         Caption         =   "Mensaje"
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   59
         Top             =   2520
         Width           =   2535
      End
      Begin VB.Label Label5 
         Caption         =   "Adjuntos"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   57
         Top             =   1920
         Width           =   2535
      End
      Begin VB.Label Label5 
         Caption         =   "Asunto"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   55
         Top             =   1320
         Width           =   2535
      End
      Begin VB.Label Label5 
         Caption         =   "Label5"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   54
         Top             =   720
         Width           =   2535
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Email CRM"
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
         Left            =   960
         TabIndex        =   53
         Top             =   240
         Width           =   5415
      End
   End
   Begin VB.Frame FrameCorreccionPrecios 
      Height          =   6375
      Left            =   0
      TabIndex        =   35
      Top             =   0
      Width           =   12975
      Begin VB.ComboBox cmbActualizarTar 
         Height          =   315
         ItemData        =   "frmMensajes.frx":06CA
         Left            =   7800
         List            =   "frmMensajes.frx":06CC
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Top             =   5960
         Width           =   2175
      End
      Begin VB.CommandButton cmdCorrecotrPrecios 
         Caption         =   "Salir"
         Height          =   375
         Index           =   1
         Left            =   11760
         TabIndex        =   38
         Top             =   5880
         Width           =   975
      End
      Begin VB.CommandButton cmdCorrecotrPrecios 
         Caption         =   "Aceptar"
         Height          =   375
         Index           =   0
         Left            =   10560
         TabIndex        =   37
         Top             =   5880
         Width           =   975
      End
      Begin MSComctlLib.ListView ListView4 
         Height          =   5175
         Left            =   120
         TabIndex        =   36
         Top             =   600
         Width           =   12660
         _ExtentX        =   22331
         _ExtentY        =   9128
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Codigo"
            Object.Width           =   3246
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Denominaci�n"
            Object.Width           =   5715
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "U.P.Compra"
            Object.Width           =   2188
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "% M"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "PVP"
            Object.Width           =   2011
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "%T"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "P.Tarifa"
            Object.Width           =   2011
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "PVP Correcto"
            Object.Width           =   2011
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "Tarifa correc."
            Object.Width           =   2011
         EndProperty
      End
      Begin VB.Label Label2 
         Caption         =   "Actualizar:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   1
         Left            =   6480
         TabIndex        =   42
         Top             =   6000
         Width           =   1215
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   2
         Left            =   11760
         Picture         =   "frmMensajes.frx":06CE
         Top             =   240
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   3
         Left            =   12360
         Picture         =   "frmMensajes.frx":0818
         Top             =   240
         Width           =   240
      End
      Begin VB.Label lblIndicadorCorregir 
         Caption         =   "Label3"
         Height          =   255
         Left            =   240
         TabIndex        =   40
         Top             =   5880
         Width           =   5055
      End
      Begin VB.Label Label2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   39
         Top             =   240
         Width           =   8175
      End
   End
   Begin VB.Frame FrameErrores 
      Height          =   5535
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   8415
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "&Cerrar"
         Height          =   375
         Left            =   7080
         TabIndex        =   29
         Top             =   4920
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   4335
         Left            =   360
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   28
         Text            =   "frmMensajes.frx":0962
         Top             =   360
         Width           =   7695
      End
   End
   Begin VB.Frame FrameComponentes 
      Height          =   3975
      Left            =   0
      TabIndex        =   14
      Top             =   120
      Width           =   5295
      Begin VB.CommandButton cmdAceptarComp 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3240
         TabIndex        =   19
         Top             =   3240
         Width           =   975
      End
      Begin VB.Frame FrameComponentes2 
         Caption         =   "Mostrar Equipos del :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   2175
         Left            =   960
         TabIndex        =   15
         Top             =   720
         Width           =   3255
         Begin VB.OptionButton OptCompXClien 
            Caption         =   "Cliente"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            TabIndex        =   18
            Top             =   1440
            Width           =   2535
         End
         Begin VB.OptionButton OptCompXDpto 
            Caption         =   "Departamento"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            TabIndex        =   17
            Top             =   960
            Width           =   2535
         End
         Begin VB.OptionButton OptCompXMant 
            Caption         =   "Mantenimiento"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            TabIndex        =   16
            Top             =   480
            Width           =   2055
         End
      End
   End
   Begin VB.Frame FrameTraspasoMante 
      Height          =   3135
      Left            =   0
      TabIndex        =   43
      Top             =   0
      Width           =   5415
      Begin VB.TextBox txtMante 
         Height          =   285
         Index           =   0
         Left            =   2640
         TabIndex        =   48
         Top             =   1080
         Width           =   855
      End
      Begin VB.CheckBox chkMante 
         Caption         =   "Copiar importes en siguiente"
         Height          =   255
         Left            =   360
         TabIndex        =   46
         Top             =   1800
         Value           =   1  'Checked
         Width           =   2655
      End
      Begin VB.CommandButton cmdMante 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   1
         Left            =   3960
         TabIndex        =   45
         Top             =   2520
         Width           =   1215
      End
      Begin VB.CommandButton cmdMante 
         Caption         =   "Aceptar"
         Height          =   375
         Index           =   0
         Left            =   2520
         TabIndex        =   44
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "A�o a traspasar"
         Height          =   195
         Left            =   1320
         TabIndex        =   49
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Pasar importes mantenimiento a historico."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   47
         Top             =   360
         Width           =   4695
      End
   End
   Begin VB.Frame FrameCobrosPtes 
      Height          =   4935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8655
      Begin VB.CommandButton cmdCancelarCobros 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   7080
         TabIndex        =   25
         Top             =   4440
         Width           =   975
      End
      Begin VB.TextBox txtParam 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   1095
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   23
         Text            =   "frmMensajes.frx":0968
         Top             =   120
         Width           =   6615
      End
      Begin VB.CommandButton cmdAceptarCobros 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   5400
         TabIndex        =   1
         Top             =   4440
         Width           =   975
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3135
         Left            =   120
         TabIndex        =   2
         Top             =   1200
         Width           =   8100
         _ExtentX        =   14288
         _ExtentY        =   5530
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
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
      Begin VB.Label Label1 
         Caption         =   "�Desea continuar?"
         Height          =   375
         Index           =   1
         Left            =   480
         TabIndex        =   26
         Top             =   4440
         Width           =   7215
      End
      Begin VB.Label Label1 
         Caption         =   "Departamento:"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   24
         Top             =   240
         Width           =   7215
      End
   End
End
Attribute VB_Name = "frmMensajes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'====================== VBLES PUBLICAS ================================

Public Event DatoSeleccionado(CadenaSeleccion As String)

Public OpcionMensaje As Byte
'======================================
'==== FACTURACION =====================
' 1 .- Mensaje de Cobros Pendientes
' 2 .- Mensaje de No hay suficiente Stock para pasar de Pedido a Albaran
' 3 .- Mensaje Acerca de...
' 4 .- Listado de los N� de Serie de un Articulo
' 5 .- Seleccionar tipo de Componente a Mostrar en Mant. de N� de Series
' 6 .- Mostrar Prefacturacion de Albaranes
' 7 .- Mostrar Prefacturacion Mantenimientos
' 8 .- Mostrar lista clientes para seleccionar los que queremos imprimir (Etiquetas)
' 9 .- Mostrar lista Proveedores para seleccionar los que queremos imprimir (Etiquetas)
'10 .- Mostrar lista de Errores de las facturas NO contabilizadas
'11 .- Mostrar lista lineas de factura a Rectificar para seleccionar las q queremos traer al Albaran de FAct. Rectificativa
'12 .- Mostrar Albaranes del Rango que no se van a Facturar. (Facturar Albaranes Venta)

'13 .- Mostrar Errores
'14 .- Mostrar Empresas existentes en el sistema



'15 .- Mostrar lista de articulos para imprimir etiquetas estanteria
'16 .- Lista de articulos para corregir importes
'17 .- Etiquetas clientes. LO MISMO QUE EL 8 pero hecho por david
'18 .- Mantenimientos. paso ejercicio siguiente a actual

'20 .- IGual que el 16. Pero los importes son de los articulos que tienen componentes

'21 .- Ver un mensaje enlazado desde el outlook para el CRM

'22 .-  Muestra clientes potenciales

'23 .- Igual que 15. Listado PVP con IVA  (para los TPVs)

'24 .- Lineas de factura sib centro de coste

'25 .- Articulos agruopados en ventas TPV

Public cadWhere As String 'Cadena para pasarle la WHERE de la SELECT de los cobros pendientes o de Pedido(para comp. stock)
                          'o CodArtic para seleccionar los N� Series
                          'para cargar el ListView
                          
Public cadWHERE2 As String

Public vCampos As String 'Articulo y cantidad Empipados para N� de Series
                         'Tambien para pasar el nombre de la tabla de lineas (sliped, slirep,...)
                         'Dependiendo desde donde llamemos, de Pedidos o Reparaciones


'====================== VBLES LOCALES ================================

Dim PulsadoSalir As Boolean 'Solo salir con el boton de Salir no con aspa del form
Dim primeravez As Boolean

'Para los N� de Serie
Dim TotalArray As Integer
Dim codArtic() As String
Dim cantidad() As Integer



Private Sub cmdAceptarCobros_Click()
    If OpcionMensaje = 12 Then vCampos = "1"
    Unload Me
End Sub


Private Sub cmdAceptarComp_Click()
'Boton Aceptar de Componentes del Mant. de N� de Series en Reparaciones
Dim H As Integer, W As Integer

    ponerFrameComponentesVisible False, H, W
    PonerFrameCobrosPtesVisible True, H, W
    Me.Height = H + 350
    Me.Width = W + 70
    
    If Me.OptCompXMant.Value Then
        'Mostrar Resumen de los N� de Serie del Mantenimiento
        Me.Caption = "Equipos del Mantenimiento"
        CargarListaComponentes (1)
    ElseIf Me.OptCompXDpto.Value Then
        'Mostrar Resumen de los N� de Serie del Departamento
        Me.Caption = "Equipos del Departamento"
        CargarListaComponentes (2)
    ElseIf Me.OptCompXClien.Value Then
        'Mostrar Resumen de los N� de Serie del Cliente
        Me.Caption = "Equipos del Cliente"
        CargarListaComponentes (3)
    End If
    PonerFocoBtn Me.cmdAceptarCobros
End Sub


Private Sub cmdAceptarNSeries_Click()
Dim I As Integer, J As Byte
Dim Seleccionados As Integer
Dim cad As String, SQL As String
Dim Articulo As String
Dim RS As ADODB.Recordset
Dim C1 As String * 10, C2 As String * 10, c3 As String * 10


    If OpcionMensaje = 4 Then
        'Comprobar que se han seleccionado el n� correcto de  N� de Serie para cada Articulo
        Seleccionados = 0
        Articulo = ""
      
        'Si se ha seleccionado la cantidad correcta de N� de series, empiparlos y
        'devolverlos al form de Albaranes(facturacion)
        cad = ""
        For J = 0 To TotalArray
            Articulo = codArtic(J)
            cad = cad & Articulo & "|"
            For I = 1 To ListView2.ListItems.Count
                If ListView2.ListItems(I).Checked Then
                    If Articulo = ListView2.ListItems(I).ListSubItems(1).Text Then
                        If Seleccionados < Abs(cantidad(J)) Then
                            Seleccionados = Seleccionados + 1
                            cad = cad & ListView2.ListItems(I).Text & "|"
                        End If
                   'cad = cad & Data1.Recordset.Fields(1) & "|"
                    End If
                End If
            Next I
            If Seleccionados < Abs(cantidad(J)) Then
                'Comprobar que si tiene N�s de serie de ese articulos cargados seleccione los
                'que corresponden
                SQL = "SELECT count(sserie.numserie)"
                SQL = SQL & " FROM sserie " 'INNER JOIN sartic ON sserie.codartic=sartic.codartic "
                SQL = SQL & " WHERE sserie.codartic=" & DBSet(Articulo, "T")
                SQL = SQL & " AND (isnull(sserie.numfactu) or sserie.numfactu='') and (isnull(sserie.numalbar) or sserie.numalbar='') "
                SQL = SQL & " ORDER BY sserie.codartic, numserie "
                Set RS = New ADODB.Recordset
                RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                
                If RS.Fields(0).Value >= Abs(cantidad(J)) - Seleccionados Then
                    MsgBox "Debe seleccionar " & cantidad(J) & " N� Series para el articulo " & codArtic(J), vbExclamation
                    Exit Sub
                Else
                    'No hay N� Serie y Pedirlos
                End If
                RS.Close
                Set RS = Nothing
            
            End If
            cad = cad & "�"
            Seleccionados = 0
        Next J
      
    ElseIf OpcionMensaje = 8 Or OpcionMensaje = 9 Or OpcionMensaje = 17 Or OpcionMensaje = 22 Then
        'concatenar todos los clientes seleccionados para imprimir etiquetas
        If OpcionMensaje = 17 Then
            
            '----------------------------------------------------------------
            '                                                      pongo numlinea cone l contador de registro como clave
            cad = "insert into tmpnlotes (codusu,numalbar,fechaalb,numlinea,codalmac,codprove) values ("
            ' ---- [30/10/2009] (LAURA) : agrupar por cliente y departamento
'            cad = cad & vUsu.Codigo & ",1,'2005-04-12',1,"
            cad = cad & vUsu.Codigo & ",1,'2005-04-12',"
            
            
            For I = 1 To ListView2.ListItems.Count
                If ListView2.ListItems(I).Checked Then
                    ' ---- [30/10/2009] (LAURA) : agrupar por cliente y departamento
'                    conn.Execute cad & (ListView2.ListItems(I).Text) & ")"

                                                    
                    conn.Execute cad & NumRegElim & "," & DBSet(ListView2.ListItems(I).ListSubItems(3).Text, "N", "S") & "," & (ListView2.ListItems(I).Text) & ")"
                    
                    NumRegElim = NumRegElim + 1
                End If
            Next I
            
            
            '----------------------------------------------------------------
            '
            ' 29/11/2010
            '
            'A partir de los datos vamos a meter en la tmpinfomres los valore
            If Not CargaDatosEtiquetas Then Exit Sub
            
        Else
            cad = ""
            NumRegElim = 0
            For I = 1 To ListView2.ListItems.Count
                If ListView2.ListItems(I).Checked Then
                    NumRegElim = NumRegElim + 1
                    cad = cad & Val(ListView2.ListItems(I).Text) & ","
                     'cad = cad & Data1.Recordset.Fields(1) & "|"
                End If
            Next I
            If NumRegElim > 1000 Then
                MsgBox "Maximo n�mero de etiquetas: 1000 (" & NumRegElim & ")", vbExclamation
                NumRegElim = 0
                cad = ""
                Exit Sub
            End If
            NumRegElim = 0
            If cad <> "" Then cad = Mid(cad, 1, Len(cad) - 1)
        End If
    ElseIf OpcionMensaje = 11 Then
    'Lineas Factura a rectificar
        'cad = "(" & cadWHERE & ")"
        cad = ""
        C1 = ""
        C2 = ""
        c3 = ""
        SQL = ""
        For I = 1 To ListView2.ListItems.Count
            If ListView2.ListItems(I).Checked Then
                If SQL = "" Then
                    C1 = DBSet(ListView2.ListItems(I), "T", "N")
                    C2 = ListView2.ListItems(I).ListSubItems(1)
'                    c3 = ListView2.ListItems(i).ListSubItems(2)
                    cad = "(codtipoa=" & Trim(C1) & " and numalbar=" & Val(C2) & " and numlinea IN (" & ListView2.ListItems(I).ListSubItems(2)

                Else
                    If Trim(DBSet(ListView2.ListItems(I), "T", "N")) = Trim(C1) And Trim(ListView2.ListItems(I).ListSubItems(1)) = Trim(C2) Then
                    'es el mismo albaran y concatenamos lineas
                        cad = "," & ListView2.ListItems(I).ListSubItems(2)

                    Else
                        If cad <> "" Then SQL = SQL & ")) "
                        C1 = DBSet(ListView2.ListItems(I), "T", "N")
                        C2 = ListView2.ListItems(I).ListSubItems(1)
'                    c3 = ListView2.ListItems(i).ListSubItems(2)
                        cad = " or (codtipoa=" & Trim(C1) & " and numalbar=" & Val(C2) & " and numlinea IN (" & ListView2.ListItems(I).ListSubItems(2)
                        
'                       cad=cad &
                    End If
                End If
                SQL = SQL & cad
'                If cad <> "" Then cad = cad & " OR "
'                cad = cad & "(codtipoa=" & DBSet(ListView2.ListItems(i), "T", "N") & " and numalbar=" & Val(ListView2.ListItems(i).ListSubItems(1)) & " and numlinea=" & ListView2.ListItems(i).ListSubItems(2) & ")"
            Else
'                cad = ""
            End If
        Next I
        If cad <> "" Then
            SQL = SQL & "))"
            cad = "(" & cadWhere & ") AND (" & SQL & ")"
        End If
'        If cad <> "" Then cad = "(" & cadWHERE & ") AND (" & cad & ")"
    ElseIf OpcionMensaje = 14 Then
        cad = RegresarCargaEmpresas
    End If
    
    
    
     'Actualizar la tabla sseries asignando los valores correspondientes a los
      'campos: codclien, coddirec, tieneman, codtipom, numalbar, fechavta, numline1
      'y Salir (Volver a Mto Albaranes Clientes (Facturacion)
      PulsadoSalir = True
      'RaiseEvent CargarNumSeries
      RaiseEvent DatoSeleccionado(cad)
      Unload Me
End Sub


Private Sub cmdArtAgrupado_Click(index As Integer)
Dim Impor As Currency
    CadenaDesdeOtroForm = ""
    If FrameSelecArtAgrupado.visible Then
        If index = 0 Then
            'OK este es el lote y las uds que quiere
            CadenaDesdeOtroForm = lwArticulosAgrupados.SelectedItem.Text & "|" & txtArtAgrupado.Text & "|"  'lote y uds
            Unload Me
        Else
            ponerframeTotaAgrupadoVisible False
        End If
    Else
        If index = 0 Then
            If txtArtAgrupado.Text = "" Then txtArtAgrupado.Text = "1"
            If Me.lwArticulosAgrupados.SelectedItem Is Nothing Then Exit Sub
            
            With lwArticulosAgrupados.SelectedItem
                Me.txtNoEditable(0).Text = .Text
                Me.txtNoEditable(1).Text = .SubItems(2)
                Me.txtNoEditable(2).Text = .SubItems(1)
                Me.txtNoEditable(3).Text = txtArtAgrupado.Text
                Me.txtNoEditable(4).Text = .SubItems(3)
                Impor = ImporteFormateado(.SubItems(3))
                Impor = Impor * CInt(Me.txtArtAgrupado.Text)
                Me.txtNoEditable(5).Text = Format(Impor, FormatoImporte)
            End With
            
            ponerframeTotaAgrupadoVisible True
            
            PonerFocoBtn cmdArtAgrupado(0)
        Else
            Unload Me
        End If
    End If
End Sub

Private Sub ponerframeTotaAgrupadoVisible(visible As Boolean)
    FrameSelecArtAgrupado.visible = visible
    Me.lwArticulosAgrupados.Enabled = Not visible
    Me.txtArtAgrupado.Enabled = Not visible
End Sub

Private Sub cmdCancelar_Click()
    If OpcionMensaje = 4 Then
        MsgBox "Debe introducir los n� de serie necesarios para el Albaran.", vbInformation
        Exit Sub
    End If
    PulsadoSalir = True
    RaiseEvent DatoSeleccionado("")
    Unload Me
End Sub

Private Sub cmdCancelarCobros_Click()
    vCampos = "0"
    Unload Me
End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub cmdCorrecotrPrecios_Click(index As Integer)
    
    If index = 0 Then
        
        If Not ActualizarPrecios Then Exit Sub
        
    End If
    Unload Me
End Sub

Private Function ActualizarPrecios() As Boolean
Dim SQL As String
    
    
    
        
        ActualizarPrecios = False
        'Compruebo si ha seleccionado algun articulo de los de precio ultima compra=0
        cadWHERE2 = ""
        SQL = ""
        For TotalArray = 1 To Me.ListView4.ListItems.Count
            If ListView4.ListItems(TotalArray).Checked Then
                If ListView4.ListItems(TotalArray).Tag = "" Then
                    SQL = SQL & "M"
                Else
                    cadWHERE2 = cadWHERE2 & "M"
                End If
            End If
        Next
    
        If SQL <> "" Then
            MsgBox "No puede actualizar los articulos cuyo precio ultima compra sea 0", vbExclamation
            Exit Function
        End If
    
        If cadWHERE2 = "" Then
            MsgBox "Seleccione algun articulo para actualizar", vbExclamation
            Exit Function
        End If
    
        'Llegado aqui todo correcto. Hacemos la pregunta de actualizar y a correr
        SQL = "art�culo"
        If Len(cadWHERE2) > 1 Then SQL = SQL & "s"
        SQL = "Va a actualizar los precios de " & Len(cadWHERE2) & " " & SQL & vbCrLf & vbCrLf & "�Desea continuar?"
        If MsgBox(SQL, vbQuestion + vbYesNo) <> vbYes Then Exit Function
        
        
        'Aqui esta el proceso de actualizacion de articulos
        Me.lblIndicadorCorregir.Caption = "Actualizaci�n precios"
        Me.Refresh
        Espera 0.5
        
       'Para el LOG
       SQL = cadWhere & vbCrLf
       For TotalArray = 1 To Me.ListView4.ListItems.Count
            If ListView4.ListItems(TotalArray).Checked Then
                If ListView4.ListItems(TotalArray).Tag <> "" Then SQL = SQL & ListView4.ListItems(TotalArray).Text & "|"
            End If
        Next
        SQL = Mid(SQL, 1, 237)
        
        '------------------------------------------------------------------------------
        '  LOG de acciones
        Set LOG = New cLOG
        LOG.Insertar 4, vUsu, "Correccion precios: " & vbCrLf & SQL
        Set LOG = Nothing
        '-----------------------------------------------------------------------------
        
        
        
        
        
        
        
        
        
        
        For TotalArray = 1 To Me.ListView4.ListItems.Count
            If ListView4.ListItems(TotalArray).Checked Then
                If ListView4.ListItems(TotalArray).Tag <> "" Then
                    
                    'lo metemos en transaccion. Si queremos vamos
                    Me.lblIndicadorCorregir.Caption = ListView4.ListItems(TotalArray).Text
                    Me.lblIndicadorCorregir.Refresh
                    
                                        
                    conn.BeginTrans
                    If ActualizaPrecios(TotalArray) Then
                        conn.CommitTrans
                    Else
                        conn.RollbackTrans
                    End If
                    
                    
                End If
            End If
        Next
    
    
        ActualizarPrecios = True
End Function


Private Function ActualizaPrecios(NumeroItem As Integer) As Boolean

On Error GoTo EActualizaPrecios
    ActualizaPrecios = False
    With ListView4.ListItems(NumeroItem)
        If OpcionMensaje = 16 Then
            'ACtualizador de precio normal
            If Me.cmbActualizarTar.ListIndex <> 2 Then
                cadWHERE2 = TransformaComasPuntos(CStr(ImporteFormateado(.SubItems(7))))
                cadWHERE2 = "UPDATE sartic set preciove=" & cadWHERE2 & " WHERE codartic = '" & ListView4.ListItems(NumeroItem).Tag & "'"
                conn.Execute cadWHERE2
            End If
            If Me.cmbActualizarTar.ListIndex <> 1 Then
                cadWHERE2 = TransformaComasPuntos(CStr(ImporteFormateado(.SubItems(8))))
                cadWHERE2 = "UPDATE slista set precioac=" & cadWHERE2 & " WHERE codartic = '" & ListView4.ListItems(NumeroItem).Tag & "' AND codlista =" & vCampos
                conn.Execute cadWHERE2
            End If
            
        Else
            'Precio articulos componentes
            '----------------------------
            vCampos = ""
            If Me.cmbActualizarTar.ListIndex <> 2 Then
                cadWHERE2 = TransformaComasPuntos(CStr(ImporteFormateado(.SubItems(7))))
                vCampos = " preciove = " & cadWHERE2
            End If
            If Me.cmbActualizarTar.ListIndex <> 1 Then
                cadWHERE2 = TransformaComasPuntos(CStr(ImporteFormateado(.SubItems(8))))
                If vCampos <> "" Then vCampos = vCampos & ","
                vCampos = vCampos & " preciouc = " & cadWHERE2
            End If
            cadWHERE2 = "UPDATE sartic set " & vCampos & " WHERE codartic = '" & ListView4.ListItems(NumeroItem).Tag & "'"
            conn.Execute cadWHERE2
            
            
                        

            
        End If
        
    End With
        
    ActualizaPrecios = True
    Exit Function
EActualizaPrecios:
    MuestraError Err.Number, ListView4.ListItems(NumeroItem).Text
End Function


Private Sub cmdDeselTodos_Click()
Dim I As Integer

    For I = 1 To ListView2.ListItems.Count
        ListView2.ListItems(I).Checked = False
    Next I
End Sub

Private Sub cmdEmail_Click()
    Unload Me
End Sub

Private Sub cmdEtiqEstan_Click(index As Integer)
    Screen.MousePointer = vbHourglass
    If index = 1 Then
        If OpcionMensaje = 23 Then
            'lISTADO PRECIOS tpv
            ImprimeListadoTPV
        Else
            GenerarEtiquetasEstanterias Me.ListView3, cadWhere
            
            
            
        End If
    Else
        If TotalArray > 0 Then
            TotalArray = -1
            Exit Sub
        End If
        NumRegElim = 0
    End If
    
    Screen.MousePointer = vbDefault
    Unload Me
End Sub



Private Sub cmdMante_Click(index As Integer)
Dim b As Boolean
    If index = 0 Then
        
        
        If Val(txtMante(0).Text) = 0 Then
            MsgBox "El campo A�o a traspasar debe ser num�rico", vbExclamation
            Exit Sub
        End If
        
        
        If MsgBox("El proceso es irreversible. Continuar de igual modo?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        
        '-------------------------------------------
        Screen.MousePointer = vbHourglass
        Set miRsAux = New ADODB.Recordset
        conn.BeginTrans
        b = TraspasarMantenimientos
        Set miRsAux = Nothing
        Screen.MousePointer = vbDefault
        If b Then
            conn.CommitTrans
        Else
            conn.RollbackTrans
        End If
        
        
    End If
    Unload Me
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdSalirCC_Click()
    Unload Me
End Sub

Private Sub cmdSelTodos_Click()
    Dim I As Integer

    For I = 1 To ListView2.ListItems.Count
        ListView2.ListItems(I).Checked = True
    Next I
End Sub

Private Sub Form_activate()
Dim OK As Boolean
    
    
    
    Select Case OpcionMensaje
        Case 4 'Mostrar N� Series
            If primeravez Then
                primeravez = False
                Me.Refresh
                Screen.MousePointer = vbHourglass
                OK = ObtenerTamanyosArray
                If OK Then OK = SeparaCampos
                If Not OK Then
                    'Error en SQL
                    'Salimos
                    Unload Me
                    Exit Sub
                End If
                CargarListaNSeries
            End If
            
        Case 8, 9, 17, 22 'Etiquetas de clientes/Proveedores
            CargarListaClientes
'        Case 10 'Errores al contabilizar facturas
'            CargarListaErrContab
        Case 11 'Lineas Factura a rectificar
            CargarListaLinFactu
            
        Case 14 'Mostrar Empresas del sistema
            CargarListaEmpresas
            
        Case 15
            'Etiquetas estanteria
            CargarArticulosEstanteria
            
        Case 16, 20
            'Articulos para corregir
            If OpcionMensaje = 16 Then
                CargarArticulosCorreccionPrecio
            Else
                CargaPVPPreciosArticulosConComponentes
            End If
            If Me.ListView4.ListItems.Count = 0 Then
                MsgBox "Ning�n dato para mostrar", vbExclamation
                Unload Me
            End If
        Case 18
            PonerFoco txtMante(0)
        Case 21
            CargarEmail
        Case 23
            CargarPVPArticulos   'aqui aqui auqi
            
        Case 24
            txtCCError.Text = vCampos
            vCampos = ""
            
        Case 25
            CargaArticulosAgrupados
            
    End Select
    
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
Dim H As Integer, W As Integer
Dim cad As String
On Error Resume Next

    Me.FrameCobrosPtes.visible = False
    Me.FrameAcercaDe.visible = False
    Me.FrameNSeries.visible = False
    Me.FrameComponentes.visible = False
    Me.FrameComponentes2.visible = False
    Me.FrameErrores.visible = False
    FrameEtiqEstant.visible = False
    FrameCorreccionPrecios.visible = False
    FrameTraspasoMante.visible = False
    FrameEMail.visible = False
    FrameErrorCC.visible = False
    FrameArticulosAgrupados.visible = False
    PulsadoSalir = True
    primeravez = True
    
    Select Case OpcionMensaje
        Case 1 'Mensaje de Cobros Pendientes
            PonerFrameCobrosPtesVisible True, H, W
            CargarListaCobrosPtes
            Me.Caption = "Cobros Pendientes"
            PonerFocoBtn Me.cmdAceptarCobros
            
        Case 2 'Mensaje de no hay suficiente Stock
            PonerFrameCobrosPtesVisible True, H, W
            CargarListaArtSinStock (vCampos)
            Me.Caption = "Art�culos sin stock suficiente"
            PonerFocoBtn Me.cmdAceptarCobros
            
        Case 3 'Mensaje ACERCA DE
            CargaImagen
            Me.Caption = "Acerca de ....."
            PonerFrameAcercaDeVisible True, H, W
            vCampos = ""
            PonerFechaArchivo
            If vCampos = "" Then
                vCampos = "Versi�n:  "
            Else
                vCampos = vCampos & "         ver:"
            End If
            Me.lblVersion.Caption = vCampos & App.Major & "." & App.Minor & "." & App.Revision & " "
        
        Case 4 'Listado N� Series Articulo
            PonerFrameNSeriesVisible True, H, W
            Me.Caption = "N� Serie"
            Me.Label7(1).Caption = "Seleccione los N� de serie para el Albaran."
            Me.Label7(1).FontSize = 12
            PulsadoSalir = False
            
        Case 5 'Seleccionar tipo de Componente que queremos mostrar en Resumen
                'En mant. de N� Series de Reparacion
            ponerFrameComponentesVisible True, H, W
            Me.Caption = "Componentes"
            Me.OptCompXMant.Value = True
            PonerFocoBtn Me.cmdAceptarComp
        
        Case 6 'Mostrar Prefacturacion de Albaranes
            PonerFrameCobrosPtesVisible True, H, W
            CargarListaPreFacturar
            Me.Caption = "Prefacturaci�n Albaranes"
            cad = RecuperaValor(vCampos, 1)
            If cad <> "" Then cad = Mid(cad, 1, Len(cad) - 1)
            Me.txtParam.Text = cad
            cad = RecuperaValor(vCampos, 2)
            If cad <> "" Then
                cad = Mid(cad, 1, Len(cad) - 1)
                If Trim(Me.txtParam.Text) <> "" Then
                    txtParam.Text = Me.txtParam.Text & vbCrLf & cad
                Else
                    txtParam.Text = cad
                End If
            End If
            cad = RecuperaValor(vCampos, 3)
            If cad <> "" Then
                cad = Mid(cad, 1, Len(cad) - 1)
                If Trim(Me.txtParam.Text) <> "" Then
                    txtParam.Text = Me.txtParam.Text & vbCrLf & cad
                Else
                    txtParam.Text = cad
                End If
            End If
            
            PonerFocoBtn Me.cmdAceptarComp
            
        Case 8, 17, 22 'Etiquetas de Clientes
            PonerFrameNSeriesVisible True, H, W
            Me.Caption = "Clientes"
            If OpcionMensaje = 22 Then Me.Caption = Me.Caption & " potenciales"
            
            
            PulsadoSalir = False
            Me.cmdSelTodos.visible = True
            Me.cmdDeselTodos.visible = True
            
        Case 9 'Etiquetas de Proveedores
            PonerFrameNSeriesVisible True, H, W
            Me.Caption = "Proveedores"
            PulsadoSalir = False
            Me.cmdSelTodos.visible = True
            Me.cmdDeselTodos.visible = True
        
        Case 10 'Errores al contabilizar facturas
            PonerFrameCobrosPtesVisible True, H, W
            CargarListaErrContab
            Me.Caption = "Facturas NO contabilizadas: "
            PonerFocoBtn Me.cmdAceptarCobros
        
        Case 11 'Lineas Factura a Rectificar
            PonerFrameNSeriesVisible True, H, W
            Me.Caption = "Lineas Factura a Rectificar"
            PulsadoSalir = False
            Me.cmdSelTodos.visible = True
            Me.cmdDeselTodos.visible = True
            Me.cmdAceptarNSeries.Left = Me.cmdAceptarNSeries.Left + 1000
            Me.cmdCancelar.Left = Me.cmdCancelar.Left + 1000
        
        Case 12 'Mensaje Albaranes que no se van a Facturar
            PonerFrameCobrosPtesVisible True, H, W
            CargarListaAlbaranes
            Me.Caption = "Facturaci�n Albaranes"
            Me.Label1(0).Caption = "Existen Albaranes que NO se van a Facturar:"
            Me.Label1(0).Top = 260
            Me.Label1(0).Left = 480
            PonerFocoBtn Me.cmdAceptarCobros
            
        Case 13 'Muestra Errores
            H = 6000
            W = 8800
            PonerFrameVisible Me.FrameErrores, True, H, W
            Me.Text1.Text = vCampos
            Me.Caption = "Errores"
        
        Case 14 'Muestra Empresas del sistema
            PonerFrameNSeriesVisible True, H, W
            Me.Caption = "Selecci�n"
            CargarListaEmpresas
        Case 15, 23
            'Etiquetas estanteria
            'PVP para TPVs
            H = FrameEtiqEstant.Height
            W = FrameEtiqEstant.Width
            PonerFrameVisible FrameEtiqEstant, True, H, W
            
            If OpcionMensaje = 23 Then
                ListView3.ColumnHeaders(3).Text = "Precio tarifa"
                ListView3.ColumnHeaders(3).Alignment = lvwColumnRight
            End If
        Case 16, 20
            
            
            Caption = "Correcci�n precios"
            H = FrameCorreccionPrecios.Height
            W = FrameCorreccionPrecios.Width
            PonerFrameVisible FrameCorreccionPrecios, True, H, W
            Me.cmdCorrecotrPrecios(1).Cancel = True
            lblIndicadorCorregir.Caption = ""
            CargaComboActualizarPrecios
            If OpcionMensaje = 20 Then
                ListView4.ColumnHeaders(9).Text = " PUC correc."
                Label2(0).Caption = " Correcci�n de precios de articulos con componentes"
            Else
                ListView4.ColumnHeaders(9).Text = "Tarifa correc."
                Label2(0).Caption = " Correcci�n de errores y actualizaci�n de tarifas"
            End If
            
        Case 18
            
            Caption = "Mantenimientos"
            H = FrameTraspasoMante.Height
            W = FrameTraspasoMante.Width
            PonerFrameVisible FrameTraspasoMante, True, H, W
            
        Case 21
            'Ver email
            limpiar Me
            H = FrameEMail.Height
            W = FrameEMail.Width
            PonerFrameVisible FrameEMail, True, H, W
            If cadWHERE2 = "0" Then
                Caption = "Enviados"
                Label5(0).Caption = "Para"
            Else
                Label5(0).Caption = "De"
                Caption = "Recibidos"
            End If
            cmdEmail.Cancel = True
            PonerFocoBtn Me.cmdEmail
            
    Case 24
        
        Caption = "Anal�tica"
        H = FrameErrorCC.Height
        W = FrameErrorCC.Width
        PonerFrameVisible FrameErrorCC, True, H, W
        PonerFocoBtn cmdSalirCC
    Case 25
        Caption = "LOTES"
        
        H = FrameArticulosAgrupados.Height
        W = FrameArticulosAgrupados.Width
        PonerFrameVisible FrameArticulosAgrupados, True, H, W
        FrameSelecArtAgrupado.visible = False
        cmdArtAgrupado(1).Cancel = True
    End Select
    'Me.cmdCancel(indFrame).Cancel = True
    Me.Height = H + 350
    Me.Width = W + 70
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub PonerFrameCobrosPtesVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Pone el Frame de Cobros Pendientes Visible y Ajustado al Formulario, y visualiza los controles
'necesario para el Informe

    H = 4600
        
    Select Case OpcionMensaje
        Case 1
            H = 5000
            W = 8600
            Me.Label1(0).Caption = "CLIENTE: " & vCampos
        Case 2
            W = 8800
            Me.cmdAceptarCobros.Top = 4000
            Me.cmdAceptarCobros.Left = 4200
        Case 5 'Componentes
            W = 6000
            H = 5000
            Me.cmdAceptarCobros.Left = 4000

        Case 6, 7 'Prefacturar Albaranes
            W = 7000
            H = 6000
            Me.cmdAceptarCobros.Top = 5400
            Me.cmdAceptarCobros.Left = 4600

        Case 10, 12 'Errores al contabilizar facturas
            H = 6000
            W = 8400
            Me.cmdAceptarCobros.Top = 5300
            Me.cmdAceptarCobros.Left = 4900
            If OpcionMensaje = 12 Then
                Me.cmdCancelarCobros.Top = 5300
                Me.cmdCancelarCobros.Left = 4600
                Me.cmdAceptarCobros.Left = 3300
                Me.Label1(1).Top = 4800
                Me.Label1(1).Left = 3400
                Me.cmdAceptarCobros.Caption = "&SI"
                Me.cmdCancelarCobros.Caption = "&NO"
            End If
    End Select
            
    PonerFrameVisible Me.FrameCobrosPtes, visible, H, W

    If visible = True Then
        Me.txtParam.visible = (OpcionMensaje = 6 Or OpcionMensaje = 7)
        Me.Label1(0).visible = (OpcionMensaje = 1) Or (OpcionMensaje = 5) Or (OpcionMensaje = 12)
        Me.cmdCancelarCobros.visible = (OpcionMensaje = 12)
        Me.Label1(1).visible = (OpcionMensaje = 12)
    End If
End Sub


Private Sub PonerFrameAcercaDeVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Pone el Frame ACERCA DE visible y Ajustado al Formulario

    Me.FrameAcercaDe.visible = visible
    If visible = True Then
        'Ajustar Tama�o del Frame para ajustar tama�o de Formulario al del Frame
        Me.FrameAcercaDe.Top = -90
        Me.FrameAcercaDe.Left = 0
        Me.FrameAcercaDe.Height = 4555
        Me.FrameAcercaDe.Width = 6600
        
        W = Me.FrameAcercaDe.Width
        H = Me.FrameAcercaDe.Height
    End If
End Sub


Private Sub PonerFrameNSeriesVisible(visible As Boolean, H As Integer, W As Integer)
'Pone el Frame de N� Serie Visible y Ajustado al Formulario, y visualiza los controles
'necesario para el Informe

    H = 5000
   
    If OpcionMensaje = 11 Then 'Lineas Factura a Rectificar
        W = 10900
    ElseIf OpcionMensaje = 14 Then
        W = 6500
        Me.Label7(1).visible = True
    ElseIf OpcionMensaje = 17 Then
        W = 10500
        Me.Label7(1).visible = False
    Else
        W = 8500
        Me.Label7(1).visible = False
    End If
    PonerFrameVisible Me.FrameNSeries, visible, H, W
End Sub


Private Sub ponerFrameComponentesVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Pone el Frame de Componentes Visible y Ajustado al Formulario, y visualiza los controles
'necesario para el Informe

'    Me.FrameComponentes.visible = visible
    Me.FrameComponentes2.visible = visible
    
    H = 4000
    W = 5300
    PonerFrameVisible Me.FrameComponentes, visible, H, W
        
    'Ajustar Tama�o del Frame para ajustar tama�o de Formulario al del Frame
    If visible Then Me.OptCompXDpto.Caption = DevuelveTextoDepto(False)
        
    
End Sub


Private Sub CargarListaCobrosPtes()
'Muestra la lista Detallada de cobros en un ListView
'Carga los valores de la tabla scobro de la Contabilidad
Dim RS As ADODB.Recordset
Dim ItmX As ListItem
Dim SQL As String
Dim Impor As Currency
Dim Borrame As Currency


    If vParamAplic.ContabilidadNueva Then
        SQL = "SELECT numserie,numfactu,fecfactu,fecvenci,impvenci,impcobro,gastos FROM "
        SQL = SQL & " cobros INNER JOIN formapago ON cobros.codforpa=formapago.codforpa "
    Else
        SQL = "SELECT numserie, codfaccl, fecfaccl, fecvenci, impvenci, impcobro ,gastos FROM "
        SQL = SQL & " scobro INNER JOIN sforpa ON scobro.codforpa=sforpa.codforpa "
    End If
    SQL = SQL & cadWhere

    Set RS = New ADODB.Recordset
    RS.Open SQL, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
     
    ListView1.Top = 900
    ListView1.Height = 3250
    ListView1.Width = 8100
    ListView1.Left = 160
    
    'Los encabezados
    ListView1.ColumnHeaders.Clear
    
    ListView1.ColumnHeaders.Add , , "Serie", 600
    ListView1.ColumnHeaders.Add , , "N� Factura", 1000, 1
    ListView1.ColumnHeaders.Add , , "Fecha Factura", 1200, 2
    ListView1.ColumnHeaders.Add , , "Fecha Venci.", 1200, 2
    ListView1.ColumnHeaders.Add , , "Imp. Venci.(�)", 1250, 1
    ListView1.ColumnHeaders.Add , , "Imp. Cobro(�)", 1250, 1
    ListView1.ColumnHeaders.Add , , "Pte. Cobro(�)", 1250, 1
   ' Borrame = 0
    While Not RS.EOF
        Set ItmX = ListView1.ListItems.Add
        ItmX.Text = RS.Fields(0).Value 'N� Serie
        ItmX.SubItems(1) = RS.Fields(1).Value 'N� Factura
        ItmX.SubItems(2) = RS.Fields(2).Value 'Fecha Factura
        ItmX.SubItems(3) = RS.Fields(3).Value 'Fecha Vencimiento
        ItmX.SubItems(4) = RS.Fields(4).Value 'Importe Vencido
        ItmX.SubItems(5) = DBLet(RS.Fields(5).Value, "N") 'Importe Cobrado
        'ItmX.SubItems(6) = RS.Fields(4).Value + DBLet(RS!gastos, "N") - DBLet(RS.Fields(5).Value, "N") 'Pendiente de cobro
        Impor = RS.Fields(4).Value + DBLet(RS!gastos, "N") - DBLet(RS.Fields(5).Value, "N") 'Pendiente de cobro
        ItmX.SubItems(6) = Impor
        If ItmX.SubItems(6) > 0 Then
            ItmX.ListSubItems(6).ForeColor = vbRed
           ' Borrame = Borrame + Impor
        End If
        
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
End Sub


Private Sub CargarListaArtSinStock(NomTabla As String)
'Muestra la lista Detallada de Articulos que no tienen stock suficiente en un ListView
Dim RS As ADODB.Recordset
Dim SQL As String
    
    SQL = "SELECT " & NomTabla & ".codalmac," & NomTabla & ".codartic, " & NomTabla & ".nomartic, salmac.canstock as canstock, SUM(cantidad) as cantidad, canstock-SUM(cantidad) as disp ,conjunto "
    SQL = SQL & " FROM ((" & NomTabla & " INNER JOIN sartic ON " & NomTabla & ".codartic=sartic.codartic) INNER JOIN sfamia ON sartic.codfamia=sfamia.codfamia) "
    SQL = SQL & " INNER JOIN salmac ON " & NomTabla & ".codalmac=salmac.codalmac and " & NomTabla & ".codartic=salmac.codartic "
    SQL = SQL & cadWhere 'Where numpedcl = 2 And sfamia.instalac = 0
    SQL = SQL & " GROUP by " & NomTabla & ".codalmac, " & NomTabla & ".codartic "
    

    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
     
    Me.ListView1.Top = 500
     
    'Los encabezados
    ListView1.Width = 8400
    ListView1.Height = 3150
    ListView1.ColumnHeaders.Clear
    
    ListView1.ColumnHeaders.Add , , "Alm.", 500
    ListView1.ColumnHeaders.Add , , "Articulo", 1800, 2
    ListView1.ColumnHeaders.Add , , "Dec. Artic", 3300
    ListView1.ColumnHeaders.Add , , "Stock", 950, 2
    ListView1.ColumnHeaders.Add , , "Cantidad", 900, 2
    ListView1.ColumnHeaders.Add , , "No Disp.", 900, 2
    
    While Not RS.EOF
        CargaItemStock RS, ""
        'Si no tiene produccion miraremos si es conjunto
        If Not vParamAplic.Produccion Then
            If RS!Conjunto = 1 Then
                SQL = RS!codAlmac & "|" & RS!codArtic & "|" & RS!cantidad & "|"
                CargaStockConjuntos SQL
            End If
        End If
        RS.MoveNext
    Wend
    RS.Close
    
    Set RS = Nothing

    
    
End Sub
    
Private Sub CargaStockConjuntos(linea As String)
    
        
        Set miRsAux = New ADODB.Recordset
            'Deberiamos cargar los elementos que tiene subconjuntos
            cadWHERE2 = "SELECT " & RecuperaValor(linea, 1) & ",sarti1.codarti1,nomartic,"
            cadWHERE2 = cadWHERE2 & " sarti1.cantidad * " & TransformaComasPuntos(RecuperaValor(linea, 3)) & " as cantidad,"
            cadWHERE2 = cadWHERE2 & " salmac.canstock as canstock,  canstock-(sarti1.cantidad * " & TransformaComasPuntos(RecuperaValor(linea, 3))
            cadWHERE2 = cadWHERE2 & ") as disp From sarti1, salmac, sartic"
            cadWHERE2 = cadWHERE2 & " Where sarti1.codarti1 = salmac.codArtic And sarti1.codarti1 = sartic.codArtic"
            cadWHERE2 = cadWHERE2 & " and sarti1.codartic='" & DevNombreSQL(RecuperaValor(linea, 2)) & "'"
            
            miRsAux.Open cadWHERE2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not miRsAux.EOF
                CargaItemStock miRsAux, " * "
                miRsAux.MoveNext
            Wend
            miRsAux.Close
        
        cadWHERE2 = ""
    Set miRsAux = Nothing
End Sub
 
    
Private Sub CargaItemStock(ByRef R As ADODB.Recordset, ByRef TxtA�adido As String)
Dim ItmX As ListItem
     If R!disp < 0 Then
            Set ItmX = ListView1.ListItems.Add
            ItmX.Text = Format(R.Fields(0).Value, "000") 'Cod Almacen
            If TxtA�adido <> "" Then TxtA�adido = "[" & TxtA�adido & "]"
            ItmX.SubItems(1) = TxtA�adido & " " & R.Fields(1).Value 'Cod Artic
            ItmX.SubItems(2) = R.Fields(2).Value 'Nom Artic
            ItmX.SubItems(3) = R.Fields(3).Value 'Stock
            ItmX.SubItems(4) = R.Fields(4).Value 'Cantidad
            ItmX.SubItems(5) = R.Fields(5).Value 'No Disp
    End If
End Sub


Private Sub CargarListaNSeries()
'Carga las lista con todos los N� de serie encontrados en la tabla:sserie
'para el articulo pasado como parametro en la cadwhere: "codartic='00012'"
'y que esten disponibles: numfactu y numalbar no tengan valor
Dim RS As ADODB.Recordset
Dim ItmX As ListItem
Dim SQL As String
Dim cadLista As String
Dim Dif As Single

    On Error GoTo ECargarLista

    If cadWHERE2 = "" Then
        'Mostramos los n� serie libres para seleccionar la cantidad
        SQL = "SELECT sserie.numserie, sserie.codartic, sartic.nomartic "
        SQL = SQL & "FROM sserie INNER JOIN sartic ON sserie.codartic=sartic.codartic "
        SQL = SQL & cadWhere 'Where codartic='000012'
        'seleccionamos los que no esten asignados a ninguna factura ni albaran
        SQL = SQL & " AND ((isnull(sserie.numfactu) or sserie.numfactu='') and (isnull(sserie.numalbar) or sserie.numalbar='')) "
        SQL = SQL & " ORDER BY sserie.codartic, numserie "
        
    Else 'venimos de modificar la cantidad y seleccionamos los ya asignados
        If InStr(1, cadWHERE2, "|") > 0 Then
            Dif = CSng(RecuperaValor(cadWHERE2, 1))
            cadWHERE2 = RecuperaValor(cadWHERE2, 2)
        
            'seleccionamos n� serie del albaran que modificamos
            SQL = "SELECT sserie.numserie, sserie.codartic, sartic.nomartic "
            SQL = SQL & "FROM sserie INNER JOIN sartic ON sserie.codartic=sartic.codartic "
            SQL = SQL & cadWHERE2
                
            
            If Dif < 0 Then
                'Si la diferencia de cantidad es < 0, mostrar en la lista los n� serie que
                'tiene la linea de albaran asignado con todos marcados y desmarcar el que no queremos
                
            Else
                'si la diferencia de cantidad es > 0, mostrar en la lista los n� de serie que
                'ya tenia asignados la linea del albaran m�s los libres para seleccionar los que a�adimos de mas
                cadLista = ""
                Set RS = New ADODB.Recordset
                RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                While Not RS.EOF
                    cadLista = cadLista & ", " & RS!numSerie
                    RS.MoveNext
                Wend
                RS.Close
                Set RS = Nothing
                
                'mostrar tambien los n� serie sin asignar
                SQL = SQL & " OR (" & Replace(cadWhere, "WHERE", "") & " and (numalbar=''or isnull(numalbar)))"
            End If
        Else
            'viene de una factura rectificativa, seleccionamos los n� de serie de
            'esa factura y marcamos los que queremos quitar
            SQL = cadWHERE2
        End If
    End If
    

    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    'Los encabezados
    ListView2.Width = 7400
    Me.ListView2.Height = 3100
    Me.ListView2.Left = 650
    ListView2.ColumnHeaders.Clear
    
    ListView2.ColumnHeaders.Add , , "N� Serie", 1800
    ListView2.ColumnHeaders.Add , , "Articulo", 1800
    ListView2.ColumnHeaders.Add , , "Desc. Artic", 3650
        
    If RS.EOF Then Unload Me
    
    While Not RS.EOF
         Set ItmX = ListView2.ListItems.Add
         ItmX.Text = RS.Fields(0).Value 'num serie
         If Dif < 0 Then
            ItmX.Checked = True
         ElseIf Dif > 0 Then
            If InStr(1, cadLista, CStr(RS!numSerie)) > 0 Then
                ItmX.Checked = True
            Else
                ItmX.Checked = False
            End If
         Else
            ItmX.Checked = False
         End If
         ItmX.SubItems(1) = RS.Fields(1).Value 'Desc Artic
         ItmX.SubItems(2) = RS.Fields(2).Value 'Nom Artic
         RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    
    
ECargarLista:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargar N� Series", Err.Description
End Sub


Private Sub CargarListaComponentes(opt As Byte)
'Muestra la lista Detallada de cobros en un ListView
'Carga los valores de la tabla scobro de la Contabilidad
Dim RS As ADODB.Recordset
Dim ItmX As ListItem
Dim SQL As String
Dim Codigo As String, cadCodigo As String

    Select Case opt
        Case 1 'Mantenimiento
            Codigo = RecuperaValor(vCampos, 1)
            If Codigo = "" Then
                cadCodigo = " isnull(nummante) "
            Else
                cadCodigo = " nummante=" & DBSet(Codigo, "T")
            End If
            SQL = ObtenerSQLcomponentes(cadWhere & " and " & cadCodigo)
            Me.Label1(0).Caption = "Mantenimiento: " & Codigo
            
        Case 2 'Departamento
            Codigo = RecuperaValor(vCampos, 2)
            If Codigo = "" Then
                cadCodigo = "isnull(coddirec)"
            Else
                cadCodigo = " coddirec=" & Codigo
            End If
            SQL = ObtenerSQLcomponentes(cadWhere & " and " & cadCodigo)
            If vParamAplic.HayDeparNuevo = 1 Then
                Me.Caption = "Equipos del Departamento"
                Me.Label1(0).Caption = " Departamento: " & RecuperaValor(vCampos, 3)
            ElseIf vParamAplic.HayDeparNuevo = 0 Then
                Me.Caption = "Equipos de la Direcci�n"
                Me.Label1(0).Caption = " Direcci�n: " & Codigo & " " & RecuperaValor(vCampos, 3)
            Else
                Me.Caption = "Equipos de la obra"
                Me.Label1(0).Caption = " Obra: " & Codigo & " " & RecuperaValor(vCampos, 3)
            End If
        
        Case 3 'Cliente
            SQL = ObtenerSQLcomponentes(cadWhere)
            Me.Caption = "Equipos del Cliente"
            Me.Label1(0).Caption = "Cliente: " & RecuperaValor(vCampos, 4)
    End Select
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    'Los encabezados
    ListView1.Top = 800
    ListView1.Left = 280
    ListView1.Width = 4900
    ListView1.Height = 3250
    ListView1.ColumnHeaders.Clear
    
    ListView1.ColumnHeaders.Add , , "TA", 760
    ListView1.ColumnHeaders.Add , , "Tipo Articulo", 2800
    ListView1.ColumnHeaders.Add , , "Cantidad", 1280, 2
    
    If Not RS.EOF Then
        While Not RS.EOF
            Set ItmX = ListView1.ListItems.Add
            ItmX.Text = RS.Fields(0).Value 'TA
            ItmX.SubItems(1) = RS.Fields(1).Value 'Tipo Articulo
            ItmX.SubItems(2) = RS.Fields(2).Value 'Cantidad
            RS.MoveNext
        Wend
    End If
    RS.Close
    Set RS = Nothing
End Sub


Private Sub CargarListaPreFacturar()
'Muestra la lista Detallada de Albaranes a Factura en un ListView
'Carga los valores de la tabla scobro de la Contabilidad
Dim RS As ADODB.Recordset
Dim ItmX As ListItem
Dim SQL As String

    On Error GoTo ECargarList
    
    SQL = "CREATE TEMPORARY TABLE tmp ( "
    SQL = SQL & "codforpa SMALLINT(3) UNSIGNED  DEFAULT '0' NOT NULL, "
    SQL = SQL & "numalbar MEDIUMINT(7) UNSIGNED  DEFAULT '0' NOT NULL, "
    SQL = SQL & "dtoppago DECIMAL(4,2) UNSIGNED  DEFAULT '0.0' NOT NULL, "
    SQL = SQL & "dtopgnral DECIMAL(4,2) UNSIGNED  DEFAULT '0.0' NOT NULL, "
    SQL = SQL & "importe DECIMAL(12,2) UNSIGNED  DEFAULT '0.0' NOT NULL, "
    SQL = SQL & "bruto DECIMAL(12,2) UNSIGNED  DEFAULT '0.0' NOT NULL) "
    conn.Execute SQL
     
'     SQL = "LOCK TABLES scaalb READ, slialb READ;"
'     Conn.Execute SQL
     
    SQL = "SELECT scaalb.codforpa, scaalb.numalbar, dtoppago, dtognral, round(sum(importel),2) as importe, round(sum(importel),2) - round(((round(sum(importel),2)*dtoppago)/100),2) - round(((round(sum(importel),2)*dtognral)/100),2) as bruto "
    SQL = SQL & " FROM (scaalb INNER JOIN slialb ON scaalb.codtipom=slialb.codtipom and scaalb.numalbar=slialb.numalbar) "
    SQL = SQL & " WHERE " & cadWhere
    SQL = SQL & " GROUP BY scaalb.numalbar "
    SQL = SQL & " ORDER BY scaalb.codforpa, scaalb.numalbar "

    SQL = " INSERT INTO tmp " & SQL
    conn.Execute SQL
     
    SQL = " SELECT tmp.codforpa, sforpa.nomforpa, sum(tmp.bruto) as bruto"
    SQL = SQL & " FROM tmp, sforpa WHERE tmp.codforpa=sforpa.codforpa "
    SQL = SQL & " GROUP BY tmp.codforpa "
    
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        ListView1.Height = 3850
        ListView1.Width = 5400
        ListView1.Left = 500
        ListView1.Top = 1200
    '    ListView1.GridLines = False
    
        'Los encabezados
        ListView1.ColumnHeaders.Clear
        
        ListView1.ColumnHeaders.Add , , " Forma de Pago", 3300
        ListView1.ColumnHeaders.Add , , "Base Imp.(�)", 2020, 1
     
        While Not RS.EOF
            Set ItmX = ListView1.ListItems.Add
            ItmX.Text = Format(RS!codforpa.Value, "000") & "  " & RS!nomforpa.Value
            
            ItmX.SubItems(1) = RS!bruto
            RS.MoveNext
        Wend
    End If
    RS.Close
    Set RS = Nothing
    
    'Borrar la tabla temporal
    SQL = " DROP TABLE IF EXISTS tmp;"
    conn.Execute SQL

ECargarList:
    If Err.Number <> 0 Then
         'Borrar la tabla temporal
        SQL = " DROP TABLE IF EXISTS tmp;"
        conn.Execute SQL
'        SQL = "UNLOCK TABLES "
'        Conn.Execute SQL
    End If
End Sub


Private Sub CargarListaClientes()
'Carga las lista con todos los clientes seleccionados en la tabla:sclien
'para imprimir etiquetas, pasando como parametro la cadwhere
Dim RS As ADODB.Recordset
Dim ItmX As ListItem
Dim SQL As String, Men As String

    On Error GoTo ECargarLista

    Select Case OpcionMensaje
    Case 8
        'CLIENTES
        SQL = "SELECT codclien,nomclien,nifclien "
        SQL = SQL & "FROM sclien "
        If cadWhere <> "" Then SQL = SQL & " WHERE " & cadWhere
        SQL = SQL & " ORDER BY codclien "
        Men = "Cliente"
    Case 9
        'PROVEEDORES
        SQL = "SELECT codprove,nomprove,nifprove "
        SQL = SQL & "FROM sprove "
        If cadWhere <> "" Then SQL = SQL & " WHERE " & cadWhere
        SQL = SQL & " ORDER BY codprove "
        Men = "Proveedor"
    Case 17
        'CLIENTES MANTENIMIENTO
        SQL = cadWhere
        Men = "Cliente"
                
    Case 22
        'CLIENTES POTENCIALES
        SQL = "SELECT codclien,nomclien,nifclien "
        SQL = SQL & "FROM sclipot "
        If cadWhere <> "" Then SQL = SQL & " WHERE " & cadWhere
        SQL = SQL & " ORDER BY codclien "
        Men = "Cli. potenciales"
    End Select
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
   
        'Los encabezados
        ListView2.Width = 9400
        ListView2.Top = 500
        ListView2.Height = 3620
        ListView2.ColumnHeaders.Clear
        
        ListView2.ColumnHeaders.Add , , Men, 1050
        ListView2.ColumnHeaders.Add , , "Nombre", 4000
        ListView2.ColumnHeaders.Add , , "NIF", 1050
       
        
        If OpcionMensaje = 17 Then
            ListView2.Width = 9400
            ListView2.Left = 500
            ListView2.ColumnHeaders.Add , , "Dpto", 550
            If vParamAplic.HayDeparNuevo = 1 Then
                ListView2.ColumnHeaders.Add , , "Departamento", 2400
            ElseIf vParamAplic.HayDeparNuevo = 0 Then
                ListView2.ColumnHeaders.Add , , "Direccion", 2400
            Else
                ListView2.ColumnHeaders.Add , , "Obra", 2400
            End If
        Else
             ListView2.Width = 7000
        End If
     If Not RS.EOF Then
        While Not RS.EOF
             Set ItmX = ListView2.ListItems.Add
             ItmX.Text = Format(RS.Fields(0).Value, "000000") 'cod clien/prove
             ItmX.Checked = False
             ItmX.SubItems(1) = RS.Fields(1).Value 'Nom clien/prove
             ItmX.SubItems(2) = DBLet(RS.Fields(2).Value, "T") 'NIF clien/prove
             
             If OpcionMensaje = 17 Then
                ItmX.SubItems(3) = DBLet(RS.Fields(3).Value, "T") 'cod dpto
                ItmX.SubItems(4) = DBLet(RS.Fields(4).Value, "T") 'nom dpto
             End If
            
             RS.MoveNext
        Wend
    End If
    RS.Close
    Set RS = Nothing
    
ECargarLista:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Cargar " & Men, Err.Description
        PulsadoSalir = True
        Unload Me
    End If
End Sub



Private Sub CargarListaErrContab()
'Muestra la lista Detallada de Facturas que dieron error al contabilizar
'en un ListView
Dim RS As ADODB.Recordset
Dim ItmX As ListItem
Dim SQL As String

    On Error GoTo ECargarList

    SQL = " SELECT  * "
    SQL = SQL & " FROM tmpErrFac "
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        ListView1.Height = 4500
        ListView1.Width = 7400
        ListView1.Left = 500
        ListView1.Top = 500

        'Los encabezados
        ListView1.ColumnHeaders.Clear

        If RS.Fields(0).Name = "codprove" Then
            'Facturas de Compra
             ListView1.ColumnHeaders.Add , , "Prove.", 700
        Else 'Facturas de Venta
            ListView1.ColumnHeaders.Add , , "Tipo", 600
        End If
        ListView1.ColumnHeaders.Add , , "Factura", 1000, 1
        ListView1.ColumnHeaders.Add , , "Fecha", 1100, 1
        ListView1.ColumnHeaders.Add , , "Error", 4620
    
        While Not RS.EOF
            Set ItmX = ListView1.ListItems.Add
            'El primer campo ser� codtipom si llamamos desde Ventas
            ' y ser� codprove si llamamos desde Compras
            ItmX.Text = RS.Fields(0).Value
            ItmX.SubItems(1) = Format(RS!Numfactu, "0000000")
            ItmX.SubItems(2) = RS!FecFactu
            ItmX.SubItems(3) = RS!Error
            RS.MoveNext
        Wend
    End If
    RS.Close
    Set RS = Nothing

ECargarList:
    If Err.Number <> 0 Then
        Err.Clear
    End If
End Sub


Private Sub CargarListaLinFactu()
'Carga las lista con todas las lineas de la factura que estamos rectificando
'seleccionamos las que nos queremos llevar al Albaran de rectificacion
Dim RS As ADODB.Recordset
Dim ItmX As ListItem
Dim SQL As String

    On Error GoTo ECargarLista

    SQL = "SELECT codtipom,numfactu,fecfactu,codtipoa,numalbar,numlinea,codalmac,codartic,nomartic,ampliaci,cantidad,precioar,dtoline1,dtoline2,importel,origpre"
    SQL = SQL & " FROM slifac "
    If cadWhere <> "" Then SQL = SQL & " WHERE " & cadWhere
        
    'HERBELCA NO DEJA traer varios para GANDIA - CASTELLONS
    If vParamAplic.NumeroInstalacion = 2 Then
        'Si el almacen es gandia y castellon NO sale si el stock es cero
        If vUsu.AlmacenPorDefecto2 = 2 Or vUsu.AlmacenPorDefecto2 = 3 Then SQL = SQL & " AND NOT codartic IN (Select codartic from sartic where artvario=1)"
    End If
    
    
    
    
    SQL = SQL & " ORDER BY codtipom,numfactu,fecfactu,codtipoa,numalbar,numlinea "
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        
        ListView2.Top = 500
        ListView2.Left = 380
        ListView2.Width = 10100
        ListView2.Height = 3620
        
        'Los encabezados
        ListView2.ColumnHeaders.Clear
    
        ListView2.ColumnHeaders.Add , , "T.Alb", 660
        ListView2.ColumnHeaders.Add , , "N� Alb", 840
        ListView2.ColumnHeaders.Add , , "Lin.", 450
         ListView2.ColumnHeaders.Item(3).Alignment = lvwColumnRight
        ListView2.ColumnHeaders.Add , , "Alm", 460
        ListView2.ColumnHeaders.Add , , "Artic", 1380
        ListView2.ColumnHeaders.Add , , "Desc. Artic.", 2500
        ListView2.ColumnHeaders.Add , , "Cant.", 600
        ListView2.ColumnHeaders.Item(7).Alignment = lvwColumnRight
        ListView2.ColumnHeaders.Add , , "Precio", 960
        ListView2.ColumnHeaders.Item(8).Alignment = lvwColumnRight
        ListView2.ColumnHeaders.Add , , "Dto 1", 600
        ListView2.ColumnHeaders.Item(9).Alignment = lvwColumnRight
        ListView2.ColumnHeaders.Add , , "Dto 2", 600
        ListView2.ColumnHeaders.Item(10).Alignment = lvwColumnRight
        ListView2.ColumnHeaders.Add , , "Importe", 950
        ListView2.ColumnHeaders.Item(11).Alignment = lvwColumnRight
    
        While Not RS.EOF
             Set ItmX = ListView2.ListItems.Add
             ItmX.Text = RS!Codtipoa 'cod tipo alb
             ItmX.Checked = False
             ItmX.SubItems(1) = Format(RS!Numalbar, "0000000") 'N� Albaran
             ItmX.SubItems(2) = RS!numlinea 'linea Albaran
             ItmX.SubItems(3) = Format(RS!codAlmac, "000") 'cod almacen
             ItmX.SubItems(4) = RS!codArtic 'Cod Articulo
             ItmX.SubItems(5) = RS!NomArtic 'Nombre del Articulo
             ItmX.SubItems(6) = RS!cantidad
             ItmX.SubItems(7) = Format(RS!precioar, FormatoPrecio)
             ItmX.SubItems(8) = RS!dtoline1
             ItmX.SubItems(9) = RS!dtoline2
             ItmX.SubItems(10) = Format(RS!ImporteL, FormatoImporte)
             RS.MoveNext
        Wend
    End If
    RS.Close
    Set RS = Nothing
    
    'si aparece la barra de desplazamiento ajustar el ancho
    If Me.ListView2.ListItems.Count > 11 Then
        Me.ListView2.ColumnHeaders(5).Width = 1200 'codartic
        Me.ListView2.ColumnHeaders(8).Width = 920  'precio
    End If
   
    If ListView2.ListItems.Count = 0 Then
        MsgBox "Ninguna linea disponible para rectificar", vbExclamation
        PulsadoSalir = True
        Unload Me
    End If
    
ECargarLista:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Cargar Lineas Factura", Err.Description
        PulsadoSalir = True
        Unload Me
    End If
End Sub




Private Sub CargarListaAlbaranes()
'Muestra la lista Detallada de Facturas que dieron error al contabilizar
'en un ListView
Dim RS As ADODB.Recordset
Dim ItmX As ListItem
Dim SQL As String

    On Error GoTo ECargarList

    SQL = cadWhere 'cadwhere ya le pasamos toda la SQL
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not RS.EOF Then
        ListView1.Height = 3900
        ListView1.Width = 7200
        ListView1.Left = 500
        ListView1.Top = 700

        'Los encabezados
        ListView1.ColumnHeaders.Clear

        ListView1.ColumnHeaders.Add , , "Tipo", 700
        ListView1.ColumnHeaders.Add , , "N� Albaran", 1000, 1
        ListView1.ColumnHeaders.Add , , "Fecha", 1100, 1
        ListView1.ColumnHeaders.Item(3).Alignment = lvwColumnCenter
        ListView1.ColumnHeaders.Add , , "Cod. Cli.", 900
        ListView1.ColumnHeaders.Add , , "Cliente", 3400
    
        While Not RS.EOF
            Set ItmX = ListView1.ListItems.Add
            ItmX.Text = RS.Fields(0).Value
            ItmX.SubItems(1) = Format(RS!Numalbar, "0000000")
            ItmX.SubItems(2) = RS!FechaAlb
            ItmX.SubItems(3) = Format(RS!codClien, "000000")
            ItmX.SubItems(4) = RS!NomClien
            RS.MoveNext
        Wend
    End If
    RS.Close
    Set RS = Nothing

ECargarList:
    If Err.Number <> 0 Then
        Err.Clear
    End If
End Sub



Private Sub CargarListaEmpresas()
'Carga las lista con todas las empresas que hay en el sistema
Dim RS As ADODB.Recordset
Dim ItmX As ListItem
Dim SQL As String
Dim I As Integer

Dim Prohibidas As String

    On Error GoTo ECargarLista

    VerEmresasProhibidas Prohibidas
    
    SQL = "Select * from usuarios.empresasariges order by codempre"
    Set ListView2.SmallIcons = frmPpal.ImageListB
    ListView2.Width = 5000
    ListView2.ColumnHeaders.Clear
    ListView2.ColumnHeaders.Add , , "Empresa", 4900
    ListView2.HideColumnHeaders = True
    ListView2.GridLines = False
    ListView2.ListItems.Clear
    
    Set RS = New ADODB.Recordset
    I = -1
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not RS.EOF
        SQL = "|" & RS!codempre & "|"
        If InStr(1, Prohibidas, SQL) = 0 Then
            Set ItmX = ListView2.ListItems.Add(, , RS!nomempre, , 5)
            ItmX.Tag = RS!codempre
            If ItmX.Tag = vEmpresa.codempre Then
                ItmX.Checked = True
                I = ItmX.index
            End If
            ItmX.ToolTipText = RS!AriGes
        End If
        RS.MoveNext
    Wend
    RS.Close
    If I > 0 Then Set ListView2.SelectedItem = ListView2.ListItems(I)

    
ECargarLista:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Cargando datos empresas", Err.Description
        PulsadoSalir = True
        Unload Me
    End If
End Sub


Private Sub VerEmresasProhibidas(ByRef VarProhibidas As String)
Dim SQL As String
Dim RS As ADODB.Recordset

On Error GoTo EVerEmresasProhibidas
    VarProhibidas = "|"
    SQL = "Select codempre from usuarios.usuarioempresasariges WHERE codusu = " & (vUsu.Codigo Mod 1000)
    SQL = SQL & " order by codempre"
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not RS.EOF
          VarProhibidas = VarProhibidas & RS!codempre & "|"
          RS.MoveNext
    Wend
    RS.Close
    Exit Sub
EVerEmresasProhibidas:
    MuestraError Err.Number, Err.Description & vbCrLf & " Consulte soporte t�cnico"
    Set RS = Nothing
End Sub



Private Sub CargaImagen()
On Error Resume Next
   ' Image2.Picture = LoadPicture(App.Path & "\minilogo.bmp")
    'Image1.Picture = LoadPicture(App.path & "\fondon.gif")
    Err.Clear
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If PulsadoSalir = False Then Cancel = 1
End Sub



Private Function ObtenerTamanyosArray() As Boolean
'Para el frame de los N� de Serie de los Articulos
'En cada indice pone en CodArtic(i) el codigo del articulo
'y en Cantidad(i) la cantidad solicitada de cada codartic
Dim I As Integer, J As Integer

    ObtenerTamanyosArray = False
    'Primero a los campos de la tabla
    TotalArray = -1
    J = 0
    Do
        I = J + 1
        J = InStr(I, vCampos, "�")
        If J > 0 Then TotalArray = TotalArray + 1
    Loop Until J = 0
    
    If TotalArray < 0 Then Exit Function
    
    'Las redimensionaremos
    ReDim codArtic(TotalArray)
    ReDim cantidad(TotalArray)
    
    ObtenerTamanyosArray = True
End Function


Private Function SeparaCampos() As Boolean
'Para el frame de los N� de Serie de los Articulos
Dim Grupo As String
Dim I As Integer
Dim J As Integer
Dim C As Integer 'Contador dentro del array

    SeparaCampos = False
    I = 0
    C = 0
    Do
        J = I + 1
        I = InStr(J, vCampos, "�")
        If I > 0 Then
            Grupo = Mid(vCampos, J, I - J)
            'Y en la martriz
            InsertaGrupo Grupo, C
            C = C + 1
        End If
    Loop Until I = 0
    SeparaCampos = True
End Function


Private Sub InsertaGrupo(Grupo As String, Contador As Integer)
Dim J As Integer
Dim cad As String

    J = 0
    cad = ""
    
    'Cod Artic
    J = InStr(1, Grupo, "|")
    If J > 0 Then
        cad = Mid(Grupo, 1, J - 1)
        Grupo = Mid(Grupo, J + 1)
        J = 1
    End If
    codArtic(Contador) = cad
    
    'Cantidad
    J = InStr(1, Grupo, "|")
    If J > 0 Then
        cad = Mid(Grupo, 1, J - 1)
        Grupo = Mid(Grupo, J + 1)
    Else
        cad = Grupo
        Grupo = ""
    End If
    cantidad(Contador) = cad
End Sub





Private Sub imgCheck_Click(index As Integer)
Dim b As Boolean
    If index < 2 Then
        'En el listview3
        b = index = 1
        For TotalArray = 1 To ListView3.ListItems.Count
            ListView3.ListItems(TotalArray).Checked = b
            If (TotalArray Mod 50) = 0 Then DoEvents
        Next TotalArray
        
    Else
        'En el listview4
        b = index = 3
        For TotalArray = 1 To ListView4.ListItems.Count
            If ListView4.ListItems(TotalArray).Tag <> "" Then
                ListView4.ListItems(TotalArray).Checked = b
            Else
                ListView4.ListItems(TotalArray).Checked = False
            End If
            If (TotalArray Mod 50) = 0 Then DoEvents
        Next TotalArray
    End If
End Sub



Private Sub Label9_Click()
        LanzaVisorMimeDocumento Me.hwnd, "http://help-ariges.ariadnasw.com/Versiones.html"
End Sub

Private Sub lwArticulosAgrupados_DblClick()
    cmdArtAgrupado_Click 0
End Sub

Private Sub lwArticulosAgrupados_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub OptCompXClien_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub OptCompXDpto_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub OptCompXMant_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim Cerrar As Boolean

    KEYpressGnral KeyAscii, 2, Cerrar
    If Cerrar Then Unload Me
End Sub



Private Function RegresarCargaEmpresas() As String
Dim SQL As String
Dim Parametros As String
Dim I As Integer

    CadenaDesdeOtroForm = ""
    
        SQL = ""
        Parametros = ""
        For I = 1 To ListView2.ListItems.Count
            If Me.ListView2.ListItems(I).Checked Then
                SQL = SQL & Me.ListView2.ListItems(I).Text & "|"
                Parametros = Parametros & "1" 'Contador
            End If
        Next I
        CadenaDesdeOtroForm = Len(Parametros) & "|" & SQL
        'Vemos las conta
        SQL = ""
        For I = 1 To ListView2.ListItems.Count
            If Me.ListView2.ListItems(I).Checked Then
                SQL = SQL & Me.ListView2.ListItems(I).Tag & "|"
            End If
        Next I
        CadenaDesdeOtroForm = CadenaDesdeOtroForm & SQL
    
    
        RegresarCargaEmpresas = CadenaDesdeOtroForm

End Function



Private Sub CargarArticulosEstanteria()
Dim SQL As String
Dim RS As ADODB.Recordset
Dim IT As ListItem
Dim RBarras As ADODB.Recordset
    
    
    Set RBarras = New ADODB.Recordset
    Label6.Caption = "Cargando"
    Label6.Refresh
    SQL = "Select * from sarti3 order by codartic,numlinea desc"
    RBarras.Open SQL, conn, adOpenKeyset, adLockPessimistic, adCmdText
    
    
    SQL = "select sartic.codartic,nomartic,preciove,codigiva,nomfamia from sartic,sfamia where sartic.codfamia=sfamia.codfamia"
    If cadWhere <> "" Then SQL = SQL & " AND " & cadWhere
    If vCampos <> "" Then SQL = SQL & " AND codartic in (Select codartic from salmac WHERE codalmac= " & vCampos & " AND  stockmin >0)"
    
    
    
    SQL = SQL & " ORDER BY sartic.codartic "
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    TotalArray = 0
    While Not RS.EOF
        Set IT = ListView3.ListItems.Add
        Label6.Caption = RS!codArtic
        Label6.Refresh
        
        RBarras.Find "codartic = " & DBSet(RS!codArtic, "T"), , adSearchForward, 1
        If RBarras.EOF Then
            SQL = ""
        Else
            SQL = RBarras!codigoea
        End If
        'Ponemos el codigo de articulo y el TIPO de IVA
        IT.Tag = "'" & DevNombreSQL(RS!codArtic) & "'," & RS!Codigiva & ",'" & SQL & "'"
        IT.Text = RS!NomArtic
        IT.SubItems(1) = Format(RS!PrecioVe, cadWHERE2)
        IT.SubItems(2) = RS!nomfamia
        IT.Checked = True
        RS.MoveNext
        TotalArray = TotalArray + 1
        If TotalArray > 300 Then
            DoEvents
            If TotalArray < 0 Then
                'Han pulsado cancelar
                
                While Not RS.EOF
                    RS.MoveNext
                Wend
                
            End If
            TotalArray = 0
        End If
    Wend
    RS.Close
    RBarras.Close
    
    
    'Febrero 2013
    'Opcion imprimir etiqetas articulo de un almacen determinado y que tengan stock minimo
    'Para ello se ha llamado al form poniendo en vCampos el codlamac
    
    
    Set RBarras = Nothing
    Set RS = Nothing
    TotalArray = 0
    Label6.Caption = ""
    
        
End Sub




Private Sub CargarArticulosCorreccionPrecio()
Dim SQL As String
Dim RS As ADODB.Recordset
Dim IT As ListItem
Dim margen As Currency
Dim MargenT As Currency
Dim ImpPVP As Currency
Dim ImpTar As Currency
Dim Aux As Currency
Dim decimales As Long
Dim precioUC As Currency
Dim SoloImporteMenor As Boolean
Dim SobreUPC As Boolean

    'El amrgen a aplicar
    'Si la tarifa es sobre el PVP es el articulo
    'si es sobre UPC entonces es sobre el de la tarifa

    lblIndicadorCorregir = "LEYENDO BD"
    lblIndicadorCorregir.Refresh
    

    
    'Si NUMREGELIM=1 entonces esta marcada la opcion(check) de solo importes menores
    If NumRegElim = 1 Then SoloImporteMenor = True
    
    
    
    'Comprobamos la tarifa donde se aplica, si sobre PVP o sobre ultima compra (%tarifa)
    SQL = DevuelveDesdeBD(conAri, "opcionINC", "starif", "codlista", vCampos)
    SobreUPC = Val(SQL) = 1
            
    
    TotalArray = InStr(1, cadWHERE2, ",")
    SQL = Mid(cadWHERE2, TotalArray + 1)
    decimales = Len(SQL)
    'Formato
    cadWHERE2 = "#,##0." & Mid(cadWHERE2, TotalArray + 1)
    
    'Sql
    SQL = " SELECT sartic.nomartic,slista.codartic,sartic.preciove,sartic.preciouc,"
    SQL = SQL & "slista.precioac, slista.codlista, starif.nomlista,"
    SQL = SQL & "sartic.margecom as margenArt,starif.margecom as margetar"
    SQL = SQL & " FROM   (slista INNER JOIN sartic ON slista.codartic=sartic.codartic)"
    SQL = SQL & " INNER JOIN starif  ON slista.codlista=starif.codlista WHERE "

    SQL = SQL & cadWhere '& " AND "
    ''SQL = SQL & " sartic.preciove <> sartic.preciouc + round((sartic.preciouc * if(isnull(sartic.margecom),0,sartic.margecom))/100," & Decimales & ")"
    
    SQL = SQL & " ORDER BY slista.codartic"
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    

    '
  
    TotalArray = 0
    
    While Not RS.EOF
        'Calculo los importes
        lblIndicadorCorregir.Caption = RS!codArtic
        lblIndicadorCorregir.Refresh
        
        margen = DBLet(RS!margenart, "N") / 100
        MargenT = DBLet(RS!margetar, "N") / 100
        precioUC = DBLet(RS!precioUC, "N")
        
        Aux = margen * precioUC
        ImpPVP = Round2(precioUC + Aux, decimales)
        
        'El de la tarifa
        If SobreUPC Then
            Aux = MargenT * precioUC
            ImpTar = Round2(precioUC + Aux, CLng(decimales))
        Else
        
            Aux = MargenT * ImpPVP
            ImpTar = Round2(ImpPVP + Aux, CLng(decimales))
        End If
        Aux = Round2(RS!PrecioVe, decimales)
        
        SQL = ""
        

        If SoloImporteMenor Then
            If Aux >= ImpPVP Then
                'El primero esta bien
                'Veamos el segundo. En la tarifa
                Aux = Round2(RS!precioac, decimales)
                If Aux < ImpTar Then SQL = "M"
            Else
                SQL = "M"
            End If
        
        
        Else
            If Aux = ImpPVP Then
                'El primero esta bien
                'Veamos el segundo. En la tarifa
                Aux = Round2(RS!precioac, decimales)
                If Aux <> ImpTar Then SQL = "M"
            Else
                SQL = "M"
            End If
        End If
        
        If SQL <> "" Then
            Set IT = ListView4.ListItems.Add
            IT.Tag = DevNombreSQL(RS!codArtic)
            IT.ToolTipText = IT.Tag
            IT.Text = IT.Tag
            IT.SubItems(1) = RS!NomArtic
            Aux = Round2(precioUC, decimales)
            IT.SubItems(2) = Format(Aux, cadWHERE2)
            
            IT.SubItems(3) = Format(margen * 100, FormatoPorcen)
            Aux = Round2(RS!PrecioVe, decimales)
            IT.SubItems(4) = Format(Aux, cadWHERE2)
            
            IT.SubItems(5) = Format(MargenT * 100, FormatoPorcen)
            Aux = Round2(RS!precioac, decimales)
            IT.SubItems(6) = Format(Aux, cadWHERE2)
            

            IT.SubItems(7) = Format(ImpPVP, cadWHERE2)
            IT.SubItems(8) = Format(ImpTar, cadWHERE2)
            
            
            
            If precioUC = 0 Then
                'Precio ultima compra =0
                'NOOOOO se puede actualizar la tarifa
                IT.Tag = "" 'para no actualizar
                IT.Checked = False
                IT.Bold = True
                IT.ForeColor = vbRed
            Else
                
            End If
            IT.Checked = False
        End If
        RS.MoveNext
        TotalArray = TotalArray + 1
        If TotalArray > 300 Then
            TotalArray = 0
            Me.Refresh
            DoEvents
        End If
    Wend
    RS.Close
    cmbActualizarTar.ListIndex = 0
    lblIndicadorCorregir.Caption = ""
End Sub




Private Function TraspasarMantenimientos() As Boolean
    
    On Error GoTo ETraspasarMantenimientos
    TraspasarMantenimientos = False

    

    cadWhere = "Select count(*) from sliman where anomante =" & txtMante(0).Text
    miRsAux.Open cadWhere, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    If Not miRsAux.EOF Then NumRegElim = DBLet(miRsAux.Fields(0), "N")
    miRsAux.Close
    
    
    If NumRegElim > 0 Then
        MsgBox "Ya existen datos para el a�o " & txtMante(0).Text, vbExclamation
        Exit Function
    End If
    
    
    
    'Se divide en 4 pasos
    '1.- Introducir una linea en la sliman con los datos para el a�o
        cadWhere = "insert into sliman (anomante,codclien,nummante,mes01man,mes02man,mes03man,mes04man,mes05man,mes06man,mes07man,mes08man,mes09man,mes10man,mes11man,mes12man)"
        cadWhere = cadWhere & " SELECT " & txtMante(0).Text & ",codclien,nummante,mes01act,mes02act,mes03act,mes04act,mes05act,mes06act,mes07act,mes08act,mes09act,mes10act,mes11act,mes12act FROM scaman"
        conn.Execute cadWhere
    '2.- Updatear los campos de actual con siguiente
        cadWhere = ""
        For TotalArray = 1 To 12
            cadWhere = cadWhere & ", mes" & Format(TotalArray, "00") & "act = mes" & Format(TotalArray, "00") & "sig"
        Next TotalArray
        cadWhere = Mid(cadWhere, 2) 'Para quitar la primera coma
        cadWhere = "UPDATE scaman SET " & cadWhere
        conn.Execute cadWhere
        
    '3.- Si no han marcado la opcion copiar datos tengo que resetear a 0
        If chkMante.Value = 0 Then
            'NO SE COPIA, luego hay que resetear
            cadWhere = ""
            For TotalArray = 1 To 12
                cadWhere = cadWhere & ", mes" & Format(TotalArray, "00") & "sig = 0 "
            Next TotalArray
            cadWhere = Mid(cadWhere, 2) 'Para quitar la primera coma
            cadWhere = "UPDATE scaman SET " & cadWhere
            conn.Execute cadWhere
        End If
        
    '4.- Ultimo mes facturado pasa a ser  cero
        conn.Execute "UPDATE scaman SET ulmesfac=0"
        
    TraspasarMantenimientos = True
    
    Exit Function
ETraspasarMantenimientos:
    MuestraError Err.Number
End Function



Private Sub txtArtAgrupado_GotFocus()
    ConseguirFoco txtArtAgrupado, 3
End Sub

Private Sub txtArtAgrupado_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtArtAgrupado_LostFocus()
    If Not PonerFormatoEntero(txtArtAgrupado) Then
        txtArtAgrupado.Text = "1"
        PonerFoco txtArtAgrupado
    End If
End Sub

Private Sub txtMante_KeyPress(index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub




Private Sub CargaPVPPreciosArticulosConComponentes()
Dim decimales As Byte
Dim SQL As String
Dim Impor As Currency
Dim IA As Currency
Dim PC As Currency
Dim PCC As Currency

    Set miRsAux = New ADODB.Recordset
    
    lblIndicadorCorregir = "LEYENDO BD"
    lblIndicadorCorregir.Refresh
    
    
    'Fomato importe
    TotalArray = InStr(1, cadWHERE2, ",")
    SQL = Mid(cadWHERE2, TotalArray + 1)
    decimales = Len(SQL)
    'Formato
    cadWHERE2 = "#,##0." & Mid(cadWHERE2, TotalArray + 1)
    
    
    'Tres columna svamos a ponerlas a tama�o 0
    ListView4.ColumnHeaders(6).Width = 0
    ListView4.ColumnHeaders(7).Width = 0
    
    SQL = "select sarti1.*,s1.nomartic,s1.preciove pre2,s1.margecom,s1.preciouc,"
    SQL = SQL & " sarti1.cantidad,s2.preciove, s2.preciouc coste"
    SQL = SQL & " from sarti1,sartic as s1,sartic as s2 where sarti1.codartic=s1.codartic and sarti1.codarti1=s2.codartic"
    'Si lleva WHERE
    If cadWhere <> "" Then
        vCampos = Replace(cadWhere, "sartic.", "s1.")
        SQL = SQL & " AND " & vCampos
        vCampos = ""
    End If
    
    SQL = SQL & " ORDER BY sarti1.codartic"
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    SQL = ""

    While Not miRsAux.EOF
        If SQL <> miRsAux!codArtic Then
            'Nuevo articulo
            lblIndicadorCorregir = miRsAux!codArtic
            lblIndicadorCorregir.Refresh
            If SQL <> "" Then
                'Si precioventa distionto   o pcompra distionto
                If IA <> Impor Or PC <> PCC Then
                    vCampos = vCampos & Format(IA, cadWHERE2) & "|" & Format(Impor, cadWHERE2) & "|"
                    vCampos = vCampos & Format(PC, cadWHERE2) & "|" & Format(PCC, cadWHERE2) & "|"
                    InsertarItemARticuloConjunto vCampos
                End If
                    
                
            End If
            SQL = miRsAux!codArtic
            vCampos = miRsAux!codArtic & "|" & miRsAux!NomArtic & "|"
            PC = DBLet(miRsAux!precioUC, "N")
            vCampos = vCampos & Format(PC, cadWHERE2)
            vCampos = vCampos & "|" & Format(DBLet(miRsAux!margecom, "N"), FormatoPorcen) & "|"
            
            IA = miRsAux!pre2
            PCC = 0 'precio compra calculado
            Impor = 0
        End If
        Impor = Impor + Round2((miRsAux!cantidad * miRsAux!PrecioVe), CLng(decimales))
        PCC = PCC + Round2((miRsAux!cantidad * DBLet(miRsAux!coste, "N")), CLng(decimales))
        miRsAux.MoveNext
    Wend
    If SQL <> "" Then
        If IA <> Impor Or PC <> PCC Then
            vCampos = vCampos & Format(IA, cadWHERE2) & "|" & Format(Impor, cadWHERE2) & "|"
            vCampos = vCampos & Format(PC, cadWHERE2) & "|" & Format(PCC, cadWHERE2) & "|"
            InsertarItemARticuloConjunto vCampos
        End If
    End If
    miRsAux.Close
    lblIndicadorCorregir = ""
End Sub



Private Sub InsertarItemARticuloConjunto(Datos As String)
Dim IT As ListItem

        Set IT = ListView4.ListItems.Add
        IT.Tag = RecuperaValor(Datos, 1)
        IT.ToolTipText = IT.Tag
        IT.Text = IT.Tag
        IT.SubItems(1) = RecuperaValor(Datos, 2)  'nomartic
    
        IT.SubItems(2) = RecuperaValor(Datos, 3)  'precio UC del articulo
        IT.SubItems(3) = RecuperaValor(Datos, 4)  ' Margen
        
        IT.SubItems(4) = RecuperaValor(Datos, 5)  'PVP articulo
        IT.SubItems(7) = RecuperaValor(Datos, 6)  'PVP calculado
        IT.SubItems(8) = RecuperaValor(Datos, 8)  'PUC calculado
        
            
End Sub

Private Sub CargaComboActualizarPrecios()
    cmbActualizarTar.Clear
    
    If OpcionMensaje = 16 Then
        'ART Y TARIFAS
        cmbActualizarTar.Tag = "Art�culos y tarifas|Solo art�culo|Solo tarifas|"
    Else
        cmbActualizarTar.Tag = "PVP y PUC|Solo PVP|Solo PUC|"
    End If
    cmbActualizarTar.AddItem RecuperaValor(cmbActualizarTar.Tag, 1)
    cmbActualizarTar.AddItem RecuperaValor(cmbActualizarTar.Tag, 2)
    cmbActualizarTar.AddItem RecuperaValor(cmbActualizarTar.Tag, 3)
    cmbActualizarTar.Tag = ""
    cmbActualizarTar.ListIndex = 0
End Sub



Private Sub CargarEmail()
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open "Select * from scrmmail WHERE " & cadWhere, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        Me.txmemail(0).Text = miRsAux!email
        
        Me.txmemail(4).Text = miRsAux!FechaHora
        Me.txmemail(1).Text = DBLet(miRsAux!asunto, "T")
        Me.txmemail(2).Text = DBLet(miRsAux!adjuntos, "T")
        Me.txmemail(3).Text = DBLet(miRsAux!cuerpo, "T")
    
    
    End If
    miRsAux.Close
    Set miRsAux = Nothing
End Sub


Private Function CargaDatosEtiquetas() As Boolean

    On Error GoTo ECargaDatosEtiquetas
    CargaDatosEtiquetas = False
    
    
    cadWhere = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute cadWhere
    
    Set miRsAux = New ADODB.Recordset
    cadWhere = "Select codprove,codalmac from tmpnlotes where codusu = " & vUsu.Codigo & " ORDER by 1,2"
    miRsAux.Open cadWhere, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    cadWhere = ""
    vCampos = "" 'Para la etiqueta
    While Not miRsAux.EOF
        'para cada cliente departamento vere el campo attetiqu
        cadWhere = "attetiqu<>"""" and coddirec "
        If IsNull(miRsAux!codAlmac) Then
            cadWhere = cadWhere & " is null"
        Else
            cadWhere = cadWhere & " = " & miRsAux!codAlmac
        End If
        cadWhere = cadWhere & " AND codclien"
        cadWhere = DevuelveDesdeBD(conAri, "attetiqu", "scaman", cadWhere, CStr(miRsAux!Codprove), "N")
        


        cadWHERE2 = "insert into `tmpinformes` (`codusu`,`codigo1`,`campo1`,`nombre1`) VALUES (" & vUsu.Codigo & ","
        cadWHERE2 = cadWHERE2 & miRsAux!Codprove & ","
        If IsNull(miRsAux!codAlmac) Then
            cadWHERE2 = cadWHERE2 & "NULL"
        Else
            cadWHERE2 = cadWHERE2 & miRsAux!codAlmac
        End If
        cadWHERE2 = cadWHERE2 & "," & DBSet(cadWhere, "T") & ")"
        NumRegElim = NumRegElim + 1
        conn.Execute cadWHERE2
            
        miRsAux.MoveNext
        
    Wend
    miRsAux.Close
        
    If NumRegElim > 0 Then CargaDatosEtiquetas = True
    
    
 
ECargaDatosEtiquetas:
    If Err.Number <> 0 Then MuestraError Err.Number, "CargaDatosEtiquetas"
    Set miRsAux = Nothing
End Function



Private Sub CargarPVPArticulos()
Dim SQL As String
Dim IT As ListItem
Dim RIVA As ADODB.Recordset
Dim Precio As Currency
Dim ImpIva As Currency

    Set RIVA = New ADODB.Recordset
    Label6.Caption = "Cargando"
    Label6.Refresh
    SQL = "Select * from tiposiva order by codigiva"
    RIVA.Open SQL, ConnConta, adOpenKeyset, adLockPessimistic, adCmdText
    
      
    SQL = "select sartic.codartic,nomartic,preciove,codigiva,nomfamia,slista.precioac,fechanue,precionu from sfamia inner join sartic"
    SQL = SQL & " on sartic.codfamia=sfamia.codfamia left join slista on slista.codartic=sartic.codartic and codlista=" & vParamAplic.CodTarifa
    SQL = SQL & " WHERE 1=1 "
    If cadWhere <> "" Then SQL = SQL & " AND " & cadWhere
    
    '
    'If vCampos <> "" Then SQL = SQL & " AND codartic in (Select codartic from salmac WHERE codalmac= " & vCampos & " AND  stockmin >0)"
    
    
    
    SQL = SQL & " ORDER BY sartic.codartic "
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    TotalArray = 0
    While Not miRsAux.EOF
        Set IT = ListView3.ListItems.Add
        Label6.Caption = miRsAux!codArtic
        Label6.Refresh
        
        
        
        '`codartic`,`numlinea`,numserie,`numlinealb`,nummante
        
        RIVA.Find "codigiva = " & miRsAux!Codigiva, , adSearchForward, 1
        ImpIva = 0
        If Not RIVA.EOF Then ImpIva = DBLet(RIVA!PorceIVA)
                
        
        
        
        
        
        cadWHERE2 = " "
        If Not IsNull(miRsAux!precioac) Then
            Precio = miRsAux!precioac
            If Not IsNull(miRsAux!fechanue) Then
                If Now >= miRsAux!fechanue Then Precio = DBLet(miRsAux!precionu, "N")
            End If
            Precio = Round(((ImpIva * Precio) / 100) + Precio, 2)
            cadWHERE2 = Format(Precio, FormatoImporte)
        End If
        
        'PVP + IVA
        Precio = Round(((ImpIva * miRsAux!PrecioVe) / 100) + miRsAux!PrecioVe, 2)
        
        
        IT.Text = miRsAux!NomArtic
        IT.SubItems(1) = Format(Precio, FormatoImporte)
        IT.SubItems(2) = cadWHERE2
        IT.Checked = True
        
        '`codartic`,`numlinea`,numserie,`numlinealb`,nummante
        SQL = "'" & DevNombreSQL(miRsAux!codArtic) & "'," & miRsAux!Codigiva & ",'" & IT.SubItems(1) & "'," & TotalArray & ",'" & IT.SubItems(2) & "')"
        IT.Tag = SQL
        
        
        miRsAux.MoveNext
        TotalArray = TotalArray + 1
        If TotalArray > 300 Then
            DoEvents
            If TotalArray < 0 Then
                'Han pulsado cancelar
                
                While Not miRsAux.EOF
                    miRsAux.MoveNext
                Wend
                
            End If
            TotalArray = 0
        End If
    Wend
    miRsAux.Close
    RIVA.Close
    
    
    'Febrero 2013
    'Opcion imprimir etiqetas articulo de un almacen determinado y que tengan stock minimo
    'Para ello se ha llamado al form poniendo en vCampos el codlamac
    
    
    Set RIVA = Nothing
    Set miRsAux = Nothing
    TotalArray = 0
    Label6.Caption = ""
    
        
End Sub




Private Sub ImprimeListadoTPV()
        
            vCampos = ""
            For NumRegElim = 1 To Me.ListView3.ListItems.Count
                '                                                En el tag YA esta grabado
                If ListView3.ListItems(NumRegElim).Checked Then
                    vCampos = vCampos & ", (" & vUsu.Codigo & "," & ListView3.ListItems(NumRegElim).Tag
                    If (NumRegElim Mod 25) = 0 Then
                        conn.Execute "insert into `tmpnseries` (`codusu`,`codartic`,`numlinea`,numserie,`numlinealb`,nummante) VALUES " & Mid(vCampos, 2) & ";"
                        vCampos = ""
                        DoEvents
                    End If
                End If
            Next NumRegElim
            If vCampos <> "" Then conn.Execute "insert into `tmpnseries` (`codusu`,`codartic`,`numlinea`,numserie,`numlinealb`,nummante) VALUES " & Mid(vCampos, 2) & ";"


End Sub

Private Sub CargaArticulosAgrupados()
Dim ItmX As ListItem
    Set miRsAux = New ADODB.Recordset
    lwArticulosAgrupados.ListItems.Clear
    cadWHERE2 = "select * from sarticAgrupado ORDER BY  idcaja"
    miRsAux.Open cadWHERE2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set ItmX = lwArticulosAgrupados.ListItems.Add()
        ItmX.Text = miRsAux.Fields(0).Value 'N� Serie
        ItmX.SubItems(1) = miRsAux.Fields(1).Value 'N� Factura
        ItmX.SubItems(2) = miRsAux.Fields(2).Value 'Fecha Factura
        ItmX.SubItems(3) = Format(miRsAux!totalmostrar, FormatoImporte) 'Fecha Vencimiento
        
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    cadWHERE2 = ""
End Sub



Private Sub PonerFechaArchivo()
    On Error GoTo ePonerFechaArchivo
    
    vCampos = App.Path & "\Ariges4.exe"
    If Dir(vCampos, vbArchive) = "" Then
        vCampos = App.Path & "\" & App.EXEName & ".exe"
        If Dir(vCampos, vbArchive) = "" Then vCampos = ""
    End If
    If vCampos <> "" Then vCampos = FileDateTime(vCampos)
        
    
    
ePonerFechaArchivo:
    If Err.Number <> 0 Then
        Err.Clear
        vCampos = ""
    End If
End Sub
