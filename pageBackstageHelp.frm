VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#17.2#0"; "Codejock.CommandBars.v17.2.0.ocx"
Begin VB.Form pageBackstageHelp 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Version 17.2.0"
   ClientHeight    =   11565
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   17115
   LinkTopic       =   "Form1"
   ScaleHeight     =   11565
   ScaleWidth      =   17115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox tabPage 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   7215
      Index           =   5
      Left            =   6200
      ScaleHeight     =   7215
      ScaleWidth      =   10815
      TabIndex        =   55
      Top             =   1320
      Visible         =   0   'False
      Width           =   10815
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   4440
         PasswordChar    =   "*"
         TabIndex        =   61
         Top             =   4320
         Width           =   3855
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   4440
         PasswordChar    =   "*"
         TabIndex        =   60
         Top             =   3360
         Width           =   3855
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   4440
         PasswordChar    =   "*"
         TabIndex        =   59
         Top             =   2040
         Width           =   3855
      End
      Begin XtremeCommandBars.BackstageSeparator BackstageSeparator4 
         Height          =   255
         Left            =   0
         TabIndex        =   56
         Top             =   0
         Width           =   6375
         _Version        =   1114114
         _ExtentX        =   11245
         _ExtentY        =   450
         _StockProps     =   2
         MarkupText      =   ""
      End
      Begin XtremeCommandBars.BackstageButton BackstageButton1 
         Height          =   630
         Left            =   8640
         TabIndex        =   62
         ToolTipText     =   "Cambiar clave acceso"
         Top             =   4200
         Width           =   690
         _Version        =   1114114
         _ExtentX        =   1217
         _ExtentY        =   1111
         _StockProps     =   79
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         EnableMarkup    =   -1  'True
         TextImageRelation=   1
         ShowShadow      =   -1  'True
         IconWidth       =   32
         Icon            =   "pageBackstageHelp.frx":0000
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Repetir contraseña"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   495
         Index           =   2
         Left            =   840
         TabIndex        =   65
         Top             =   4335
         Width           =   3495
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Nueva contraseña "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   495
         Index           =   1
         Left            =   840
         TabIndex        =   64
         Top             =   3375
         Width           =   3495
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Contraseña actual"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   495
         Index           =   0
         Left            =   840
         TabIndex        =   63
         Top             =   2055
         Width           =   3495
      End
      Begin VB.Label Label20 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cambiar claves acceso programas Ariadna Software"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005B5B5B&
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   58
         Top             =   240
         Width           =   4455
      End
      Begin VB.Label Label19 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   4440
         TabIndex        =   57
         Top             =   5280
         Width           =   4695
      End
   End
   Begin VB.PictureBox tabPage 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   9735
      Index           =   3
      Left            =   6200
      ScaleHeight     =   9735
      ScaleWidth      =   10575
      TabIndex        =   33
      Top             =   1440
      Width           =   10575
      Begin XtremeCommandBars.BackstageButton btnProtectDocument 
         Height          =   1230
         Left            =   3120
         TabIndex        =   34
         Top             =   4800
         Width           =   1290
         _Version        =   1114114
         _ExtentX        =   2275
         _ExtentY        =   2170
         _StockProps     =   79
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         EnableMarkup    =   -1  'True
         TextImageRelation=   1
         ShowShadow      =   -1  'True
         IconWidth       =   32
         Icon            =   "pageBackstageHelp.frx":106A
      End
      Begin XtremeCommandBars.BackstageButton btnManageVersions 
         Height          =   1230
         Left            =   8280
         TabIndex        =   35
         Top             =   4800
         Width           =   1290
         _Version        =   1114114
         _ExtentX        =   2275
         _ExtentY        =   2170
         _StockProps     =   79
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         EnableMarkup    =   -1  'True
         TextImageRelation=   1
         ShowShadow      =   -1  'True
         IconWidth       =   32
         Icon            =   "pageBackstageHelp.frx":20D4
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "fecver"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   4800
         TabIndex        =   53
         Top             =   5160
         Width           =   4215
      End
      Begin VB.Label Label18 
         BackColor       =   &H00FFFFFF&
         Caption         =   "www.ariadnasw.com"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005B5B5B&
         Height          =   375
         Left            =   720
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   52
         Top             =   3000
         Width           =   2535
      End
      Begin VB.Label Label16 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Acerca de ..."
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005B5B5B&
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   240
         Width           =   4695
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Segoe UI Light"
            Size            =   9.75
            Charset         =   204
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005B5B5B&
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   4920
         Width           =   4695
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tel: +34 963 805 579"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   720
         TabIndex        =   42
         Top             =   2520
         Width           =   4215
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "46007 Valencia"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   720
         TabIndex        =   41
         Top             =   2280
         Width           =   4215
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Pasaje Ventura Feliu 13, Entresuelo 2 Izquierda"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   720
         TabIndex        =   40
         Top             =   2040
         Width           =   4215
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Licencia usuario final"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   11.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005B5B5B&
         Height          =   300
         Left            =   720
         TabIndex        =   39
         Top             =   4320
         Width           =   2190
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   11.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005B5B5B&
         Height          =   300
         Left            =   4800
         TabIndex        =   38
         Top             =   4320
         Width           =   795
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "There are no previo"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   4800
         TabIndex        =   37
         Top             =   4800
         Width           =   4215
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ariadna Software "
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   21.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   615
         Left            =   720
         TabIndex        =   36
         Top             =   1320
         Width           =   3975
      End
   End
   Begin VB.PictureBox tabPage 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   6375
      Index           =   2
      Left            =   6200
      ScaleHeight     =   6375
      ScaleWidth      =   10575
      TabIndex        =   24
      Top             =   2160
      Visible         =   0   'False
      Width           =   10575
      Begin MSComctlLib.ListView ListView1 
         Height          =   4035
         Left            =   120
         TabIndex        =   27
         Top             =   720
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   7117
         View            =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
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
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Segoe UI Light"
            Size            =   9.75
            Charset         =   204
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005B5B5B&
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   4920
         Width           =   4695
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Documentos de interes"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005B5B5B&
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   4695
      End
   End
   Begin VB.PictureBox tabPage 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   6375
      Index           =   0
      Left            =   6120
      ScaleHeight     =   6375
      ScaleWidth      =   10695
      TabIndex        =   9
      Top             =   1920
      Visible         =   0   'False
      Width           =   10695
      Begin MSComctlLib.ListView ListView3 
         Height          =   4905
         Left            =   240
         TabIndex        =   11
         Top             =   1065
         Width           =   10395
         _ExtentX        =   18336
         _ExtentY        =   8652
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
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
         NumItems        =   0
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Segoe UI Light"
            Size            =   9.75
            Charset         =   204
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005B5B5B&
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   6120
         Width           =   4695
      End
      Begin VB.Label Label46 
         BackStyle       =   0  'Transparent
         Caption         =   "Concepto"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   720
         Width           =   2355
      End
      Begin VB.Label Label47 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ejercicio Actual"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2985
         TabIndex        =   16
         Top             =   360
         Width           =   3705
      End
      Begin VB.Label Label48 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ejercicio Siguiente"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6735
         TabIndex        =   15
         Top             =   360
         Width           =   3435
      End
      Begin VB.Label Label45 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Porcentaje"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5685
         TabIndex        =   14
         Top             =   720
         Width           =   1005
      End
      Begin VB.Label Label49 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Porcentaje"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9015
         TabIndex        =   13
         Top             =   720
         Width           =   1005
      End
      Begin VB.Label Label51 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8145
         TabIndex        =   12
         Top             =   720
         Width           =   795
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Informacion BBDD"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005B5B5B&
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   4695
      End
   End
   Begin VB.PictureBox tabPage 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   6375
      Index           =   1
      Left            =   6200
      ScaleHeight     =   6375
      ScaleWidth      =   10575
      TabIndex        =   20
      Top             =   2040
      Visible         =   0   'False
      Width           =   10575
      Begin MSComctlLib.ListView ListView4 
         Height          =   5145
         Left            =   120
         TabIndex        =   23
         Top             =   720
         Width           =   10395
         _ExtentX        =   18336
         _ExtentY        =   9075
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
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
         NumItems        =   0
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Usuarios conectados"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005B5B5B&
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   4695
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Segoe UI Light"
            Size            =   9.75
            Charset         =   204
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005B5B5B&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   21
         Top             =   6120
         Width           =   4695
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   6480
      Top             =   480
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin XtremeCommandBars.BackstageButton btnAcciones 
      Height          =   1095
      Index           =   3
      Left            =   360
      TabIndex        =   0
      Top             =   9000
      Width           =   5295
      _Version        =   1114114
      _ExtentX        =   9340
      _ExtentY        =   1931
      _StockProps     =   79
      Caption         =   $"pageBackstageHelp.frx":313E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextAlignment   =   4
      EnableMarkup    =   -1  'True
      ImageAlignment  =   0
      IconWidth       =   64
      Icon            =   "pageBackstageHelp.frx":31E8
   End
   Begin XtremeCommandBars.BackstageButton btnAcciones 
      Height          =   975
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   6120
      Width           =   5295
      _Version        =   1114114
      _ExtentX        =   9340
      _ExtentY        =   1720
      _StockProps     =   79
      Caption         =   $"pageBackstageHelp.frx":7252
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextAlignment   =   4
      EnableMarkup    =   -1  'True
      ImageAlignment  =   0
      IconWidth       =   64
      Icon            =   "pageBackstageHelp.frx":731A
   End
   Begin XtremeCommandBars.BackstageButton btnAcciones 
      Height          =   975
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Top             =   7320
      Width           =   5295
      _Version        =   1114114
      _ExtentX        =   9340
      _ExtentY        =   1720
      _StockProps     =   79
      Caption         =   $"pageBackstageHelp.frx":B384
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextAlignment   =   4
      EnableMarkup    =   -1  'True
      ImageAlignment  =   0
      IconWidth       =   64
      Icon            =   "pageBackstageHelp.frx":B463
   End
   Begin XtremeCommandBars.BackstageButton btnAcciones 
      Height          =   975
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   4920
      Width           =   5295
      _Version        =   1114114
      _ExtentX        =   9340
      _ExtentY        =   1720
      _StockProps     =   79
      Caption         =   $"pageBackstageHelp.frx":F4CD
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextAlignment   =   4
      EnableMarkup    =   -1  'True
      ImageAlignment  =   0
      IconWidth       =   64
      Icon            =   "pageBackstageHelp.frx":F586
   End
   Begin XtremeCommandBars.BackstageButton btnAcciones 
      Height          =   1095
      Index           =   4
      Left            =   360
      TabIndex        =   6
      Top             =   10200
      Width           =   5295
      _Version        =   1114114
      _ExtentX        =   9340
      _ExtentY        =   1931
      _StockProps     =   79
      Caption         =   $"pageBackstageHelp.frx":135F0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextAlignment   =   4
      EnableMarkup    =   -1  'True
      ImageAlignment  =   0
      IconWidth       =   64
      Icon            =   "pageBackstageHelp.frx":136B4
   End
   Begin XtremeCommandBars.BackstageSeparator BackstageSeparator1 
      Height          =   255
      Left            =   6360
      TabIndex        =   8
      Top             =   1320
      Width           =   10335
      _Version        =   1114114
      _ExtentX        =   18230
      _ExtentY        =   450
      _StockProps     =   2
      MarkupText      =   ""
   End
   Begin XtremeCommandBars.BackstageSeparator lblBackstageSeparator4 
      Height          =   11415
      Left            =   5760
      TabIndex        =   7
      Top             =   120
      Width           =   135
      _Version        =   1114114
      _ExtentX        =   238
      _ExtentY        =   20135
      _StockProps     =   2
      Vertical        =   -1  'True
      MarkupText      =   ""
   End
   Begin XtremeCommandBars.BackstageSeparator lblBackstageSeparator2 
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   8760
      Width           =   5175
      _Version        =   1114114
      _ExtentX        =   9128
      _ExtentY        =   450
      _StockProps     =   2
      MarkupText      =   ""
   End
   Begin XtremeCommandBars.BackstageButton btnAcciones 
      Height          =   975
      Index           =   5
      Left            =   240
      TabIndex        =   29
      Top             =   600
      Width           =   5295
      _Version        =   1114114
      _ExtentX        =   9340
      _ExtentY        =   1720
      _StockProps     =   79
      Caption         =   $"pageBackstageHelp.frx":1771E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextAlignment   =   4
      EnableMarkup    =   -1  'True
      ImageAlignment  =   0
      IconWidth       =   48
      Icon            =   "pageBackstageHelp.frx":177E3
   End
   Begin XtremeCommandBars.BackstageSeparator BackstageSeparator2 
      Height          =   255
      Left            =   240
      TabIndex        =   30
      Top             =   360
      Width           =   5175
      _Version        =   1114114
      _ExtentX        =   9128
      _ExtentY        =   450
      _StockProps     =   2
      MarkupText      =   ""
   End
   Begin XtremeCommandBars.BackstageButton btnAcciones 
      Height          =   975
      Index           =   6
      Left            =   240
      TabIndex        =   32
      Top             =   1800
      Width           =   5295
      _Version        =   1114114
      _ExtentX        =   9340
      _ExtentY        =   1720
      _StockProps     =   79
      Caption         =   $"pageBackstageHelp.frx":19E03
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextAlignment   =   4
      EnableMarkup    =   -1  'True
      ImageAlignment  =   0
      IconWidth       =   48
      Icon            =   "pageBackstageHelp.frx":19EC4
   End
   Begin XtremeCommandBars.BackstageSeparator BackstageSeparator3 
      Height          =   255
      Left            =   240
      TabIndex        =   50
      Top             =   4680
      Width           =   5175
      _Version        =   1114114
      _ExtentX        =   9128
      _ExtentY        =   450
      _StockProps     =   2
      MarkupText      =   ""
   End
   Begin XtremeCommandBars.BackstageButton btnAcciones 
      Height          =   975
      Index           =   7
      Left            =   240
      TabIndex        =   54
      Top             =   3240
      Width           =   5295
      _Version        =   1114114
      _ExtentX        =   9340
      _ExtentY        =   1720
      _StockProps     =   79
      Caption         =   $"pageBackstageHelp.frx":1C32E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextAlignment   =   4
      EnableMarkup    =   -1  'True
      ImageAlignment  =   0
      IconWidth       =   48
      Icon            =   "pageBackstageHelp.frx":1C3F8
   End
   Begin VB.PictureBox tabPage 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   7215
      Index           =   4
      Left            =   6200
      ScaleHeight     =   7215
      ScaleWidth      =   10815
      TabIndex        =   45
      Top             =   1320
      Visible         =   0   'False
      Width           =   10815
      Begin XtremeCommandBars.BackstageSeparator lblBackstageSeparator111 
         Height          =   255
         Left            =   0
         TabIndex        =   46
         Top             =   0
         Width           =   4695
         _Version        =   1114114
         _ExtentX        =   8281
         _ExtentY        =   450
         _StockProps     =   2
         MarkupText      =   ""
      End
      Begin MSComctlLib.ListView ListViewEmpresa 
         Height          =   6015
         Left            =   120
         TabIndex        =   47
         Top             =   840
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   10610
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
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID"
            Object.Width           =   317
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   317
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Resumido"
            Object.Width           =   317
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Inicio"
            Object.Width           =   317
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Fin"
            Object.Width           =   317
         EndProperty
      End
      Begin VB.Label Label17 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Segoe UI Light"
            Size            =   9.75
            Charset         =   204
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005B5B5B&
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   4920
         Width           =   4695
      End
      Begin VB.Label Label111 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cambiar empresa"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005B5B5B&
         Height          =   255
         Left            =   0
         TabIndex        =   48
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ayuda"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H005B5B5B&
      Height          =   255
      Left            =   240
      TabIndex        =   51
      Top             =   4440
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   1065
      Left            =   12840
      Picture         =   "pageBackstageHelp.frx":1E862
      Stretch         =   -1  'True
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Acerca de ..."
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H005B5B5B&
      Height          =   255
      Left            =   240
      TabIndex        =   31
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Segoe UI Light"
         Size            =   9.75
         Charset         =   204
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H005B5B5B&
      Height          =   255
      Left            =   6360
      TabIndex        =   28
      Top             =   1680
      Visible         =   0   'False
      Width           =   8175
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Leyendo datos"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H005B5B5B&
      Height          =   255
      Left            =   6360
      TabIndex        =   18
      Top             =   2760
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Base de datos"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H005B5B5B&
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   8520
      Width           =   2535
   End
End
Attribute VB_Name = "pageBackstageHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BackstageButton1_Click()
    Msg = ""
    For i = 0 To 2
        If Trim(Text1(i).Text) = "" Then Msg = "Campos obligatorios"
    Next
    
    If Msg = "" Then
        If Trim(Text1(1).Text) <> Trim(Text1(2).Text) Then
            Msg = "Valores distintos en nueva contraseña"
        Else
            If Trim(Text1(0).Text) <> vUsu.PasswdPROPIO Then Msg = "Clave actual incorrecta"
        End If
            
    End If
    
    If Msg <> "" Then
        Me.Label19.Caption = Msg
    Else
        'Actualizamos password y limpiamos
         Msg = "UPDATE usuarios.usuarios SET passwordpropio = " & DBSet(Text1(1).Text, "T")
         Msg = Msg & " WHERE codusu =" & vUsu.Id
            
        Screen.MousePointer = vbHourglass
        Me.Label19.Caption = "Cambiando contraseña"
        Label19.Refresh
        Espera 0.25
        If ejecutar(CStr(Msg), False) Then
            Label19.ForeColor = vbBlue
            Me.Label19.Caption = "Cambio OK .."
            Label19.Refresh
            Espera 0.5
            Me.Label19.Caption = "Cambio OK ...."
            Label19.Refresh
            Espera 0.25
            vUsu.PasswdPROPIO = Text1(1).Text
            Label19.Caption = ""
            btnAcciones_Click 5
            Label19.ForeColor = vbRed
           
        End If
        Screen.MousePointer = vbDefault
    End If
    
End Sub

Private Sub btnManageVersions_Click()
    frmppalN.OpcionesMenuInformacion 1 'ID_Ver_CambiosVersion  FALTA###
End Sub

Private Sub btnProtectDocument_Click()
    frmppalN.OpcionesMenuInformacion 2 'ID_Licencia_Usuario_Final_web FALTA###
End Sub

Private Sub Form_Load()
    Label9.Caption = vEmpresa.nomempre
    Label6(0).Caption = App.Major & "." & App.Minor & "." & App.Revision
    FechaModfichero
    
End Sub

Private Sub Form_Resize()
On Error Resume Next
    lblBackstageSeparator4.Height = Me.ScaleHeight
    i = Me.Width - tabPage(4).Left - 240
    If i < 300 Then i = 300
    tabPage(4).Width = i
    
    
    i = i - ListViewEmpresa.Left - 120
    If i < 100 Then i = 100
    ListViewEmpresa.Width = i
    
    
    Me.Image1.Left = Me.Width - Image1.Width - 120
End Sub


Private Sub btnAcciones_Click(Index As Integer)
    Screen.MousePointer = vbHourglass
    tabPage(0).visible = False
    tabPage(1).visible = False
    tabPage(2).visible = False
    tabPage(3).visible = False
    tabPage(4).visible = False
    tabPage(5).visible = False
    
    Label9.visible = False
    Select Case Index
       
        Case 2
            'CAlendario del contribuyente
            LanzaVisorMimeDocumento Me.hwnd, "http://www.agenciatributaria.es/AEAT.internet/Bibl_virtual/folletos/calendario_contribuyente.shtml"
            tabPage(3).visible = True
            
        Case 1 ' documentos
            
            Label3.visible = True
            Label3.Refresh
            cargaDocumentos
            ListView1.Refresh
            tabPage(2).visible = True
            Label3.visible = False
        Case 0 ' ayuda
            tabPage(3).visible = True
            LanzaVisorMimeDocumento Me.hwnd, DireccionAyuda & "/Ariconta-6.html"  ' "http://www.ariadnasw.com/clientes/"
            
        Case 10 ' arimailges.exe
       '
        Case 3 ' Informacion de la base de datos
            Label3.visible = True
            Label3.Refresh
            If CargarInformacionBBDD Then
                CargaInformeBBDD
                Label9.visible = True
                tabPage(0).visible = True
            End If
            Label3.visible = False
            
        Case 4 'Usuarios activos
            Label3.visible = True
            Label3.Refresh
            CargaShowProcessList
            Label9.visible = True
            tabPage(1).visible = True
            Label3.visible = False
 
        Case 5
            
             tabPage(3).visible = True
 
 
        Case 6
            'Label3.Visible = True
            'Label3.Refresh
            BuscaEmpresas
            'Label9.Visible = True
            tabPage(4).visible = True
            Label3.visible = False
            
        Case 7
            Text1(0).Text = "": Text1(1).Text = "": Text1(2).Text = ""
            Label19.Caption = ""
            tabPage(5).visible = True
            PonerFoco Text1(0)
    End Select
    
    Screen.MousePointer = vbDefault
End Sub


Private Function CargarInformacionBBDD() As String
Dim Sql As String
Dim Sql2 As String
Dim CadValues As String
Dim NroRegistros As Long
Dim NroRegistrosSig As Long
Dim NroRegistrosTot As Long
Dim NroRegistrosTotSig As Long
Dim FecIniSig As Date
Dim FecFinSig As Date
Dim Porcen1 As Currency
Dim Porcen2 As Currency
Dim RS As ADODB.Recordset

    On Error GoTo eCargarInformacionBBDD
    
    CargarInformacionBBDD = False
    
'    SQL = "delete from tmpinfbbdd where codusu = " & vUsu.Codigo
'    conn.Execute SQL
'
'    FecIniSig = DateAdd("yyyy", 1, vParam.FechaIni)
'    FecFinSig = DateAdd("yyyy", 1, vParam.FechaFin)
'
'    Sql2 = "insert into tmpinfbbdd (codusu,posicion,concepto,nactual,poractual,nsiguiente,porsiguiente) values "
'
'    'asientos
'    SQL = "select count(*) from hcabapu where fechaent between " & DBSet(vParam.FechaIni, "F") & " and " & DBSet(vParam.FechaFin, "F")
'    NroRegistros = DevuelveValor(SQL)
'    SQL = "select count(*) from hcabapu where fechaent between " & DBSet(FecIniSig, "F") & " and " & DBSet(FecFinSig, "F")
'    NroRegistrosSig = DevuelveValor(SQL)
'
'    CadValues = "(" & vUsu.Codigo & ",1,'Asientos'," & DBSet(NroRegistros, "N") & ",0," & DBSet(NroRegistrosSig, "N") & ",0)"
'    conn.Execute Sql2 & CadValues
'
'    'apuntes
'    SQL = "select count(*) from hlinapu where fechaent between " & DBSet(vParam.FechaIni, "F") & " and " & DBSet(vParam.FechaFin, "F")
'    NroRegistros = DevuelveValor(SQL)
'    SQL = "select count(*) from hlinapu where fechaent between " & DBSet(FecIniSig, "F") & " and " & DBSet(FecFinSig, "F")
'    NroRegistrosSig = DevuelveValor(SQL)
'
'    CadValues = "(" & vUsu.Codigo & ",2,'Apuntes'," & DBSet(NroRegistros, "N") & ",0," & DBSet(NroRegistrosSig, "N") & ",0)"
'    conn.Execute Sql2 & CadValues
'
'
'    If vEmpresa.TieneContabilidad Then
'            'facturas de venta
'            SQL = "select count(*) from factcli where "
'            SQL = SQL & " fecfactu between " & DBSet(vParam.FechaIni, "F") & " and " & DBSet(vParam.FechaFin, "F")
'
'            NroRegistrosTot = DevuelveValor(SQL)
'
'
'            SQL = "select count(*) from factcli where "
'            SQL = SQL & " fecfactu between " & DBSet(FecIniSig, "F") & " and " & DBSet(FecFinSig, "F")
'
'            NroRegistrosTotSig = DevuelveValor(SQL)
'
'            i = 3
'
'            SQL = "select * from contadores where not tiporegi in ('0','1')"
'
'            Set Rs = New ADODB.Recordset
'            Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'            While Not Rs.EOF
'
'                SQL = "select count(*) from factcli where numserie = " & DBSet(Rs!tiporegi, "T")
'                SQL = SQL & " and fecfactu between " & DBSet(vParam.FechaIni, "F") & " and " & DBSet(vParam.FechaFin, "F")
'
'                NroRegistros = DevuelveValor(SQL)
'                Porcen1 = 0
'                If NroRegistrosTot <> 0 Then
'                    Porcen1 = Round(NroRegistros * 100 / NroRegistrosTot, 2)
'                End If
'
'                SQL = "select count(*) from factcli where numserie = " & DBSet(Rs!tiporegi, "T")
'                SQL = SQL & " and fecfactu between " & DBSet(FecIniSig, "F") & " and " & DBSet(FecFinSig, "F")
'
'                NroRegistrosSig = DevuelveValor(SQL)
'                Porcen2 = 0
'                If NroRegistrosTotSig <> 0 Then
'                    Porcen2 = Round(NroRegistrosSig * 100 / NroRegistrosTotSig, 2)
'                End If
'
'                CadValues = "(" & vUsu.Codigo & "," & DBSet(i, "N") & "," & DBSet(Rs!nomregis, "T") & "," & DBSet(NroRegistros, "N") & "," & DBSet(Porcen1, "N") & ","
'                CadValues = CadValues & DBSet(NroRegistrosSig, "N") & "," & DBSet(Porcen2, "N") & ")"
'                conn.Execute Sql2 & CadValues
'
'                i = i + 1
'
'                Rs.MoveNext
'            Wend
'
'            Set Rs = Nothing
'
'            'facturas de proveedor
'            i = i + 1
'
'            SQL = "select count(*) from factpro where fecharec between " & DBSet(vParam.FechaIni, "F") & " and " & DBSet(vParam.FechaFin, "F")
'            NroRegistros = DevuelveValor(SQL)
'            SQL = "select count(*) from factpro where fecharec between " & DBSet(FecIniSig, "F") & " and " & DBSet(FecFinSig, "F")
'            NroRegistrosSig = DevuelveValor(SQL)
'
'            CadValues = "(" & vUsu.Codigo & "," & DBSet(i, "N") & ",'Facturas Proveedores'," & DBSet(NroRegistros, "N") & ",0,"
'            CadValues = CadValues & DBSet(NroRegistrosSig, "N") & ",0)"
'
'            conn.Execute Sql2 & CadValues
'
'
'    End If
'
    
    CargarInformacionBBDD = True
    Exit Function


eCargarInformacionBBDD:
    MuestraError Err.Number, "Cargar Temporal de BBDD", Err.Description
End Function






Private Sub CargaInformeBBDD()
Dim IT As ListItem
Dim TotalArray  As Long
    On Error GoTo ECargaInformeBBDD
    
    Set RS = New ADODB.Recordset
    
    Sql = "select * from tmpinfbbdd where codusu = " & vUsu.Codigo & " order by posicion "
    
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    ListView3.ListItems.Clear
    ListView3.ColumnHeaders.Clear
    ListView3.ColumnHeaders.Add , , "CONCEPTO", 3500.0631
    ListView3.ColumnHeaders.Add , , "count ACTUAL", 2250.2522, 1
    ListView3.ColumnHeaders.Add , , "porcen ACTUAL", 1000.2522, 1
    ListView3.ColumnHeaders.Add , , "count siguiente", 2250.2522, 1
    ListView3.ColumnHeaders.Add , , "porcen siguiente", 1000.2522, 1
    
    
    
    
    TotalArray = 0
    While Not RS.EOF
        Set IT = ListView3.ListItems.Add
        
        IT.Text = UCase(DBLet(RS!Concepto, "T"))
        
        If DBLet(RS!posicion, "N") > 2 Then
            IT.SubItems(1) = Format(DBLet(RS!nactual, "N"), "###,###,###,##0")
            IT.SubItems(2) = Format(DBLet(RS!Poractual, "N"), "##0.00") & "%"
            IT.SubItems(3) = Format(DBLet(RS!nsiguiente, "N"), "###,###,###,##0")
            IT.SubItems(4) = Format(DBLet(RS!Porsiguiente, "N"), "##0.00") & "%"
        Else
            IT.SubItems(1) = Format(DBLet(RS!nactual, "N"), "###,###,###,##0")
            IT.SubItems(3) = Format(DBLet(RS!nsiguiente, "N"), "###,###,###,##0")
        End If
        
        RS.MoveNext
    Wend
    
    RS.Close
    
    Label5.Caption = Format(Now, "dd/mm/yyyy hh:nn:ss")
    Exit Sub
    
ECargaInformeBBDD:
    MuestraError Err.Number, Err.Description
    Set RS = Nothing
End Sub




Private Sub CargaShowProcessList()
Dim IT As ListItem
Dim TotalArray  As Long
Dim SERVER As String
Dim EquipoConBD As Boolean
Dim Cad As String
Dim Equipo As String

    On Error GoTo ECargaShowProcessList
    
    Set RS = New ADODB.Recordset
    
    ListView4.ListItems.Clear
    ListView4.ColumnHeaders.Clear
    
    ListView4.ColumnHeaders.Add , , "ID", 1500.0631
    ListView4.ColumnHeaders.Add , , "User", 2250.2522, 1
    ListView4.ColumnHeaders.Add , , "Host", 3000.2522, 1
    ListView4.ColumnHeaders.Add , , "Tiempo espera", 3050.2522, 1
    
    
    Set RS = New ADODB.Recordset
    
    SERVER = Mid(conn.ConnectionString, InStr(LCase(conn.ConnectionString), "server=") + 7)
    SERVER = Mid(SERVER, 1, InStr(1, SERVER, ";"))
    
    EquipoConBD = (UCase(vUsu.PC) = UCase(SERVER)) Or (LCase(SERVER) = "localhost")
    
    Cad = "show full processlist"
    RS.Open Cad, conn, adOpenKeyset, adLockOptimistic, adCmdText
    Cad = ""
    While Not RS.EOF
        If Not IsNull(RS.Fields(3)) Then
            If InStr(1, RS.Fields(3), "ariconta") <> 0 Then
                If UCase(RS.Fields(3)) = UCase(vUsu.CadenaConexion) Then
                    Equipo = RS.Fields(2)
                    'Primero quitamos los dos puntos del puerto
                    NumRegElim = InStr(1, Equipo, ":")
                    If NumRegElim > 0 Then Equipo = Mid(Equipo, 1, NumRegElim - 1)
                    
                    'El punto del dominio
                    NumRegElim = InStr(1, Equipo, ".")
                    If NumRegElim > 0 Then Equipo = Mid(Equipo, 1, NumRegElim - 1)
                    
                    Equipo = UCase(Equipo)
                    
                    
                    Set IT = ListView4.ListItems.Add
                    
                    IT.Text = RS.Fields(0)
                    IT.SubItems(1) = RS.Fields(1)
                    IT.SubItems(2) = Equipo
                    
                    'tiempo de espera
                    Dim FechaAnt As Date
                    FechaAnt = DateAdd("s", RS.Fields(5), Now)
                    IT.SubItems(3) = Format((Now - FechaAnt), "hh:mm:ss")
                End If
            End If
        End If
        
        'Siguiente
        RS.MoveNext
    Wend
    NumRegElim = 0
    RS.Close
    Set RS = Nothing
    'Label6.Caption = Format(Now, "dd/mm/yyyy hh:nn:ss")
    
    
    Exit Sub
    
ECargaShowProcessList:
    MuestraError Err.Number, Err.Description
    Set RS = Nothing
End Sub


Private Sub cargaDocumentos()
Dim RN As ADODB.Recordset
Dim Cad As String
Dim IT As ListItem

    'FALTA###
    'Set Me.ListView1.SmallIcons = frmppal.ImageList1 'frmPpal.ImageListDocumentos
    On Error GoTo eCargadocumentos
    Cad = "select iddocumento,nombrearchi from usuarios.wfichdocs WHERE aplicacion='ariconta' order by orden "
    Set RN = New ADODB.Recordset
    RN.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RN.EOF
        ListView1.ListItems.Add , "D" & Format(RN!iddocumento, "00000"), RN!nombrearchi, , 7   '1:PDF
       
        RN.MoveNext
    Wend
    RN.Close
    
eCargadocumentos:
    Err.Clear
    If ListView1.ListItems.Count = 0 Then
        Label8.Caption = "No hay documentacion disponible"
        Label8.visible = True
    Else
        Label8.visible = False
    End If
End Sub

Private Sub Label18_Click()
    LanzaVisorMimeDocumento Me.hwnd, "http://www.ariadnasw.com"
                
End Sub

Private Sub ListView1_DblClick()
Dim abrir As Boolean

    If Me.ListView1.SelectedItem Is Nothing Then Exit Sub
    
    abrir = False 'antes \ImgFicFT
    If Dir(App.Path & "\temp\" & ListView1.SelectedItem & ".pdf", vbArchive) = "" Then
        Adodc1.ConnectionString = conn
        Adodc1.RecordSource = "Select * from usuarios.wfichdocs where idDocumento=" & Mid(ListView1.SelectedItem.Key, 2)
        Adodc1.Refresh

        If LeerBinary(Adodc1.Recordset!campo, App.Path & "\temp\" & ListView1.SelectedItem.Text & ".pdf") Then abrir = True
    Else
        abrir = True
        
    End If
    
    If abrir Then LanzaVisorMimeDocumento Me.hwnd, App.Path & "\temp\" & ListView1.SelectedItem & ".pdf"
        
End Sub




Private Sub BuscaEmpresas()
Dim Prohibidas As String
Dim RS As ADODB.Recordset
Dim Cad As String
Dim Sql As String
Dim ItmX As ListItem

ListView1.ListItems.Clear

Dim i As Integer


    If Me.ListViewEmpresa.Tag = "1" Then Exit Sub
    
    
    'Ajustamos las columnas "Inamovibles"
    ListViewEmpresa.ColumnHeaders(1).Width = 800
    ListViewEmpresa.ColumnHeaders(4).Width = 1400
    ListViewEmpresa.ColumnHeaders(5).Width = 1400
    
    i = ListViewEmpresa.Width - 3840
    If i < 0 Then i = 180
    i = CInt(i / 4)
    ListViewEmpresa.ColumnHeaders(2).Width = i * 3
    ListViewEmpresa.ColumnHeaders(3).Width = i
        
 
    
    
    
    
    'Cargamos las prohibidas
    Prohibidas = DevuelveProhibidas
    
    'Cargamos las empresas
    
    Set RS = New ADODB.Recordset
    
    '[Monica]11/04/2014: solo debe de salir las ariconta
    RS.Open "Select * from usuarios.empresasariconta where conta like 'ariconta%' ORDER BY Codempre", conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    While Not RS.EOF
        Cad = "|" & RS!codempre & "|"
        If InStr(1, Prohibidas, Cad) = 0 Then
            Cad = RS!nomempre
            Set ItmX = ListViewEmpresa.ListItems.Add()
            ItmX.Text = RS!codempre
            
            
            ItmX.SubItems(1) = RS!nomempre
            ItmX.SubItems(2) = RS!nomresum
            Cad = "fechafin"
            Sql = DevuelveDesdeBD(conAri, "fechaini", "ariconta" & RS!codempre & ".parametros", "1", "1", "N", Cad)
            ItmX.SubItems(3) = Sql
            ItmX.SubItems(4) = Cad
            
                
            Cad = RS!CONTA & "|" & RS!nomresum '& "|" & Rs!Usuario & "|" & Rs!Pass & "|"
            
            If RS!codempre = vEmpresa.codempre Then
                ItmX.Bold = True
                Set ListView1.SelectedItem = ItmX
            End If
                
           ' ItmX.Tag = Cad
           ' ItmX.ToolTipText = Rs!CONTA
            
            
            'Si el codconta > 100 son empresas que viene del cambio del plan contable.
            'Atenuare su visibilidad
            If RS!codempre > 100 Then
                ItmX.ForeColor = &H808080
                ItmX.ListSubItems(1).ForeColor = &H808080
                ItmX.ListSubItems(2).ForeColor = &H808080
                ItmX.ListSubItems(3).ForeColor = &H808080
                'ItmX.SmallIcon = 2
            Else
                'normal
                'ItmX.SmallIcon = 1
            End If
        End If
        RS.MoveNext
    Wend
    RS.Close
    ListViewEmpresa.Tag = "1"
End Sub


Private Function DevuelveProhibidas() As String
Dim RS As ADODB.Recordset
Dim Cad As String
Dim i As Integer
    On Error GoTo EDevuelveProhibidas
    DevuelveProhibidas = ""
    Set RS = New ADODB.Recordset
    i = vUsu.Codigo Mod 1000
    RS.Open "Select * from usuarios.usuarioempresasariconta WHERE codusu =" & i, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Cad = ""
    While Not RS.EOF
        Cad = Cad & RS.Fields(1) & "|"
        RS.MoveNext
    Wend
    If Cad <> "" Then Cad = "|" & Cad
    RS.Close
    DevuelveProhibidas = Cad
EDevuelveProhibidas:
    Err.Clear
    Set RS = Nothing
End Function


Private Sub ListViewEmpresa_DblClick()
    ' If Not ItemCheck Then Exit Sub
    If ListViewEmpresa.SelectedItem Is Nothing Then Exit Sub
   frmppalN.CambiarEmpresa_ CInt(ListViewEmpresa.SelectedItem.Text)
End Sub


Private Sub FechaModfichero()
    On Error GoTo efech
    Label6(5).Caption = ""
    
    Dim fso, Archivo As Object


Set fso = CreateObject("Scripting.FileSystemObject")
Set Archivo = fso.GetFile(App.Path & "\" & App.EXEName & ".exe")


  Label6(5).Caption = Archivo.DateLastModified



    
    
    
efech:
    Set Archivo = Nothing
    Set fso = Nothing
    
    Err.Clear
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 3, False
End Sub
