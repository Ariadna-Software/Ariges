VERSION 5.00
Begin VB.Form frmFacEulerDatosRep 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   8550
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12795
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8550
   ScaleWidth      =   12795
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboRecepAgenClien 
      Height          =   315
      ItemData        =   "frmFacEulerDatosRep.frx":0000
      Left            =   4800
      List            =   "frmFacEulerDatosRep.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   99
      Top             =   600
      Width           =   2895
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   255
      Left            =   10200
      TabIndex        =   96
      Top             =   600
      Width           =   3255
      Begin VB.OptionButton optEule_R 
         Caption         =   "Cliente"
         Height          =   195
         Index           =   1
         Left            =   1800
         TabIndex        =   98
         Top             =   0
         Width           =   975
      End
      Begin VB.OptionButton optEule_R 
         Caption         =   "Agencia"
         Height          =   195
         Index           =   0
         Left            =   0
         TabIndex        =   97
         Top             =   0
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdBuscar 
      Cancel          =   -1  'True
      Caption         =   "Buscar"
      Height          =   375
      Left            =   10080
      TabIndex        =   95
      Top             =   8040
      Width           =   1215
   End
   Begin VB.TextBox txtEule_R 
      Height          =   315
      Index           =   27
      Left            =   10440
      MaxLength       =   16
      TabIndex        =   92
      Text            =   "Text1"
      Top             =   7440
      Width           =   2175
   End
   Begin VB.TextBox txtEule_R 
      Height          =   315
      Index           =   26
      Left            =   10440
      MaxLength       =   16
      TabIndex        =   91
      Text            =   "Text1"
      Top             =   7080
      Width           =   2175
   End
   Begin VB.TextBox txtEule_R 
      Height          =   315
      Index           =   25
      Left            =   5880
      MaxLength       =   16
      TabIndex        =   88
      Text            =   "Text1"
      Top             =   7440
      Width           =   2175
   End
   Begin VB.TextBox txtEule_R 
      Height          =   315
      Index           =   24
      Left            =   5880
      MaxLength       =   16
      TabIndex        =   83
      Text            =   "Text1"
      Top             =   7080
      Width           =   2175
   End
   Begin VB.TextBox txtEule_R 
      Height          =   315
      Index           =   23
      Left            =   2280
      MaxLength       =   16
      TabIndex        =   79
      Text            =   "Text1"
      Top             =   7440
      Width           =   1575
   End
   Begin VB.TextBox txtEule_R 
      Height          =   315
      Index           =   22
      Left            =   2280
      MaxLength       =   50
      TabIndex        =   78
      Text            =   "Text1"
      Top             =   7080
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   11400
      TabIndex        =   76
      Top             =   8040
      Width           =   1215
   End
   Begin VB.TextBox txtEule_R 
      Height          =   315
      Index           =   11
      Left            =   4080
      MaxLength       =   16
      TabIndex        =   72
      Text            =   "Text1"
      Top             =   6120
      Width           =   855
   End
   Begin VB.ComboBox cboEulerUdR 
      Height          =   315
      ItemData        =   "frmFacEulerDatosRep.frx":003D
      Left            =   5040
      List            =   "frmFacEulerDatosRep.frx":004A
      Style           =   2  'Dropdown List
      TabIndex        =   71
      Top             =   6120
      Width           =   735
   End
   Begin VB.TextBox txtEule_R 
      Height          =   315
      Index           =   17
      Left            =   7920
      MaxLength       =   16
      TabIndex        =   70
      Text            =   "Text1"
      Top             =   6000
      Width           =   855
   End
   Begin VB.TextBox txtEule_R 
      Height          =   315
      Index           =   18
      Left            =   9840
      MaxLength       =   16
      TabIndex        =   69
      Text            =   "Text1"
      Top             =   6000
      Width           =   855
   End
   Begin VB.TextBox txtEule_R 
      Height          =   315
      Index           =   19
      Left            =   11760
      MaxLength       =   16
      TabIndex        =   68
      Text            =   "Text1"
      Top             =   6000
      Width           =   855
   End
   Begin VB.OptionButton optEule_R 
      Caption         =   "C"
      Height          =   195
      Index           =   4
      Left            =   1560
      TabIndex        =   67
      Top             =   6240
      Width           =   615
   End
   Begin VB.OptionButton optEule_R 
      Caption         =   "N"
      Height          =   195
      Index           =   5
      Left            =   960
      TabIndex        =   66
      Top             =   6240
      Width           =   615
   End
   Begin VB.OptionButton optEule_R 
      Caption         =   "Otro"
      Height          =   195
      Index           =   6
      Left            =   2760
      TabIndex        =   65
      Top             =   6240
      Width           =   615
   End
   Begin VB.OptionButton optEule_R 
      Caption         =   "V"
      Height          =   195
      Index           =   7
      Left            =   2160
      TabIndex        =   64
      Top             =   6240
      Value           =   -1  'True
      Width           =   615
   End
   Begin VB.TextBox txtEule_R 
      Height          =   315
      Index           =   7
      Left            =   960
      MaxLength       =   50
      TabIndex        =   47
      Text            =   "Text1"
      Top             =   4560
      Width           =   4815
   End
   Begin VB.TextBox txtEule_R 
      Height          =   315
      Index           =   5
      Left            =   960
      MaxLength       =   50
      TabIndex        =   46
      Text            =   "Text1"
      Top             =   4080
      Width           =   2175
   End
   Begin VB.TextBox txtEule_R 
      Height          =   315
      Index           =   6
      Left            =   4080
      MaxLength       =   50
      TabIndex        =   45
      Text            =   "Text1"
      Top             =   4080
      Width           =   1695
   End
   Begin VB.TextBox txtEule_R 
      Height          =   315
      Index           =   8
      Left            =   960
      MaxLength       =   50
      TabIndex        =   44
      Text            =   "Text1"
      Top             =   5040
      Width           =   4815
   End
   Begin VB.TextBox txtEule_R 
      Height          =   315
      Index           =   10
      Left            =   3960
      MaxLength       =   100
      TabIndex        =   43
      Text            =   "Text1"
      Top             =   5520
      Width           =   1815
   End
   Begin VB.TextBox txtEule_R 
      Height          =   315
      Index           =   9
      Left            =   960
      MaxLength       =   50
      TabIndex        =   42
      Text            =   "Text1"
      Top             =   5520
      Width           =   1575
   End
   Begin VB.TextBox txtEule_R 
      Height          =   315
      Index           =   12
      Left            =   7920
      MaxLength       =   50
      TabIndex        =   41
      Text            =   "Text1"
      Top             =   4080
      Width           =   2175
   End
   Begin VB.TextBox txtEule_R 
      Height          =   315
      Index           =   13
      Left            =   7920
      MaxLength       =   50
      TabIndex        =   40
      Text            =   "Text1"
      Top             =   4560
      Width           =   4695
   End
   Begin VB.TextBox txtEule_R 
      Height          =   315
      Index           =   14
      Left            =   7920
      MaxLength       =   50
      TabIndex        =   39
      Text            =   "Text1"
      Top             =   5040
      Width           =   4695
   End
   Begin VB.TextBox txtEule_R 
      Height          =   315
      Index           =   16
      Left            =   10920
      MaxLength       =   50
      TabIndex        =   38
      Text            =   "Text1"
      Top             =   5520
      Width           =   1695
   End
   Begin VB.TextBox txtEule_R 
      Height          =   315
      Index           =   15
      Left            =   7920
      MaxLength       =   50
      TabIndex        =   37
      Text            =   "Text1"
      Top             =   5520
      Width           =   1575
   End
   Begin VB.CheckBox chkEuler 
      Caption         =   "chkEuler"
      Height          =   255
      Index           =   0
      Left            =   2760
      TabIndex        =   25
      Top             =   2640
      Width           =   255
   End
   Begin VB.CheckBox chkEuler 
      Caption         =   "chkEuler"
      Height          =   255
      Index           =   1
      Left            =   3840
      TabIndex        =   24
      Top             =   2640
      Width           =   255
   End
   Begin VB.CheckBox chkEuler 
      Caption         =   "chkEuler"
      Height          =   255
      Index           =   2
      Left            =   5160
      TabIndex        =   23
      Top             =   2640
      Width           =   255
   End
   Begin VB.CheckBox chkEuler 
      Caption         =   "chkEuler"
      Height          =   255
      Index           =   3
      Left            =   6240
      TabIndex        =   22
      Top             =   2640
      Width           =   255
   End
   Begin VB.CheckBox chkEuler 
      Caption         =   "chkEuler"
      Height          =   255
      Index           =   4
      Left            =   7800
      TabIndex        =   21
      Top             =   2640
      Width           =   255
   End
   Begin VB.CheckBox chkEuler 
      Caption         =   "chkEuler"
      Height          =   255
      Index           =   5
      Left            =   2760
      TabIndex        =   20
      Top             =   3000
      Width           =   255
   End
   Begin VB.CheckBox chkEuler 
      Caption         =   "chkEuler"
      Height          =   255
      Index           =   6
      Left            =   3840
      TabIndex        =   19
      Top             =   3000
      Width           =   255
   End
   Begin VB.CheckBox chkEuler 
      Caption         =   "chkEuler"
      Height          =   255
      Index           =   7
      Left            =   5160
      TabIndex        =   18
      Top             =   3000
      Width           =   255
   End
   Begin VB.CheckBox chkEuler 
      Caption         =   "chkEuler"
      Height          =   255
      Index           =   8
      Left            =   6240
      TabIndex        =   17
      Top             =   3000
      Width           =   255
   End
   Begin VB.CheckBox chkEuler 
      Caption         =   "chkEuler"
      Height          =   255
      Index           =   9
      Left            =   7800
      TabIndex        =   16
      Top             =   3000
      Width           =   255
   End
   Begin VB.TextBox txtEule_R 
      Height          =   315
      Index           =   3
      Left            =   8640
      MaxLength       =   16
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   2640
      Width           =   3975
   End
   Begin VB.TextBox txtEule_R 
      Height          =   315
      Index           =   4
      Left            =   8640
      MaxLength       =   16
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   3000
      Width           =   3975
   End
   Begin VB.Frame Frame4R 
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   1320
      Width           =   3015
      Begin VB.OptionButton optEule_R 
         Caption         =   "Debidos"
         Height          =   195
         Index           =   2
         Left            =   600
         TabIndex        =   12
         Top             =   0
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton optEule_R 
         Caption         =   "Pagados"
         Height          =   195
         Index           =   3
         Left            =   1680
         TabIndex        =   11
         Top             =   0
         Width           =   975
      End
      Begin VB.Label Label3E 
         Caption         =   "Portes"
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
         Index           =   21
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   1935
      End
   End
   Begin VB.TextBox txtEule_R 
      Height          =   315
      Index           =   0
      Left            =   4800
      MaxLength       =   100
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   1200
      Width           =   4215
   End
   Begin VB.TextBox txtEule_R 
      Height          =   315
      Index           =   1
      Left            =   9240
      MaxLength       =   50
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1200
      Width           =   2055
   End
   Begin VB.TextBox txtEule_R 
      Height          =   315
      Index           =   2
      Left            =   11640
      MaxLength       =   16
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox txtEule_R 
      Height          =   315
      Index           =   20
      Left            =   240
      MaxLength       =   16
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox txtEule_R 
      Height          =   315
      Index           =   21
      Left            =   2160
      MaxLength       =   16
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label3E 
      Caption         =   "Marca"
      Height          =   195
      Index           =   47
      Left            =   9720
      TabIndex        =   94
      Top             =   7080
      Width           =   705
   End
   Begin VB.Label Label3E 
      Caption         =   "Modelo"
      Height          =   195
      Index           =   46
      Left            =   9720
      TabIndex        =   93
      Top             =   7440
      Width           =   705
   End
   Begin VB.Label Label3E 
      Caption         =   "Marca"
      Height          =   195
      Index           =   45
      Left            =   5280
      TabIndex        =   90
      Top             =   7080
      Width           =   705
   End
   Begin VB.Label Label3E 
      AutoSize        =   -1  'True
      Caption         =   "Fecha"
      Height          =   195
      Index           =   44
      Left            =   1200
      TabIndex        =   89
      Top             =   7440
      Width           =   450
   End
   Begin VB.Label Label3E 
      Caption         =   "Motor"
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
      Index           =   43
      Left            =   8640
      TabIndex        =   87
      Top             =   7080
      Width           =   975
   End
   Begin VB.Label Label3E 
      Caption         =   "Bomba"
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
      Index           =   42
      Left            =   4320
      TabIndex        =   86
      Top             =   7080
      Width           =   975
   End
   Begin VB.Label Label3E 
      Caption         =   "Pedido"
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
      Index           =   41
      Left            =   120
      TabIndex        =   85
      Top             =   7080
      Width           =   975
   End
   Begin VB.Label Label3E 
      AutoSize        =   -1  'True
      Caption         =   "Referencia"
      Height          =   195
      Index           =   40
      Left            =   1200
      TabIndex        =   84
      Top             =   7080
      Width           =   780
   End
   Begin VB.Label Label3E 
      Caption         =   "Nº Expedicion"
      Height          =   195
      Index           =   39
      Left            =   4800
      TabIndex        =   82
      Top             =   960
      Width           =   2865
   End
   Begin VB.Label Label3E 
      Caption         =   "Modelo"
      Height          =   195
      Index           =   38
      Left            =   5280
      TabIndex        =   81
      Top             =   7440
      Width           =   705
   End
   Begin VB.Label Label3E 
      Caption         =   "Marca"
      Height          =   195
      Index           =   22
      Left            =   4680
      TabIndex        =   80
      Top             =   7080
      Width           =   705
   End
   Begin VB.Label Label3E 
      Caption         =   "ORDEN DE TRABAJO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   77
      Top             =   6600
      Width           =   3015
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00008000&
      BorderWidth     =   3
      X1              =   240
      X2              =   12720
      Y1              =   6720
      Y2              =   6720
   End
   Begin VB.Label Label3E 
      Caption         =   "Pot(CV)"
      Height          =   195
      Index           =   33
      Left            =   7080
      TabIndex        =   75
      Top             =   6000
      Width           =   705
   End
   Begin VB.Label Label3E 
      Caption         =   "Pot (Kw)"
      Height          =   195
      Index           =   34
      Left            =   9120
      TabIndex        =   74
      Top             =   6000
      Width           =   705
   End
   Begin VB.Label Label3E 
      Caption         =   "RPM"
      Height          =   195
      Index           =   35
      Left            =   10920
      TabIndex        =   73
      Top             =   6000
      Width           =   705
   End
   Begin VB.Label Label3E 
      Caption         =   "Tipo de rodete"
      Height          =   195
      Index           =   31
      Left            =   120
      TabIndex        =   63
      Top             =   6000
      Width           =   1035
   End
   Begin VB.Label Label3E 
      Caption         =   "Caudal"
      Height          =   195
      Index           =   32
      Left            =   3480
      TabIndex        =   62
      Top             =   6000
      Width           =   705
   End
   Begin VB.Label Label3E 
      Caption         =   "Marca"
      Height          =   195
      Index           =   12
      Left            =   120
      TabIndex        =   61
      Top             =   4080
      Width           =   705
   End
   Begin VB.Label Label3E 
      Caption         =   "NºCurva"
      Height          =   195
      Index           =   13
      Left            =   3240
      TabIndex        =   60
      Top             =   4080
      Width           =   705
   End
   Begin VB.Label Label3E 
      Caption         =   "Modelo"
      Height          =   195
      Index           =   14
      Left            =   120
      TabIndex        =   59
      Top             =   4560
      Width           =   705
   End
   Begin VB.Label Label3E 
      Caption         =   "Bombas(Parte hidraulica)"
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
      Index           =   16
      Left            =   1920
      TabIndex        =   58
      Top             =   3720
      Width           =   2655
   End
   Begin VB.Label Label3E 
      Caption         =   "Nº Serie"
      Height          =   195
      Index           =   17
      Left            =   120
      TabIndex        =   57
      Top             =   5040
      Width           =   705
   End
   Begin VB.Label Label3E 
      Caption         =   "H (m.c.a)"
      Height          =   195
      Index           =   18
      Left            =   3000
      TabIndex        =   56
      Top             =   5520
      Width           =   705
   End
   Begin VB.Label Label3E 
      Caption         =   "Marca"
      Height          =   195
      Index           =   19
      Left            =   120
      TabIndex        =   55
      Top             =   5520
      Width           =   705
   End
   Begin VB.Label Label3E 
      Caption         =   "Datos equipo / bomba recepcionado"
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
      Height          =   255
      Index           =   11
      Left            =   120
      TabIndex        =   54
      Top             =   3360
      Width           =   4095
   End
   Begin VB.Label Label3E 
      Caption         =   "Motor"
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
      Index           =   25
      Left            =   8880
      TabIndex        =   53
      Top             =   3720
      Width           =   2655
   End
   Begin VB.Label Label3E 
      Caption         =   "Modelo"
      Height          =   195
      Index           =   26
      Left            =   7080
      TabIndex        =   52
      Top             =   4560
      Width           =   705
   End
   Begin VB.Label Label3E 
      Caption         =   "Marca"
      Height          =   195
      Index           =   27
      Left            =   7080
      TabIndex        =   51
      Top             =   4080
      Width           =   705
   End
   Begin VB.Label Label3E 
      Caption         =   "Nº Serie"
      Height          =   195
      Index           =   28
      Left            =   7080
      TabIndex        =   50
      Top             =   5040
      Width           =   705
   End
   Begin VB.Label Label3E 
      Caption         =   "I (A)"
      Height          =   195
      Index           =   29
      Left            =   10320
      TabIndex        =   49
      Top             =   5520
      Width           =   705
   End
   Begin VB.Label Label3E 
      Caption         =   "V"
      Height          =   195
      Index           =   30
      Left            =   7080
      TabIndex        =   48
      Top             =   5520
      Width           =   705
   End
   Begin VB.Label Label3E 
      Caption         =   "Tipo de bombas recepcionadas"
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
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   36
      Top             =   1920
      Width           =   3615
   End
   Begin VB.Label Label3E 
      Caption         =   "Aguas residuales"
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
      Index           =   2
      Left            =   120
      TabIndex        =   35
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label Label3E 
      Caption         =   "Aguas limpias"
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
      Index           =   3
      Left            =   120
      TabIndex        =   34
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label Label3E 
      Caption         =   "Bombas superficie"
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
      Index           =   4
      Left            =   2640
      TabIndex        =   33
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label Label3E 
      Caption         =   "Bombas sumegibles"
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
      Index           =   5
      Left            =   5040
      TabIndex        =   32
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Label Label3E 
      Caption         =   "Agitador"
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
      Index           =   6
      Left            =   7560
      TabIndex        =   31
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label3E 
      Caption         =   "Horizontal"
      Height          =   195
      Index           =   7
      Left            =   2520
      TabIndex        =   30
      Top             =   2400
      Width           =   705
   End
   Begin VB.Label Label3E 
      Caption         =   "Vertical"
      Height          =   195
      Index           =   8
      Left            =   3720
      TabIndex        =   29
      Top             =   2400
      Width           =   525
   End
   Begin VB.Label Label3E 
      AutoSize        =   -1  'True
      Caption         =   "Pozo"
      Height          =   195
      Index           =   9
      Left            =   5040
      TabIndex        =   28
      Top             =   2400
      Width           =   360
   End
   Begin VB.Label Label3E 
      AutoSize        =   -1  'True
      Caption         =   "Vertical"
      Height          =   195
      Index           =   10
      Left            =   6120
      TabIndex        =   27
      Top             =   2400
      Width           =   525
   End
   Begin VB.Label Label3E 
      Caption         =   "Otros equipos / tipos"
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
      Index           =   15
      Left            =   9240
      TabIndex        =   26
      Top             =   2280
      Width           =   2655
   End
   Begin VB.Label Label3E 
      Caption         =   "Recepcion del equipo"
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
      Height          =   255
      Index           =   20
      Left            =   240
      TabIndex        =   9
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label Label3E 
      Caption         =   "Nº Expedicion"
      Height          =   195
      Index           =   23
      Left            =   9240
      TabIndex        =   8
      Top             =   960
      Width           =   2865
   End
   Begin VB.Label Label3E 
      Caption         =   "F. Alb"
      Height          =   195
      Index           =   24
      Left            =   11640
      TabIndex        =   7
      Top             =   960
      Width           =   405
   End
   Begin VB.Label Label3E 
      AutoSize        =   -1  'True
      Caption         =   "Orden de trabajo"
      Height          =   195
      Index           =   36
      Left            =   240
      TabIndex        =   6
      Top             =   600
      Width           =   1185
   End
   Begin VB.Label Label3E 
      Caption         =   "T. Externo"
      Height          =   195
      Index           =   37
      Left            =   2160
      TabIndex        =   5
      Top             =   600
      Width           =   1005
   End
   Begin VB.Line Line4 
      BorderWidth     =   3
      X1              =   240
      X2              =   12720
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line3 
      BorderWidth     =   3
      X1              =   120
      X2              =   12600
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line5 
      BorderWidth     =   3
      X1              =   240
      X2              =   12720
      Y1              =   3480
      Y2              =   3480
   End
End
Attribute VB_Name = "frmFacEulerDatosRep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Buscar As Boolean

Private Sub cmdBuscar_Click()
    RealizarBuscar
End Sub

Private Sub Command1_Click()
    CadenaDesdeOtroForm = ""
    Unload Me
End Sub

Private Sub Form_Load()
Dim Bloquea As Boolean
Dim N As Byte
        Me.Icon = frmPpal.Icon
    Me.Caption = "Datos albaran Euler"
    
        
    'If Buscar Then
    'Bloquea = True
    Bloquea = Not Buscar
            
    cboEulerUdR.Enabled = Not Bloquea
     
    For N = 0 To Me.txtEule_R.Count - 1
        BloquearTxt txtEule_R(N), Bloquea
    Next
    
    For N = 0 To Me.optEule_R.Count - 1
        Me.optEule_R(N).Enabled = Not Bloquea
    Next N
    
    For N = 0 To chkEuler.Count - 1
        chkEuler(N).Enabled = Not Bloquea
    Next

    If Buscar Then
        Me.Height = 9015
        limpiar Me
        For N = 0 To Me.optEule_R.Count - 1
           ' Me.optEule_R(N).Enabled = Not Bloquea
            Me.optEule_R(N).Value = False
        Next N
        For N = 0 To chkEuler.Count - 1
            chkEuler(N).Value = 0
        Next
        cboEulerUdR.ListIndex = -1
    Else
        PonerCamposFichaReparacion
        Me.Height = 6990
    End If
End Sub



'*************************************************************************************
' Ficha reparacion

Private Function CamposSQlFichaReparacion() As String
    'Primero iran todos los txts juntos y por orden de index
    CamposSQlFichaReparacion = "RecepAgenCliMat,RecpNumExp,FechaAlb,TipoBomResOtrosEqu,TipoBomLimOtrosEqu,DatosBommarca"
    CamposSQlFichaReparacion = CamposSQlFichaReparacion & ",DatosBomNumCurva,DatosBomModelo,DatosBomNumSerie,DatosBomAno,DatosBomH,DatosBomCaudal"
    CamposSQlFichaReparacion = CamposSQlFichaReparacion & ",DatosMotorMarca , DatosMotorModelo, DatosMotorNumSerie, DatosMotorV, DatosMotorI"
    CamposSQlFichaReparacion = CamposSQlFichaReparacion & ",DatosMotorCV, DatosMotorKw, DatosMotorrpm,NumTrabajExterno,NumParteTrabajo"

    'Tipo bomba recepcionada
    'Son los check. Tambien vmos con el ordern
    CamposSQlFichaReparacion = CamposSQlFichaReparacion & ", TipoBombResSuperHor,TipoBombResSuperVer,TipoBombResSumPoz, TipoBombResSumVer, TipoBomAgitadorRes"
    CamposSQlFichaReparacion = CamposSQlFichaReparacion & ", TipoBombLimSuperHor,TipoBombLimSuperVer,TipoBombLimSumPoz, TipoBombLimSumVer, TipoBomAgitadorLim "
    

    'Luego resto campos
    CamposSQlFichaReparacion = CamposSQlFichaReparacion & ", RecepAgenClien,RecepPortes, DatosBomUdCaudal,DatosBomTipoRodete"
    CamposSQlFichaReparacion = CamposSQlFichaReparacion & ",codtipom,numalbar"
    
    
    If Buscar Then CamposSQlFichaReparacion = CamposSQlFichaReparacion & ", ReferPedido , FechaPed, bombamarca, bombaModelo, motormarca, motorModelo,"

    
End Function

Private Sub PonerCamposFichaReparacion()
Dim N As Byte
Dim SQL As String
    
    SQL = CamposSQlFichaReparacion()
    SQL = "Select " & SQL & " FROM scafac_eu   " & CadenaDesdeOtroForm

        
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If miRsAux.EOF Then
        limpiar Me
        
    Else
        
        
        
        'EL SQL estara montaddo para que coincida el orden del columna con el index
        For N = 0 To txtEule_R.Count - 1
            txtEule_R(N).Text = DBLet(miRsAux.Fields(CInt(N)), "T")
            If N = 20 Or N = 21 Then
                'NUmerico
                If txtEule_R(N).Text <> "" Then txtEule_R(N).Text = Format(txtEule_R(N).Text, "000000")
            End If
        Next
    
        'Agencia cliente
        'N = 1
        'If DBLet(miRsAux!RecepAgenClien, "N") = 0 Then N = 0
        'optEule_R(N).Value = True
        
        cboRecepAgenClien.ListIndex = -1
        If Not IsNull(miRsAux!RecepAgenClien) Then cboRecepAgenClien.ListIndex = miRsAux!RecepAgenClien
        
        
        N = 3
        If DBLet(miRsAux!RecepPortes, "N") = 1 Then N = 2
        optEule_R(N).Value = True
        
        'Empieza en la 20
        For N = 1 To Me.chkEuler.Count
            chkEuler(N - 1).Value = DBLet(miRsAux.Fields(CInt(N) + 21), "N")
        Next
        
        ' DatosBomUdCaudal,DatosBomTipoRodete"
        SQL = DBLet(miRsAux!DatosBomTipoRodete, "N")
        If SQL = 0 Then SQL = 6  'OTROS
        For N = 4 To 7
            If N = Val(SQL) Then Me.optEule_R(N).Value = True
        Next
        

        cboEulerUdR.ListIndex = -1
        If Not IsNull(miRsAux!DatosBomUdCaudal) Then cboEulerUdR.ListIndex = miRsAux!DatosBomUdCaudal
        
       
        'Combo1.ListIndex = kCampo
        
        
    End If
    miRsAux.Close
    Set miRsAux = Nothing
End Sub




Private Sub RealizarBuscar()
Dim N As Byte
Dim SQL As String
Dim Campos As String
Dim Aux As String
Dim K As Integer


    If Not Buscar Then
        Unload Me
        Exit Sub
    End If


    Campos = CamposSQlFichaReparacion()
    Campos = Replace(Campos, ",", "|")
    SQL = ""
    'EL SQL estara montaddo para que coincida el orden del columna con el index
    For N = 0 To 27
        If txtEule_R(N).Text <> "" Then
            K = N + 1
            If N > 21 Then K = N + 17
            
            If N = 20 Or N = 21 Then
                If InStr(1, txtEule_R(N).Text, ">") > 0 Or InStr(1, txtEule_R(N).Text, "<") Then
                    Aux = InStr(1, txtEule_R(N).Text, ">")
                Else
                    Aux = " = " & txtEule_R(N).Text
                End If
                Aux = RecuperaValor(Campos, K) & Aux
            Else
                If InStr(1, txtEule_R(N).Text, "*") > 0 Then
                    Aux = Replace(txtEule_R(N).Text, "*", "%")
                Else
                    Aux = "%" & txtEule_R(N).Text & "%"
                End If
                Aux = RecuperaValor(Campos, K) & " like '" & Aux & "'"
            End If
            If SQL <> "" Then SQL = SQL & " AND "
            SQL = SQL & Aux
            
        End If
    Next
    
    'Agencia cliente
   ' N = 1
   ' If DBLet(miRsAux!RecepAgenClien, "N") = 0 Then N = 0
   ' optEule_R(N).Value = True
        
   '     N = 3
   '     If DBLet(miRsAux!RecepPortes, "N") = 1 Then N = 2
   '     optEule_R(N).Value = True
        
        'Empieza en la 20
  '      For N = 1 To Me.chkEuler.Count
  '          chkEuler(N - 1).Value = DBLet(miRsAux.Fields(CInt(N) + 21), "N")
  '      Next
        
        ' DatosBomUdCaudal,DatosBomTipoRodete"
  '      SQL = DBLet(miRsAux!DatosBomTipoRodete, "N")
  '      If SQL = 0 Then SQL = 6  'OTROS
  '      For N = 4 To 7
  '          If N = Val(SQL) Then Me.optEule_R(N).Value = True
  '      Next
        

  '      cboEulerUdR.ListIndex = -1
  '      If Not IsNull(miRsAux!DatosBomUdCaudal) Then cboEulerUdR.ListIndex = miRsAux!DatosBomUdCaudal
        
       
        'Combo1.ListIndex = kCampo
        
        
    If SQL <> "" Then
        Aux = "Select codtipom,numfactu,fecfactu FROM scafac_eu where " & SQL
        Set miRsAux = New ADODB.Recordset
        miRsAux.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        SQL = ""
        While Not miRsAux.EOF
            SQL = SQL & ", ('" & miRsAux!codtipom & "'," & miRsAux!Numfactu & "," & DBSet(miRsAux!FecFactu, "F") & ")"
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        Set miRsAux = Nothing
        If SQL = "" Then
            MsgBox "Ningun dato con estos valores", vbExclamation
        Else
            CadenaDesdeOtroForm = Mid(SQL, 2)
            Unload Me
        End If
    End If
End Sub



